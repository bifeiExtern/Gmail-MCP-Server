import { gmail_v1 } from 'googleapis';

/**
 * Gmail-style quoting for drafts and sends that continue an existing thread.
 *
 * Gmail's web UI appends the previous message ("On ... wrote:" for replies,
 * "---------- Forwarded message ---------" for forwards) when you click
 * Reply/Forward. Messages created through the API get none of that, so a
 * reply sent from an API draft arrives with no history below it. This module
 * rebuilds that block from the source message and attaches it to the outgoing
 * text and HTML bodies.
 */

interface MessagePart {
    mimeType?: string | null;
    filename?: string | null;
    headers?: Array<{ name?: string | null; value?: string | null }> | null;
    body?: { attachmentId?: string | null; size?: number | null; data?: string | null } | null;
    parts?: MessagePart[] | null;
}

export interface QuoteAttachment {
    filename: string;
    content: Buffer;
    contentType?: string;
    cid?: string;
}

export interface QuoteResult {
    applied: boolean;
    note: string;
}

// Skip original attachments beyond this total, Gmail's own send limit is 25 MB
const MAX_ATTACHMENT_BYTES = 20 * 1024 * 1024;

const FORWARD_SUBJECT = /^\s*(fwd?|fw)\s*:/i;

function header(part: MessagePart | undefined | null, name: string): string {
    const h = part?.headers?.find(x => x.name?.toLowerCase() === name.toLowerCase());
    return h?.value || '';
}

function decode(data: string): string {
    return Buffer.from(data, 'base64').toString('utf8');
}

function escapeHtml(s: string): string {
    return s.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');
}

function linkify(escaped: string): string {
    return escaped.replace(/(https?:\/\/[^\s<>"]+[^\s<>".,;:!?)\]])/g, '<a href="$1">$1</a>');
}

/** Plain text to Gmail-like HTML: one div per line, blank lines as <div><br></div> */
export function plainToHtml(text: string): string {
    const lines = text.replace(/\r\n/g, '\n').split('\n');
    return lines.map(l => l.trim() === '' ? '<div><br></div>' : `<div>${linkify(escapeHtml(l))}</div>`).join('');
}

function htmlToPlain(html: string): string {
    return html
        .replace(/<style[\s\S]*?<\/style>/gi, '')
        .replace(/<br\s*\/?>/gi, '\n')
        .replace(/<\/(p|div|tr|li|h\d)>/gi, '\n')
        .replace(/<[^>]+>/g, '')
        .replace(/&nbsp;/g, ' ').replace(/&lt;/g, '<').replace(/&gt;/g, '>').replace(/&quot;/g, '"').replace(/&#39;/g, "'").replace(/&amp;/g, '&')
        .replace(/\n{3,}/g, '\n\n')
        .trim();
}

/** Keep only what sits inside <body>, so the quote can be nested in our own HTML */
function bodyInner(html: string): string {
    const m = html.match(/<body[^>]*>([\s\S]*)<\/body>/i);
    let inner = m ? m[1] : html;
    inner = inner.replace(/<\/?(html|head)[^>]*>/gi, '').replace(/<meta[^>]*>/gi, '').replace(/<title>[\s\S]*?<\/title>/gi, '');
    return inner.trim();
}

function collect(part: MessagePart, out: { text: string; html: string; files: MessagePart[] }) {
    const type = (part.mimeType || '').toLowerCase();
    const isAttachment = !!part.filename || !!header(part, 'Content-ID') && !!part.body?.attachmentId;
    if (isAttachment && part.body?.attachmentId) {
        out.files.push(part);
    } else if (part.body?.data) {
        if (type === 'text/plain' && !out.text) out.text = decode(part.body.data);
        else if (type === 'text/html' && !out.html) out.html = decode(part.body.data);
    }
    for (const p of part.parts || []) collect(p, out);
}

/** Gmail's attribution date, e.g. "Mon, Oct 26, 2026 at 2:29 PM" */
function formatDate(dateHeader: string, internalDate?: string | null): string {
    const d = internalDate ? new Date(Number(internalDate)) : new Date(dateHeader);
    if (isNaN(d.getTime())) return dateHeader;
    const timeZone = process.env.GMAIL_QUOTE_TZ || undefined;
    const day = new Intl.DateTimeFormat('en-US', { weekday: 'short', month: 'short', day: 'numeric', year: 'numeric', timeZone }).format(d);
    const time = new Intl.DateTimeFormat('en-US', { hour: 'numeric', minute: '2-digit', timeZone }).format(d);
    return `${day} at ${time}`;
}

function addressHtml(list: string): string {
    return escapeHtml(list).replace(/([\w.+-]+@[\w-]+(\.[\w-]+)+)/g, '<a href="mailto:$1">$1</a>');
}

/** "Name <a@b>" → { name, email } */
function splitAddress(from: string): { name: string; email: string } {
    const m = from.match(/^\s*"?([^"<]*?)"?\s*<([^>]+)>\s*$/);
    return m ? { name: m[1].trim(), email: m[2].trim() } : { name: '', email: from.trim() };
}

async function findSource(gmail: gmail_v1.Gmail, args: any): Promise<string | null> {
    if (args.quoteMessageId) return args.quoteMessageId;
    if (args.inReplyTo && !args.inReplyTo.startsWith('<')) return args.inReplyTo;
    if (args.inReplyTo) {
        const res = await gmail.users.messages.list({ userId: 'me', q: `rfc822msgid:${args.inReplyTo.replace(/[<>]/g, '')}`, maxResults: 1 });
        if (res.data.messages?.[0]?.id) return res.data.messages[0].id;
    }
    if (args.threadId) {
        const t = await gmail.users.threads.get({ userId: 'me', id: args.threadId, format: 'minimal' });
        const msgs = (t.data.messages || []).filter(m => !(m.labelIds || []).includes('DRAFT') && !(m.labelIds || []).includes('TRASH'));
        msgs.sort((a, b) => Number(a.internalDate || 0) - Number(b.internalDate || 0));
        return msgs.length ? msgs[msgs.length - 1].id || null : null;
    }
    return null;
}

/**
 * Appends the quoted source message to args.body / args.htmlBody and fills in
 * threading headers. Mutates args; returns what it did so the caller can report it.
 */
export async function applyQuote(gmail: gmail_v1.Gmail, args: any): Promise<QuoteResult> {
    if (args.quoteOriginal === false) return { applied: false, note: '' };
    if (!args.threadId && !args.inReplyTo && !args.quoteMessageId) return { applied: false, note: '' };

    const forward = FORWARD_SUBJECT.test(args.subject || '');
    const existing = `${args.body || ''}\n${args.htmlBody || ''}`;
    if (/-{5,}\s*Forwarded message\s*-{5,}/i.test(existing) || /^On .+wrote:\s*$/m.test(existing) || /gmail_quote/.test(existing)) {
        return { applied: false, note: 'Quote skipped: the body already contains quoted text.' };
    }

    const sourceId = await findSource(gmail, args);
    if (!sourceId) return { applied: false, note: 'Quote skipped: no earlier message found in the thread.' };

    const src = await gmail.users.messages.get({ userId: 'me', id: sourceId, format: 'full' });
    const payload = src.data.payload as MessagePart;
    const parts = { text: '', html: '', files: [] as MessagePart[] };
    collect(payload, parts);

    const from = header(payload, 'From');
    const to = header(payload, 'To');
    const cc = header(payload, 'Cc');
    const subject = header(payload, 'Subject');
    const date = formatDate(header(payload, 'Date'), src.data.internalDate);
    const messageId = header(payload, 'Message-ID');
    const references = header(payload, 'References');

    const srcText = parts.text || (parts.html ? htmlToPlain(parts.html) : '');
    const srcHtml = parts.html ? bodyInner(parts.html) : plainToHtml(srcText);

    // Attachments: inline images always (so the quoted HTML renders), regular files only on forwards
    const attachments: QuoteAttachment[] = [];
    let total = 0;
    let skipped = 0;
    for (const f of parts.files) {
        const cid = header(f, 'Content-ID').replace(/[<>]/g, '');
        const inline = !!cid && srcHtml.includes(`cid:${cid}`);
        if (!inline && !forward) continue;
        const size = f.body?.size || 0;
        if (total + size > MAX_ATTACHMENT_BYTES) { skipped++; continue; }
        const a = await gmail.users.messages.attachments.get({ userId: 'me', messageId: sourceId, id: f.body!.attachmentId! });
        total += size;
        attachments.push({
            filename: f.filename || cid || 'attachment',
            content: Buffer.from(a.data.data || '', 'base64'),
            contentType: f.mimeType || undefined,
            ...(inline ? { cid } : {}),
        });
    }

    const body = (args.body || '').replace(/\s+$/, '');
    const ownHtml = args.htmlBody ? bodyInner(args.htmlBody) : plainToHtml(body);
    const { name, email } = splitAddress(from);
    const fromHtml = name ? `${escapeHtml(name)} &lt;<a href="mailto:${escapeHtml(email)}">${escapeHtml(email)}</a>&gt;` : addressHtml(email);

    if (forward) {
        const headerLines = [`From: ${from}`, `Date: ${date}`, `Subject: ${subject}`, `To: ${to}`, ...(cc ? [`Cc: ${cc}`] : [])];
        args.body = `${body}\n\n---------- Forwarded message ---------\n${headerLines.join('\n')}\n\n\n${srcText}\n`;
        const headerHtml = [
            '---------- Forwarded message ---------',
            `From: <strong class="gmail_sendername" dir="auto">${escapeHtml(name || email)}</strong> <span dir="auto">&lt;<a href="mailto:${escapeHtml(email)}">${escapeHtml(email)}</a>&gt;</span>`,
            `Date: ${escapeHtml(date)}`,
            `Subject: ${escapeHtml(subject)}`,
            `To: ${addressHtml(to)}`,
            ...(cc ? [`Cc: ${addressHtml(cc)}`] : []),
        ].join('<br>');
        args.htmlBody = `<div dir="ltr">${ownHtml}</div><br><div class="gmail_quote gmail_quote_container"><div dir="ltr" class="gmail_attr">${headerHtml}<br></div><br><br>${srcHtml}</div>`;
    } else {
        const quoted = srcText.replace(/\r\n/g, '\n').split('\n').map(l => (l.startsWith('>') ? `>${l}` : l ? `> ${l}` : '>')).join('\n');
        args.body = `${body}\n\nOn ${date} ${from} wrote:\n\n${quoted}\n`;
        args.htmlBody = `<div dir="ltr">${ownHtml}</div><br><div class="gmail_quote gmail_quote_container"><div dir="ltr" class="gmail_attr">On ${escapeHtml(date)} ${fromHtml} wrote:<br></div><blockquote class="gmail_quote" style="margin:0px 0px 0px 0.8ex;border-left:1px solid rgb(204,204,204);padding-left:1ex">${srcHtml}</blockquote></div>`;
    }
    args.mimeType = 'multipart/alternative';

    // Threading: reply to the quoted message and carry its reference chain
    if (messageId) {
        args.inReplyTo = messageId;
        args.references = [...references.split(/\s+/).filter(Boolean), messageId];
    }
    if (!args.threadId && src.data.threadId) args.threadId = src.data.threadId;
    args.quoteAttachments = [...(args.quoteAttachments || []), ...attachments];

    const fileCount = attachments.filter(a => !a.cid).length;
    const note = `${forward ? 'Forwarded' : 'Quoted'} message ${sourceId} from ${from} (${date})` +
        (fileCount ? `, with ${fileCount} original attachment(s)` : '') +
        (skipped ? `, ${skipped} attachment(s) skipped for size` : '') + '.';
    return { applied: true, note };
}
