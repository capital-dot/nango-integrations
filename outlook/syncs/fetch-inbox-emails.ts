import { createSync } from 'nango';
import * as z from 'zod';

// ─────────────────────────────────────────────────────────────────────────
// MODEL: one record per email. This is what gets saved to Nango's record
// store and what triggers the "Syncs: new records" webhook to your worker.
// ─────────────────────────────────────────────────────────────────────────
const EmailAddress = z.object({
    name: z.string().nullable(),
    address: z.string().nullable()
});

const Email = z.object({
    id: z.string(),
    subject: z.string(),
    from: EmailAddress,
    to: z.array(EmailAddress),
    cc: z.array(EmailAddress),
    bodyContentType: z.string(), // 'text' or 'html'
    body: z.string(),            // full body content
    bodyPreview: z.string(),     // short plaintext preview, useful for lists/notifications
    receivedAt: z.string(),      // ISO datetime
    isRead: z.boolean(),
    hasAttachments: z.boolean()
});

// Fields requested from Graph — trimmed to exactly what we map below, to
// keep each page's payload small. Add fields here AND to the mapper below
// if you need more later (e.g. attachments list, importance, categories).
const SELECT_FIELDS = [
    'id',
    'subject',
    'from',
    'toRecipients',
    'ccRecipients',
    'body',
    'bodyPreview',
    'receivedDateTime',
    'isRead',
    'hasAttachments'
].join(',');

function mapAddress(a: any): { name: string | null; address: string | null } {
    return {
        name: a?.emailAddress?.name ?? null,
        address: a?.emailAddress?.address ?? null
    };
}

function mapMessage(m: any) {
    return {
        id: m.id,
        subject: m.subject || '(no subject)',
        from: mapAddress(m.from),
        to: (m.toRecipients || []).map(mapAddress),
        cc: (m.ccRecipients || []).map(mapAddress),
        bodyContentType: m.body?.contentType || 'text',
        body: m.body?.content || '',
        bodyPreview: m.bodyPreview || '',
        receivedAt: m.receivedDateTime,
        isRead: Boolean(m.isRead),
        hasAttachments: Boolean(m.hasAttachments)
    };
}

export default createSync({
    description: 'Pulls inbox emails (full body, from/to/cc, subject) from Outlook via Microsoft Graph.',
    version: '1.0.0',
    frequency: 'every 1 minute',
    autoStart: true,
    syncType: 'incremental',
    trackDeletes: false, // we only care about new mail arriving, not deletions
    scopes: ['Mail.Read'],
    models: { Email },

    // Checkpoint = last receivedDateTime we've already saved, so every run
    // only asks Graph for mail newer than that instead of re-fetching the
    // whole inbox each time.
    checkpoint: z.object({
        lastReceivedISO: z.string().nullable()
    }),

    exec: async (nango) => {
        const checkpoint = await nango.getCheckpoint();
        const since = checkpoint?.lastReceivedISO;

        let endpoint = '/v1.0/me/mailFolders/inbox/messages';
        let params: Record<string, string> | undefined = {
            $select: SELECT_FIELDS,
            $orderby: 'receivedDateTime desc',
            $top: '25',
            ...(since ? { $filter: `receivedDateTime gt ${since}` } : {})
        };

        let newestSeen: string | null = since ?? null;
        let hasMore = true;

        while (hasMore) {
            const response = await nango.get({
                baseUrlOverride: 'https://graph.microsoft.com',
                endpoint,
                params
            });

            const messages: any[] = response.data.value || [];

            if (messages.length > 0) {
                const mapped = messages.map(mapMessage);
                await nango.batchSave(mapped, 'Email');

                const pageNewest = mapped[0].receivedAt;
                if (!newestSeen || pageNewest > newestSeen) {
                    newestSeen = pageNewest;
                }
            }

            const nextLink: string | undefined = response.data['@odata.nextLink'];
            if (nextLink) {
                endpoint = nextLink.replace('https://graph.microsoft.com', '');
                params = undefined;
                hasMore = true;
            } else {
                hasMore = false;
            }
        }

        if (newestSeen) {
            await nango.saveCheckpoint({ lastReceivedISO: newestSeen });
        }

        await nango.log(`Inbox sync completed. Newest receivedAt: ${newestSeen ?? '(none)'}`);
    }
});
