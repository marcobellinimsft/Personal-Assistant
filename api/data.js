const { verifyToken } = require('./_auth');
const { kv } = require('@vercel/kv');

const DATA_KEY = 'pa_data';
const BACKUP_KEY = 'pa_data_backup';

module.exports = async function handler(req, res) {
    // Auth check
    const payload = await verifyToken(req);
    if (!payload) {
        return res.status(401).json({ error: 'Not authenticated' });
    }

    if (req.method === 'GET') {
        const data = await kv.get(DATA_KEY);
        return res.status(200).json(data || {});
    }

    if (req.method === 'PUT') {
        const data = req.body;
        if (!data || typeof data !== 'object') {
            return res.status(400).json({ error: 'Invalid data' });
        }

        // Save current data as backup before overwriting
        const existing = await kv.get(DATA_KEY);
        if (existing) {
            await kv.set(BACKUP_KEY, existing);
        }

        data._savedAt = new Date().toISOString();
        await kv.set(DATA_KEY, data);

        return res.status(200).json({ ok: true, savedAt: data._savedAt });
    }

    return res.status(405).json({ error: 'Method not allowed' });
};
