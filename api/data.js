const { verifyToken } = require('./_auth');
const { Redis } = require('@upstash/redis');

const redis = new Redis({
    url: process.env.KV_REST_API_URL,
    token: process.env.KV_REST_API_TOKEN,
});

const DATA_KEY = 'pa_data';
const BACKUP_KEY = 'pa_data_backup';

module.exports = async function handler(req, res) {
    // Auth check
    const payload = await verifyToken(req);
    if (!payload) {
        return res.status(401).json({ error: 'Not authenticated' });
    }

    if (req.method === 'GET') {
        const data = await redis.get(DATA_KEY);
        return res.status(200).json(data || {});
    }

    if (req.method === 'PUT' || req.method === 'POST') {
        // Support both PUT (fetch) and POST (sendBeacon)
        let data = req.body;
        if (typeof data === 'string') {
            try { data = JSON.parse(data); } catch { return res.status(400).json({ error: 'Invalid JSON' }); }
        }
        if (!data || typeof data !== 'object') {
            return res.status(400).json({ error: 'Invalid data' });
        }

        // Save current data as backup before overwriting
        const existing = await redis.get(DATA_KEY);
        if (existing) {
            await redis.set(BACKUP_KEY, existing);
        }

        data._savedAt = new Date().toISOString();
        await redis.set(DATA_KEY, data);

        return res.status(200).json({ ok: true, savedAt: data._savedAt });
    }

    return res.status(405).json({ error: 'Method not allowed' });
};
