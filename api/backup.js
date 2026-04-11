const { verifyToken } = require('./_auth');
const { Redis } = require('@upstash/redis');

const redis = new Redis({
    url: process.env.KV_REST_API_URL,
    token: process.env.KV_REST_API_TOKEN,
});

module.exports = async function handler(req, res) {
    const payload = await verifyToken(req);
    if (!payload) {
        return res.status(401).json({ error: 'Not authenticated' });
    }

    if (req.method === 'GET') {
        // Return backup data
        const backup = await redis.get('pa_data_backup');
        const current = await redis.get('pa_data');
        return res.status(200).json({
            backup: backup || null,
            current: current || null,
            backupSavedAt: backup?._savedAt || null,
            currentSavedAt: current?._savedAt || null
        });
    }

    if (req.method === 'POST') {
        // Restore from backup
        const backup = await redis.get('pa_data_backup');
        if (!backup) {
            return res.status(404).json({ error: 'No backup found' });
        }
        // Save current as a secondary backup first
        const current = await redis.get('pa_data');
        if (current) {
            await redis.set('pa_data_backup2', current);
        }
        // Restore backup as current
        await redis.set('pa_data', backup);
        return res.status(200).json({ ok: true, restoredFrom: backup._savedAt });
    }

    return res.status(405).json({ error: 'Method not allowed' });
};
