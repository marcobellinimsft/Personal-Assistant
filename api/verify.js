const { verifyToken } = require('./_auth');

module.exports = async function handler(req, res) {
    if (req.method !== 'GET') {
        return res.status(405).json({ error: 'Method not allowed' });
    }

    const payload = await verifyToken(req);
    if (!payload) {
        return res.status(401).json({ authenticated: false });
    }

    return res.status(200).json({ authenticated: true, user: payload.sub });
};
