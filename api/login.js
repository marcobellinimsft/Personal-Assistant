const { createToken, setSessionCookie, timingSafeCompare } = require('./_auth');

module.exports = async function handler(req, res) {
    if (req.method !== 'POST') {
        return res.status(405).json({ error: 'Method not allowed' });
    }

    const { password } = req.body || {};
    const expected = process.env.AUTH_PASSWORD;

    if (!expected) {
        return res.status(500).json({ error: 'AUTH_PASSWORD not configured' });
    }

    if (!password || !timingSafeCompare(password, expected)) {
        return res.status(401).json({ error: 'Invalid password' });
    }

    const token = await createToken();
    res.setHeader('Set-Cookie', setSessionCookie(token));
    return res.status(200).json({ ok: true });
};
