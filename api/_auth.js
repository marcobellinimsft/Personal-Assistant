const { jwtVerify, SignJWT } = require('jose');

const JWT_SECRET_KEY = () => new TextEncoder().encode(process.env.JWT_SECRET);
const COOKIE_NAME = 'pa_session';

async function createToken() {
    return new SignJWT({ sub: 'marco' })
        .setProtectedHeader({ alg: 'HS256' })
        .setIssuedAt()
        .setExpirationTime('7d')
        .sign(JWT_SECRET_KEY());
}

async function verifyToken(req) {
    const cookies = parseCookies(req.headers.cookie || '');
    const token = cookies[COOKIE_NAME];
    if (!token) return null;
    try {
        const { payload } = await jwtVerify(token, JWT_SECRET_KEY());
        return payload;
    } catch {
        return null;
    }
}

function parseCookies(cookieHeader) {
    const cookies = {};
    cookieHeader.split(';').forEach(pair => {
        const [key, ...vals] = pair.trim().split('=');
        if (key) cookies[key.trim()] = vals.join('=').trim();
    });
    return cookies;
}

function setSessionCookie(token) {
    return `${COOKIE_NAME}=${token}; HttpOnly; Secure; SameSite=Strict; Path=/; Max-Age=${7 * 24 * 60 * 60}`;
}

function clearSessionCookie() {
    return `${COOKIE_NAME}=; HttpOnly; Secure; SameSite=Strict; Path=/; Max-Age=0`;
}

function timingSafeCompare(a, b) {
    const crypto = require('crypto');
    if (a.length !== b.length) {
        // Compare against itself to avoid timing leak on length
        crypto.timingSafeEqual(Buffer.from(a), Buffer.from(a));
        return false;
    }
    return crypto.timingSafeEqual(Buffer.from(a), Buffer.from(b));
}

module.exports = {
    createToken,
    verifyToken,
    setSessionCookie,
    clearSessionCookie,
    timingSafeCompare,
    COOKIE_NAME
};
