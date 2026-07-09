/**
 * 登录认证 Netlify Function
 *
 * GET  /api/auth → 检查 cookie 是否有效，返回 { authenticated: true/false }
 * POST /api/auth → 校验密码，成功则种 HttpOnly signed cookie
 *
 * 环境变量（在 Netlify Dashboard → Site settings → Environment variables 设置）：
 *   SITE_PASSWORD  — 访问密码（必填）
 *   COOKIE_SECRET  — Cookie 签名密钥（可选，默认与 SITE_PASSWORD 相同）
 */

const crypto = require("crypto");

const SITE_PASSWORD = process.env.SITE_PASSWORD;
const COOKIE_SECRET = process.env.COOKIE_SECRET || SITE_PASSWORD;
const COOKIE_MAX_AGE = 24 * 60 * 60; // 24 小时

/** 用 HMAC-SHA256 签名 + base64 编码 */
function signToken(payload) {
  const data = Buffer.from(JSON.stringify(payload)).toString("base64");
  const sig = crypto.createHmac("sha256", COOKIE_SECRET).update(data).digest("hex");
  return `${data}.${sig}`;
}

/** 验证 token：比对签名、检查过期 */
function verifyToken(token) {
  try {
    const idx = token.lastIndexOf(".");
    if (idx < 0) return null;
    const data = token.slice(0, idx);
    const sig = token.slice(idx + 1);
    const expected = crypto.createHmac("sha256", COOKIE_SECRET).update(data).digest("hex");
    if (sig !== expected) return null;
    const payload = JSON.parse(Buffer.from(data, "base64").toString("utf-8"));
    if (!payload.exp || payload.exp < Date.now()) return null;
    return payload;
  } catch (_) {
    return null;
  }
}

/** 从请求头里取出 auth_token cookie */
function getAuthCookie(headers) {
  const cookieStr = headers.cookie || "";
  const match = cookieStr.match(/(?:^|;\s*)auth_token=([^;]+)/);
  return match ? match[1] : null;
}

exports.handler = async (event) => {
  // GET: 检查登录状态
  if (event.httpMethod === "GET") {
    const token = getAuthCookie(event.headers);
    const ok = !!(token && verifyToken(token));
    return {
      statusCode: 200,
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ authenticated: ok }),
    };
  }

  // POST: 登录
  if (event.httpMethod === "POST") {
    // 未配置密码
    if (!SITE_PASSWORD) {
      return {
        statusCode: 500,
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ success: false, error: "服务器未配置访问密码，请联系管理员。" }),
      };
    }

    let body;
    try {
      body = JSON.parse(event.body || "{}");
    } catch (_) {
      return { statusCode: 400, body: JSON.stringify({ success: false, error: "请求格式错误" }) };
    }

    if (!body.password || body.password !== SITE_PASSWORD) {
      return {
        statusCode: 401,
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ success: false, error: "密码错误，请重试。" }),
      };
    }

    const payload = { exp: Date.now() + COOKIE_MAX_AGE * 1000 };
    const token = signToken(payload);

    return {
      statusCode: 200,
      headers: {
        "Content-Type": "application/json",
        "Set-Cookie": [
          `auth_token=${token}; HttpOnly; Secure; SameSite=Lax; Path=/; Max-Age=${COOKIE_MAX_AGE}`,
        ].join(", "),
      },
      body: JSON.stringify({ success: true }),
    };
  }

  return { statusCode: 405, body: "Method Not Allowed" };
};
