/**
 * 数据代理 Netlify Function
 *
 * GET /api/data?file=page_data.json → 校验 cookie 后返回 data/ 目录下的文件内容
 *
 * 环境变量：
 *   COOKIE_SECRET — 与 auth.js 中的签名密钥一致
 */

const crypto = require("crypto");
const fs = require("fs");
const path = require("path");

const COOKIE_SECRET = process.env.COOKIE_SECRET || process.env.SITE_PASSWORD;

/** 和 auth.js 完全一致的验证逻辑 */
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

function getAuthCookie(headers) {
  const cookieStr = headers.cookie || "";
  const match = cookieStr.match(/(?:^|;\s*)auth_token=([^;]+)/);
  return match ? match[1] : null;
}

exports.handler = async (event) => {
  // 1. 校验登录状态
  const token = getAuthCookie(event.headers);
  if (!token || !verifyToken(token)) {
    return {
      statusCode: 401,
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ error: "未登录或登录已过期，请刷新页面重新登录。" }),
    };
  }

  // 2. 读取文件参数
  const params = event.queryStringParameters || {};
  const filename = params.file;

  if (!filename) {
    return { statusCode: 400, body: JSON.stringify({ error: "缺少 file 参数" }) };
  }

  // 路径穿越防护
  if (filename.includes("..") || filename.includes("/") || filename.includes("\\")) {
    return { statusCode: 403, body: JSON.stringify({ error: "非法文件名" }) };
  }

  // 3. 读取文件
  // Netlify Function 的工作目录是 publish 目录（dist/），里面已有 data/ 文件夹
  const filePath = path.resolve("data", filename);
  try {
    const content = fs.readFileSync(filePath, "utf-8");
    return {
      statusCode: 200,
      headers: { "Content-Type": "application/json; charset=utf-8" },
      body: content,
    };
  } catch (err) {
    if (err.code === "ENOENT") {
      return { statusCode: 404, body: JSON.stringify({ error: `文件 ${filename} 不存在` }) };
    }
    return { statusCode: 500, body: JSON.stringify({ error: "读取文件失败" }) };
  }
};
