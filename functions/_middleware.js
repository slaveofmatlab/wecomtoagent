// Cloudflare Pages Functions — 密码保护中间件
export async function onRequest(context) {
  const { request, next } = context;
  const url = new URL(request.url);

  // 登录接口和登录页面直接放行
  if (url.pathname === "/api/auth" || url.pathname === "/login.html") {
    return next();
  }

  // 静态资源也放行（登录页需要 CSS/JS）
  if (url.pathname.startsWith("/vendor/")) {
    return next();
  }

  // 检查 auth cookie
  const cookie = request.headers.get("Cookie") || "";
  if (cookie.includes("wecom_auth=1")) {
    return next();
  }

  // 未登录 -> 返回登录页面
  return new Response(LOGIN_HTML, {
    headers: { "Content-Type": "text/html; charset=utf-8" },
  });
}

const LOGIN_HTML = `<!doctype html>
<html lang="zh-CN">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width,initial-scale=1">
<title>企业微信看板 — 登录</title>
<style>
  * { box-sizing:border-box; }
  body { margin:0; display:flex; align-items:center; justify-content:center; min-height:100vh; background:#f6f8fb; font-family:-apple-system,"Microsoft YaHei",sans-serif; }
  .box { background:#fff; border-radius:14px; box-shadow:0 10px 30px rgba(15,23,42,.08); padding:32px 28px; width:360px; max-width:90vw; text-align:center; }
  h1 { margin:0 0 6px; font-size:20px; color:#172033; }
  p { margin:0 0 20px; font-size:13px; color:#64748b; }
  input { width:100%; padding:10px 14px; border:1px solid #d9e0ea; border-radius:8px; font-size:15px; text-align:center; margin-bottom:14px; }
  button { width:100%; padding:10px; border:0; border-radius:8px; background:#6d8f3c; color:#fff; font-size:15px; font-weight:600; cursor:pointer; }
  button:hover { opacity:.9; }
  #err { color:#b91c1c; font-size:13px; margin-top:8px; }
</style>
</head>
<body>
<div class="box">
  <h1>🔐 企业微信转单看板</h1>
  <p>请输入密码查看数据</p>
  <input type="password" id="pwd" placeholder="输入密码" autofocus>
  <button onclick="login()">进入看板</button>
  <div id="err"></div>
</div>
<script>
async function login() {
  var pwd = document.getElementById("pwd").value;
  if (!pwd) return;
  var res = await fetch("/api/auth", {
    method: "POST",
    headers: {"Content-Type":"application/json"},
    body: JSON.stringify({password:pwd})
  });
  var data = await res.json();
  if (data.ok) { window.location.href = "/"; }
  else { document.getElementById("err").textContent = data.error || "密码错误"; }
}
document.getElementById("pwd").addEventListener("keydown", function(e) {
  if (e.key === "Enter") login();
});
</script>
</body>
</html>`;
