// Cloudflare Pages Functions — 密码验证
export async function onRequestPost(context) {
  const { request, env } = context;
  try {
    const body = await request.json();
    if (body.password === env.SITE_PASSWORD) {
      return new Response(JSON.stringify({ ok: true }), {
        headers: {
          "Content-Type": "application/json",
          "Set-Cookie": "wecom_auth=1; Path=/; Max-Age=86400; HttpOnly; SameSite=Lax",
        },
      });
    }
  } catch (e) {}
  return new Response(JSON.stringify({ ok: false, error: "密码错误" }), {
    status: 401,
    headers: { "Content-Type": "application/json" },
  });
}
