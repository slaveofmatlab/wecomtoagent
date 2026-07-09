#!/bin/bash
# 启动企业微信看板本地服务器 + 公网隧道
# 用法: bash start-tunnel.sh

ROOT="$(cd "$(dirname "$0")" && pwd)"
PORT="${PORT:-3000}"

echo "=== 企业微信转单看板 ==="

# 1. 启动本地服务器
node "$ROOT/server.js" &
SERVER_PID=$!
sleep 1

# 2. 启动隧道（优先用 ngrok，没有则用 localtunnel）
if command -v ngrok &> /dev/null; then
  echo "使用 ngrok 创建隧道..."
  ngrok http $PORT --log=stdout &
  TUNNEL_PID=$!
else
  echo "使用 localtunnel 创建隧道..."
  npx localtunnel --port $PORT --subdomain tezan001-wecom 2>&1 &
  TUNNEL_PID=$!
fi

echo ""
echo "本地地址: http://localhost:$PORT"
echo "按 Ctrl+C 停止所有服务"

trap "kill $SERVER_PID $TUNNEL_PID 2>/dev/null; exit 0" INT TERM
wait
