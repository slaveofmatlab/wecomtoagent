# 企业微信看板 — 维护交接文档

> 面向接手维护的技术同学。使用教程见 `README.md`，这里只讲"怎么改代码、怎么部署、有哪些坑"。

---

## 部署架构

```
tezan001（主仓库，Lucy 本地 + GitHub）
  └── 03-wecom-tracking/wecomToAgent/  ← 源码在这里
        │  git subtree split + push
        ▼
slaveofmatlab/wecomtoagent（独立 GitHub 仓库）
        │  Render 监听这个仓库，有 push 就自动重新部署
        ▼
Render（免费托管，Node.js 服务）
  URL: https://wecomtoagent.onrender.com
```

- **源码**在 `tezan001` 主仓库的 `03-wecom-tracking/wecomToAgent/` 目录下维护
- 每次要更新线上版本，从主仓库用 `git subtree split + push --force` 推到 `wecomtoagent` 仓库
- `wecomtoagent` 仓库本身**不直接编辑**，只作为 Render 的部署源

---

## 环境变量（Render 后台配置）

在 Render → Dashboard → wecomtoagent → Environment 里设置：

| 变量名 | 说明 | 示例 |
|--------|------|------|
| `SITE_PASSWORD` | 看板访问密码 | `yihai2026` |
| `GITHUB_TOKEN` | GitHub 细粒度 PAT，用于把数据持久化写回 GitHub | `github_pat_xxx` |
| `PORT` | Render 自动注入，不需要手动设 | — |

### 如何生成 GITHUB_TOKEN

1. GitHub → Settings → Developer settings → Fine-grained personal access tokens → Generate new token
2. Resource owner：`slaveofmatlab`
3. Repository access：只选 `wecomtoagent`
4. Permissions → Contents：Read and write
5. 生成后复制，粘贴到 Render 环境变量

---

## 数据持久化机制

Render 免费版的磁盘是**临时的**（每次重启/重新部署都会重置），所以真正的持久化靠 GitHub：

```
用户上传 Excel
  → POST /api/upload
  → 服务器内存（currentPageData / currentTrends）
  → 写磁盘（data/page_data.json、data/trends.json）
  → 通过 GitHub Contents API 写回 wecomtoagent/data/ 目录
        page_data.json：最新一次上传的汇总（companySummary、KPI 等，< 1MB）
        trends.json：所有日期的历史趋势数据

服务器启动时（init()）：
  1. 先从 GitHub 拉 page_data.json 和 trends.json
  2. 拉不到（GitHub TOKEN 没配 / 文件不存在）则读磁盘
  3. 磁盘也没有则内存为空，看板显示空白（需要用户重新上传）
```

**注意**：`data/page_data.json` 和 `data/trends.json` 同时被 git 追踪（作为兜底快照）。每次代码部署，Render 的磁盘会从 git 仓库恢复这两个文件作为基准，之后用户的上传会通过 API 覆盖 GitHub 上的文件。

---

## 已知坑 / 容易出问题的地方

### ⚠️ 坑 1：推代码后看板可能显示"旧数据"

**原因**：`data/trends.json` 和 `data/page_data.json` 是 git 追踪的文件。每次 `git subtree push --force` 会把 git 里记录的版本推到 wecomtoagent，覆盖掉 GitHub API 写进去的更新版本。

**具体场景**：
- 有人上传了 7月10日数据 → GitHub 上 trends.json 新增了 07-10 条目
- 你推了一次代码（subtree push）→ git 里的 trends.json 只有 07-09 → 07-10 被覆盖掉了
- Render 重启 → 服务器从 GitHub 读 trends.json → 07-10 不见了

**处理方法**：
1. 推代码之前，先把 GitHub 上的最新 `data/trends.json` 下载下来，覆盖本地的，然后 `git add + git commit`，再推
2. 或者：推代码后，告知用户重新上传一次最新数据（最简单）

**长期建议**：如果频繁改代码又频繁更新数据，考虑在部署脚本里加一步"推代码前从 GitHub 同步数据文件"（见下方"推代码操作步骤"）。

---

### ⚠️ 坑 2：GitHub API 写入 409 SHA 冲突

**原因**：GitHub Contents API 的 PUT 要求提供文件当前的 blob SHA。如果在我们 GET SHA 和 PUT 之间，文件被另一个操作（代码 push / 另一个并发上传）改变了，SHA 就对不上，返回 409。

**现状**：`server.js` 的 `githubPutFile` 函数已加了 **retry-on-409** 逻辑（失败后重新 GET SHA 再重试一次），正常情况下会自动处理。

**如果还是持续 409**：大概率是有其他操作在同时改这个文件（比如刚刚推过代码），等 30 秒再上传一次通常就好了。

---

### ⚠️ 坑 3：Render 免费版会"睡眠"

Render 免费版 **15 分钟无流量会自动休眠**，下次有人访问时需要 30～60 秒重新启动，期间页面会加载很慢或报错。这是免费版限制，不是 bug。

处理方法：
- 可以用 cron-job.org 之类的服务每 10 分钟 ping 一次 `https://wecomtoagent.onrender.com/api/auth`，防止休眠（免费）
- 或者升级 Render 付费版（$7/月）

---

### ⚠️ 坑 4：page_data.json 不能超过 1MB（GitHub Contents API 限制）

GitHub Contents API 对超过 1MB 的文件只返回 SHA、不返回内容（`content` 字段为空）。服务器 `init()` 会读不到内容，回退到磁盘版。

**现在**：服务器只把 `companySummary`、`pendingTotals`、`logStats` 等渲染所需字段（slim 版）推到 GitHub，原始的 salesRows/pendingRows 不推，所以 page_data.json 通常只有几 KB，不会有问题。

**注意**：如果将来往 slim 版里加字段，确保大小保持在 1MB 以内。

---

### ⚠️ 坑 5：前端上传的"备注列"依赖消息日志

"备注"列（下单方式：图片/PDF/Excel）来自**企业微信消息日志 Excel**，不是从销售订单里算出来的。

- 如果上传时没选日志文件，服务器会**保留上一次的备注**（从内存里的旧数据复用）
- Render 重启后内存清空，如果 GitHub 上的 page_data.json 里没有备注，下次启动后备注就没了
- 如果需要备注持久化，每次上传时记得带上日志文件

---

## 推代码操作步骤（从主仓库 tezan001）

```bash
# 在 tezan001 主仓库目录下执行

# 【推荐】推代码前先同步最新 trends.json（防止覆盖上传的数据）
git fetch wecom master:wecom-latest 2>/dev/null
git show wecom-latest:data/trends.json > 03-wecom-tracking/wecomToAgent/data/trends.json 2>/dev/null && \
  git add 03-wecom-tracking/wecomToAgent/data/trends.json && \
  git diff --cached --quiet || git commit -m "chore: sync trends.json before deploy [skip render]"
git branch -D wecom-latest 2>/dev/null || true

# 推代码
git subtree split --prefix="03-wecom-tracking/wecomToAgent" -b temp-wecom-deploy
git push wecom temp-wecom-deploy:master --force
git branch -D temp-wecom-deploy
```

其中 `wecom` 是指向 `git@github.com:slaveofmatlab/wecomtoagent.git` 的 remote，可以用 `git remote -v` 确认。

如果还没加 remote：
```bash
git remote add wecom git@github.com:slaveofmatlab/wecomtoagent.git
```

---

## 代码结构速览

```
wecomToAgent/
├── index.html              # 前端全部逻辑（单文件，Chart.js + html2canvas）
├── server.js               # Node.js HTTP 服务器（无 Express）
│                             ├── 静态文件服务
│                             ├── POST /api/upload  ← 核心：解析 Excel + 写 GitHub
│                             ├── GET  /api/data    ← 返回内存中的 page_data / trends
│                             └── POST /api/auth    ← 密码登录 + Cookie
├── scripts/
│   ├── export_page_data.js  # 本地导出单日数据（命令行用，线上不用）
│   ├── export_trend_data.js # 本地扫描所有日期生成趋势（命令行用）
│   └── lib/page_logic.js    # 核心计算逻辑（被 server.js 和上面两个脚本共用）
├── basicData/               # 推进表 Excel（随代码一起 git 管理）
├── 示例数据/                # 本地测试用的历史 Excel（随代码一起 git 管理）
├── vendor/xlsx.full.min.js  # 前端 Excel 解析库（绕过认证直接返回，见 server.js）
├── data/
│   ├── page_data.json       # 最新汇总快照（git 追踪，兜底用）
│   └── trends.json          # 历史趋势（git 追踪，兜底用）
└── MAINTAINER.md            # 本文档
```

---

## 本地开发调试

```bash
cd 03-wecom-tracking/wecomToAgent
npm install

# 不接 GitHub（纯本地测试）
node server.js
# → http://localhost:3000  无密码直接进

# 带密码 + 带 GitHub 持久化
SITE_PASSWORD=test GITHUB_TOKEN=xxx node server.js
```

本地开发时不需要 GITHUB_TOKEN，上传的数据只写磁盘，刷新后从磁盘读。

---

## 联系人

| 角色 | 说明 |
|------|------|
| Lucy / 陆心怡 | 原开发接手人，益海嘉里业务方 |
| 何应钦 | 001/004 部署对接人（Render / Phoenix 发版），遇到服务器问题可找他 |
| 利拉 / Nina | 业务负责人，数字口径有疑问找她确认 |
