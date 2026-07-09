# 企业微信转单看板 — 使用教程

> 面向任何拿到这个文件夹的人，不需要懂代码。

---

## 这是什么

一个本地小网页，把「销售订单全链路」「待转单」「企业微信AI转单推进表」这三份 Excel 汇总成一张表：每个运营公司已经登记了多少个企业微信群项目点、IT配置了多少、配置率多少、订单转单率多少。

计算逻辑和验证结果记录在 `../企业微信看板实现方案.md`，遇到的数据问题和待确认事项记录在 `../数据核实与待确认事项.md`，这份文档只讲"怎么用"。

---

## 第一次使用：装环境（只需要做一次）

1. 确认电脑装了 [Node.js](https://nodejs.org/)（装好之后打开终端/命令行，输入 `node -v` 能看到版本号就算装好了）
2. 打开终端，进到这个文件夹，跑一次：

   ```bash
   npm install
   ```

   （只是把 Excel 解析用的小工具装好，装一次就行，以后不用再装）

---

## 每天要做的事：换成当天的三份表

### 第一步：把当天的三份 Excel 放进对应文件夹，文件名不用改，直接覆盖

| Excel | 放在哪个文件夹 |
|---|---|
| 销售订单全链路.xlsx | `示例数据/` |
| 待转单-全量.xlsx | `示例数据/` |
| 企业微信AI转单推进表.xlsx | `basicData/` |

文件名不用完全一致，脚本是按"文件名里包含哪几个字"去找的（比如文件名里有"销售订单全链路"这几个字就认得），换成当天日期开头的文件名（比如"7.3日销售订单全链路.xlsx"）也没关系，只要旧的那份被换掉/删掉，避免同一个文件夹里有两份同类文件让脚本挑错。

### 第二步：重新生成数据

终端里跑：

```bash
node scripts/export_page_data.js
```

跑完会打印一行类似这样的结果，确认一下数字：

```
已导出: .../data/page_data.json
快照截止日期: 0702（晚于这天的推进表 OK-MMDD 确认状态不计入）
销售订单全链路 5002 行
...
```

**注意"快照截止日期"这一行**——它决定了推进表里哪些"OK-日期"状态算数。默认是 `0702`，如果今天换成了 7.3 的数据，要加上 `--cutoff` 手动指定：

```bash
node scripts/export_page_data.js --cutoff 0703
```

（这一步是必须的，不指定的话默认还是按 7.2 那天算，会漏掉 7.3 当天新确认的配置进度。规则是：填当天的日期，格式是月份+日期共4位数字，比如 7 月 3 日就是 `0703`。）

### 第三步：打开看板

第一次打开，终端里跑：

```bash
npm start
```

会打印一行 `企业微信看板已启动: http://localhost:3000/`，把这个网址复制到浏览器打开就能看到表格。

**这个终端窗口不要关**——关了看板就打不开了。如果不小心关了，重新跑一次 `npm start` 就行。

如果看板已经开着，只是换了新的 Excel 重新跑了第二步，浏览器里点一下页面上方的「重新加载 page_data.json」按钮就能看到新数字，不用重启服务。

---

## 分享给别人：生成公网链接

如果要把看板发给同事（比如利拉姐、虎哥），可以用隧道工具把本地 3000 端口暴露到公网：

### 用 localtunnel（默认，无需安装额外软件）

```bash
bash start-tunnel.sh
```

会打印一个 `https://xxx.loca.lt` 链接，复制发给别人就行。

### 用 ngrok（需要先下载 ngrok 客户端）

```bash
ngrok config add-authtoken <你的authtoken>
ngrok http 3000
```

### 注意事项

- localtunnel 链接在重启后会变化（随机子域名），如果每天需要固定链接，建议用 ngrok
- localtunnel 首次访问会显示一个 "Click to Continue" 页面，点击即可进入看板
- **看板页面里的上传功能在公网链接下同样可用**，对方上传 Excel 后能直接在浏览器里看到计算结果

---

## 更快的方式：不跑命令，直接在网页里上传（适合临时看一眼）

如果不想开终端敲命令，也可以：

1. 跑一次 `npm start` 把网页打开（这一步还是要做一次，之后就不用再动终端了）
2. 在页面**最上方**「上传区域」把当天的三份 Excel 分别选好
3. 填一下「快照截止日期」（比如 `0703`）
4. 点击「用上传的文件重新计算」

网页会当场重新算一遍，不用等命令行。**但这种方式算出来的结果只在当前这次打开的网页里看得到，关掉浏览器/刷新页面就没了**，不会保存成文件。如果想让这份数据长期留档（比如给别人转发链接、明天还能看到今天的），还是要走上面"每天要做的事"里的命令行方式。

---

## 常见问题

**Q: 网页打开是空的/加载失败？**
先看看有没有跑过 `node scripts/export_page_data.js`（这一步会生成 `data/page_data.json`，网页默认读这个文件）。如果这个文件本来就没有，就先跑一次这个命令。

**Q: 网页上"运营公司"这张表，多了一家没见过的公司/少了一家？**
八成是公司名字写法变了（比如括号从"（）"变成"()"，或者多了个空格）导致脚本没认出这是同一家公司，把它当成了新公司。把这个情况记到 `../数据核实与待确认事项.md` 里，找技术同学看一下 `normalizeCompanyName` 这个函数要不要加规则。

**Q: 数字看起来不对，怎么排查？**
1. 先确认「快照截止日期」有没有填对（是不是忘了改成当天日期）
2. 确认三份 Excel 有没有真的换成当天的（文件夹里是不是还是旧文件）
3. 如果还是不对，去看 `../数据核实与待确认事项.md`，看是不是已知问题

**Q: 想同时保留每天的历史数据，方便以后回看趋势？**
目前 `data/page_data.json` 每次都会被覆盖，不会自动存历史。如果需要留档，跑完 `export_page_data.js` 之后手动把 `data/page_data.json` 复制一份改名（比如 `page_data_0703.json`）存起来即可。

---

## 文件夹里都是什么

```
wecomToAgent/
├── index.html                # 看板网页本身
├── server.js                 # 本地小服务器，负责把网页发出来
├── package.json
├── scripts/
│   ├── export_page_data.js   # 把 Excel 转成网页能读的 JSON，每天换数据后要跑这个
│   └── lib/page_logic.js     # 核心计算逻辑（配置率、AI转单占比怎么算的都在这）
├── basicData/                 # 放"企业微信AI转单推进表.xlsx"
├── 示例数据/                  # 放"销售订单全链路.xlsx"和"待转单-全量.xlsx"
├── vendor/xlsx.full.min.js    # 网页里手动上传 Excel 用到的解析库，不用管
├── data/page_data.json        # 脚本算出来的结果，网页默认读这个文件
├── netlify.toml               # Netlify 部署配置
├── netlify/functions/         # Netlify 云函数（密码保护 + 数据代理）
│   ├── auth.js                # 登录认证（校验密码、签发 Cookie）
│   └── data.js                # 数据代理（校验登录后才能读取 data/ 下的文件）
└── README.md
```

---

## Netlify 线上部署（带密码保护）

### 前置准备

1. 注册 [Netlify](https://www.netlify.com/) 账号（免费）
2. 把代码推到 GitHub 仓库（当前仓库：`https://github.com/slaveofmatlab/tezan001`）

### 部署步骤

**第一步：连接仓库**

在 Netlify 里点 "Add new site" → "Import an existing project" → 选 GitHub → 授权并选择 `slaveofmatlab/tezan001`。

**第二步：关键配置**

Netlify 会让你填部署设置，填这两项：

| 设置 | 值 |
|------|-----|
| Base directory | `03-企业微信追踪/wecomToAgent/` |
| Build command | `rm -rf dist && mkdir -p dist && cp index.html dist/ && cp -r data dist/ && cp -r vendor dist/` |
| Publish directory | `dist` |

（这些配置已经写在 `netlify.toml` 里了，Netlify 会自动读取，但第一步创建时手动确认一下。）

**第三步：设置环境变量（密码）**

在 Netlify 站点后台 → **Site configuration** → **Environment variables**，添加：

| 变量名 | 值 | 说明 |
|--------|-----|------|
| `SITE_PASSWORD` | 你设定的密码 | 访问看板时输入的密码 |
| `COOKIE_SECRET` | 另一串随机字符（可选，不设则等于 SITE_PASSWORD） | Cookie 签名密钥，建议设成和密码不一样 |

> ⚠️ **重要**：设置完环境变量后，需要重新部署一次才能生效（Netlify 后台点 "Retry deploy" 或 "Trigger deploy"）。

**第四步：部署**

Netlify 会自动开始第一次部署。部署完成后会给你一个 `https://xxx.netlify.app` 的域名。

**之后每次更新看板，只需要 push 代码**，Netlify 会自动重新部署。

### 访问效果

打开 Netlify 分配的域名 → 先看到一个密码输入页面 → 输入正确密码后进入看板。

- 登录状态保持 **24 小时**，之后需要重新输入密码
- 密码不会出现在网页源代码里，即使查看源码也看不到
- data/ 下的数据文件同样受保护，未登录无法直接访问

### 如何更新看板数据

1. 本地更新三份 Excel，放到对应文件夹
2. 跑 `node scripts/export_page_data.js --cutoff MMDD`
3. `git add` + `git commit` + `git push`
4. Netlify 检测到 push → 自动重新部署 → 线上看板更新完成
