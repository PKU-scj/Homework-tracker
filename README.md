# 作业收集与公开查询

一个轻量的课程作业工具：本地读取 IMAP 邮件、下载附件、生成跨周 Excel，再生成可搜索的静态网页。将更新后的网页 push 到 GitHub，GitHub Pages 自动发布。

学生通过网址查询，无需 GitHub 账号。交了显示“已交”，其余留空。网站不提供登录、留言或后台收邮件功能。

## 工作流程

```text
本地收邮件 / 更新 Excel
          ↓
生成 site/index.html
          ↓
git commit + git push 到 main
          ↓
GitHub Actions 发布 GitHub Pages
```

GitHub Actions 仅部署已生成的网页，不连接邮箱，不需要邮箱授权码，不运行定时任务。只 push 源代码不会重新读取本地 Excel；更新数据后要先生成并提交 `site/index.html`。

## 项目结构

| 文件 | 作用 | 是否提交 |
| --- | --- | --- |
| `收作业.py` | 本地下载附件，累计统计，同时生成网页 | 是 |
| `生成网页.py` | 只从 Excel 生成公开网页 | 是 |
| `网页模板.html` | 页面样式、搜索和周次筛选 | 是 |
| `site/index.html` | 可直接发布的独立网页，包含公开学生数据 | 是 |
| `.github/workflows/pages.yml` | push 到 main 后自动发布 | 是 |
| `邮箱配置.example.ps1` | 无真实凭据的配置示例 | 是 |
| `邮箱配置.local.ps1` | 本地邮箱与授权码 | 否 |
| `作业收集_2026秋/` 等输出目录 | Excel、原始邮件记录、作业附件 | 否 |
| 选课名单、旧版 `回作业.py`、本地依赖 | 本地资料及旧工具 | 否 |

`.gitignore` 使用明确的文件清单：默认忽略本地文件，只纳入列出的源码、说明、工作流及 `site/index.html`。新增源码时，请同步修改该清单。忽略规则不会移除已经提交过的文件；不要用 `git add -f` 强行添加本地资料。

公开页面只导出姓名、学号、各周已交状态和统计更新时间，不导出邮件明细、附件路径及 Excel 备注。公开仓库中的网页数据及历史版本可被查看。

## 安装与本地配置

需要 Python 3.10 或更新版本、Git。以下命令在项目根目录执行：

```powershell
python -m pip install -r requirements.txt
```

新用户复制 `邮箱配置.example.ps1` 为 `邮箱配置.local.ps1`，在本地副本中填写邮箱和 IMAP 授权码，然后加载：

```powershell
. .\邮箱配置.local.ps1
```

已有 `邮箱配置.local.ps1` 的用户直接加载，不要覆盖。也可不加载配置，运行时按提示输入；QQ 邮箱需开启 IMAP 服务，使用授权码而非登录密码。

若 `python` 不在 PATH 中，可把命令中的 `python` 换成完整解释器路径。当前电脑可用：

```powershell
& "D:\Users\10338\anaconda3\python.exe" .\生成网页.py
```

## 本地更新作业

只收第三周、仅收取 2026-09-01 起的邮件（含当天）：

```powershell
python 收作业.py --since 2026-09-01 --assignment 3 --output 作业收集_2026秋
```

- 第四周改为 `--assignment 4`；同一学期始终使用同一个 `--output` 目录。
- 日期按邮箱服务器收件日期筛选。参数只限制本次下载，不删除历史记录。
- Excel 始终汇总该目录的所有周次，网页默认生成到 `site/index.html`。
- 自定义课程名：增加 `--title "课程名称"`。更换网页目录用 `--web-output`，但内置发布工作流固定读取 `site/index.html`。

邮件主题示例：`第3周作业+1234567890+张三`，也支持 `第三周作业 1234567890 张三`。学号为10位。

只有成功保存附件才算已交；不自动判断附件内容正确性。重复提交保留不同版本，同一周只计一次已交。默认只读收件箱，不发送、删除邮件或改变已读状态。网盘链接和正文不作为附件下载。

### 名单与统计

将选课名单放在项目根目录。自动识别唯一的文件名含“选课名单”的 `.xls` / `.xlsx` 文件，也可增加 `--roster "名单.xlsx"`。表头需包含“学号”“姓名”；xls 读取第一个工作表，xlsx 读取活动工作表。

有名单才会列出全班学生；没有名单只能列出邮件中识别到的学生。当前本地原名单已移出文件夹，重新生成全班 Excel 前请放回名单。已有 Excel 可直接用下一节的命令生成网页，无需名单。

Excel 第一列为姓名、第二列为学号，之后每周一列。名单外学生和姓名不一致会在备注中标注。“邮件明细”保存所有历史周次的处理记录。空白可能表示未交，也可能尚未收取该周数据。

`收取记录.json` 是累计记录，请保留。Excel 每次重新生成，手动编辑的内容不会保留；成绩及批改备注请另存。运行前请关闭 Excel。

### 只更新网页

现有 Excel 已正确时，直接运行，不登录邮箱、不重新生成 Excel：

```powershell
python 生成网页.py
```

默认读取 `作业收集_2026秋/作业统计.xlsx`，输出 `site/index.html`。自定义示例：

```powershell
python 生成网页.py --excel "其他学期/作业统计.xlsx" --title "我的课程"
```

双击 `site/index.html` 可本地预览，支持手机布局、姓名/学号搜索和周次筛选。更新网页模板后也需要重新运行此命令。页面更新时间来自源 Excel 的修改时间。

## 首次上传 GitHub 并发布

1. 在 GitHub 创建一个空的公开仓库，例如 `homework-tracker`。不要勾选自动创建 README，以便直接推送本项目。免费账号可使用公开仓库的 GitHub Pages。
2. 确认本地 `site/index.html` 是准备公开的版本，然后执行：

```powershell
git init -b main
git add .
git status --short
git commit -m "Set up homework tracker and Pages"
git remote add origin https://github.com/你的用户名/homework-tracker.git
git push -u origin main
```

替换仓库地址。首次使用 Git 如提示缺少身份，按提示设置提交用的 `user.name` 和 `user.email`。通过 GitHub Desktop 的 Publish repository 也可以完成首次上传，但默认分支需为 `main`。

3. 在仓库中进入 **Settings → Pages → Build and deployment → Source**，选择 **GitHub Actions**。
4. 进入 **Actions → Publish homework site → Run workflow**，选择 `main` 运行一次。如果首次 push 因尚未启用 Pages 而失败，完成第3步后重跑即可。
5. 等发布成功，在 **Settings → Pages** 查看网站地址，通常为 `https://你的用户名.github.io/homework-tracker/`，将实际地址发给同学。

工作流只部署 `site/index.html`。部署使用 GitHub 自动提供的权限，不需要手动创建访问令牌或填写邮箱 Secrets。默认部署 main，其他分支不会发布。

官方操作说明：[GitHub Pages 自定义工作流](https://docs.github.com/en/pages/getting-started-with-github-pages/using-custom-workflows-with-github-pages)。

## 日常更新：生成、提交、推送

```powershell
python 生成网页.py
git add site/index.html
git commit -m "Update homework submissions"
git push
```

如果刚运行过 `收作业.py`，它已经生成新版网页，可从 `git add` 开始。改了代码或说明时，也把相关文件一起提交。

push 到 main 后，等待 Actions 发布完成，同学刷新原网址即可看到新数据。无需重新创建网站。若 Git 提示无变更，说明没有新的网页内容可提交。

## 离线验证

```powershell
python 验证跨周统计.py
python 验证网页.py
```

测试使用临时示例数据，不依赖真实名单，不连接邮箱。验证跨周保留、重复提交计数、空白状态及网页数据转义和导出范围。