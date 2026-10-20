# 作业收集与公开查询

本地收取 PDF 作业、更新跨周 Excel、平铺打包 ZIP，并生成学号查询网页。push 到 `main` 后，GitHub Pages 自动发布 `site/index.html`。

## 目录结构

```text
scripts/          收作业、回作业、生成网页的运行脚本
tests/            离线验证脚本
docs/             使用指南、收作业说明
templates/        网页模板
邮箱配置*.ps1     根目录邮箱配置；仅示例上传
site/             公开查询网页，上传 GitHub
.github/          自动发布工作流
作业收集_2026秋/   Excel、收取记录、附件、ZIP，仅本地保存
本地依赖/         本地 Python 依赖，不上传
```

真实邮箱配置位于 `邮箱配置.local.ps1`，不会上传 GitHub。网站仅公开学号和各周状态。

## 两台电脑共享作业数据

`作业收集_2026秋/` 可作为独立私有 Git 仓库使用，外层公开仓库继续忽略整个目录。两台电脑分别克隆外层项目，再将同一个私有数据仓库克隆到该子目录。不要把数据目录强制加入外层仓库。

每次收取前执行 `git -C 作业收集_2026秋 pull --ff-only`；完成后在该子仓库中 add、commit、push。两台电脑避免同时修改；遇到冲突先核对，不强制覆盖。数据仓库仅追踪 `作业统计.xlsx`、`收取记录.json` 和 `学生邮箱.xlsx`（另保留仓库说明与 Git 配置）。原始附件、批改文件、返还清单与记录、异常日志和 ZIP 均只保留本地，不同步。另一台电脑进行打包或返还前需另行准备附件和相关资料。邮箱配置每台电脑单独保存，选课名单也需在两台电脑上保持一致。首次连接私有远端以及每台电脑的 GitHub 登录须分别完成。

## 常用命令

以下命令均在项目根目录执行。需要 Python 3.10 或更新版本，安装依赖：`python -m pip install -r requirements.txt`。

收取第三周 PDF，自动打包并更新 Excel、网页：

```powershell
. .\邮箱配置.local.ps1
python scripts/收作业.py --since 2026-09-01 --assignment 3 --output 作业收集_2026秋
```

兼容 QQ 邮箱超大附件：自动尝试下载正文中的中转站 PDF 链接；需要网页登录的项目列在本地 `超大附件待处理.txt`。支持 `--retry-large` 重试，以及 `--import-large "任务编号" "PDF路径"` 补录网页下载的文件并更新统计。具体用法见 [超大附件说明](docs/收作业使用说明.md#qq-邮箱超大附件)。

只打包本地 PDF：

```powershell
python scripts/收作业.py --pack-only --assignment 3 --output 作业收集_2026秋
```

只从现有 Excel 更新网页：

```powershell
python scripts/生成网页.py
```

当前电脑如果找不到 `python`，改用完整路径，例如：

```powershell
& "D:\Users\10338\anaconda3\python.exe" .\scripts\生成网页.py
```

运行验证：

```powershell
python tests/验证跨周统计.py
python tests/验证网页.py
python tests/验证超大附件.py
```

发布修改：

```powershell
git add .
git status --short
git commit -m "Organize project files"
git push
```

## 详细说明

## 返还批改作业

每次收作业会自动更新本地 `学生邮箱.xlsx`：仅使用主题符合收取格式且学号在收取范围内的邮件，按学号覆盖为本轮最新邮件的唯一发件地址，不汇总历史邮箱。主题不规范或发件地址不唯一时跳过；本轮未匹配到的学生保留原邮箱。邮箱表纳入私有数据仓库同步。

批改 PDF 可统一重命名为 `学号_姓名_第4周作业批改.pdf`（姓名后无空格）：

```powershell
python scripts/重命名批改作业.py --week 4 --dry-run
python scripts/重命名批改作业.py --week 4
```

默认处理 `作业收集_2026秋/第4周作业批改/`，可用 `--folder` 指定目录。同一目录同一学生的多份 PDF 自动加 `_2`、`_3`，已有规范文件名保持不变，不覆盖文件。旧名映射保存于批改目录的 `重命名记录.json`。返还脚本会通过该记录追溯原邮件标识，核对学号、姓名和周次后匹配多个邮箱。旧邮件的姓名备注或非标准主题备注可由本次原邮件证据核实；邮箱关联多个学号仍须人工核对。请保留重命名记录。

将批改 PDF 放在 `作业收集_2026秋/第3周作业批改/`，文件名以 `学号_姓名_` 开头。收件地址读取本地 `学生邮箱.xlsx`，不按学号猜测邮箱。邮箱表中只有一个地址时以该地址为准；多个地址时尝试用附件文件名里的原邮件标识匹配。

先预览返还清单：

```powershell
python scripts/回作业.py --week 3
```

确认后实际发送：

```powershell
python scripts/回作业.py --week 3 --send
```

自动读取根目录 `邮箱配置.local.ps1`，收作业和回作业统一使用 `HOMEWORK_AUTH_CODE` 中的 SMTP/IMAP 授权码，不再使用单独的 `HOMEWORK_SMTP_AUTH_CODE`。同一学生同一地址的 PDF 合并一封邮件，主题为“第3周作业批改”。其他周修改 `--week`，也可用 `--folder` 指定批改目录。

`第3周返还清单.json` 保存匹配清单和待核对项，`第3周返还记录.json` 保存发送结果。重复运行跳过同一批附件的成功发送记录；网络中断导致“结果不确定”时也会跳过，应先查发件箱核实。修改附件或收件地址会视为新邮件。以上记录及批改文件不上传 GitHub，回作业不会修改提交统计。

整理发件人邮箱（以现有作业统计为名单，不连接邮箱）：

```powershell
python scripts/整理学生邮箱.py --output 作业收集_2026秋
```

根据本地收取记录中的规范主题，按学号更新 `学生邮箱.xlsx`，每人采用最新匹配邮件的唯一发件地址，覆盖旧邮箱；未匹配到的学生保留原邮箱，不汇总历史邮箱。正常收作业时会自动更新，无需单独运行。邮箱表仅同步到私有数据仓库，不发布到公开网站或公开仓库。

- [完整使用指南与 GitHub Pages 发布步骤](docs/使用指南.md)
- [收作业说明](docs/收作业使用说明.md)

现有 Excel 中的学生、已交标记、备注及周次会保留，已交周数重新计算。原表的自定义列、其他工作表与样式暂不保留。运行前请保存并关闭 Excel。
