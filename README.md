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

## 常用命令

以下命令均在项目根目录执行。需要 Python 3.10 或更新版本，安装依赖：`python -m pip install -r requirements.txt`。

收取第三周 PDF，自动打包并更新 Excel、网页：

```powershell
. .\邮箱配置.local.ps1
python scripts/收作业.py --since 2026-09-01 --assignment 3 --output 作业收集_2026秋
```

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

将批改 PDF 放在 `作业收集_2026秋/第3周作业批改/`，文件名以 `学号_姓名_` 开头。收件地址读取本地 `学生邮箱.xlsx`，不按学号猜测邮箱。邮箱表中只有一个地址时以该地址为准；多个地址时尝试用附件文件名里的原邮件标识匹配。

先预览返还清单：

```powershell
python scripts/回作业.py --week 3
```

确认后实际发送：

```powershell
python scripts/回作业.py --week 3 --send
```

自动读取根目录 `邮箱配置.local.ps1`，发件使用 `HOMEWORK_SMTP_AUTH_CODE`，没有单独设置时使用 `HOMEWORK_AUTH_CODE`。同一学生同一地址的 PDF 合并一封邮件，主题为“第3周作业批改”。其他周修改 `--week`，也可用 `--folder` 指定批改目录。

`第3周返还清单.json` 保存匹配清单和待核对项，`第3周返还记录.json` 保存发送结果。重复运行跳过同一批附件的成功发送记录；网络中断导致“结果不确定”时也会跳过，应先查发件箱核实。修改附件或收件地址会视为新邮件。以上记录及批改文件不上传 GitHub，回作业不会修改提交统计。

整理发件人邮箱（以现有作业统计为名单，不连接邮箱）：

```powershell
python scripts/整理学生邮箱.py --output 作业收集_2026秋
```

生成本地 `学生邮箱.xlsx`，包含“学生邮箱”“来源明细”“待人工核对”。保留多个邮箱，不按学号猜测邮箱；非标准主题、姓名不一致和同一邮箱关联多个学号会提示确认。只使用本地收取记录，不代表邮箱中的全部历史发件人。邮箱表不发布到网站或 GitHub。

- [完整使用指南与 GitHub Pages 发布步骤](docs/使用指南.md)
- [收作业说明](docs/收作业使用说明.md)

现有 Excel 中的学生、已交标记、备注及周次会保留，已交周数重新计算。原表的自定义列、其他工作表与样式暂不保留。运行前请保存并关闭 Excel。
