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

- [完整使用指南与 GitHub Pages 发布步骤](docs/使用指南.md)
- [收作业说明](docs/收作业使用说明.md)

现有 Excel 中的学生、已交标记、备注及周次会保留，已交周数重新计算。原表的自定义列、其他工作表与样式暂不保留。运行前请保存并关闭 Excel。
