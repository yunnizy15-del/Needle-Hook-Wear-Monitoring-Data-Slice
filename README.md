# 针钩摩擦系数数据切片工具

本项目用于从 `xlsx` 中提取 `mu_true`（摩擦系数）数据，按 `tlife` 规则分为有效/失效两组，并按固定秒数切片保存为 CSV。

## 功能说明

1. 支持 GUI 图形界面操作（中文界面）。
2. 支持多工作表选择（多选框，默认选前 2 个）。
3. 支持命令行处理（CLI）。
4. 输出统一保存到 `valid/` 和 `invalid/` 两个文件夹。
5. 多工作表处理时，文件序号连续递增，不按工作表分目录。
6. 支持先丢弃每个工作表前 `x` 小时数据（默认 `1h`）。

## 数据划分规则

设 `drop_minutes` 默认为 30 分钟：

1. 先丢弃每个工作表前 `drop_initial_hours * 3600` 秒（默认 `1h`）。
2. 对剩余数据：`t < tlife - drop_minutes * 60` -> 有效数据（`valid`）。
3. 对剩余数据：`t > tlife + drop_minutes * 60` -> 失效数据（`invalid`）。
4. 中间区间（含边界）丢弃。

切片列：

1. 时间列：列名包含 `t_s` 或 `time`
2. 摩擦系数列：列名包含 `mu_true`

## 依赖安装

项目依赖文件为 `requirement.txt`：

```powershell
& .\.venv\Scripts\python.exe -m pip install -r requirement.txt
```

如果不使用虚拟环境，可改为：

```powershell
python -m pip install -r requirement.txt
```

## GUI 使用

启动：

```powershell
& .\.venv\Scripts\python.exe .\gui_mu_splitter.py
```

操作步骤：

1. 选择 `xlsx` 文件（默认会尝试读取当前目录 `data.xlsx`）。
2. 在工作表多选框里勾选要处理的工作表（默认前 2 个已选中）。
3. 输入 `tlife`（秒）。
4. 设置切片秒数（默认 `5`）。
5. 设置剔除窗口分钟数（默认 `30`）。
6. 设置丢弃前 `x` 小时（默认 `1`）。
7. 选择输出目录并点击“开始处理”。

输出结构：

```text
输出目录/
  valid/
    000001.csv
    000002.csv
    ...
  invalid/
    000001.csv
    000002.csv
    ...
```

## CLI 使用

查看帮助：

```powershell
& .\.venv\Scripts\python.exe .\split_mu_by_tlife.py --help
```

示例：

```powershell
& .\.venv\Scripts\python.exe .\split_mu_by_tlife.py --input .\data.xlsx --tlife 36000 --slice-seconds 5 --drop-minutes 30 --drop-initial-hours 1 --sheet closed_loop_1 --output-dir .\result
```

## PyInstaller 打包

已提供配置文件 `main.spec`，并配置了软件图标 `app.ico`。

1. 安装 PyInstaller：

```powershell
& .\.venv\Scripts\python.exe -m pip install pyinstaller
```

2. 开始打包：

```powershell
& .\.venv\Scripts\python.exe -m PyInstaller .\main.spec --clean
```

3. 产物位置：

`dist/Needle-Hook-Wear-Monitoring-Data-Slice.exe`

## 主要文件

1. `gui_mu_splitter.py`：图形界面入口。
2. `split_mu_by_tlife.py`：核心处理逻辑 + CLI。
3. `main.spec`：PyInstaller 打包配置。
4. `app.ico`：程序图标。
5. `requirement.txt`：Python 依赖列表。
