# KML 鸟类样线记录转换工具

本地运行的 Python 3 工具。把 KML 点标记转换为**待人工复核**的 Excel 鸟类记录表，再把 Excel 按日期和样线整理为 Word 样线数据表。只使用 Python 标准库，不上传数据，也不修改输入文件。

## Windows 使用

1. 安装 Python 3，确认 `py -3` 或 `python` 能运行。
2. 把一个 `.kml` 文件拖到 `拖入KML生成Excel.cmd` 图标上。
3. 窗口会显示 Excel 路径；文件保存在本目录 `outputs/` 的新子目录。
4. 在 Excel 中核对物种、数量、行为、生境和“复核状态”。
5. 把该 `.xlsx` 文件拖到 `拖入Excel生成Word.cmd`，得到 Word 样线表。

也可在命令行运行：

```text
py -3 kml_to_excel.py "C:\path\to\survey.kml"
py -3 excel_to_word.py "C:\path\to\鸟类观测拆分_待复核.xlsx"
```

## 识别范围

- 支持 KML 点标记的 `TimeStamp/when`（UTC，转换为 UTC+8），以及描述文本中的 `时间：YYYY-MM-DD HH:MM:SS`（按原样读取）。Excel 时间为 `hh:mm`。
- 对“物种 数量 生境 行为”的空格写法，按数字边界拆分，一条标记可形成多行；没有原始时间的记录留空。
- 对紧凑写法（如“黄眉柳莺1鸣叫”），可提供已有 Excel 的物种词表：`py -3 kml_to_excel.py "C:\path\to\survey.kml" --reference "C:\path\to\species.xlsx"`。词表只用于识别原文，不自动认可旧表的物种调整。没有词表时，此类标记保留原文并待人工拆分。
- 经纬度直接取自 KML 点坐标。轨迹线、无名称点及起终点不当作鸟类观测。Word 的天气、调查人、调查起止时间及样带长度缺少来源时标为待填写。

**输出不是自动鉴定或正式调查成果。** 物种缩写、错字、复合记录、重复计数及零记录均需研究者核对。带敏感坐标的文件仅应在获授权的本地环境处理。

## 文件

- `kml_to_excel.py`：单个 KML 的便捷入口。
- `split_observations.py`：鸟类标记拆分与 Excel 输出。
- `generate_forms.py`：KML 读取和 Office 文件生成辅助函数。
- `excel_to_word.py`：Excel 到 Word 样线表。
- 两个 `.cmd` 文件：Windows 拖放入口。
