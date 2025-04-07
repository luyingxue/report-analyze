# XLSX文件分析工具

一个简单的命令行工具，用于分析Excel文件并将内容导出为TXT文本文件。

## 功能

- 读取指定的XLSX文件
- 提取文件内的所有工作表及数据
- 生成带时间戳的TXT文件，包含Excel文件的所有内容

## 环境要求

- Python 3.6+
- openpyxl

## 安装

1. 克隆或下载此仓库
2. 进入项目目录
3. 安装依赖库：

```bash
pip install -r requirements.txt
```

## 使用方法

```bash
python src/xlsx_analyzer.py <xlsx文件路径>
```

### 示例

```bash
python src/xlsx_analyzer.py ./reports/sample.xlsx
```

## 输出

程序会在与XLSX文件相同的目录下生成一个TXT文件，文件名格式为：
`原文件名_时间戳.txt`

例如：`sample_20250407_152030.txt`

## 项目结构

```
report-analyze/
├── requirements.txt     # 项目依赖
├── README.md           # 项目说明文档
└── src/                # 源代码目录
    └── xlsx_analyzer.py # 主程序
```

## 扩展功能

如需添加更复杂的分析功能，可在`analyze_xlsx`函数中添加相应的数据处理逻辑。 