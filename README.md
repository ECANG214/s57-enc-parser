# S-57 ENC Parser — 电子海图（ENC）二进制解析器

基于 Python 实现 **ISO 8211** 二进制解析，直接读取 **S-57 电子海图（ENC）** 文件，逐层解析记录结构、按 FOID 定位目标物标、关联空间矢量坐标，并自动生成解析报告。

![Python](https://img.shields.io/badge/Python-3.7%2B-blue)
![S-57](https://img.shields.io/badge/IHO-S--57%20Ed.3.1-0b6b78)
![ISO-8211](https://img.shields.io/badge/ISO-8211-1994-lightgrey)

## 项目定位

不依赖任何第三方海图库（如 GDAL 的 S-57 驱动），根据 **ISO 8211:1994** 规范自行实现二进制记录解析逻辑，完成对 S-57 ENC 文件（`.000`）的结构解读，包括物标属性提取、空间矢量记录关联与坐标解码。

> 说明：本项目为课程小组协作项目（见文末「项目分工」），解析逻辑由本人负责实现。

## 解析流程

```
.000 二进制文件
   │  从 offset 0 起切 24 字节 Leader，读记录长度与 Base Address
   ▼
ISO 8211 记录流
   ├─ 记录 #1  DDR         解析目录表，构建父-子字段树
   ├─ 记录 #2  首条数据记录   DSID（生产机构/版本/日期）+ DSSI（各类记录计数）
   ├─ 后续记录              定位 DSPM（坐标乘数 COMF），并按 FOID 匹配目标物标
   │                        命中后解析 FRID（物标类型）/ ATTF / NATF / FSPT
   └─ 空间记录扫描          匹配 VRID → 解码 SG2D/SG3D 经纬度（3/4 字节自动识别，÷COMF）
   ▼
s57_parse_report.docx（未装 python-docx 时降级为 .txt）
```

## 解析能力

对目标物标及其关联的空间记录做深度解析，覆盖以下结构：

| 阶段 | 字段 | 说明 |
|------|------|------|
| **Leader / DDR** | 24 字节记录头 | 记录长度、交换级别、Entry Map（字段长度/位置/标签尺寸）|
| **DDR Directory** | 目录表 | 各字段的 Tag / Length / Position |
| **DSID** | 数据集标识 | 生产机构 AGEN、版本、更新日期 |
| **DSSI** | 数据结构信息 | 各类记录数量统计（NOMR/NOCR/NOGR/NOLR…）|
| **DSPM** | 数据集参数 | 坐标乘数 COMF、比例尺 CSCL、坐标单位 |
| **FRID / FOID** | 物标记录标识 | 按 FOID（AGEN+FIDN+FIDS）定位目标物标 |
| **ATTF / NATF** | 物标属性 | ASCII / UTF-16LE 属性值（含中文名）|
| **FSPT** | 物标→空间指针 | 关联空间矢量记录 |
| **VRID** | 矢量记录标识 | 定位关联的空间记录 |
| **SG2D / SG3D** | 坐标字段 | 解析 YCOO/XCOO 经纬度（自动识别 3/4 字节、COMF 缩放）；SG3D 的 VE3D 高程字段未解码 |

> 说明：主流程对**目标 FOID** 的物标深度解析（属性、空间指针），其余记录在空间扫描阶段按需遍历，并非对全文件所有物标的完整属性解析。

## 内置 S-57 知识库

- **物标类（OBJL）**：160 个键（编号 1–160，含少量重复名称，唯一名称 158 个），覆盖 Administration Area、Anchor Berth、Wreck 等常用物标
- **属性字典（ATTR_NAMES）**：81 个属性键（OBJNAM、CATWRK、VALSOU、QUASOU 等）
- **枚举值映射（ATTVAL）**：17 条枚举值（CATWRK 沉船类型、WATLEV 水深等级等）

## 快速开始

```bash
cd s57-enc-parser
pip install -r requirements.txt   # python-docx 可选，用于生成 Word 报告

python s57_parser.py path/to/CN335002.000
# 或交互式：python s57_parser.py 后手动输入 .000 文件路径
```

程序会：

1. 控制台逐层打印解析结果（DDR → DSID/DSSI → DSPM → 目标 FOID → 关联矢量坐标）
2. 在脚本所在目录生成 `s57_parse_report.docx` 结构化 Word 报告（无 python-docx 时降级为 `.txt`）

## 目标物标

示例以 `CN 0013753099 22501`（AGEN=70 中国，FIDN=13753099，FIDS=22501）为目标 FOID，解析其物标类型、属性、中文名及关联的空间坐标。

> 解析目标可在源码顶部 `TARGET_AGENCY / TARGET_FIDN / TARGET_FIDS` 常量中修改。

## 数据说明

ENC 数据文件（`.000`）**未随代码提供**（真实海图数据可能受分发协议限制）。测试数据可自行获取：

- NOAA 免费 ENC（美国水域，公开下载）
- 课程 / 单位提供的中国海图 ENC 文件

代码支持任意标准 ISO 8211 格式的 S-57 ENC 文件。

## 技术要点

- **ISO 8211 核心**：实现 Leader（24B）、Directory、Entry Map 解析，正确处理 field length/position/tag size
- **小端序二进制**：`int.from_bytes(..., 'little')` 解析 RCNM/RCID/OBJL/AGEN 等字段
- **多字节坐标**：SG2D/SG3D 自动识别 3 字节 vs 4 字节坐标，按 COMF 缩放
- **字节序处理**：FSPT 指针中的 RCID 按大端读出后重排为小端，再与 VRID 匹配
- **中文属性解码**：NATF 的 UTF-16LE 编码 + 对齐字节处理
- **报告生成**：python-docx 输出结构化 Word 报告，无依赖时降级 txt

## 项目分工

小组协作项目。本人负责**核心解析逻辑的实现**（ISO 8211 二进制解析、物标提取、空间记录关联、坐标解码），团队成员负责解析结果的整理与报告格式排版。

## 目录结构

```
s57-enc-parser/
├── s57_parser.py      # 主程序（ISO 8211 解析 + S-57 知识库 + 报告生成）
├── requirements.txt
├── LICENSE
└── README.md
```

## 效果截图

<!-- TODO: 运行 `python s57_parser.py 某海图.000` 后，截图以下输出粘贴至此：
     1. 控制台逐层解析结果（Leader → DDR → DSID/DSSI → DSPM）
     2. 目标物标属性 + 中文名解析结果
     3. 关联的空间矢量坐标输出
     4. 生成的 Word / txt 报告内容
-->

## License

MIT
