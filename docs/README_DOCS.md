# Excel Config Tool - 设计文档索引

> 📚 本项目的完整设计文档导航

---

## 📖 核心设计文档

| 文档 | 说明 | 阅读优先级 |
|------|------|------------|
| [FINAL_DESIGN.md](./FINAL_DESIGN.md) | **最终设计方案** - 整合所有核心设计 | ⭐⭐⭐ |
| [ARCHITECTURE.md](./ARCHITECTURE.md) | 系统架构设计 - 前后端模块划分 | ⭐⭐⭐ |
| [FRONTEND_DESIGN.md](./FRONTEND_DESIGN.md) | 前端组件设计 - Univer + Vue3 | ⭐⭐ |

---

## 📋 模式详解文档

### 提取模式（Import）

| 文档 | 说明 | 模式数量 |
|------|------|----------|
| [EXTRACT_MODES.md](./EXTRACT_MODES.md) | 基础提取模式详解 | 5 种 |
| [EXTENDED_MODES.md](./EXTENDED_MODES.md) | 扩展提取模式详解 | 10 种 |

**基础模式**：
- `SINGLE` - 提取单个单元格
- `DOWN` - 向下提取列数据
- `RIGHT` - 向右提取行数据
- `BLOCK` - 提取区域矩阵
- `UNTIL_EMPTY` - 提取到空行停止

**扩展模式**：
- `KEY_VALUE` - A 列 key，B 列 value
- `TABLE` - 表头 + 数据行
- `CROSS_TAB` - 交叉统计表
- `GROUPED` - 分组数据
- `HIERARCHY` - 层级树形结构
- `MERGED_CELLS` - 合并单元格处理
- `MULTI_SHEET` - 多工作表
- `PIVOT` - 透视表
- `FORMULA` - 公式计算
- `CONDITIONAL` - 条件过滤

### 导出模式（Export）

| 文档 | 说明 | 模式数量 |
|------|------|----------|
| [FILL_MODES.md](./FILL_MODES.md) | 导出/填充模式详解 | 10 种 |

**模式列表**：
- `FILL_CELL` - 填充单个单元格
- `FILL_DOWN` - 向下填充列数据
- `FILL_RIGHT` - 向右填充行数据
- `FILL_BLOCK` - 填充区域矩阵
- `FILL_TABLE` - 填充表格（带表头）
- `APPEND_ROWS` - 追加行
- `APPEND_COLS` - 追加列
- `REPLACE_AREA` - 替换区域
- `FILL_TEMPLATE` - 模板填充（占位符）
- `MULTI_SHEET_FILL` - 多工作表填充

---

## 🔧 核心机制文档

| 文档 | 说明 | 解决的核心问题 |
|------|------|----------------|
| [HEADER_MATCHING.md](./HEADER_MATCHING.md) | 表头匹配与动态扩展 | 表头位置不固定，数据量超过模板空间 |
| [COLUMN_ISOLATION.md](./COLUMN_ISOLATION.md) | 列隔离与行偏移 | 多列同时填充时互不干扰 |
| [DYNAMIC_ROW_COUNT.md](./DYNAMIC_ROW_COUNT.md) | 动态行数确定机制 | 配置时无法预知数据量 |

---

## 🔍 边界检测方案（备选）

| 文档 | 方案类型 | 状态 |
|------|----------|------|
| [CONFIG_DRIVEN_BOUNDARY.md](./CONFIG_DRIVEN_BOUNDARY.md) | 配置驱动边界 | ✅ 已采用 |
| [BOUNDARY_DETECTION.md](./BOUNDARY_DETECTION.md) | 自动边界检测 | ⚠️ 备选方案 |

---

## 📊 调研分析文档

| 文档 | 说明 |
|------|------|
| [ANALYSIS.md](./ANALYSIS.md) | 初始调研 - Java Excel 工具包分析 |
| [GRAPECITY_ANALYSIS.md](./GRAPECITY_ANALYSIS.md) | GrapeCity GcExcel + SpreadJS 分析 |

---

## 🗂️ 文档结构

```
excel-config-tool/docs/
│
├── 📘 核心设计
│   ├── FINAL_DESIGN.md          # 最终设计方案（入口）
│   ├── ARCHITECTURE.md          # 系统架构
│   └── FRONTEND_DESIGN.md       # 前端设计
│
├── 📋 模式详解
│   ├── EXTRACT_MODES.md         # 提取模式（5 基础 +10 扩展）
│   ├── EXTENDED_MODES.md        # 扩展模式详解
│   └── FILL_MODES.md            # 导出模式（10 种）
│
├── 🔧 核心机制
│   ├── HEADER_MATCHING.md       # 表头匹配定位
│   ├── COLUMN_ISOLATION.md      # 列隔离机制
│   └── DYNAMIC_ROW_COUNT.md     # 动态行数机制
│
├── 🔍 边界检测
│   ├── CONFIG_DRIVEN_BOUNDARY.md # 配置驱动（已采用）
│   └── BOUNDARY_DETECTION.md     # 自动检测（备选）
│
└── 📊 调研分析
    ├── ANALYSIS.md              # 初始调研
    └── GRAPECITY_ANALYSIS.md    # GrapeCity 分析
```

---

## 🚀 快速开始

### 新阅读者建议路径

```
1. FINAL_DESIGN.md          ← 先读这个，了解整体设计
   ↓
2. EXTRACT_MODES.md         ← 了解提取模式
   ↓
3. FILL_MODES.md            ← 了解导出模式
   ↓
4. HEADER_MATCHING.md       ← 理解核心机制
   ↓
5. ARCHITECTURE.md          ← 了解技术架构
```

### 开发者建议路径

```
1. FINAL_DESIGN.md          ← 了解设计目标
   ↓
2. ARCHITECTURE.md          ← 了解技术架构
   ↓
3. EXTRACT_MODES.md         ← 实现提取逻辑
   ↓
4. FILL_MODES.md            ← 实现导出逻辑
   ↓
5. 开始编码实现
```

---

## 📝 版本历史

| 版本 | 日期 | 变更说明 |
|------|------|----------|
| 1.0 | 2024-04-16 | 初始设计完成 |

---

## 🔗 相关链接

- [GitHub 仓库](https://github.com/xxx/excel-config-tool)（待创建）
- [Apache POI 官方文档](https://poi.apache.org/)
- [Univer 官方文档](https://univer.ai/)
- [Vue 3 官方文档](https://vuejs.org/)
