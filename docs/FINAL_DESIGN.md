# Excel Config Tool - 最终设计方案

> **版本**: 1.0  
> **最后更新**: 2026-04-18  
> **状态**: 已实现

---

## 一、核心设计理念

### 1.1 设计原则

| 原则 | 说明 |
|------|------|
| **配置驱动** | 边界由配置决定，不是自动检测 |
| **表头匹配** | 通过表头文字匹配定位，不依赖固定位置 |
| **数据驱动** | 行数由实际数据决定，不是配置写死 |
| **自动扩展** | 模板空间不足时，自动下移下方内容 |
| **列隔离** | 每列独立处理，互不干扰 |
| **零侵入** | 用户只需关注配置，无需关心实现细节 |

### 1.2 核心价值

```
┌─────────────────────────────────────────────────────────────────┐
│                    Excel Config Tool 核心价值                    │
├─────────────────────────────────────────────────────────────────┤
│                                                                 │
│  用户痛点                          解决方案                      │
│  ───────────────────────────────────────────────────────────    │
│  ❌ 硬编码处理 Excel 数据           ✅ 配置驱动，零代码            │
│  ❌ 表头位置变化导致代码失效        ✅ 表头匹配，自动定位          │
│  ❌ 数据量变化导致覆盖/截断         ✅ 动态扩展，自动调整          │
│  ❌ 多表共存时处理复杂              ✅ 列隔离 + 边界保护           │
│  ❌ 内存溢出处理大文件              ✅ SAX 流式读取                │
│                                                                 │
└─────────────────────────────────────────────────────────────────┘
```

---

## 二、配置模型（JSON 格式）

### 2.1 配置结构

```json
{
  "version": "1.0",
  "templateName": "配置名称",
  
  "extractions": [
    {
      "key": "字段名",
      "header": {
        "match": "表头文字"
      },
      "mode": "DOWN",
      "range": {
        "skipEmpty": true,
        "maxRows": 1000
      },
      "parser": {
        "type": "string"
      }
    }
  ],
  
  "exports": [
    {
      "key": "字段名",
      "header": {
        "match": "表头文字"
      },
      "mode": "FILL_DOWN"
    }
  ]
}
```

### 2.2 配置定位方式

| 方式 | 配置 | 适用场景 |
|------|------|------|
| **表头匹配**（推荐） | `"header": { "match": "订单号" }` | 表头位置不固定 |
| **固定位置** | `"position": { "cellRef": "A1" }` | 格式固定的模板 |

---

## 三、提取模式（Import）

### 3.1 基础模式（5 种）

| 模式 | 说明 | 配置示例 |
|------|------|----------|
| **SINGLE** | 提取单个单元格 | `position: A1, mode: SINGLE` |
| **DOWN** | 向下提取列数据 | `position: A2, mode: DOWN` |
| **RIGHT** | 向右提取行数据 | `position: A1, mode: RIGHT` |
| **BLOCK** | 提取区域矩阵 | `position: A2:D10, mode: BLOCK` |
| **UNTIL_EMPTY** | 提取到空行停止 | `position: A2, mode: UNTIL_EMPTY` |

### 3.2 扩展模式（8 种）

| 模式 | 说明 | 典型场景 |
|------|------|----------|
| **KEY_VALUE** | A 列 key，B 列 value | 配置表、字典 |
| **TABLE** | 表头 + 数据行 | 标准数据表 |
| **CROSS_TAB** | 行头 + 列头 + 数据 | 交叉统计表 |
| **GROUPED** | 分组数据 | 分类汇总 |
| **HIERARCHY** | 层级数据 | 树形结构 |
| **MERGED_CELLS** | 合并单元格处理 | 复杂报表 |
| **MULTI_SHEET** | 多工作表 | 分月报表 |
| **PIVOT** | 透视表 | 数据汇总 |

### 3.3 行数确定机制

```json
{
  "extractions": [
    {
      "key": "orderNos",
      "header": { "match": "订单号" },
      "mode": "DOWN",
      "range": {
        "skipEmpty": true,
        "maxRows": 1000
      }
    },
    {
      "key": "monthHeaders",
      "header": { "match": "月份" },
      "mode": "RIGHT",
      "range": {
        "cols": 12
      }
    }
  ]
}
```

### 3.4 边界检测

```
边界条件（优先级从高到低）：
1. 达到 maxRows 限制 → 停止
2. 遇到已知配置点 → 停止
3. 遇到空行 (skipEmpty) → 停止
4. 到达 sheet 末尾 → 停止
```

---

## 四、导出模式（Export）

### 4.1 基础模式（4 种）

| 模式 | 说明 | 配置示例 |
|------|------|----------|
| **FILL_CELL** | 填充单个单元格 | `"position": "A1", "mode": "FILL_CELL"` |
| **FILL_DOWN** | 向下填充列数据 | `"position": "A2", "mode": "FILL_DOWN"` |
| **FILL_RIGHT** | 向右填充行数据 | `"position": "A1", "mode": "FILL_RIGHT"` |
| **FILL_BLOCK** | 填充区域矩阵 | `"position": "A2:D10", "mode": "FILL_BLOCK"` |

### 4.2 表格模式（3 种）

| 模式 | 说明 | 配置示例 |
|------|------|----------|
| **FILL_TABLE** | 填充表格（带表头） | `"mode": "FILL_TABLE", "columns": [...]` |
| **APPEND_ROWS** | 追加行 | `"mode": "APPEND_ROWS"` |
| **APPEND_COLS** | 追加列 | `"mode": "APPEND_COLS"` |

### 4.3 高级模式（3 种）

| 模式 | 说明 | 配置示例 |
|------|------|----------|
| **REPLACE_AREA** | 替换区域 | `"mode": "REPLACE_AREA"` |
| **FILL_TEMPLATE** | 模板填充（占位符） | `"mode": "FILL_TEMPLATE"` |
| **MULTI_SHEET_FILL** | 多工作表填充 | `"mode": "MULTI_SHEET_FILL"` |

### 4.4 动态扩展机制

```
导出流程：
1. 定位表头（通过 header.match 匹配）
2. 检查下方是否有其他配置点
3. 计算：需要行数 vs 可用行数
4. 如果需要 > 可用：下移下方内容
5. 填充数据
```

```json
{
  "exports": [
    {
      "key": "orderNos",
      "header": { "match": "订单号" },
      "mode": "FILL_DOWN"
    }
  ]
}
```

---

## 五、核心机制详解

### 5.1 表头定位器

```java
public class HeaderLocator {
    
    public Position locate(Sheet sheet, HeaderConfig config) {
        // 1. 确定搜索范围
        int startRow = config.getInRows() != null ? config.getInRows()[0] : 1;
        int endRow = config.getInRows() != null ? config.getInRows()[1] : sheet.getLastRowNum();
        
        // 2. 在范围内搜索
        for (int rowNum = startRow; rowNum <= endRow; rowNum++) {
            Row row = sheet.getRow(rowNum);
            if (row == null) continue;
            
            for (Cell cell : row) {
                String value = getCellValueAsString(cell);
                if (matches(value, config)) {
                    return new Position(rowNum, cell.getColumnIndex());
                }
            }
        }
        
        throw new HeaderNotFoundException("未找到表头：" + config.getMatch());
    }
}
```

### 5.2 列隔离与行偏移

```
场景：
A 列填充 5 行数据，B 列填充 3 行数据

处理：
- A 列：A2-A6 填充 5 行，原 A7 数据下移到 A7
- B 列：B2-B4 填充 3 行，原 B5 数据下移到 B5
- 每列独立计算，互不影响
```

### 5.3 配置点管理

```java
public class ConfigPointManager {
    
    // 收集所有配置点
    public List<ConfigPoint> collect(ExcelConfig config) {
        List<ConfigPoint> points = new ArrayList<>();
        
        for (ExtractConfig extract : config.getExtractions()) {
            Position pos = resolvePosition(extract);
            points.add(new ConfigPoint(pos, extract, ConfigType.EXTRACT));
        }
        
        for (ExportConfig export : config.getExports()) {
            Position pos = resolvePosition(export);
            points.add(new ConfigPoint(pos, export, ConfigType.EXPORT));
        }
        
        // 按列分组，每列内按行排序
        return sortByColumnAndRow(points);
    }
    
    // 找到同列的下一个配置点
    public ConfigPoint findNextInColumn(int column, int currentRow) {
        return points.stream()
            .filter(p -> p.getPosition().getColumn() == column)
            .filter(p -> p.getPosition().getRow() > currentRow)
            .min(Comparator.comparingInt(p -> p.getPosition().getRow()))
            .orElse(null);
    }
}
```

---

## 六、架构设计

### 6.1 模块结构

```
excel-config-tool/
├── packages/
│   └── core/                        # 核心引擎（已实现）
│       ├── src/main/java/
│       │   └── com/excelconfig/
│       │       ├── ExcelConfigHelper.java       # 门面 API（推荐）
│       │       ├── ExcelConfigService.java      # Service API
│       │       ├── ExcelConfigException.java    # 异常类
│       │       ├── model/                       # 配置模型
│       │       │   ├── ExcelConfig.java
│       │       │   ├── ExtractConfig.java
│       │       │   ├── ExportConfig.java
│       │       │   └── ...
│       │       ├── spi/                         # SPI 接口
│       │       │   ├── ExtractMode.java
│       │       │   ├── FillMode.java
│       │       │   └── ...
│       │       ├── extract/                     # 提取引擎
│       │       ├── export/                      # 填充引擎
│       │       ├── locator/                     # 表头定位
│       │       ├── config/                      # JSON 配置解析
│       │       └── sax/                         # SAX 流式读取
│       └── src/test/java/                       # 单元测试
│
├── docs/                                        # 文档
├── examples/                                    # 使用示例
└── pom.xml                                      # 父 POM
```

### 6.2 核心接口

```java
// 提取策略接口
public interface ExtractStrategy {
    List<Object> extract(Sheet sheet, ExtractContext context);
    ExtractMode getSupportedMode();
}

// 数据解析器接口
public interface CellParser {
    Object parse(Cell cell, ParserConfig config);
}

// 导出策略接口
public interface FillStrategy {
    void fill(Workbook workbook, FillContext context);
    FillMode getSupportedMode();
}

// 表头定位器接口
public interface HeaderLocator {
    Position locate(Sheet sheet, HeaderConfig config);
}
```

### 6.3 数据流

```
导入流程：
Excel 文件 → SAX 流式读取 → 表头定位 → 数据提取 → 解析器 → 处理器 → Map<String, Object>

导出流程：
Map<String, Object> → 表头定位 → 空间检查 → 动态扩展 → 数据填充 → Excel 文件
```

---

## 七、配置示例大全（JSON 格式）

### 7.1 基础导入配置

```json
{
  "version": "1.0",
  "templateName": "订单导入",
  "extractions": [
    {
      "key": "orderNos",
      "header": { "match": "订单号" },
      "mode": "DOWN",
      "range": { "skipEmpty": true },
      "parser": { "type": "string" }
    },
    {
      "key": "amounts",
      "header": { "match": "金额" },
      "mode": "DOWN",
      "range": { "skipEmpty": true },
      "parser": {
        "type": "number",
        "format": "#,##0.00"
      }
    },
    {
      "key": "dates",
      "header": { "match": "日期" },
      "mode": "DOWN",
      "range": { "skipEmpty": true },
      "parser": {
        "type": "date",
        "format": "yyyy-MM-dd"
      }
    }
  ]
}
```

### 7.2 基础导出配置

```json
{
  "version": "1.0",
  "templateName": "订单导出",
  "exports": [
    {
      "key": "orderNos",
      "header": { "match": "订单号" },
      "mode": "FILL_DOWN"
    },
    {
      "key": "amounts",
      "header": { "match": "金额" },
      "mode": "FILL_DOWN",
      "style": {
        "format": "#,##0.00",
        "horizontalAlign": "RIGHT"
      }
    }
  ]
}
```

### 7.3 表格导出配置

```json
{
  "version": "1.0",
  "exports": [
    {
      "key": "orders",
      "header": { "match": "订单号" },
      "mode": "FILL_TABLE",
      "columns": [
        {
          "key": "orderNo",
          "header": "订单号",
          "width": 15
        },
        {
          "key": "amount",
          "header": "金额",
          "width": 12,
          "format": "#,##0.00"
        },
        {
          "key": "orderDate",
          "header": "日期",
          "width": 12,
          "format": "yyyy-MM-dd"
        }
      ],
      "headerStyle": {
        "bold": true,
        "background": "#4472C4",
        "fontColor": "#FFFFFF"
      },
      "alternateRows": true,
      "autoWidth": true
    }
  ]
}
```

### 7.4 多表配置

```json
{
  "version": "1.0",
  "templateName": "多表处理",
  "extractions": [
    {
      "key": "orderTable",
      "header": {
        "match": "订单号",
        "inRows": [1, 10]
      },
      "mode": "DOWN"
    },
    {
      "key": "customerTable",
      "header": {
        "match": "客户",
        "inRows": [15, 25]
      },
      "mode": "DOWN"
    }
  ]
}
```

### 7.5 模板填充配置

```json
{
  "version": "1.0",
  "exports": [
    {
      "key": "orderInfo",
      "mode": "FILL_TEMPLATE",
      "placeholder": {
        "prefix": "{{",
        "suffix": "}}"
      },
      "fields": [
        {
          "key": "customerName",
          "type": "string"
        },
        {
          "key": "orderNo",
          "type": "string"
        },
        {
          "key": "amount",
          "type": "number",
          "format": "#,##0.00"
        }
      ]
    }
  ]
}
```

---

## 八、技术栈

| 组件 | 技术 | 说明 |
|------|------|------|
| Java | 21 | 现代化 Java 栈 |
| Excel 处理 | Apache POI 5.2.5 | 行业标准 |
| JSON 处理 | Jackson 2.16.1 | 高性能 JSON 解析 |
| 日志 | SLF4J 2.0.11 + Logback | 灵活日志 |
| 测试 | JUnit 5 + Mockito | 单元测试 |

### 内存优化

```
SAX 流式读取（大文件）：
- 不一次性加载整个 Excel
- 逐行读取，处理，丢弃
- 内存占用：O(1) 而非 O(n)

适用场景：
- 文件 > 100MB
- 行数 > 10 万行
```

---

## 九、API 使用方式

### 9.1 门面 API（推荐）

```java
import com.excelconfig.ExcelConfigHelper;

// 从 Excel 提取数据
Map<String, Object> data = ExcelConfigHelper.read("template.xlsx")
    .config("config.json")
    .extract();

// 填充数据到 Excel
Map<String, Object> inputData = Map.of(
    "orderNos", Arrays.asList("ORD001", "ORD002", "ORD003"),
    "amounts", Arrays.asList(100.0, 200.0, 150.0)
);
ExcelConfigHelper.write("template.xlsx")
    .config("config.json")
    .data(inputData)
    .writeTo("output.xlsx");

// 或输出到 OutputStream
ExcelConfigHelper.write(templateInputStream)
    .configJson(configJson)
    .data(inputData)
    .writeTo(outputStream);
```

### 9.2 Service API

```java
import com.excelconfig.ExcelConfigService;

ExcelConfigService service = new ExcelConfigService();

// 从 Excel 提取数据
Map<String, Object> data = service.extract(
    new FileInputStream("template.xlsx"),
    configJson
);

// 填充数据到 Excel
byte[] result = service.fill(
    new FileInputStream("template.xlsx"),
    inputData,
    configJson
);
```

---

## 十、测试策略

### 10.1 单元测试

```java
// 表头定位器测试
@Test
public void testLocateHeader_Success() {
    // 准备：Excel 中有"订单号"表头
    // 执行：locator.locate(sheet, config)
    // 验证：返回正确位置 A1
}

// 提取引擎测试
@Test
public void testExtract_DownMode() {
    // 准备：Excel 中有 10 行数据
    // 执行：engine.extract(input, config)
    // 验证：返回 10 条数据
}
```

### 10.2 集成测试

```java
// 完整导入流程测试
@Test
public void testFullImport() {
    // 准备：JSON 配置 + Excel 文件
    // 执行：导入整个流程
    // 验证：生成正确的 Map<String, Object>
}

// 完整导出流程测试
@Test
public void testFullExport() {
    // 准备：JSON 配置 + 数据 + 模板
    // 执行：导出整个流程
    // 验证：生成正确的 Excel 文件
}
```

### 10.3 边界测试

```java
// 数据量超过模板空间
@Test
public void testExport_DataExceedsTemplate() {
    // 准备：模板预留 5 行，数据有 20 条
    // 执行：导出
    // 验证：下方内容自动下移，数据完整填充
}

// 表头不存在
@Test
public void testLocateHeader_NotFound() {
    // 准备：Excel 中没有指定表头
    // 执行：locator.locate()
    // 验证：抛出 HeaderNotFoundException
}
```

---

## 十一、文档索引

| 文档 | 说明 |
|------|------|
| [EXTRACT_MODES.md](./EXTRACT_MODES.md) | 提取模式详解 |
| [FILL_MODES.md](./FILL_MODES.md) | 填充模式详解 |
| [HEADER_MATCHING.md](./HEADER_MATCHING.md) | 表头匹配与动态扩展机制 |
| [COLUMN_ISOLATION.md](./COLUMN_ISOLATION.md) | 列隔离与行偏移机制 |
| [DYNAMIC_ROW_COUNT.md](./DYNAMIC_ROW_COUNT.md) | 动态行数确定机制 |
| [SAX_READER.md](./SAX_READER.md) | SAX 流式读取 |
| [ARCHITECTURE.md](./ARCHITECTURE.md) | 系统架构设计 |

---

## 十二、总结

### 核心特性

1. **表头匹配定位** - 配置通过表头文字匹配，不依赖固定位置
2. **数据量驱动** - 提取/填充行数由实际数据决定
3. **自动扩展** - 模板空间不足时自动下移下方内容
4. **列隔离** - 每列独立处理，互不干扰
5. **配置驱动边界** - 配置点即边界，自动检测
6. **SAX 流式读取** - 内存优化，支持大文件
7. **简洁 API** - ExcelConfigHelper 门面类，类似 EasyExcel

### 配置示例（一分钟上手）

```json
// 导入：3 行配置搞定
{
  "extractions": [
    {
      "key": "orderNos",
      "header": { "match": "订单号" },
      "mode": "DOWN"
    }
  ]
}

// 导出：3 行配置搞定
{
  "exports": [
    {
      "key": "orderNos",
      "header": { "match": "订单号" },
      "mode": "FILL_DOWN"
    }
  ]
}
```

### Maven 依赖

```xml
<dependency>
    <groupId>com.excelconfig</groupId>
    <artifactId>excel-config-core</artifactId>
    <version>1.0.0</version>
</dependency>
```
