# Excel Config Tool

> 📊 配置驱动的 Excel 导入导出工具 - 通过表头匹配定位，数据量驱动，自动扩展空间

[![License](https://img.shields.io/badge/license-MIT-blue.svg)](LICENSE)
[![Java](https://img.shields.io/badge/java-21-orange.svg)](https://openjdk.java.net/)
[![Maven Central](https://img.shields.io/badge/maven-central-com.excelconfig:excel-config-core-green.svg)](https://central.sonatype.com/)

---

## 🎯 核心特性

| 特性 | 说明 |
|------|------|
| **表头匹配** | 通过表头文字匹配定位，不依赖固定单元格位置 |
| **数据驱动** | 提取/填充行数由实际数据决定，不是配置写死 |
| **自动扩展** | 模板空间不足时，自动下移下方内容，不会覆盖或截断 |
| **列隔离** | 每列独立处理，互不干扰 |
| **SAX 流式读取** | 内存优化，支持大文件处理 |
| **简洁 API** | 类似 EasyExcel 的链式调用，开箱即用 |

---

## 🚀 快速上手

### 1. 添加依赖

```xml
<dependency>
    <groupId>io.github.cynosure-tech</groupId>
    <artifactId>excel-config-core</artifactId>
    <version>1.0.1</version>
</dependency>
```

### 2. 配置 JSON

创建 Excel 配置文件（`config.json`）：

```json
{
  "version": "1.0",
  "templateName": "订单管理",
  "extractions": [
    {
      "key": "orderNos",
      "header": { "match": "订单号" },
      "mode": "DOWN",
      "range": { "skipEmpty": true }
    },
    {
      "key": "amounts",
      "header": { "match": "金额" },
      "mode": "DOWN"
    }
  ],
  "exports": [
    {
      "key": "orderNos",
      "header": { "match": "订单号" },
      "mode": "FILL_DOWN"
    }
  ]
}
```

### 3. Java 代码使用

#### 方式一：简洁门面 API（推荐）

```java
import com.excelconfig.ExcelConfigHelper;

// 从 Excel 提取数据
Map<String, Object> data = ExcelConfigHelper.read("template.xlsx")
    .config("config.json")
    .extract();
List<Object> orderNos = (List<Object>) data.get("orderNos");

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

#### 方式二：Service API

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

## 📋 配置说明

### 提取配置 (ExtractConfig)

| 字段 | 类型 | 必填 | 说明 |
|------|------|------|------|
| key | String | 是 | 数据键名（映射到结果 Map 的 key） |
| header | HeaderConfig | 是 | 表头匹配配置 |
| mode | String | 是 | 提取模式 |
| range | RangeConfig | 否 | 范围配置 |
| parser | ParserConfig | 否 | 单元格解析器配置 |

#### HeaderConfig

| 字段 | 类型 | 说明 |
|------|------|------|
| match | String | 表头匹配文字 |
| row | Integer | 表头所在行（从 0 开始，可选） |

#### RangeConfig

| 字段 | 类型 | 说明 |
|------|------|------|
| skipEmpty | Boolean | 是否跳过空行 |
| startRow | Integer | 起始行 |
| endRow | Integer | 结束行 |

### 导出配置 (ExportConfig)

| 字段 | 类型 | 必填 | 说明 |
|------|------|------|------|
| key | String | 是 | 数据键名 |
| header | HeaderConfig | 是 | 表头匹配配置 |
| mode | String | 是 | 填充模式 |
| columns | List<ColumnConfig> | 否 | 列配置（FILL_TABLE 模式需要） |
| headerStyle | StyleConfig | 否 | 表头样式 |
| style | StyleConfig | 否 | 单元格样式 |
| maxRows | Integer | 否 | 最大填充行数 |
| alternateRows | Boolean | 否 | 是否隔行换色 |
| autoWidth | Boolean | 否 | 是否自动列宽 |
| merge | MergeConfig | 否 | 合并单元格配置 |

### 提取模式 (ExtractMode)

**基础模式：**

| 模式 | 说明 | 示例 |
|------|------|------|
| `SINGLE` | 单个单元格 | 提取一个单元格的值 |
| `DOWN` | 向下提取 | 提取一列数据 |
| `RIGHT` | 向右提取 | 提取一行数据 |
| `BLOCK` | 区域矩阵 | 提取一个矩形区域 |
| `UNTIL_EMPTY` | 提取到空行 | 提取直到遇到空行 |

**扩展模式：**

| 模式 | 说明 |
|------|------|
| `KEY_VALUE` | 键值对提取 |
| `TABLE` | 表格提取 |
| `CROSS_TAB` | 交叉表提取 |
| `GROUPED` | 分组提取 |
| `HIERARCHY` | 层级提取 |
| `MERGED_CELLS` | 合并单元格提取 |
| `MULTI_SHEET` | 多工作表提取 |

### 导出模式 (FillMode)

**基础模式：**

| 模式 | 说明 |
|------|------|
| `FILL_CELL` | 填充单个单元格 |
| `FILL_DOWN` | 向下填充 |
| `FILL_RIGHT` | 向右填充 |
| `FILL_BLOCK` | 填充区域 |

**表格模式：**

| 模式 | 说明 |
|------|------|
| `FILL_TABLE` | 填充表格（带表头） |
| `APPEND_ROWS` | 追加行 |
| `APPEND_COLS` | 追加列 |

**高级模式：**

| 模式 | 说明 |
|------|------|
| `REPLACE_AREA` | 替换区域 |
| `FILL_TEMPLATE` | 模板填充 |
| `MULTI_SHEET_FILL` | 多工作表填充 |

---

## 💡 使用场景

### 场景 1：订单导入

```java
// 从 Excel 导入订单数据
Map<String, Object> orders = ExcelConfigHelper.read("orders.xlsx")
    .config("order-import.json")
    .extract();

List<Object> orderNos = (List<Object>) orders.get("orderNos");
List<Object> amounts = (List<Object>) orders.get("amounts");
```

### 场景 2：报表导出

```java
// 填充报表数据
Map<String, Object> reportData = Map.of(
    "dates", Arrays.asList("2024-01", "2024-02", "2024-03"),
    "revenues", Arrays.asList(100000, 150000, 120000),
    "costs", Arrays.asList(60000, 80000, 70000)
);

ExcelConfigHelper.write("report-template.xlsx")
    .config("report-export.json")
    .data(reportData)
    .writeTo("report-2024-q1.xlsx");
```

### 场景 3：模板填充

```java
// 使用复杂表格模板
Map<String, Object> invoiceData = Map.of(
    "invoiceNo", "INV-2024-001",
    "customer", "某某公司",
    "items", Arrays.asList(
        Map.of("name", "商品 A", "qty", 10, "price", 100),
        Map.of("name", "商品 B", "qty", 5, "price", 200)
    ),
    "total", 2000
);

byte[] excelBytes = ExcelConfigHelper.write("invoice-template.xlsx")
    .config("invoice-config.json")
    .data(invoiceData)
    .write();

// 发送给客户端下载
response.setContentType("application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
response.getOutputStream().write(excelBytes);
```

---

## 🏗️ 项目结构

```
excel-config-tool/
├── src/
│   ├── main/java/
│   │   └── com/excelconfig/
│   │       ├── ExcelConfigService.java    # Service API
│   │       ├── ExcelConfigHelper.java     # 简洁门面 API
│   │       ├── ExcelConfigException.java  # 异常类
│   │       ├── model/                     # 配置模型
│   │       │   ├── ExcelConfig.java
│   │       │   ├── ExtractConfig.java
│   │       │   ├── ExportConfig.java
│   │       │   └── ...
│   │       ├── spi/                       # SPI 接口
│   │       │   ├── ExtractMode.java
│   │       │   ├── FillMode.java
│   │       │   └── ...
│   │       ├── extract/                   # 提取引擎
│   │       ├── export/                    # 填充引擎
│   │       ├── locator/                   # 表头定位
│   │       ├── config/                    # JSON 配置解析
│   │       └── sax/                       # SAX 流式读取
│   └── test/java/                         # 单元测试
├── docs/                                  # 设计文档
├── examples/                              # 使用示例
└── pom.xml                                # Maven 配置
```

---

## 📦 技术栈

| 模块 | 技术/版本 |
|------|-----------|
| Java | 21 |
| Excel 处理 | Apache POI 5.2.5 |
| JSON 处理 | Jackson 2.16.1 |
| 日志 | SLF4J 2.0.11 |
| 测试 | JUnit 5 + Mockito |

---

## 🔨 构建与测试

```bash
# 设置 Java 21
export JAVA_HOME=$(/usr/libexec/java_home -v 21)

# 构建并安装到本地 Maven 仓库
mvn clean install

# 运行所有测试
mvn test
```

### 测试结果

```
[INFO] Tests run: 57, Failures: 0, Errors: 0, Skipped: 0
[INFO] BUILD SUCCESS
```

---

## 📚 文档

完整的设计文档位于 [`docs/`](./docs) 文件夹：

| 文档 | 说明 |
|------|------|
| [FINAL_DESIGN.md](./docs/FINAL_DESIGN.md) | 最终设计方案 - 整合所有核心设计 |
| [ARCHITECTURE.md](./docs/ARCHITECTURE.md) | 系统架构设计 |
| [EXTRACT_MODES.md](./docs/EXTRACT_MODES.md) | 提取模式详解 |
| [FILL_MODES.md](./docs/FILL_MODES.md) | 填充模式详解 |
| [HEADER_MATCHING.md](./docs/HEADER_MATCHING.md) | 表头匹配与动态扩展 |
| [COLUMN_ISOLATION.md](./docs/COLUMN_ISOLATION.md) | 列隔离与行偏移 |
| [SAX_READER.md](./docs/SAX_READER.md) | SAX 流式读取 |

使用示例：[`examples/USAGE_EXAMPLES.md`](./examples/USAGE_EXAMPLES.md)

---

## 📚 与类似项目对比

| 项目 | 类型 | 特点 |
|------|------|------|
| [Alibaba EasyExcel](https://github.com/alibaba/easyexcel) | 注解驱动 | 流式读写，大文件处理 |
| [Apache POI](https://poi.apache.org/) | 底层 API | 功能全面，API 复杂 |
| [JXL](https://github.com/jxls-team/jxls) | 模板驱动 | 基于 XML 模板 |
| **Excel Config Tool** | 配置驱动 | JSON 配置，表头自动定位，列隔离，自动扩展 |

### 本项目优势

1. **配置驱动** - 无需注解，JSON 配置更灵活
2. **表头匹配** - 不依赖固定位置，表头文字匹配
3. **自动扩展** - 空间不足时自动下移，不会覆盖
4. **列隔离** - 每列独立处理，互不干扰
5. **简洁 API** - 类似 EasyExcel 的链式调用

---

## 🔧 常见问题

### Q: 如何处理大文件？

A: 核心引擎使用 SAX 流式读取，内存占用低，支持大文件处理。

### Q: 支持哪些 Excel 格式？

A: 支持 `.xlsx` 格式（Excel 2007+），基于 Apache POI。

### Q: 如何自定义单元格解析器？

A: 在配置中指定 `parser` 配置，或实现 `CellParser` 接口。

### Q: 支持 Spring Boot 集成吗？

A: 核心模块是纯 Java，无框架依赖。Spring Boot 集成可由用户自行封装。

---

## 📄 License

MIT License
