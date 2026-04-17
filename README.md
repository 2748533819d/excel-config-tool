# Excel Config Tool

> 📊 配置驱动的 Excel 导入导出工具 - 通过表头匹配定位，数据量驱动，自动扩展空间

---

## 🎯 核心特性

| 特性 | 说明 |
|------|------|
| **表头匹配** | 通过表头文字匹配定位，不依赖固定单元格位置 |
| **数据驱动** | 提取/填充行数由实际数据决定，不是配置写死 |
| **自动扩展** | 模板空间不足时，自动下移下方内容，不会覆盖或截断 |
| **列隔离** | 每列独立处理，互不干扰 |
| **SAX 流式读取** | 内存优化，支持大文件处理 |

---

## 🚀 快速上手

### 1. 添加依赖

只需依赖 `excel-config-spring-boot-starter`：

```xml
<dependency>
    <groupId>com.excelconfig</groupId>
    <artifactId>excel-config-spring-boot-starter</artifactId>
    <version>1.0.0-SNAPSHOT</version>
</dependency>
```

或者只使用核心模块：

```xml
<dependency>
    <groupId>com.excelconfig</groupId>
    <artifactId>excel-config-core</artifactId>
    <version>1.0.0-SNAPSHOT</version>
</dependency>
```

### 2. 配置 JSON

创建 Excel 配置文件：

```json
{
  "version": "1.0",
  "templateName": "订单提取",
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

**方式一：简洁门面 API（推荐）**

```java
// 从 Excel 提取数据
Map<String, Object> data = ExcelConfigHelper.read("template.xlsx")
    .config("config.json")
    .extract();
List<Object> orderNos = (List<Object>) data.get("orderNos");

// 填充数据到 Excel
Map<String, Object> inputData = Map.of(
    "orderNos", Arrays.asList("ORD001", "ORD002", "ORD003")
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

**方式二：Service API**

```java
// 创建服务
ExcelConfigService service = new ExcelConfigService();

// 从 Excel 提取数据
Map<String, Object> data = service.extract(
    new FileInputStream("template.xlsx"),
    configJson
);
List<Object> orderNos = (List<Object>) data.get("orderNos");

// 填充数据到 Excel
Map<String, Object> inputData = Map.of(
    "orderNos", Arrays.asList("ORD001", "ORD002", "ORD003")
);
byte[] result = service.fill(
    new FileInputStream("template.xlsx"),
    inputData,
    configJson
);

// 或输出到 OutputStream
service.fill(templateInputStream, data, configJson, outputStream);
```

### 4. Spring Boot 集成

```java
@Service
public class OrderService {
    
    @Autowired
    private com.excelconfig.starter.ExcelConfigService excelConfigService;
    
    // 导入 Excel
    public List<Object> importOrders(MultipartFile file) throws IOException {
        ExcelConfig config = excelConfigService.loadConfig(
            "classpath:config/order-import.json"
        );
        return excelConfigService.extract(config, file.getInputStream())
            .get("orderNos");
    }
    
    // 导出 Excel
    public byte[] exportOrders() throws IOException {
        ExcelConfig config = excelConfigService.loadConfig(
            "classpath:config/order-export.json"
        );
        Map<String, Object> data = Map.of("orderNos", getOrderList());
        try (InputStream template = getClass()
            .getResourceAsStream("/templates/order-template.xlsx")) {
            return excelConfigService.fill(config, data, template);
        }
    }
}
```

---

## 📋 配置说明

### 提取配置 (ExtractConfig)

| 字段 | 类型 | 必填 | 说明 |
|------|------|------|------|
| key | String | 是 | 数据键名（映射到结果 Map 的 key） |
| header | HeaderConfig | 是 | 表头匹配配置 |
| mode | String | 是 | 提取模式：SINGLE/DOWN/RIGHT/BLOCK/UNTIL_EMPTY/TABLE/KEY_VALUE 等 |
| range | RangeConfig | 否 | 范围配置 |
| parser | ParserConfig | 否 | 单元格解析器配置 |

### 导出配置 (ExportConfig)

| 字段 | 类型 | 必填 | 说明 |
|------|------|------|------|
| key | String | 是 | 数据键名 |
| header | HeaderConfig | 是 | 表头匹配配置 |
| mode | String | 是 | 填充模式：FILL_CELL/FILL_DOWN/FILL_RIGHT/FILL_BLOCK/FILL_TABLE/APPEND_ROWS 等 |
| columns | List<ColumnConfig> | 否 | 列配置（FILL_TABLE 模式需要） |
| headerStyle | StyleConfig | 否 | 表头样式 |
| style | StyleConfig | 否 | 单元格样式 |
| maxRows | Integer | 否 | 最大填充行数 |
| alternateRows | Boolean | 否 | 是否隔行换色 |
| autoWidth | Boolean | 否 | 是否自动列宽 |
| merge | MergeConfig | 否 | 合并单元格配置 |

### 提取模式 (ExtractMode)

**基础模式：**
- `SINGLE` - 单个单元格
- `DOWN` - 向下提取
- `RIGHT` - 向右提取
- `BLOCK` - 区域矩阵
- `UNTIL_EMPTY` - 提取到空行

**扩展模式：**
- `KEY_VALUE` - 键值对
- `TABLE` - 表格
- `CROSS_TAB` - 交叉表
- `GROUPED` - 分组
- `HIERARCHY` - 层级
- `MERGED_CELLS` - 合并单元格
- `MULTI_SHEET` - 多工作表
- `PIVOT` - 透视表

### 导出模式 (FillMode)

**基础模式：**
- `FILL_CELL` - 填充单个单元格
- `FILL_DOWN` - 向下填充
- `FILL_RIGHT` - 向右填充
- `FILL_BLOCK` - 填充区域

**表格模式：**
- `FILL_TABLE` - 填充表格（带表头）
- `APPEND_ROWS` - 追加行
- `APPEND_COLS` - 追加列

**高级模式：**
- `REPLACE_AREA` - 替换区域
- `FILL_TEMPLATE` - 模板填充
- `MULTI_SHEET_FILL` - 多工作表填充

---

## 🏗️ 项目结构

```
excel-config-tool/
├── packages/
│   ├── core/                        # 核心引擎
│   │   ├── src/main/java/
│   │   │   └── com/excelconfig/
│   │   │       ├── ExcelConfigService.java    # 门面服务
│   │   │       ├── model/                     # 配置模型
│   │   │       ├── spi/                       # SPI 接口（ExtractMode, FillMode）
│   │   │       ├── extract/                   # 提取引擎
│   │   │       ├── export/                    # 填充引擎
│   │   │       ├── locator/                   # 表头定位
│   │   │       ├── config/                    # JSON 配置解析
│   │   │       └── sax/                       # SAX 流式读取
│   │   └── src/test/java/                     # 单元测试
│   └── spring-boot-starter/                   # Spring Boot 自动配置
├── docs/                                      # 设计文档
├── examples/                                  # 使用示例
└── pom.xml                                    # 父 POM
```

---

## 📦 技术栈

| 模块 | 技术 |
|------|------|
| Java 版本 | Java 21 |
| Excel 处理 | Apache POI 5.2.5 |
| JSON 处理 | Jackson 2.16.1 |
| Spring Boot | Spring Boot 3.2.1 |
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

# 只运行 core 模块测试
cd packages/core && mvn test
```

### 测试结果

```
[INFO] Tests run: 51, Failures: 0, Errors: 0, Skipped: 0 - Core 模块
[INFO] Tests run: 3, Failures: 0, Errors: 0, Skipped: 0 - Spring Boot Starter 模块
[INFO] BUILD SUCCESS
```

---

## 📚 文档

完整的设计文档位于 [`docs/`](./docs) 文件夹：

| 文档 | 说明 |
|------|------|
| [FINAL_DESIGN.md](./docs/FINAL_DESIGN.md) | 最终设计方案 - 整合所有核心设计 |
| [ARCHITECTURE.md](./docs/ARCHITECTURE.md) | 系统架构设计 |
| [EXTRACT_MODES.md](./docs/EXTRACT_MODES.md) | 5 种基础提取模式详解 |
| [EXTENDED_MODES.md](./docs/EXTENDED_MODES.md) | 10 种扩展提取模式详解 |
| [FILL_MODES.md](./docs/FILL_MODES.md) | 10 种导出/填充模式详解 |
| [HEADER_MATCHING.md](./docs/HEADER_MATCHING.md) | 表头匹配与动态扩展 |
| [COLUMN_ISOLATION.md](./docs/COLUMN_ISOLATION.md) | 列隔离与行偏移 |
| [SAX_READER.md](./docs/SAX_READER.md) | SAX 流式读取 |

使用示例：[`examples/USAGE_EXAMPLES.md`](./examples/USAGE_EXAMPLES.md)

---

## 📚 与类似项目对比

| 项目 | 类型 | Stars | 差异化 |
|------|------|-------|--------|
| [Alibaba EasyExcel](https://github.com/alibaba/easyexcel) | 注解驱动 | 33,750+ | 流式读写，大文件处理 |
| **本项目** | 配置驱动 | - | JSON 配置，表头自动定位，列隔离，自动扩展 |

---

## 📄 License

MIT License
