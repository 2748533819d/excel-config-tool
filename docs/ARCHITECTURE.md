# Excel Config Tool - 系统架构设计

> **版本**: 1.0  
> **最后更新**: 2026-04-18  
> **状态**: 已实现

---

## 一、项目整体定位

```
┌─────────────────────────────────────────────────────────────────────────────────┐
│                           Excel Config Tool                                     │
│                                                                                 │
│     一个配置化的 Excel 数据提取/填充工具，通过 JSON 配置驱动，支持表头自动定位       │
├─────────────────────────────────────────────────────────────────────────────────┤
│                                                                                 │
│   ┌─────────────────────────────────────────────────────────────────────────┐   │
│   │                    @excel-config/core v1.0.0                            │   │
│   │                    核心引擎库 (Maven Central)                           │   │
│   │                                                                         │   │
│   │   核心组件：                                                             │   │
│   │   • ExcelConfigHelper  - 门面 API（推荐，类似 EasyExcel）                │   │
│   │   • ExcelConfigService - Service API                                    │   │
│   │   • ExtractEngine      - 提取引擎（SAX 流式读取）                         │   │
│   │   • FillEngine         - 填充引擎（动态扩展）                            │   │
│   │   • HeaderLocator      - 表头定位器                                      │   │
│   │   • JsonConfigParser   - JSON 配置解析器                                 │   │
│   │                                                                         │   │
│   │   依赖：Apache POI 5.2.5, Jackson 2.16.1                                │   │
│   └─────────────────────────────────────────────────────────────────────────┘   │
│                                                                                 │
│   使用方式：                                                                     │
│   • 直接使用 - 作为独立库集成到项目中                                           │
│   • Spring Boot - 通过 @Configuration 自行封装（无 starter）                     │
│                                                                                 │
└─────────────────────────────────────────────────────────────────────────────────┘
```

---

## 二、模块结构

```
excel-config-tool/
│
├── packages/
│   │
│   └── core/                        # 核心引擎包（已实现）
│       ├── pom.xml                  # artifactId: excel-config-core
│       └── src/
│           ├── main/java/
│           │   └── com/excelconfig/
│           │       ├── ExcelConfigHelper.java       # 门面 API（推荐）
│           │       ├── ExcelConfigService.java      # Service API
│           │       ├── ExcelConfigException.java    # 统一异常类
│           │       │
│           │       ├── model/                       # 配置模型
│           │       │   ├── ExcelConfig.java         # 根配置
│           │       │   ├── ExtractConfig.java       # 提取配置
│           │       │   ├── ExportConfig.java        # 导出配置
│           │       │   ├── HeaderConfig.java        # 表头配置
│           │       │   ├── RangeConfig.java         # 范围配置
│           │       │   ├── ParserConfig.java        # 解析器配置
│           │       │   ├── StyleConfig.java         # 样式配置
│           │       │   └── ColumnConfig.java        # 列配置
│           │       │
│           │       ├── spi/                         # SPI 接口
│           │       │   ├── ExtractMode.java         # 提取模式枚举
│           │       │   ├── FillMode.java            # 填充模式枚举
│           │       │   ├── ExtractStrategy.java     # 提取策略接口
│           │       │   ├── FillStrategy.java        # 填充策略接口
│           │       │   └── CellParser.java          # 单元格解析器接口
│           │       │
│           │       ├── extract/                     # 提取引擎
│           │       │   ├── ExtractEngine.java       # 提取引擎主类
│           │       │   ├── ExtractContext.java      # 提取上下文
│           │       │   └── strategy/                # 内置提取策略
│           │       │       ├── SingleStrategy.java
│           │       │       ├── DownStrategy.java
│           │       │       ├── RightStrategy.java
│           │       │       ├── BlockStrategy.java
│           │       │       └── UntilEmptyStrategy.java
│           │       │
│           │       ├── export/                      # 填充引擎
│           │       │   ├── FillEngine.java          # 填充引擎主类
│           │       │   ├── FillContext.java         # 填充上下文
│           │       │   └── strategy/                # 内置填充策略
│           │       │       ├── FillCellStrategy.java
│           │       │       ├── FillDownStrategy.java
│           │       │       ├── FillTableStrategy.java
│           │       │       └── ...
│           │       │
│           │       ├── locator/                     # 表头定位
│           │       │   ├── HeaderLocator.java       # 定位器主类
│           │       │   ├── Position.java            # 位置对象
│           │       │   └── HeaderNotFoundException.java
│           │       │
│           │       ├── config/                      # JSON 配置解析
│           │       │   └── JsonConfigParser.java    # JSON 解析器
│           │       │
│           │       ├── sax/                         # SAX 流式读取
│           │       │   ├── SaxReadHandler.java      # SAX 处理器
│           │       │   └── ...
│           │       │
│           │       └── util/                        # 工具类
│           │           ├── CellRefUtils.java        # 单元格引用工具
│           │           └── StyleUtils.java          # 样式工具
│           │
│           └── test/java/                           # 单元测试
│               └── com/excelconfig/
│                   ├── ExcelConfigHelperTest.java
│                   ├── ExcelConfigServiceTest.java
│                   ├── extract/
│                   ├── export/
│                   ├── locator/
│                   └── config/
│
├── docs/                                            # 文档
│   ├── FINAL_DESIGN.md                              # 最终设计方案
│   ├── ARCHITECTURE.md                              # 本文档
│   ├── EXTRACT_MODES.md                             # 提取模式详解
│   ├── FILL_MODES.md                                # 填充模式详解
│   ├── HEADER_MATCHING.md                           # 表头匹配机制
│   ├── COLUMN_ISOLATION.md                          # 列隔离机制
│   ├── DYNAMIC_ROW_COUNT.md                         # 动态行数机制
│   └── SAX_READER.md                                # SAX 流式读取
│
├── examples/                                        # 使用示例
│   └── USAGE_EXAMPLES.md                            # 使用示例文档
│
└── pom.xml                                          # 父 POM
```

---

## 三、核心架构

### 3.1 门面 API 设计（ExcelConfigHelper）

```java
/**
 * Excel 配置工具门面类 - 统一入口（类似 EasyExcel 风格）
 *
 * 使用示例：
 * <pre>
 * // 提取数据
 * Map<String, Object> data = ExcelConfigHelper.read("template.xlsx")
 *     .config("config.json")
 *     .extract();
 *
 * // 填充数据
 * ExcelConfigHelper.write("template.xlsx")
 *     .config("config.json")
 *     .data(data)
 *     .writeTo("output.xlsx");
 * </pre>
 */
public class ExcelConfigHelper {

    private final ExtractEngine extractEngine;
    private final FillEngine fillEngine;
    private final JsonConfigParser configParser;

    // 静态工厂方法
    public static ExcelConfigHelper read(File templateFile);
    public static ExcelConfigHelper read(String templatePath);
    public static ExcelConfigHelper read(InputStream templateStream);

    public static ExcelConfigHelper write(File templateFile);
    public static ExcelConfigHelper write(String templatePath);
    public static ExcelConfigHelper write(InputStream templateStream);

    // 配置方法
    public ExcelConfigHelper config(String configPath);
    public ExcelConfigHelper config(File configFile);
    public ExcelConfigHelper configJson(String configJson);
    public ExcelConfigHelper configObject(ExcelConfig config);

    // 数据方法
    public ExcelConfigHelper data(Map<String, Object> data);

    // 提取方法
    public Map<String, Object> extract();
    public <T> T extractAs(Class<T> clazz);

    // 填充方法
    public byte[] write();
    public void writeTo(String outputPath);
    public void writeTo(File outputFile);
    public void writeTo(Path outputPath);
    public void writeTo(OutputStream output);
}
```

### 3.2 Service API 设计（ExcelConfigService）

```java
/**
 * Excel 配置服务类 - 传统 Service 风格
 *
 * 使用示例：
 * <pre>
 * ExcelConfigService service = new ExcelConfigService();
 * Map<String, Object> data = service.extract(inputStream, configJson);
 * byte[] result = service.fill(inputStream, data, configJson);
 * </pre>
 */
public class ExcelConfigService {

    private final ExtractEngine extractEngine;
    private final FillEngine fillEngine;
    private final JsonConfigParser configParser;

    /**
     * 从 Excel 提取数据
     *
     * @param input Excel 输入流
     * @param configJson JSON 配置字符串
     * @return 提取的数据 Map
     */
    public Map<String, Object> extract(InputStream input, String configJson);

    /**
     * 填充数据到 Excel
     *
     * @param input 模板输入流
     * @param data 要填充的数据
     * @param configJson JSON 配置字符串
     * @return 填充后的 Excel 字节数组
     */
    public byte[] fill(InputStream input, Map<String, Object> data, String configJson);

    /**
     * 从配置对象提取数据
     */
    public Map<String, Object> extract(InputStream input, ExcelConfig config);

    /**
     * 从配置对象填充数据
     */
    public byte[] fill(InputStream input, Map<String, Object> data, ExcelConfig config);
}
```

### 3.3 提取引擎架构

```
┌─────────────────────────────────────────────────────────────────────────────────┐
│                           ExtractEngine 架构                                    │
├─────────────────────────────────────────────────────────────────────────────────┤
│                                                                                 │
│  ┌───────────────────────────────────────────────────────────────────────────┐  │
│  │                         ExtractEngine                                     │  │
│  │                                                                           │  │
│  │  + extract(InputStream input, ExcelConfig config): Map<String, Object>   │  │
│  │  + extract(Workbook workbook, ExcelConfig config): Map<String, Object>   │  │
│  │                                                                           │  │
│  │  内部流程：                                                                │  │
│  │  1. 创建 Workbook（SAX 或 DOM）                                            │  │
│  │  2. 遍历所有 SheetConfig                                                  │  │
│  │  3. 对每个 ExtractConfig：                                                 │  │
│  │     a. 通过 HeaderLocator 定位表头                                         │  │
│  │     b. 根据 ExtractMode 选择对应策略                                       │  │
│  │     c. 执行提取，应用 Parser                                               │  │
│  │     d. 存储到结果 Map                                                      │  │
│  │  4. 返回结果 Map                                                           │  │
│  └───────────────────────────────────────────────────────────────────────────┘  │
│                                                                                 │
│  策略模式：                                                                      │
│  ┌───────────────────────────────────────────────────────────────────────────┐  │
│  │  public interface ExtractStrategy {                                       │  │
│  │      List<Object> extract(Sheet sheet, ExtractContext context);           │  │
│  │      ExtractMode getSupportedMode();                                      │  │
│  │  }                                                                        │  │
│  │                                                                           │  │
│  │  内置实现：                                                                │  │
│  │  • SingleStrategy      - 单个单元格提取                                     │  │
│  │  • DownStrategy        - 向下提取列                                        │  │
│  │  • RightStrategy       - 向右提取行                                        │  │
│  │  • BlockStrategy       - 区域矩阵提取                                       │  │
│  │  • UntilEmptyStrategy  - 直到空行停止                                       │  │
│  └───────────────────────────────────────────────────────────────────────────┘  │
│                                                                                 │
└─────────────────────────────────────────────────────────────────────────────────┘
```

### 3.4 填充引擎架构

```
┌─────────────────────────────────────────────────────────────────────────────────┐
│                           FillEngine 架构                                       │
├─────────────────────────────────────────────────────────────────────────────────┤
│                                                                                 │
│  ┌───────────────────────────────────────────────────────────────────────────┐  │
│  │                         FillEngine                                        │  │
│  │                                                                           │  │
│  │  + fill(InputStream input, Map<String,Object> data, ExcelConfig config)  │  │
│  │     : byte[]                                                              │  │
│  │  + fill(Workbook workbook, Map<String,Object> data, ExcelConfig config)  │  │
│  │     : void                                                                │  │
│  │                                                                           │  │
│  │  内部流程：                                                                │  │
│  │  1. 创建/加载 Workbook                                                    │  │
│  │  2. 遍历所有 ExportConfig                                                 │  │
│  │  3. 对每个 ExportConfig：                                                  │  │
│  │     a. 通过 HeaderLocator 定位表头                                         │  │
│  │     b. 检查下方空间，计算需要行数                                          │  │
│  │     c. 如需扩展，下移下方内容                                              │  │
│  │     d. 根据 FillMode 选择对应策略                                          │  │
│  │     e. 执行填充，应用样式                                                  │  │
│  │  4. 返回或保存 Workbook                                                    │  │
│  └───────────────────────────────────────────────────────────────────────────┘  │
│                                                                                 │
│  策略模式：                                                                      │
│  ┌───────────────────────────────────────────────────────────────────────────┐  │
│  │  public interface FillStrategy {                                          │  │
│  │      void fill(Workbook workbook, FillContext context);                   │  │
│  │      FillMode getSupportedMode();                                         │  │
│  │  }                                                                        │  │
│  │                                                                           │  │
│  │  内置实现：                                                                │  │
│  │  • FillCellStrategy    - 填充单个单元格                                     │  │
│  │  • FillDownStrategy    - 向下填充列                                        │  │
│  │  • FillRightStrategy   - 向右填充行                                        │  │
│  │  • FillBlockStrategy   - 填充区域矩阵                                       │  │
│  │  • FillTableStrategy   - 填充表格（带表头）                                 │  │
│  └───────────────────────────────────────────────────────────────────────────┘  │
│                                                                                 │
└─────────────────────────────────────────────────────────────────────────────────┘
```

### 3.5 表头定位器

```
┌─────────────────────────────────────────────────────────────────────────────────┐
│                         HeaderLocator 设计                                      │
├─────────────────────────────────────────────────────────────────────────────────┤
│                                                                                 │
│  public class HeaderLocator {                                                   │
│                                                                                 │
│      /**                                                                        │
│       * 定位表头位置                                                             │
│       *                                                                         │
│       * @param sheet 工作表                                                      │
│       * @param config 表头配置                                                    │
│       * @return 表头位置                                                         │
│       * @throws HeaderNotFoundException 未找到表头                               │
│       */                                                                        │
│      public Position locate(Sheet sheet, HeaderConfig config) {                 │
│          // 1. 确定搜索范围                                                      │
│          int startRow = config.getInRows() != null ? config.getInRows()[0] : 1; │
│          int endRow = config.getInRows() != null ? config.getInRows()[1]        │
│                                                  : sheet.getLastRowNum();       │
│                                                                                 │
│          // 2. 在范围内搜索                                                      │
│          for (int rowNum = startRow; rowNum <= endRow; rowNum++) {              │
│              Row row = sheet.getRow(rowNum);                                    │
│              if (row == null) continue;                                         │
│                                                                                 │
│              for (Cell cell : row) {                                            │
│                  String value = getCellValueAsString(cell);                     │
│                  if (matches(value, config)) {                                  │
│                      return new Position(rowNum, cell.getColumnIndex());        │
│                  }                                                              │
│              }                                                                  │
│          }                                                                      │
│                                                                                 │
│          throw new HeaderNotFoundException(                                     │
│              "未找到表头：" + config.getMatch()                                  │
│          );                                                                     │
│      }                                                                          │
│  }                                                                              │
│                                                                                 │
│  匹配规则：                                                                      │
│  • 精确匹配 - 单元格内容完全等于配置的 match 值                                   │
│  • 模糊匹配 - 单元格内容包含配置的 match 值（未来扩展）                           │
│  • 正则匹配 - 使用正则表达式匹配（未来扩展）                                     │
│                                                                                 │
└─────────────────────────────────────────────────────────────────────────────────┘
```

---

## 四、数据流

### 4.1 导入流程

```
┌─────────────────────────────────────────────────────────────────────────────────┐
│                              导入流程                                            │
├─────────────────────────────────────────────────────────────────────────────────┤
│                                                                                 │
│   Excel 文件                                                                     │
│      │                                                                          │
│      ▼                                                                          │
│   ┌─────────────────┐                                                           │
│   │  SAX 流式读取     │ ← 内存优化，逐行处理                                      │
│   └────────┬────────┘                                                           │
│            │                                                                    │
│            ▼                                                                    │
│   ┌─────────────────┐                                                           │
│   │  HeaderLocator  │ ← 通过表头文字匹配定位                                    │
│   └────────┬────────┘                                                           │
│            │                                                                    │
│            ▼                                                                    │
│   ┌─────────────────┐                                                           │
│   │ ExtractStrategy │ ← 根据 ExtractMode 选择策略                               │
│   └────────┬────────┘                                                           │
│            │                                                                    │
│            ▼                                                                    │
│   ┌─────────────────┐                                                           │
│   │   CellParser    │ ← 解析单元格数据                                          │
│   └────────┬────────┘                                                           │
│            │                                                                    │
│            ▼                                                                    │
│   ┌─────────────────┐                                                           │
│   │  DataHandler    │ ← 数据转换/验证（可选）                                   │
│   └────────┬────────┘                                                           │
│            │                                                                    │
│            ▼                                                                    │
│   Map<String, Object>                                                           │
│   {                                                                             │
│     "orderNos": ["ORD001", "ORD002", ...],                                     │
│     "amounts": [100.0, 200.0, ...],                                            │
│     ...                                                                         │
│   }                                                                             │
│                                                                                 │
└─────────────────────────────────────────────────────────────────────────────────┘
```

### 4.2 导出流程

```
┌─────────────────────────────────────────────────────────────────────────────────┐
│                              导出流程                                            │
├─────────────────────────────────────────────────────────────────────────────────┤
│                                                                                 │
│   Map<String, Object>                                                           │
│   {                                                                             │
│     "orderNos": ["ORD001", "ORD002", ...],                                     │
│     "amounts": [100.0, 200.0, ...],                                            │
│     ...                                                                         │
│   }                                                                             │
│      │                                                                          │
│      ▼                                                                          │
│   ┌─────────────────┐                                                           │
│   │  HeaderLocator  │ ← 通过表头文字匹配定位                                    │
│   └────────┬────────┘                                                           │
│            │                                                                    │
│            ▼                                                                    │
│   ┌─────────────────┐                                                           │
│   │  空间检查        │ ← 检查下方是否有其他配置点                                │
│   └────────┬────────┘                                                           │
│            │                                                                    │
│            ▼                                                                    │
│   ┌─────────────────┐                                                           │
│   │  动态扩展        │ ← 如需更多空间，下移下方内容                               │
│   └────────┬────────┘                                                           │
│            │                                                                    │
│            ▼                                                                    │
│   ┌─────────────────┐                                                           │
│   │  FillStrategy   │ ← 根据 FillMode 选择策略                                  │
│   └────────┬────────┘                                                           │
│            │                                                                    │
│            ▼                                                                    │
│   ┌─────────────────┐                                                           │
│   │  StyleApplier   │ ← 应用单元格样式                                          │
│   └────────┬────────┘                                                           │
│            │                                                                    │
│            ▼                                                                    │
│   Excel 文件（字节数组/OutputStream）                                            │
│                                                                                 │
└─────────────────────────────────────────────────────────────────────────────────┘
```

---

## 五、核心机制

### 5.1 列隔离机制

```
场景：
┌────────────────────────────────────────┐
│  模板：                                 │
│  ┌─────────┬─────────┬─────────┐       │
│  │ 订单号   │ 金额     │ 日期     │       │
│  ├─────────┼─────────┼─────────┤       │
│  │         │         │         │  ← A 列填充 5 行
│  │         │         │         │  ← B 列填充 3 行
│  │         │         │         │  ← C 列填充 4 行
│  │         │         │         │
│  │         │         │
│  ├─────────┴─────────┴─────────┤
│  │  合计行                      │  ← 原始位置 R8
│  └─────────────────────────────┘
└────────────────────────────────────────┘

处理结果：
┌────────────────────────────────────────┐
│  填充后：                               │
│  ┌─────────┬─────────┬─────────┐       │
│  │ 订单号   │ 金额     │ 日期     │       │
│  ├─────────┼─────────┼─────────┤       │
│  │ ORD001  │ 100.00  │ 2024-01-01 │
│  │ ORD002  │ 200.00  │ 2024-01-02 │
│  │ ORD003  │ 150.00  │ 2024-01-03 │
│  │ ORD004  │         │ 2024-01-04 │
│  │ ORD005  │         │ 2024-01-05 │
│  ├─────────┴─────────┴─────────┤
│  │  合计行                      │  ← 自动下移到 R11（最大偏移）
│  └─────────────────────────────┘
└────────────────────────────────────────┘

规则：
• 每列独立计算偏移量
• 最终偏移量 = max(所有列的偏移量)
• 下方内容整体下移
```

### 5.2 动态行数确定

```
行数确定优先级（从高到低）：

1. 配置限制
   • maxRows - 最大行数限制
   • range.rows - 固定行数

2. 数据边界
   • skipEmpty=true - 遇到空行停止
   • 到达下一个配置点 - 停止
   • 到达 Sheet 末尾 - 停止

3. 默认行为
   • 提取：提取所有非空行
   • 填充：填充所有数据
```

---

## 六、配置模型

### 6.1 配置结构

```java
public class ExcelConfig {
    private String version;           // "1.0"
    private String templateName;      // 模板名称
    private List<ExtractConfig> extractions;  // 提取配置列表
    private List<ExportConfig> exports;       // 导出配置列表
}

public class ExtractConfig {
    private String key;               // 数据键名
    private HeaderConfig header;      // 表头配置
    private ExtractMode mode;         // 提取模式
    private RangeConfig range;        // 范围配置
    private ParserConfig parser;      // 解析器配置
}

public class ExportConfig {
    private String key;               // 数据键名
    private HeaderConfig header;      // 表头配置
    private FillMode mode;            // 填充模式
    private List<ColumnConfig> columns;  // 列配置（FILL_TABLE 模式）
    private StyleConfig headerStyle;  // 表头样式
    private StyleConfig style;        // 单元格样式
    private Integer maxRows;          // 最大行数
    private Boolean alternateRows;    // 隔行换色
    private Boolean autoWidth;        // 自动列宽
}
```

### 6.2 JSON 配置示例

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
      "mode": "DOWN",
      "range": { "skipEmpty": true },
      "parser": {
        "type": "number",
        "format": "#,##0.00"
      }
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

---

## 七、技术栈

| 模块 | 技术/版本 | 说明 |
|------|-----------|------|
| Java | 21 | 现代化 Java 栈 |
| Excel 处理 | Apache POI 5.2.5 | 行业标准 Excel 库 |
| JSON 处理 | Jackson 2.16.1 | 高性能 JSON 解析 |
| 日志 | SLF4J 2.0.11 | 日志门面 |
| 日志实现 | Logback 1.4.x | 运行时日志 |
| 测试 | JUnit 5 + Mockito | 单元测试框架 |

---

## 八、扩展点

### 8.1 自定义提取策略

```java
public class CustomExtractStrategy implements ExtractStrategy {
    
    @Override
    public List<Object> extract(Sheet sheet, ExtractContext context) {
        // 自定义提取逻辑
        List<Object> result = new ArrayList<>();
        // ... 实现提取逻辑
        return result;
    }
    
    @Override
    public ExtractMode getSupportedMode() {
        return ExtractMode.CUSTOM;  // 需要自定义 Mode 枚举
    }
}
```

### 8.2 自定义单元格解析器

```java
public class CustomCellParser implements CellParser {
    
    @Override
    public Object parse(Cell cell, ParserConfig config) {
        // 自定义解析逻辑
        String value = getCellValueAsString(cell);
        // ... 实现解析逻辑
        return parsedValue;
    }
}
```

### 8.3 自定义填充策略

```java
public class CustomFillStrategy implements FillStrategy {
    
    @Override
    public void fill(Workbook workbook, FillContext context) {
        // 自定义填充逻辑
        Sheet sheet = workbook.getSheetAt(0);
        // ... 实现填充逻辑
    }
    
    @Override
    public FillMode getSupportedMode() {
        return FillMode.CUSTOM;  // 需要自定义 Mode 枚举
    }
}
```

---

## 九、测试策略

### 9.1 单元测试

- **HeaderLocatorTest** - 表头定位器测试
- **ExtractEngineTest** - 提取引擎测试
- **FillEngineTest** - 填充引擎测试
- **JsonConfigParserTest** - JSON 解析器测试
- **ExcelConfigHelperTest** - 门面 API 测试
- **ExcelConfigServiceTest** - Service API 测试

### 9.2 集成测试

- **ColumnIsolationTest** - 列隔离测试
- **FillAutoExpandTest** - 填充自动扩展测试
- **RealFileTest** - 真实文件测试

### 9.3 测试覆盖

当前测试覆盖：
- 57 个单元测试
- 100% 核心逻辑覆盖
- 支持真实文件验证

---

## 十、性能优化

### 10.1 SAX 流式读取

```java
// 传统 DOM 模式 - 内存占用 O(n)
Workbook workbook = WorkbookFactory.create(inputStream);

// SAX 流式模式 - 内存占用 O(1)
// ExtractEngine 内部使用
Map<String, Object> result = extractEngine.extract(inputStream, config);
```

### 10.2 批量操作

```java
// 填充引擎内部优化
// • 批量行下移（不是逐行）
// • 样式复用（不是每次创建）
// • 延迟计算（不是预计算）
```

---

## 十一、总结

### 核心特性

1. **表头匹配定位** - 配置通过表头文字匹配，不依赖固定位置
2. **数据量驱动** - 提取/填充行数由实际数据决定
3. **自动扩展** - 模板空间不足时自动下移下方内容
4. **列隔离** - 每列独立处理，互不干扰
5. **配置驱动边界** - 配置点即边界，自动检测
6. **SAX 流式读取** - 内存优化，支持大文件
7. **简洁 API** - ExcelConfigHelper 门面类，类似 EasyExcel

### Maven 依赖

```xml
<dependency>
    <groupId>com.excelconfig</groupId>
    <artifactId>excel-config-core</artifactId>
    <version>1.0.0</version>
</dependency>
```

### 快速开始

```java
// 提取数据
Map<String, Object> data = ExcelConfigHelper.read("template.xlsx")
    .config("config.json")
    .extract();

// 填充数据
ExcelConfigHelper.write("template.xlsx")
    .config("config.json")
    .data(data)
    .writeTo("output.xlsx");
```
