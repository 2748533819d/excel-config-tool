# Excel 动态配置工具 - 可行性分析与设计思想

## 零、拓展性设计目标 ⭐

### 核心设计理念

参考 EasyExcel 的设计思想：
1. **策略可插拔** - 每个提取模式独立实现，通过 SPI 注册
2. **解析器可配置** - 每个单元格配置可以指定专属解析器
3. **数据处理器可链式** - 提取后的数据可以经过多个处理器链式处理

### 支持的读取模式

| 模式 | 配置示例 | 输出 |
|------|----------|------|
| **单一单元格** | `pos: "A1"` | `{"key": "value"}` |
| **下方列表** | `pos: "A1", extract: DOWN` | `{"key": ["v1","v2","v3"]}` |
| **右侧列表** | `pos: "A1", extract: RIGHT` | `{"key": ["v1","v2","v3"]}` |
| **区域块** | `pos: "A1:C10", extract: BLOCK` | `{"key": [[v11,v12],[v21,v22]]}` |
| **多行多列** | `pos: "A1", rows: 5, cols: 3` | `{"key": [[...], ...]}` |
| **动态范围** | `pos: "A1", until: "EMPTY"` | `{"key": ["v1","v2",...]}` |

### 拓展性设计

```
┌─────────────────────────────────────────────────────────────────┐
│                        可拓展点设计                              │
├─────────────────────────────────────────────────────────────────┤
│                                                                 │
│  1. 提取策略 (ExtractStrategy) - SPI 拓展                        │
│     用户可以实现自定义的提取逻辑                                 │
│                                                                 │
│  2. 单元格解析器 (CellParser) - 每个配置独立                     │
│     原始值 → 目标类型的转换逻辑                                  │
│                                                                 │
│  3. 数据处理器 (DataHandler) - 链式处理                          │
│     提取后的数据可以经过过滤、转换、验证等处理                    │
│                                                                 │
│  4. 位置解析器 (PositionResolver) - 自定义定位                   │
│     支持自定义的定位方式                                         │
│                                                                 │
└─────────────────────────────────────────────────────────────────┘
```

---

## 一、优秀 Excel 工具的核心思想分析

### 1. Apache POI 的设计思想

#### 两种解析模式

**DOM 模式 (User Model)**
- 将整个 Excel 文件加载到内存中
- 提供完整的对象树，可以随机访问任意单元格
- 优点：功能完整，支持所有 Excel 特性
- 缺点：内存消耗大，大文件容易 OOM

**SAX 模式 (Event User Model)**
- 基于事件驱动的流式解析
- 逐行读取，处理完一行立即释放
- 核心类：`XSSFReader` + `XSSFSheetXMLHandler`
- 内存占用：几 MB vs DOM 模式的几百 MB

```
┌─────────────────────────────────────────────────────────┐
│                    SAX 解析流程                          │
├─────────────────────────────────────────────────────────┤
│  文件输入流 → XMLParser → ContentHandler → 你的处理器     │
│       ↓              ↓           ↓            ↓         │
│    .xlsx 文件   解析 XML   触发事件     接收行/单元格数据   │
│                                                        │
│  核心接口：ContentHandler                              │
│  - startDocument()   - 文档开始                         │
│  - startElement()    - 元素开始 (row/cell)              │
│  - characters()      - 字符内容 (单元格值)               │
│  - endElement()      - 元素结束                         │
│  - endDocument()     - 文档结束                         │
└─────────────────────────────────────────────────────────┘
```

### 2. EasyExcel 的核心思想

#### 核心创新点

1. **重写解析逻辑**
   - 基于 POI 的 SAX 模式，但重写了核心解析逻辑
   - 针对 07+ (.xlsx) 格式优化
   - 避免 POI 的一些冗余处理

2. **分析器 (Analyzer) 模式**
   - 定义 `AnalysisEventListener` 接口
   - 用户实现 `invoke(Map<Integer, String> rowData)` 方法
   - 每解析一行，触发一次回调

3. **懒加载 + 流式处理**
   - 不构建完整的对象树
   - 按需解析，处理完即释放

```
┌─────────────────────────────────────────────────────────┐
│                  EasyExcel 架构                          │
├─────────────────────────────────────────────────────────┤
│                                                        │
│  Excel 文件 → 解压 → XML 流 → 解析器 → 事件监听器          │
│                              ↓                         │
│                        RowMapper                       │
│                              ↓                         │
│                        用户定义的处理逻辑                 │
│                                                        │
│  关键类：                                               │
│  - ExcelReader: 读取器入口                              │
│  - AnalysisEventListener: 事件监听器                    │
│  - RowModel: 行数据模型                                │
└─────────────────────────────────────────────────────────┘
```

### 3. 其他工具的借鉴点

| 工具 | 借鉴点 |
|------|--------|
| **Poiji2** | 注解映射：@ExcelRow, @ExcelCell |
| **NanoXLSX4j** | 轻量级设计，专注核心功能 |
| **Windmill** | Fluent API，链式调用 |

---

## 二、你的想法可行性分析

### 需求描述
```
根据配置信息（key + position），提取指定单元格的值，生成 Map
支持将值填入指定单元格位置
配置结构：key → position → value
```

### ✅ 可行性评估：**完全可行**

#### 技术可行性
1. **读取场景**：基于 POI 的 SAX 模式，可以高效定位并提取指定单元格
2. **写入场景**：基于 POI 的 DOM 模式或流式写入，可以填充指定单元格

#### 商业价值
1. 配置化：非开发人员可以通过配置调整 Excel 模板
2. 解耦：业务代码与 Excel 结构解耦
3. 可维护性：配置集中管理，易于修改

---

## 三、核心设计思路

### 3.1 配置模型设计（支持拓展）

```java
/**
 * 单元格配置 - 支持多种提取模式和独立解析器
 */
public class CellConfig {
    private String key;                    // 业务键名
    private Position position;             // 起始位置
    private ExtractMode extractMode;       // 提取模式
    private ExtractRange range;            // 提取范围配置
    
    // ===== 核心拓展点：独立的解析器和处理器 =====
    private String parserType;             // 解析器类型 (SPI 名称或类名)
    private Map<String, Object> parserParams; // 解析器参数
    private List<DataHandlerConfig> dataHandlers; // 数据处理器链
    
    // 基础字段
    private CellType type;                 // 单元格类型
    private String defaultValue;           // 默认值
    private boolean required;              // 是否必填
    
    // 拓展字段
    private Map<String, Object> extensions; 
}

/**
 * 数据处理器配置 - 支持链式配置
 */
public class DataHandlerConfig {
    private String type;                   // 处理器类型 (SPI 名称或类名)
    private Map<String, Object> params;    // 处理器参数
    private String order;                  // 执行顺序
}

/**
 * 提取模式
 */
public enum ExtractMode {
    SINGLE,          // 单一单元格 → Object
    DOWN,            // 向下提取 → List<Object>
    RIGHT,           // 向右提取 → List<Object>
    BLOCK,           // 区域块 → List<List<Object>> (二维数组)
    UNTIL_EMPTY,     // 直到空值 → List<Object>
    CUSTOM           // 自定义策略
}

/**
 * 提取范围配置
 */
public class ExtractRange {
    private Integer rows;              // 向下提取多少行 (用于 DOWN/BLOCK 模式)
    private Integer cols;              // 向右提取多少列 (用于 RIGHT/BLOCK 模式)
    private String untilCondition;     // 终止条件：如 "EMPTY", "BLANK_ROW", 或自定义表达式
    private Boolean skipEmpty;         // 是否跳过空值
}

/**
 * 位置信息 - 支持多种定位方式
 */
public class Position {
    // 方式 1: 绝对坐标
    private Integer row;               // 行号 (0-based)
    private Integer col;               // 列号 (0-based)
    
    // 方式 2: Excel 风格坐标
    private String cellRef;            // 如 "A1", "B2"
    
    // 方式 3: 相对定位
    private String anchorKey;          // 锚点键 (相对于某个已定位的单元格)
    private RowOffset rowOffset;       // 行偏移
    private ColOffset colOffset;       // 列偏移
    
    // 方式 4: 表头定位
    private String headerName;         // 表头名称 (自动查找该列)
    
    // 方式 5: 命名范围
    private String namedRange;         // Excel 命名范围
    
    // 方式 6: 区域
    private String areaRef;            // 如 "A1:C10"
}

/**
 * 完整的 Excel 配置
 */
public class ExcelTemplateConfig {
    private String templateName;
    private Map<String, CellConfig> cellConfigs;   // key → config
    private List<SheetConfig> sheets;              // 多工作表支持
    private Map<String, ExtractStrategy> customStrategies; // 自定义提取策略
    private Map<String, CellParser> customParsers; // 自定义解析器
    private Map<String, DataHandler> customHandlers; // 自定义处理器
}
```

### 3.2 配置文件格式 (YAML 示例)

```yaml
template:
  name: "订单导入模板"
  file: "order_template.xlsx"

# ===== 全局解析器定义 =====
parsers:
  # 内置解析器快捷配置
  dateParser:
    type: "date"
    pattern: "yyyy-MM-dd"
    locale: "zh_CN"
    
  # 自定义解析器
  customMoneyParser:
    type: "com.example.MoneyParser"
    params:
      currency: "CNY"
      scale: 2

# ===== 全局数据处理器定义 =====
handlers:
  trimString:
    type: "trim"
  notEmpty:
    type: "notEmpty"
    params:
      errorMessage: "字段不能为空"

cells:
  # ===== 单一单元格 + 内置解析器 =====
  orderNo:
    position: { cellRef: "A2" }
    extractMode: SINGLE
    parserType: "string"
    dataHandlers:
      - { type: "trim" }
      - { type: "notEmpty" }
    required: true
    
  # ===== 向下提取列表 + 自定义解析器链 =====
  productList:
    position: { cellRef: "B2" }
    extractMode: DOWN
    range: { rows: 10, skipEmpty: true }
    parserType: "string"
    dataHandlers:
      - { type: "trim" }
      - { type: "deduplicate" }  # 去重
    
  # ===== 金额字段 + 自定义解析器 =====
  amounts:
    position: { cellRef: "C2" }
    extractMode: DOWN
    range: { rows: 10 }
    parserType: "com.example.MoneyParser"
    parserParams:
      currency: "CNY"
      scale: 2
    dataHandlers:
      - { type: "positive" }  # 验证正数
      
  # ===== 日期字段 + 内置日期解析器 =====
  orderDates:
    position: { cellRef: "D2" }
    extractMode: DOWN
    range: { rows: 10 }
    parserType: "date"
    parserParams:
      pattern: "yyyy-MM-dd"
    defaultValue: "2024-01-01"
    
  # ===== 向右提取 (表头) =====
  monthHeaders:
    position: { cellRef: "E1" }
    extractMode: RIGHT
    range: { cols: 12 }
    parserType: "string"
    
  # ===== 区域块 + 数据处理器 =====
  dataMatrix:
    position: { areaRef: "A1:C10" }
    extractMode: BLOCK
    parserType: "string"
    dataHandlers:
      - { type: "flatten" }  # 二维转一维
      
  # ===== 直到空值 + 链式处理 =====
  dynamicList:
    position: { cellRef: "F2" }
    extractMode: UNTIL_EMPTY
    parserType: "string"
    dataHandlers:
      - { type: "trim" }
      - { type: "filter", params: { pattern: "^[A-Z].*" } }  # 正则过滤
      - { type: "transform", params: { expr: "toUpperCase()" } }

# ===== 自定义提取策略 =====
customStrategies:
  # 提取键值对 (A 列是 key, B 列是 value)
  keyValuePairs:
    type: "com.example.ConfigKVStrategy"
    params:
      keyColumn: 0
      valueColumn: 1
      startRow: 1
```

### 3.3 核心架构设计 - 参考 EasyExcel 的易用性设计

```
┌─────────────────────────────────────────────────────────────────┐
│                     Excel Config Tool 架构                       │
│                    (参考 EasyExcel 设计思想)                      │
├─────────────────────────────────────────────────────────────────┤
│                                                                 │
│  使用示例 (Java API):                                           │
│  ┌───────────────────────────────────────────────────────────┐ │
│  │  Map<String, Object> data = ExcelReader.read(            │ │
│  │      new File("template.xlsx"),                           │ │
│  │      ExcelConfig.fromYaml("config.yaml")                  │ │
│  │  );                                                       │ │
│  └───────────────────────────────────────────────────────────┘ │
│                                                                 │
│  ┌──────────────┐    ┌──────────────┐    ┌──────────────┐      │
│  │  配置解析器   │    │   Excel 引擎   │    │  结果映射器   │      │
│  │              │    │              │    │              │      │
│  │ YAML/JSON   │───→│ POI Core    │───→│ Map<K,V>     │      │
│  │ 配置解析    │    │ 读取/写入    │    │ 对象绑定     │      │
│  └──────────────┘    └──────────────┘    └──────────────┘      │
│         │                   │                    │              │
│         ▼                   ▼                    ▼              │
│  ┌─────────────────────────────────────────────────────────┐   │
│  │  SPI 拓展层：                                             │   │
│  │  - ExtractStrategy (提取策略)                            │   │
│  │  - CellParser (单元格解析器)                             │   │
│  │  - DataHandler (数据处理器)                              │   │
│  │  - PositionResolver (位置解析器)                         │   │
│  └─────────────────────────────────────────────────────────┘   │
└─────────────────────────────────────────────────────────────────┘
```

### 核心 SPI 接口设计

```java
/**
 * 提取策略接口 - 核心 SPI (类似 EasyExcel 的 AnalysisEventListener)
 * 每个提取模式对应一个策略实现
 */
public interface ExtractStrategy {
    /**
     * 执行提取
     */
    Object extract(ExtractContext context);
    
    /**
     * 支持的提取模式
     */
    Set<ExtractMode> supportedModes();
    
    /**
     * 设置解析器工厂
     */
    void setParserFactory(ParserFactory factory);
    
    /**
     * 设置处理器链
     */
    void setDataHandlerChain(List<DataHandler> handlers);
}

/**
 * 单元格解析器接口 - 每个配置可独立指定
 */
public interface CellParser<T> {
    /**
     * 解析原始单元格值
     * @param rawValue 原始值
     * @param context 上下文
     * @return 解析后的值
     */
    T parse(String rawValue, ParseContext context);
    
    /**
     * 解析器类型标识
     */
    String getName();
}

/**
 * 数据处理器接口 - 链式处理
 */
public interface DataHandler<T> {
    /**
     * 处理数据
     * @param data 输入数据
     * @param context 上下文
     * @return 处理后的数据
     */
    T handle(T data, HandlerContext context);
    
    /**
     * 处理器顺序
     */
    default int getOrder() { return 0; }
}

/**
 * 位置解析器接口
 */
public interface PositionResolver {
    /**
     * 解析位置
     * @param sheet Excel Sheet
     * @param position 位置配置
     * @return 实际的单元格引用
     */
    Cell resolve(Sheet sheet, Position position);
}
```

### 3.4 提取策略接口设计（SPI 拓展）

```java
/**
 * 数据提取策略接口 - 核心 SPI
 */
public interface ExtractStrategy {
    /**
     * 执行提取
     * @param sheet Excel Sheet
     * @param startPosition 起始位置
     * @param config 提取配置
     * @return 提取结果
     */
    Object extract(Sheet sheet, Position startPosition, CellConfig config);
    
    /**
     * 支持的提取模式
     */
    Set<ExtractMode> supportedModes();
}

/**
 * 基础抽象策略 - 提供通用工具方法
 */
public abstract class BaseExtractStrategy implements ExtractStrategy {
    protected CellValueExtractor cellValueExtractor;
    protected PositionResolver positionResolver;
    
    // 工具方法：获取单元格值
    protected Object getCellValue(Cell cell, CellType type) { ... }
    
    // 工具方法：判断是否为空
    protected boolean isEmpty(Cell cell) { ... }
}

/**
 * 策略注册表 - 管理所有提取策略
 */
public class StrategyRegistry {
    private Map<ExtractMode, ExtractStrategy> strategies = new HashMap<>();
    
    public void register(ExtractMode mode, ExtractStrategy strategy) { ... }
    public ExtractStrategy getStrategy(ExtractMode mode) { ... }
}
```

### 3.4 内置提取策略实现 - 每个策略独立可配置

```java
/**
 * 基础抽象策略 - 提供通用工具方法和链式处理支持
 */
public abstract class BaseExtractStrategy implements ExtractStrategy {
    protected ParserFactory parserFactory;
    protected List<DataHandler> handlerChain;
    
    @Override
    public void setParserFactory(ParserFactory factory) {
        this.parserFactory = factory;
    }
    
    @Override
    public void setDataHandlerChain(List<DataHandler> handlers) {
        this.handlerChain = handlers;
    }
    
    /**
     * 解析单元格值 - 支持独立配置的解析器
     */
    protected <T> T parseCellValue(Cell cell, String parserType, Map<String, Object> params) {
        CellParser<T> parser = parserFactory.getParser(parserType, params);
        String rawValue = getRawCellValue(cell);
        return parser.parse(rawValue, createParseContext(cell));
    }
    
    /**
     * 应用数据处理器链
     */
    @SuppressWarnings("unchecked")
    protected <T> T applyHandlerChain(T data) {
        if (handlerChain == null) return data;
        
        T result = data;
        for (DataHandler handler : handlerChain) {
            result = (T) handler.handle(result, createHandlerContext());
        }
        return result;
    }
    
    // 工具方法
    protected String getRawCellValue(Cell cell) { ... }
    protected boolean isEmpty(Cell cell) { ... }
    protected ParseContext createParseContext(Cell cell) { ... }
    protected HandlerContext createHandlerContext() { ... }
}

/**
 * 单一单元格提取策略
 */
public class SingleCellStrategy extends BaseExtractStrategy {
    @Override
    public Object extract(ExtractContext context) {
        Cell cell = context.getResolver().resolve(context.getSheet(), context.getPosition());
        Object value = parseCellValue(cell, context.getParserType(), context.getParserParams());
        return applyHandlerChain(value);
    }
    
    @Override
    public Set<ExtractMode> supportedModes() {
        return Set.of(ExtractMode.SINGLE);
    }
}

/**
 * 向下提取列表策略
 */
public class DownListStrategy extends BaseExtractStrategy {
    @Override
    public Object extract(ExtractContext context) {
        List<Object> result = new ArrayList<>();
        Position startPos = context.getPosition();
        ExtractRange range = context.getRange();
        
        Cell startCell = context.getResolver().resolve(context.getSheet(), startPos);
        int startRow = startCell.getRowIndex();
        int col = startCell.getColumnIndex();
        
        for (int i = 0; i < range.getRows(); i++) {
            Row row = context.getSheet().getRow(startRow + i);
            if (row == null) break;
            
            Cell cell = row.getCell(col);
            if (isEmpty(cell)) {
                if ("EMPTY".equals(range.getUntilCondition())) break;
                if (range.getSkipEmpty()) continue;
            }
            
            Object value = parseCellValue(cell, context.getParserType(), context.getParserParams());
            result.add(value);
        }
        
        return applyHandlerChain(result);
    }
    
    @Override
    public Set<ExtractMode> supportedModes() {
        return Set.of(ExtractMode.DOWN);
    }
}

/**
 * 向右提取列表策略
 */
public class RightListStrategy extends BaseExtractStrategy {
    @Override
    public Object extract(ExtractContext context) {
        List<Object> result = new ArrayList<>();
        // 实现类似 DownListStrategy，只是遍历列
        // ...
        return applyHandlerChain(result);
    }
    
    @Override
    public Set<ExtractMode> supportedModes() {
        return Set.of(ExtractMode.RIGHT);
    }
}

/**
 * 区域块提取策略 (二维数组)
 */
public class BlockStrategy extends BaseExtractStrategy {
    @Override
    public Object extract(ExtractContext context) {
        List<List<Object>> result = new ArrayList<>();
        AreaRef areaRef = parseAreaRef(context.getPosition().getAreaRef());
        
        for (int r = areaRef.firstRow; r <= areaRef.lastRow; r++) {
            List<Object> rowList = new ArrayList<>();
            Row row = context.getSheet().getRow(r);
            if (row == null) continue;
            
            for (int c = areaRef.firstCol; c <= areaRef.lastCol; c++) {
                Cell cell = row.getCell(c);
                Object value = parseCellValue(cell, context.getParserType(), context.getParserParams());
                rowList.add(value);
            }
            result.add(rowList);
        }
        
        return applyHandlerChain(result);
    }
    
    @Override
    public Set<ExtractMode> supportedModes() {
        return Set.of(ExtractMode.BLOCK);
    }
}

/**
 * 直到空值提取策略
 */
public class UntilEmptyStrategy extends BaseExtractStrategy {
    @Override
    public Object extract(ExtractContext context) {
        List<Object> result = new ArrayList<>();
        // 实现：循环提取直到遇到空值
        // ...
        return applyHandlerChain(result);
    }
    
    @Override
    public Set<ExtractMode> supportedModes() {
        return Set.of(ExtractMode.UNTIL_EMPTY);
    }
}
```

### 3.6 读取流程 (基于 SAX 思想，支持多种提取模式)

```java
// 伪代码展示核心思想
public class ConfigBasedExcelReader {
    
    private ExcelTemplateConfig config;
    private StrategyRegistry strategyRegistry;
    private Map<String, Object> result = new HashMap<>();
    
    // 基于 SAX 的事件驱动
    public void read(InputStream inputStream) {
        // 1. 创建 SAX 解析器
        SAXParser parser = createParser();
        
        // 2. 设置 ContentHandler
        parser.setContentHandler(new SimpleXSSFSheetXMLHandler(
            new EventListeningSheetContentsHandler()
        ));
        
        // 3. 解析过程中根据配置提取目标单元格
        class EventListeningSheetContentsHandler implements SheetContentsHandler {
            @Override
            public void row(Row row) {
                // 检查该行是否有配置的单元格
                for (CellConfig cellConfig : config.getCellConfigs()) {
                    if (row.getRowNum() == cellConfig.getPosition().getRow()) {
                        Cell cell = row.getCell(cellConfig.getPosition().getCol());
                        result.put(cellConfig.getKey(), extractValue(cell));
                    }
                }
            }
        }
    }
    
    // 统一的提取入口
    public Map<String, Object> extractAll(Sheet sheet) {
        Map<String, Object> result = new HashMap<>();
        
        for (CellConfig cellConfig : config.getCellConfigs()) {
            ExtractStrategy strategy = strategyRegistry.getStrategy(
                cellConfig.getExtractMode()
            );
            Position pos = cellConfig.getPosition();
            result.put(cellConfig.getKey(), strategy.extract(sheet, pos, cellConfig));
        }
        
        return result;
    }
}
```

### 使用示例

```java
// 配置
YAML 配置:
products:
  position: { cellRef: "A2" }
  extractMode: DOWN
  range: { rows: 10, skipEmpty: true }
  
prices:
  position: { cellRef: "B2" }
  extractMode: DOWN
  range: { rows: 10 }

// Java 调用
ExcelConfigReader reader = ExcelConfigReader.fromYaml("config.yaml");
Map<String, Object> data = reader.extractAll(workbook.getSheetAt(0));

// 结果
// data.get("products") → List<String> ["产品 A", "产品 B", ...]
// data.get("prices")   → List<Decimal> [100.00, 200.00, ...]
```

### 3.5 写入流程

```java
public class ConfigBasedExcelFiller {
    
    private ExcelTemplateConfig config;
    
    public void fill(InputStream template, 
                     Map<String, Object> data,
                     OutputStream output) {
        // 1. 加载模板 (DOM 模式，因为需要修改)
        XSSFWorkbook workbook = new XSSFWorkbook(template);
        
        // 2. 根据配置填充数据
        for (Map.Entry<String, CellConfig> entry : config.getCells().entrySet()) {
            String key = entry.getKey();
            CellConfig cellConfig = entry.getValue();
            Object value = data.get(key);
            
            Position pos = cellConfig.getPosition();
            XSSFSheet sheet = workbook.getSheetAt(0);
            XSSFRow row = getOrCreateRow(sheet, pos.getRow());
            XSSFCell cell = getOrCreateCell(row, pos.getCol());
            
            // 3. 设置值
            setCellValue(cell, value, cellConfig.getType());
        }
        
        // 4. 写出
        workbook.write(output);
    }
}
```

---

## 四、关键技术点

### 4.1 单元格定位策略

| 定位方式 | 实现思路 | 适用场景 |
|----------|----------|----------|
| **绝对坐标** | 直接通过行号列号定位 | 固定格式模板 |
| **表头查找** | 遍历第一行找表头名 | 动态列顺序 |
| **命名范围** | 利用 Excel NamedRange | 复杂模板 |
| **锚点 + 偏移** | 先找锚点单元格再偏移 | 相对位置场景 |

### 4.2 类型转换系统

```java
public interface TypeConverter<T> {
    T convert(String rawValue, CellType targetType);
}

// 内置转换器
- StringConverter
- IntegerConverter  
- DecimalConverter
- DateConverter (支持 pattern)
- BooleanConverter
- EnumConverter
```

### 4.3 优化策略

1. **读取优化**
   - 只解析配置中涉及的 Sheet
   - 使用 SAX 模式，跳过无关行
   - 缓存已解析的配置

2. **写入优化**
   - 流式写入 (SXSSFWorkbook) 适合大数据量
   - 复用模板文件，避免重复创建样式

---

## 五、推荐技术栈

```xml
<dependencies>
    <!-- 核心依赖：只依赖 POI -->
    <dependency>
        <groupId>org.apache.poi</groupId>
        <artifactId>poi-ooxml</artifactId>
        <version>5.2.5</version>
    </dependency>
    
    <!-- 配置解析 -->
    <dependency>
        <groupId>org.yaml</groupId>
        <artifactId>snakeyaml</artifactId>
        <version>2.2</version>
    </dependency>
    
    <!-- 可选：JSON 配置支持 -->
    <dependency>
        <groupId>com.fasterxml.jackson.core</groupId>
        <artifactId>jackson-databind</artifactId>
    </dependency>
    
    <!-- 工具类 -->
    <dependency>
        <groupId>com.google.guava</groupId>
        <artifactId>guava</artifactId>
    </dependency>
</dependencies>
```

---

## 七、项目结构建议

```
excel-config-tool/
├── src/main/java/com/excelconfig/
│   ├── core/
│   │   ├── ExcelConfigReader.java       # 配置读取入口
│   │   ├── ExcelExtractor.java          # 数据提取器 (统一入口)
│   │   ├── ExcelFiller.java             # 数据填充器
│   │   └── CellPositionResolver.java    # 位置解析器
│   ├── config/
│   │   ├── ExcelTemplateConfig.java     # 配置模型
│   │   ├── CellConfig.java              # 单元格配置
│   │   ├── Position.java                # 位置定义
│   │   ├── ExtractMode.java             # 提取模式枚举
│   │   └── ExtractRange.java            # 提取范围配置
│   ├── strategy/
│   │   ├── ExtractStrategy.java         # 提取策略接口 (SPI)
│   │   ├── BaseExtractStrategy.java     # 基础抽象类
│   │   ├── StrategyRegistry.java        # 策略注册表
│   │   └── builtin/                     # 内置策略实现
│   │       ├── SingleCellStrategy.java  # 单一单元格
│   │       ├── DownListStrategy.java    # 向下列表
│   │       ├── RightListStrategy.java   # 向右列表
│   │       ├── BlockStrategy.java       # 区域块
│   │       └── UntilEmptyStrategy.java  # 直到空值
│   ├── parser/
│   │   ├── SaxBasedReader.java          # SAX 读取实现
│   │   ├── DomBasedWriter.java          # DOM 写入实现
│   │   └── YamlConfigParser.java        # YAML 配置解析
│   ├── converter/
│   │   ├── TypeConverter.java           # 类型转换接口
│   │   └── converters/                  # 具体转换器实现
│   │       ├── StringConverter.java
│   │       ├── IntegerConverter.java
│   │       ├── DecimalConverter.java
│   │       ├── DateConverter.java
│   │       └── BooleanConverter.java
│   ├── spi/
│   │   └── ExtractStrategyProvider.java # SPI 服务提供者接口
│   └── util/
│       ├── ExcelUtils.java              # 工具方法
│       └── CellRefParser.java           # Excel 坐标解析 (A1 ↔ row/col)
├── src/test/java/
│   └── com/excelconfig/
│       ├── strategy/
│       │   ├── SingleCellStrategyTest.java
│       │   ├── DownListStrategyTest.java
│       │   └── BlockStrategyTest.java
│       └── integration/
│           └── ExtractorIntegrationTest.java
├── src/main/resources/
│   └── META-INF/services/
│       └── com.excelconfig.spi.ExtractStrategyProvider  # SPI 配置
├── pom.xml
└── README.md
```

---

## 九、EasyExcel 内存优化思想深度分析 ⭐

### 9.1 传统 POI 的内存问题

```
传统 POI DOM 模式的问题:
┌─────────────────────────────────────────────────────────┐
│  Excel 文件 (100MB)                                     │
│       ↓                                                 │
│  完全加载到内存 → Workbook (对象树)                     │
│       ↓                                                 │
│  内存占用：300MB - 500MB (可能 OOM)                     │
└─────────────────────────────────────────────────────────┘

原因:
1. DOM 模式需要构建完整的对象树
2. XML 完全解析后全部驻留内存
3. 样式、公式等元数据也全部加载
```

### 9.2 EasyExcel 的内存优化核心思想

```
EasyExcel 的 SAX 流式处理:
┌─────────────────────────────────────────────────────────┐
│  Excel 文件 (.xlsx 本质是 ZIP 压缩包)                    │
│       ↓ 解压流式读取                                    │
│  XML 片段流 (SheetXML)                                  │
│       ↓  SAX 解析器                                     │
│  事件回调 (startElement, characters, endElement)        │
│       ↓  逐行处理                                       │
│  用户代码 (invoke 方法)                                 │
│       ↓  处理完立即释放                                 │
│  内存占用：< 10MB (恒定)                                │
└─────────────────────────────────────────────────────────┘

核心思想:
1. 不构建完整对象树
2. 流式解压 + 流式解析
3. 处理完一行立即释放，不保留引用
4. 按需读取，而不是预先加载
```

### 9.3 EasyExcel 的关键优化点

| 优化点 | 传统 POI | EasyExcel | 内存对比 |
|--------|----------|-----------|----------|
| **解析模式** | DOM (全量加载) | SAX (流式) | 100MB vs 5MB |
| **数据保留** | 全部保留 | 用完即弃 | 持久 vs 瞬时 |
| **样式处理** | 全部加载 | 按需加载 | 大量 vs 少量 |
| **行对象创建** | 复用少 | 高度复用 | 多对象 vs 单对象 |

### 9.4 本项目如何借鉴 EasyExcel 的内存优化

#### 方案一：基于 POI 的 SXSSF/SAX 混合模式

```java
/**
 * 流式读取器 - 借鉴 EasyExcel 思想
 * 只保留当前处理的行在内存中
 */
public class StreamingExcelReader {
    
    private final int rowCacheSize = 100;  // 只缓存 100 行
    
    public void readLargeFile(InputStream input, RowHandler handler) {
        // 使用 POI 的 SXSSF 流式 API
        try (OPCPackage opc = OPCPackage.open(input)) {
            XSSFReader reader = new XSSFReader(opc);
            
            // 获取共享字符串表 (懒加载)
            SharedStringsTable sst = reader.getSharedStringsTable();
            
            // 遍历所有 sheet
            Iterator<InputStream> sheets = reader.getSheetsData();
            while (sheets.hasNext()) {
                InputStream sheetStream = sheets.next();
                
                // SAX 解析
                SAXParser parser = SAXParserFactory.newInstance().newSAXParser();
                parser.setProperty(
                    "http://xml.org/sax/properties/lexical-handler",
                    new StreamingSheetHandler(sst, handler)
                );
                parser.parse(new InputSource(sheetStream));
                
                // 解析完成后 sheetStream 可以关闭，内存释放
            }
        }
    }
}

/**
 * 流式行处理器 - 类似 EasyExcel 的 AnalysisEventListener
 */
@FunctionalInterface
public interface RowHandler {
    /**
     * 处理一行数据
     * @param rowNum 行号
     * @param rowData 行数据
     * @param isLast 是否最后一行
     */
    void invoke(int rowNum, Map<String, Object> rowData, boolean isLast);
}
```

#### 方案二：配置驱动的按需解析（我们的创新）

```java
/**
 * 配置驱动的按需解析 - 只解析配置的单元格
 * 比 EasyExcel 更进一步：不需要的行/列直接跳过
 */
public class SmartStreamingReader {
    
    private ExcelTemplateConfig config;
    
    public Map<String, Object> readSelective(InputStream input) {
        // 1. 分析配置，计算需要解析的行范围
        RowRange neededRows = analyzeNeededRows(config);
        
        // 2. 只解析需要的行
        // 如果配置只关心 A2 和 B2，那么第 3-10000 行直接跳过
        skipUnneededRows(input, neededRows);
        
        // 3. 提取配置单元格的值
        return extractConfiguredCells(input, config);
    }
    
    private RowRange analyzeNeededRows(ExcelTemplateConfig config) {
        int minRow = Integer.MAX_VALUE;
        int maxRow = Integer.MIN_VALUE;
        
        for (CellConfig cell : config.getCellConfigs()) {
            if (cell.getExtractMode() == ExtractMode.DOWN) {
                minRow = Math.min(minRow, cell.getPosition().getRow());
                maxRow = Math.max(maxRow, 
                    cell.getPosition().getRow() + cell.getRange().getRows());
            }
            // ... 分析其他模式
        }
        
        return new RowRange(minRow, maxRow);
    }
}
```

#### 方案三：对象池化 + 复用

```java
/**
 * 行对象池 - 减少 GC 压力
 */
public class RowObjectPool {
    private final ArrayDeque<Map<String, Object>> pool = new ArrayDeque<>(16);
    private final int maxSize = 16;
    
    public Map<String, Object> borrow() {
        Map<String, Object> row = pool.pollFirst();
        if (row == null) {
            row = new HashMap<>(32);  // 预设容量
        }
        return row;
    }
    
    public void returnRow(Map<String, Object> row) {
        row.clear();  // 清空数据
        if (pool.size() < maxSize) {
            pool.offerFirst(row);
        }
    }
}

// 使用
public void processRow() {
    Map<String, Object> rowData = rowPool.borrow();
    try {
        // 填充数据
        extractCells(sheet, rowData);
        // 交给用户处理
        handler.invoke(rowData);
    } finally {
        rowPool.returnRow(rowData);  // 归还对象
    }
}
```

### 9.5 内存优化检查清单

在实现时需要考虑:

```
□ 使用 SAX 模式而不是 DOM 模式
□ 及时关闭 InputStream 和 OPCPackage
□ 不保留已处理行的引用 (让用户负责保存需要的数据)
□ 使用对象池复用临时对象
□ 大文件分批处理，避免单次处理过多
□ 共享字符串表懒加载
□ 样式信息按需读取
□ 支持流式写入 (SXSSFWorkbook)
```

### 9.6 EasyExcel 的易用性 API 设计

EasyExcel 之所以流行，除了内存优化，还有简洁的 API 设计：

```java
// EasyExcel 的 API 设计
EasyExcel.read(fileName, DemoData.class, new PageReadListener<DemoData>(dataList -> {
    for (DemoData demoData : dataList) {
        log.info("读取到一条数据:{}", demoData.toString());
    }
})).sheet().doRead();

// 核心设计思想:
// 1. 流式 API (Builder 模式)
// 2. 函数式回调 (Listener)
// 3. 注解映射 (@ExcelProperty)
// 4. 链式调用
```

### 本项目参考设计

```java
// ===== 场景 1: 简单读取 (类似 EasyExcel) =====
// 读取配置，提取数据
Map<String, Object> data = ExcelReader.read(
    new File("order.xlsx"),
    ExcelConfig.fromYaml("config.yaml")
);

// ===== 场景 2: 流式读取大文件 (内存优化) =====
ExcelReader.readStreaming(
    new File("large_order.xlsx"),
    ExcelConfig.fromYaml("config.yaml"),
    new RowHandler() {
        @Override
        public void invoke(int rowNum, Map<String, Object> rowData, boolean isLast) {
            // 处理每一行，处理完立即释放内存
            processRow(rowData);
        }
    }
);

// ===== 场景 3: 填充模板 =====
ExcelFiller.fill(
    new File("template.xlsx"),
    new File("output.xlsx"),
    ExcelConfig.fromYaml("config.yaml"),
    data  // Map<String, Object>
);

// ===== 场景 4: 流式 API 链式调用 =====
ExcelReader.read(
    new File("order.xlsx")
)
.config("config.yaml")
.sheet(0)
.extractMode(ExtractMode.DOWN)
.parser("string")
.addHandler(new TrimHandler())
.addHandler(new NotEmptyHandler())
.doRead();  // 返回 Map<String, Object>
```

### 9.7 易用性设计检查清单

```
□ 流式 API (Builder 模式)
□ 支持链式调用
□ 提供函数式接口 (Lambda 支持)
□ 配置与代码分离 (YAML/JSON)
□ 内置常用解析器和处理器
□ 提供默认值，减少配置
□ 错误信息清晰易懂
□ 提供详细的日志/调试模式
```

### 核心思想借鉴

| 来源 | 借鉴内容 | 本项目应用 |
|------|----------|------------|
| **POI SAX** | 事件驱动、流式解析 | 低内存读取 |
| **EasyExcel** | 分析器模式、懒加载 | 配置驱动的按需解析 |
| **Poiji2** | 注解/配置映射 | Key-Position-Value 配置模型 |
| **策略模式** | SPI 可扩展设计 | 支持自定义提取策略 |

### 差异化优势

1. **配置驱动** - 不需要写代码，改配置即可
2. **按需解析** - 只处理配置的单元格，效率更高
3. **多种提取模式** - SINGLE/DOWN/RIGHT/BLOCK/UNTIL_EMPTY
4. **SPI 拓展** - 支持自定义提取策略
5. **轻量级** - 只依赖 POI，不依赖其他重量级库

### 支持的提取模式总览

| 模式 | 配置示例 | 输出类型 | 适用场景 |
|------|----------|----------|----------|
| SINGLE | `pos: "A1"` | Object | 单一值 |
| DOWN | `pos: "A1", mode: DOWN, rows: 10` | List | 纵向列表 |
| RIGHT | `pos: "A1", mode: RIGHT, cols: 12` | List | 横向列表 (表头) |
| BLOCK | `pos: "A1:C10"` | List<List> | 数据矩阵 |
| UNTIL_EMPTY | `pos: "A1", mode: UNTIL_EMPTY` | List | 动态长度列表 |

### 下一步行动

1. ✅ 搭建项目骨架
2. ✅ 实现基础配置模型 (CellConfig, Position, ExtractMode, ExtractRange)
3. ✅ 实现策略注册表和内置策略 (SINGLE, DOWN, RIGHT, BLOCK, UNTIL_EMPTY)
4. ✅ 实现基于 SAX 的读取器
5. ✅ 实现基于 DOM 的填充器
6. ✅ 编写测试用例
7. ✅ 完善文档

---

**结论**：这个想法完全可行，技术路线清晰。通过 SPI 策略模式设计，可以支持灵活的数据提取方式，用户不仅可以配置位置，还可以配置提取范围、提取方向、终止条件等，满足复杂场景需求。

---

## 十、前后端配合方案 - Univer + 后端配置 ⭐

### 10.1 整体架构

```
┌─────────────────────────────────────────────────────────────────┐
│                        前后端分离架构                            │
├─────────────────────────────────────────────────────────────────┤
│                                                                 │
│  ┌─────────────────┐         HTTP/JSON         ┌─────────────┐ │
│  │   前端 (Univer) │ ←────────────────────────→ │  后端服务   │ │
│  │                 │                           │             │ │
│  │  - Excel 预览    │     1. 上传 Excel 模板      │  - 配置存储  │ │
│  │  - 可视化配置    │     2. 解析返回结构       │  - 配置管理  │ │
│  │  - 单元格选择    │     3. 保存配置           │  - 数据提取  │ │
│  │  - 拖拽映射      │     4. 执行导入/导出      │  - 数据填充  │ │
│  └─────────────────┘                           └─────────────┘ │
│                                                                 │
└─────────────────────────────────────────────────────────────────┘
```

### 10.2 Univer 简介

**Univer** (https://github.com/dream-num/univer) 是阿里开源的在线表格解决方案，类似腾讯文档/Google Sheets。

**核心能力**：
- 完整的在线表格编辑能力
- 支持单元格选择、区域选择
- 支持命名范围
- 丰富的 API 用于数据交互
- 插件系统可扩展

### 10.3 前端配置界面设计

```
┌────────────────────────────────────────────────────────────────┐
│                   Excel 模板配置界面                             │
├────────────────────────────────────────────────────────────────┤
│                                                                │
│  ┌──────────────────┐    ┌─────────────────────────────────┐  │
│  │                  │    │  配置面板                        │  │
│  │   Excel 预览区     │    │                                 │  │
│  │   (Univer)       │    │  模板名称：[订单导入模板____]     │  │
│  │                  │    │                                 │  │
│  │  [A] [B] [C] [D] │    │  字段映射:                       │  │
│  │  ┌───┬───┬───┬───┐│    │  ┌─────────────────────────────┐ │  │
│  │  │订单│金额│日期│备注││    │  │ 字段名  │ 位置  │ 提取模式│  │ │  │
│  │  ├───┼───┼───┼───┤│    │  ├─────────────────────────────┤ │  │
│  │  │A2 │B2 │C2 │D2 ││    │  │ orderNo │ [A2▼] │ [DOWN ▼]│  │ │  │
│  │  │A3 │B3 │C3 │D3 ││    │  │ amount  │ [B2▼] │ [DOWN ▼]│  │ │  │
│  │  │...│...│...│...││    │  │ date    │ [C2▼] │ [DOWN ▼]│  │ │  │
│  │                  │    │  └─────────────────────────────┘ │  │
│  │  [选中区域 A2:D10]│    │                                 │  │
│  │                  │    │  [+ 添加字段]  [保存配置]        │  │
│  └──────────────────┘    └─────────────────────────────────┘  │
│                                                                │
└────────────────────────────────────────────────────────────────┘
```

### 10.4 交互流程

```
1. 用户上传 Excel 模板
   ↓
2. 前端用 Univer 渲染预览
   ↓
3. 用户在表格中选择单元格/区域
   ↓
4. 前端调用 API 获取选中区域信息
   ↓
5. 配置字段名、提取模式等
   ↓
6. 保存配置到后端
   ↓
7. 后端根据配置执行导入/导出
```

### 10.5 后端 API 设计

```java
@RestController
@RequestMapping("/api/excel-config")
public class ExcelConfigController {
    
    // 上传模板并解析结构
    @PostMapping("/upload")
    public ResponseEntity<ExcelStructure> uploadTemplate(
        @RequestParam("file") MultipartFile file
    ) {
        ExcelStructure structure = excelService.analyzeStructure(file);
        return ResponseEntity.ok(structure);
    }
    
    // 保存配置
    @PostMapping("/config")
    public ResponseEntity<Void> saveConfig(
        @RequestBody ExcelTemplateConfigDTO config
    ) {
        configService.save(config);
        return ResponseEntity.ok().build();
    }
    
    // 执行导入
    @PostMapping("/import")
    public ResponseEntity<ImportResult> doImport(
        @RequestParam("file") MultipartFile file,
        @RequestParam("configId") String configId
    ) {
        ImportResult result = excelService.importData(file, configId);
        return ResponseEntity.ok(result);
    }
    
    // 执行导出
    @PostMapping("/export")
    public ResponseEntity<byte[]> doExport(
        @RequestBody ExportRequest request
    ) {
        byte[] excelData = excelService.exportData(request);
        return ResponseEntity.ok()
            .header("Content-Disposition", "attachment; filename=export.xlsx")
            .body(excelData);
    }
}
```

### 10.6 配置数据格式 (前后端通用)

```typescript
// TypeScript (前端) / Java (后端) 共用结构
interface ExcelTemplateConfig {
  templateName: string;
  cells: Record<string, CellConfig>;
}

interface CellConfig {
  key: string;                    // 字段名
  position: {
    cellRef?: string;             // "A1"
    areaRef?: string;             // "A1:C10"
    headerName?: string;          // 表头名
  };
  extractMode: 'SINGLE' | 'DOWN' | 'RIGHT' | 'BLOCK' | 'UNTIL_EMPTY';
  range?: {
    rows?: number;
    cols?: number;
    skipEmpty?: boolean;
  };
  parserType?: string;
  type?: 'STRING' | 'NUMBER' | 'DATE' | 'BOOLEAN';
  required?: boolean;
}
```

### 10.7 项目模块划分

```
excel-config-tool/
├── excel-config-core/           # 核心模块 (后端)
│   └── 策略、解析器、处理器等核心逻辑
│
├── excel-config-web/            # Web 服务模块 (后端)
│   └── Controller、DTO、API 接口
│
└── excel-config-frontend/       # 前端模块
    └── React + Univer + TypeScript
```

### 10.8 技术栈推荐

| 层级 | 技术选型 |
|------|----------|
| **前端框架** | React 18 + TypeScript |
| **表格引擎** | Univer (或 Luckysheet) |
| **UI 组件** | Ant Design |
| **后端框架** | Spring Boot 3.x |
| **Excel 处理** | Apache POI 5.x |

### 10.9 核心价值

| 角色 | 价值 |
|------|------|
| **业务人员** | 可视化配置，不需要写代码 |
| **前端开发** | Univer 提供成熟的表格能力 |
| **后端开发** | 配置驱动，不需要硬编码 |
| **运维** | 配置变更不需要重新发布 |
