# Excel Config Tool 使用示例

## 示例 1：基础数据提取

### 场景
从 Excel 表格中提取"订单号"列的数据。

### Excel 模板
```
| 订单号 | 金额   | 日期       |
|--------|--------|------------|
| ORD001 | 100.00 | 2024-01-01 |
| ORD002 | 200.00 | 2024-01-02 |
| ORD003 | 150.00 | 2024-01-03 |
```

### 配置文件 (extract-config.json)
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
      "mode": "DOWN",
      "range": { "skipEmpty": true }
    }
  ]
}
```

### Java 代码
```java
// 创建引擎
ExtractEngine engine = new ExtractEngine();

// 加载配置
String configJson = Files.readString(Paths.get("extract-config.json"));
JsonConfigParser parser = new JsonConfigParser();
ExcelConfig config = parser.parse(configJson);

// 执行提取
Map<String, Object> result = engine.extract(
    new FileInputStream("template.xlsx"),
    config
);

// 获取结果
List<Object> orderNos = (List<Object>) result.get("orderNos");
// [ORD001, ORD002, ORD003]
```

---

## 示例 2：数据填充（FILL_DOWN 模式）

### 场景
将订单号列表填充到 Excel 模板中。

### Excel 模板
```
| 订单号 |
|--------|
|        |
|        |
```

### 配置文件 (fill-config.json)
```json
{
  "version": "1.0",
  "exports": [
    {
      "key": "orderNos",
      "header": { "match": "订单号" },
      "mode": "FILL_DOWN"
    }
  ]
}
```

### Java 代码
```java
// 准备数据
Map<String, Object> data = new HashMap<>();
data.put("orderNos", Arrays.asList("ORD001", "ORD002", "ORD003", "ORD004"));

// 执行填充
FillEngine engine = new FillEngine();
byte[] result = engine.fill(
    new FileInputStream("template.xlsx"),
    data,
    config
);

// 写入文件
Files.write(Paths.get("output.xlsx"), result);
```

### 输出结果
```
| 订单号 |
|--------|
| ORD001 |
| ORD002 |
| ORD003 |
| ORD004 |
```

---

## 示例 3：表格填充（FILL_TABLE 模式）

### 场景
填充完整的订单表格，包括表头和多列数据。

### Excel 模板（只需要有表头）
```
| 订单号 |
|--------|
```

### 配置文件 (table-config.json)
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
          "key": "date",
          "header": "日期",
          "width": 15
        }
      ],
      "headerStyle": {
        "bold": true,
        "background": "#4472C4",
        "horizontalAlign": "CENTER"
      },
      "alternateRows": true,
      "autoWidth": true
    }
  ]
}
```

### Java 代码
```java
// 准备数据
Map<String, Object> data = new HashMap<>();
List<Map<String, Object>> orders = Arrays.asList(
    Map.of("orderNo", "ORD001", "amount", 100.00, "date", new Date()),
    Map.of("orderNo", "ORD002", "amount", 200.00, "date", new Date()),
    Map.of("orderNo", "ORD003", "amount", 150.00, "date", new Date())
);
data.put("orders", orders);

// 执行填充
FillEngine engine = new FillEngine();
byte[] result = engine.fill(
    new FileInputStream("template.xlsx"),
    data,
    config
);
```

### 输出结果
```
| 订单号 | 金额     | 日期       |
|--------|----------|------------|
| ORD001 | 100.00   | 2024-01-01 |
| ORD002 | 200.00   | 2024-01-02 |
| ORD003 | 150.00   | 2024-01-03 |
```

---

## 示例 4：Spring Boot 集成

### 添加依赖
```xml
<dependency>
    <groupId>com.excelconfig</groupId>
    <artifactId>excel-config-spring-boot-starter</artifactId>
    <version>1.0.0-SNAPSHOT</version>
</dependency>
```

### 配置文件 (application.yml)
```yaml
excel:
  config:
    enabled: true
    template-location: classpath:templates/
    output-location: classpath:output/
```

### Service 类
```java
@Service
public class OrderService {
    
    @Autowired
    private ExcelConfigService excelConfigService;
    
    public byte[] exportOrders() throws IOException {
        // 加载配置
        ExcelConfig config = excelConfigService.loadConfig(
            "classpath:config/order-export.json"
        );
        
        // 准备数据
        Map<String, Object> data = new HashMap<>();
        data.put("orders", orderRepository.findAll());
        
        // 执行填充
        try (InputStream template = getClass()
            .getResourceAsStream("/templates/order-template.xlsx")) {
            return excelConfigService.fill(config, data, template);
        }
    }
    
    public List<Object> importOrders(MultipartFile file) throws IOException {
        // 加载配置
        ExcelConfig config = excelConfigService.loadConfig(
            "classpath:config/order-import.json"
        );
        
        // 执行提取
        return excelConfigService.extract(config, file.getInputStream())
            .get("orders");
    }
}
```

---

## 示例 5：复杂场景 - 多表格处理

### 场景
一个 Excel 文件包含多个表格（订单表、明细表），需要分别处理。

### Excel 模板
```
Sheet: 订单
| 订单号 | 客户   | 日期       |
|--------|--------|------------|
| ORD001 | 客户 A | 2024-01-01 |

Sheet: 明细
| 订单号 | 商品   | 数量 | 单价   |
|--------|--------|------|--------|
| ORD001 | 商品 A | 10   | 10.00  |
```

### 配置文件
```json
{
  "version": "1.0",
  "extractions": [
    {
      "key": "orders",
      "header": { "match": "订单号" },
      "mode": "DOWN"
    },
    {
      "key": "details",
      "header": { "match": "商品" },
      "mode": "DOWN"
    }
  ],
  "exports": [
    {
      "key": "orders",
      "header": { "match": "订单号" },
      "mode": "FILL_TABLE",
      "columns": [
        {"key": "orderNo", "header": "订单号"},
        {"key": "customer", "header": "客户"},
        {"key": "date", "header": "日期"}
      ]
    },
    {
      "key": "details",
      "header": { "match": "订单号" },
      "mode": "FILL_TABLE",
      "columns": [
        {"key": "orderNo", "header": "订单号"},
        {"key": "product", "header": "商品"},
        {"key": "quantity", "header": "数量"},
        {"key": "price", "header": "单价", "format": "#,##0.00"}
      ]
    }
  ]
}
```

### Java 代码
```java
// 提取数据
Map<String, Object> extracted = engine.extract(inputStream, config);
List<Object> orders = (List<Object>) extracted.get("orders");
List<Object> details = (List<Object>) extracted.get("details");

// 处理数据...

// 填充数据
byte[] result = engine.fill(templateInputStream, processedData, config);
```

---

## 配置参考

### ExtractConfig 配置项

| 字段 | 类型 | 必填 | 说明 |
|------|------|------|------|
| key | String | 是 | 数据键名 |
| header | HeaderConfig | 是 | 表头配置 |
| mode | String | 是 | 提取模式 (SINGLE/DOWN/RIGHT/BLOCK/UNTIL_EMPTY) |
| range | RangeConfig | 否 | 范围配置 |
| maxRows | Integer | 否 | 最大行数 |
| skipEmpty | Boolean | 否 | 是否跳过空行 |

### ExportConfig 配置项

| 字段 | 类型 | 必填 | 说明 |
|------|------|------|------|
| key | String | 是 | 数据键名 |
| header | HeaderConfig | 是 | 表头配置 |
| mode | String | 是 | 填充模式 (FILL_CELL/FILL_DOWN/FILL_TABLE) |
| columns | List<ColumnConfig> | 否 | 列配置（FILL_TABLE 模式需要） |
| headerStyle | StyleConfig | 否 | 表头样式 |
| alternateRows | Boolean | 否 | 是否隔行换色 |
| autoWidth | Boolean | 否 | 是否自动列宽 |

### ColumnConfig 配置项

| 字段 | 类型 | 必填 | 说明 |
|------|------|------|------|
| key | String | 是 | 数据字段名 |
| header | String | 是 | 列头显示文字 |
| width | Integer | 否 | 列宽 |
| format | String | 否 | 数字格式（如 "#,##0.00"） |

### StyleConfig 配置项

| 字段 | 类型 | 必填 | 说明 |
|------|------|------|------|
| bold | Boolean | 否 | 是否加粗 |
| background | String | 否 | 背景色（十六进制，如 "#4472C4"） |
| horizontalAlign | String | 否 | 水平对齐 (LEFT/CENTER/RIGHT) |
