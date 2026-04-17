# Excel Config Tool - 导出/填充模式详解

> 导出是提取的逆过程，但有独特的设计考虑

---

## 一、导出模式总览

```
┌─────────────────────────────────────────────────────────────────────────────────┐
│                           导出模式总览                                           │
├─────────────────────────────────────────────────────────────────────────────────┤
│                                                                                 │
│  模式              |  输入数据类型    |  典型场景                               │
│  ─────────────────────────────────────────────────────────────────────────────  │
│  FILL_CELL         |  Object          |  填充单个单元格                          │
│  FILL_DOWN         |  Array           |  向下填充列数据                          │
│  FILL_RIGHT        |  Array           |  向右填充行数据                          │
│  FILL_BLOCK        |  Array<Array>    |  填充区域矩阵                            │
│  FILL_TABLE        |  Array<Object>   |  填充表格 (带表头)                        │
│  APPEND_ROWS       |  Array           |  追加行 (动态扩展)                        │
│  APPEND_COLS       |  Array           |  追加列 (动态扩展)                        │
│  REPLACE_AREA      |  Any             |  替换区域 (先清空再填充)                   │
│  FILL_TEMPLATE     |  Object          |  模板填充 (占位符替换)                     │
│  MULTI_SHEET       |  Object          |  多工作表填充                            │
│                                                                                 │
└─────────────────────────────────────────────────────────────────────────────────┘
```

---

## 二、基础导出模式

### 1. FILL_CELL - 填充单个单元格

#### 配置
```yaml
exports:
  - key: "title"
    position:
      cellRef: "A1"
    mode: FILL_CELL
    parser:
      type: string
```

#### 数据
```json
{
  "title": "2024 年销售报表"
}
```

#### 结果
```
A1 单元格被填充为："2024 年销售报表"
```

---

### 2. FILL_DOWN - 向下填充

#### 配置
```yaml
exports:
  - key: "orderNos"
    position:
      cellRef: "A2"
    mode: FILL_DOWN
    parser:
      type: string
    # 不指定 rows，根据 orderNos 数组长度自动填充
    # 可选：maxRows: 1000 限制最大填充行数
```

#### 数据
```json
{
  "orderNos": ["ORD001", "ORD002", "ORD003", "ORD004"]
}
```

#### 结果
```
┌────────┐
│ A2: ORD001 │
│ A3: ORD002 │
│ A4: ORD003 │
│ A5: ORD004 │
└────────┘
```

#### 动态行数说明
- FILL_DOWN 模式根据**实际数据量**自动计算填充行数
- 如果 `orderNos` 有 100 个元素，则填充 100 行
- 如果 `orderNos` 有 3 个元素，则填充 3 行
- 可选的 `maxRows` 配置用于限制最大填充行数（安全考虑）

---

### 3. FILL_RIGHT - 向右填充

#### 配置
```yaml
exports:
  - key: "monthHeaders"
    position:
      cellRef: "B1"
    mode: FILL_RIGHT
    parser:
      type: string
```

#### 数据
```json
{
  "monthHeaders": ["1 月", "2 月", "3 月", "4 月", "5 月", "6 月"]
}
```

#### 结果
```
┌─────┬─────┬─────┬─────┬─────┬─────┐
│ B1  │ C1  │ D1  │ E1  │ F1  │ G1  │
│ 1 月 │ 2 月 │ 3 月 │ 4 月 │ 5 月 │ 6 月 │
└─────┴─────┴─────┴─────┴─────┴─────┘
```

---

### 4. FILL_BLOCK - 填充区域

#### 配置
```yaml
exports:
  - key: "dataMatrix"
    position:
      areaRef: "B2:E4"
    mode: FILL_BLOCK
    parser:
      type: number
```

#### 数据
```json
{
  "dataMatrix": [
    [100, 120, 150, 130],
    [200, 220, 250, 230],
    [300, 320, 350, 330]
  ]
}
```

#### 结果
```
┌─────┬─────┬─────┬─────┐
│ B2  │ C2  │ D2  │ E2  │
│ 100 │ 120 │ 150 │ 130 │
├─────┼─────┼─────┼─────┤
│ B3  │ C3  │ D3  │ E3  │
│ 200 │ 220 │ 250 │ 230 │
├─────┼─────┼─────┼─────┤
│ B4  │ C4  │ D4  │ E4  │
│ 300 │ 320 │ 350 │ 330 │
└─────┴─────┴─────┴─────┘
```

---

## 三、表格导出模式

### 5. FILL_TABLE - 填充表格（带表头）

#### 配置
```yaml
exports:
  - key: "orders"
    position:
      cellRef: "A1"
    mode: FILL_TABLE
    columns:
      - key: "orderNo"
        header: "订单号"
        type: string
      - key: "amount"
        header: "金额"
        type: number
        format: "#,##0.00"
      - key: "orderDate"
        header: "日期"
        type: date
        format: "yyyy-MM-dd"
      - key: "customerName"
        header: "客户"
        type: string
    options:
      headerStyle:
        bold: true
        background: "#4472C4"
        fontColor: "#FFFFFF"
      alternateRows: true  # 隔行换色
      autoWidth: true      # 自动列宽
```

#### 数据
```json
{
  "orders": [
    { "orderNo": "ORD001", "amount": 100.00, "orderDate": "2024-01-01", "customerName": "张三" },
    { "orderNo": "ORD002", "amount": 200.00, "orderDate": "2024-01-02", "customerName": "李四" },
    { "orderNo": "ORD003", "amount": 300.00, "orderDate": "2024-01-03", "customerName": "王五" }
  ]
}
```

#### 结果
```
┌──────────┬──────────┬────────────┬──────────┐
│ 订单号    │   金额    │    日期     │   客户    │ ← 表头 (加粗，蓝色背景)
├──────────┼──────────┼────────────┼──────────┤
│ ORD001   │ 100.00   │ 2024-01-01 │ 张三      │ ← 隔行换色
│ ORD002   │ 200.00   │ 2024-01-02 │ 李四      │
│ ORD003   │ 300.00   │ 2024-01-03 │ 王五      │
└──────────┴──────────┴────────────┴──────────┘
```

---

### 6. APPEND_ROWS - 追加行

#### 场景
在现有数据下方追加新数据，自动检测追加位置

#### 配置
```yaml
exports:
  - key: "newOrders"
    position:
      cellRef: "A2"    # 从 A2 开始检测
    mode: APPEND_ROWS
    detectEmptyRow: true  # 自动检测第一个空行
    columns:
      - key: "orderNo"
        type: string
      - key: "amount"
        type: number
      - key: "orderDate"
        type: date
```

#### 数据
```json
{
  "newOrders": [
    { "orderNo": "ORD004", "amount": 400.00, "orderDate": "2024-01-04" },
    { "orderNo": "ORD005", "amount": 500.00, "orderDate": "2024-01-05" }
  ]
}
```

#### 结果（追加前）
```
┌──────────┬──────────┬────────────┐
│ 订单号    │   金额    │    日期     │
├──────────┼──────────┼────────────┤
│ ORD001   │ 100.00   │ 2024-01-01 │
│ ORD002   │ 200.00   │ 2024-01-02 │
│ ORD003   │ 300.00   │ 2024-01-03 │
│          │          │            │ ← 空行，从这里开始追加
└──────────┴──────────┴────────────┘
```

#### 结果（追加后）
```
┌──────────┬──────────┬────────────┐
│ 订单号    │   金额    │    日期     │
├──────────┼──────────┼────────────┤
│ ORD001   │ 100.00   │ 2024-01-01 │
│ ORD002   │ 200.00   │ 2024-01-02 │
│ ORD003   │ 300.00   │ 2024-01-03 │
│ ORD004   │ 400.00   │ 2024-01-04 │ ← 新增
│ ORD005   │ 500.00   │ 2024-01-05 │ ← 新增
└──────────┴──────────┴────────────┘
```

---

### 7. APPEND_COLS - 追加列

#### 场景
在现有数据右侧追加新列

#### 配置
```yaml
exports:
  - key: "newMetrics"
    position:
      cellRef: "A1"    # 从 A1 开始检测
    mode: APPEND_COLS
    detectEmptyCol: true
    columns:
      - key: "profit"
        header: "利润"
        type: number
      - key: "profitMargin"
        header: "利润率"
        type: number
        format: "0.00%"
```

#### 结果（追加后）
```
┌──────────┬──────────┬────────────┬──────────┬────────────┐
│ 订单号    │   金额    │    日期     │   利润    │   利润率     │ ← 新增列
├──────────┼──────────┼────────────┼──────────┼────────────┤
│ ORD001   │ 100.00   │ 2024-01-01 │  40.00   │   40.00%   │
│ ORD002   │ 200.00   │ 2024-01-02 │  80.00   │   40.00%   │
└──────────┴──────────┴────────────┴──────────┴────────────┘
```

---

## 四、高级导出模式

### 8. REPLACE_AREA - 替换区域

#### 场景
先清空区域，再填充新数据

#### 配置
```yaml
exports:
  - key: "updatedData"
    position:
      areaRef: "A2:D100"
    mode: REPLACE_AREA
    clearFirst: true   # 先清空
    clearOptions:
      clearValues: true
      clearStyles: false  # 保留样式
      clearFormats: false # 保留格式
    columns:
      - key: "orderNo"
        type: string
      - key: "amount"
        type: number
      - key: "orderDate"
        type: date
      - key: "customer"
        type: string
```

#### 行为
1. 清空 A2:D100 区域的值（保留样式和格式）
2. 填充新数据

---

### 9. FILL_TEMPLATE - 模板填充（占位符替换）

#### 场景
基于模板文件，替换占位符生成最终文档

#### 模板文件内容
```
┌─────────────────────────────────────────┐
│           订单确认函                      │
│                                         │
│  尊敬的 {{customerName}} 先生/女士：       │
│                                         │
│  您的订单 (编号：{{orderNo}}) 已确认。   │
│                                         │
│  订单金额：¥{{amount}}                  │
│  下单日期：{{orderDate}}                │
│                                         │
│  感谢您的惠顾！                          │
└─────────────────────────────────────────┘
```

#### 配置
```yaml
exports:
  - key: "orderInfo"
    mode: FILL_TEMPLATE
    placeholder:
      prefix: "{{"
      suffix: "}}"
    fields:
      - key: "customerName"
        type: string
      - key: "orderNo"
        type: string
      - key: "amount"
        type: number
        format: "#,##0.00"
      - key: "orderDate"
        type: date
        format: "yyyy 年 MM 月 dd 日"
```

#### 数据
```json
{
  "orderInfo": {
    "customerName": "张三",
    "orderNo": "ORD20240101001",
    "amount": 1299.00,
    "orderDate": "2024-01-01"
  }
}
```

#### 结果
```
┌─────────────────────────────────────────┐
│           订单确认函                      │
│                                         │
│  尊敬的 张三 先生/女士：                   │
│                                         │
│  您的订单 (编号：ORD20240101001) 已确认。│
│                                         │
│  订单金额：¥1,299.00                    │
│  下单日期：2024 年 01 月 01 日              │
│                                         │
│  感谢您的惠顾！                          │
└─────────────────────────────────────────┘
```

#### 占位符类型

```
1. 简单占位符：{{fieldName}}
   替换为对应的值

2. 列表占位符：{{#items}}...{{/items}}
   用于循环列表数据

3. 条件占位符：{{#if condition}}...{{/if}}
   根据条件显示/隐藏

4. 计算占位符：{{=field1+field2}}
   支持简单计算
```

#### 列表占位符示例

模板：
```
┌─────────────────────────────────────┐
│  订单明细                             │
│  {{#items}}                         │
│  - {{itemName}}: ¥{{price}}        │
│  {{/items}}                         │
│                                     │
│  总计：¥{{total}}                   │
└─────────────────────────────────────┘
```

配置：
```yaml
exports:
  - key: "orderDetail"
    mode: FILL_TEMPLATE
    fields:
      - key: "items"
        type: array
        itemType: object
        itemFields:
          - key: "itemName"
            type: string
          - key: "price"
            type: number
      - key: "total"
        type: number
```

数据：
```json
{
  "orderDetail": {
    "items": [
      { "itemName": "产品 A", "price": 100 },
      { "itemName": "产品 B", "price": 200 },
      { "itemName": "产品 C", "price": 300 }
    ],
    "total": 600
  }
}
```

结果：
```
┌─────────────────────────────────────┐
│  订单明细                             │
│  - 产品 A: ¥100                      │
│  - 产品 B: ¥200                      │
│  - 产品 C: ¥300                      │
│                                     │
│  总计：¥600                         │
└─────────────────────────────────────┘
```

---

### 10. MULTI_SHEET_FILL - 多工作表填充

#### 场景
同时填充多个工作表

#### 配置
```yaml
exports:
  - key: "monthlyData"
    mode: MULTI_SHEET_FILL
    sheets:
      - name: "1 月"
        areaRef: "A2:D100"
        mode: FILL_TABLE
        columns: [...]
      - name: "2 月"
        areaRef: "A2:D100"
        mode: FILL_TABLE
        columns: [...]
      - name: "3 月"
        areaRef: "A2:D100"
        mode: FILL_TABLE
        columns: [...]
```

#### 数据
```json
{
  "monthlyData": {
    "1 月": [...],
    "2 月": [...],
    "3 月": [...]
  }
}
```

---

## 五、样式和格式配置

### 完整样式配置示例

```yaml
exports:
  - key: "reportData"
    position:
      cellRef: "A1"
    mode: FILL_TABLE
    
    # 列配置
    columns:
      - key: "orderNo"
        header: "订单号"
        width: 15        # 列宽
        style:
          horizontalAlign: CENTER
          verticalAlign: CENTER
      
      - key: "amount"
        header: "金额"
        width: 12
        format: "#,##0.00"
        style:
          horizontalAlign: RIGHT
          format:
            type: CURRENCY
            currencySymbol: "¥"
      
      - key: "orderDate"
        header: "日期"
        width: 12
        format: "yyyy-MM-dd"
        style:
          horizontalAlign: CENTER
      
      - key: "customerName"
        header: "客户"
        width: 10
        style:
          horizontalAlign: LEFT
    
    # 表头样式
    headerStyle:
      bold: true
      fontSize: 12
      fontColor: "#FFFFFF"
      background:
        color: "#4472C4"
        pattern: SOLID
      borders:
        top: { style: THIN, color: "#000000" }
        bottom: { style: THIN, color: "#000000" }
        left: { style: THIN, color: "#CCCCCC" }
        right: { style: THIN, color: "#CCCCCC" }
    
    # 数据行样式
    rowStyle:
      fontSize: 11
      height: 20       # 行高
      borders:
        bottom: { style: HAIR, color: "#EEEEEE" }
    
    # 隔行换色
    alternateRows:
      enabled: true
      evenStyle:
        background: "#F8F8F8"
      oddStyle:
        background: "#FFFFFF"
    
    # 条件格式
    conditionalFormatting:
      - column: "amount"
        rules:
          - type: GREATER_THAN
            value: 10000
            style:
              fontColor: "#FF0000"
              bold: true
          - type: LESS_THAN
            value: 0
            style:
              fontColor: "#FF0000"
              background: "#FFCCCC"
    
    # 自动列宽
    autoWidth:
      enabled: true
      maxWidth: 50
      minWidth: 8
    
    # 冻结窗格
    freeze:
      rows: 1    # 冻结第 1 行 (表头)
      cols: 0    # 不冻结列
    
    # 打印设置
    printSettings:
      fitToPage: true
      orientation: LANDSCAPE
      paperSize: A4
      headerRows: 1   # 每页重复表头
```

---

## 六、导出模式对比表

| 模式 | 输入类型 | 是否覆盖 | 是否扩展 | 典型场景 |
|------|----------|----------|----------|----------|
| FILL_CELL | Object | ✓ | ✗ | 填充单个值 |
| FILL_DOWN | Array | ✓ | ✗ | 填充列数据 |
| FILL_RIGHT | Array | ✓ | ✗ | 填充行数据 |
| FILL_BLOCK | Array<Array> | ✓ | ✗ | 填充矩阵 |
| FILL_TABLE | Array<Object> | ✓ | ✗ | 填充表格 |
| APPEND_ROWS | Array | ✗ | ✓ | 追加行 |
| APPEND_COLS | Array | ✗ | ✓ | 追加列 |
| REPLACE_AREA | Any | ✓ | ✗ | 替换区域 |
| FILL_TEMPLATE | Object | ✓ | ✗ | 模板填充 |
| MULTI_SHEET | Object | ✓ | ✗ | 多工作表 |

---

## 七、导出配置完整示例

```yaml
version: "1.0"
templateName: "销售报表导出配置"

# 模板文件路径
templateFile: "templates/sales_report.xlsx"

# 导出配置
exports:
  # 1. 填充报表标题
  - key: "reportTitle"
    position:
      cellRef: "A1"
    mode: FILL_CELL
    parser:
      type: string
    style:
      bold: true
      fontSize: 18
      horizontalAlign: CENTER
  
  # 2. 填充报表日期
  - key: "reportDate"
    position:
      cellRef: "A2"
    mode: FILL_CELL
    parser:
      type: date
      format: "yyyy 年 MM 月 dd 日"
  
  # 3. 填充表格数据
  - key: "salesData"
    position:
      cellRef: "A4"
    mode: FILL_TABLE
    columns:
      - key: "orderNo"
        header: "订单号"
        width: 15
      - key: "customerName"
        header: "客户"
        width: 12
      - key: "amount"
        header: "金额"
        width: 12
        format: "#,##0.00"
      - key: "orderDate"
        header: "日期"
        width: 12
        format: "yyyy-MM-dd"
      - key: "status"
        header: "状态"
        width: 10
        style:
          horizontalAlign: CENTER
    headerStyle:
      bold: true
      background: "#4472C4"
      fontColor: "#FFFFFF"
    alternateRows: true
    conditionalFormatting:
      - column: "status"
        rules:
          - type: EQUALS
            value: "已完成"
            style:
              fontColor: "#00AA00"
          - type: EQUALS
            value: "进行中"
            style:
              fontColor: "#FF9900"
          - type: EQUALS
            value: "已取消"
            style:
              fontColor: "#CC0000"
  
  # 4. 填充汇总数据
  - key: "summary"
    position:
      cellRef: "A100"
    mode: FILL_TABLE
    columns:
      - key: "metric"
        header: "指标"
      - key: "value"
        header: "数值"
        format: "#,##0.00"
    style:
      bold: true

# 输出配置
output:
  format: XLSX
  fileName: "销售报表_${reportDate}.xlsx"
  autoSizeColumns: true
  freezeTopRow: 1
```

---

## 八、导出模式与提取模式对应关系

```
┌─────────────────────────────────────────────────────────────────┐
│                    导出模式与提取模式对应关系                    │
├─────────────────────────────────────────────────────────────────┤
│                                                                 │
│  提取模式          →     导出模式                               │
│  ───────────────────────────────────────────────────────────    │
│  SINGLE            →     FILL_CELL                              │
│  DOWN              →     FILL_DOWN / APPEND_ROWS                │
│  RIGHT             →     FILL_RIGHT / APPEND_COLS               │
│  BLOCK             →     FILL_BLOCK                             │
│  TABLE             →     FILL_TABLE                             │
│  KEY_VALUE         →     FILL_TEMPLATE (键值替换)                │
│  CROSS_TAB         →     FILL_BLOCK + 样式                       │
│  MULTI_SHEET       →     MULTI_SHEET_FILL                       │
│                                                                 │
│  特有导出模式：                                                  │
│  • REPLACE_AREA    - 替换区域                                   │
│  • FILL_TEMPLATE   - 模板填充 (占位符)                            │
│  • APPEND_ROWS     - 追加行                                     │
│  • APPEND_COLS     - 追加列                                     │
│                                                                 │
└─────────────────────────────────────────────────────────────────┘
```

---

## 九、实现建议

### 策略接口

```java
public interface FillStrategy {
    /**
     * 执行填充
     */
    void fill(Workbook workbook, FillContext context);
    
    /**
     * 支持的填充模式
     */
    Set<FillMode> supportedModes();
}
```

### 填充引擎

```java
public class FillEngine {
    
    // 基于模板填充
    public byte[] fillTemplate(InputStream template, 
                                Map<String, Object> data, 
                                ExcelConfig config) {
        Workbook workbook = WorkbookFactory.create(template);
        
        for (ExportConfig export : config.getExports()) {
            FillStrategy strategy = getStrategy(export.getMode());
            FillContext context = new FillContext(workbook, export, data);
            strategy.fill(workbook, context);
        }
        
        ByteArrayOutputStream output = new ByteArrayOutputStream();
        workbook.write(output);
        return output.toByteArray();
    }
    
    // 生成新 Excel
    public byte[] generate(List<List<Object>> data, 
                           ExcelConfig config) {
        Workbook workbook = new XSSFWorkbook();
        // ... 类似逻辑
    }
}
```

---

## 十、总结

### 导出模式优先级

| 优先级 | 模式 | 理由 |
|--------|------|------|
| P0 | FILL_CELL | 最基础 |
| P0 | FILL_TABLE | 最常用 |
| P0 | FILL_TEMPLATE | 模板场景多 |
| P1 | FILL_DOWN/RIGHT | 列表填充 |
| P1 | APPEND_ROWS/COLS | 动态扩展 |
| P1 | REPLACE_AREA | 替换数据 |
| P2 | FILL_BLOCK | 矩阵填充 |
| P2 | MULTI_SHEET | 多 sheet 场景 |

### 核心特性

1. **样式支持**: 完整的样式配置能力
2. **条件格式**: 根据值自动应用样式
3. **模板填充**: 占位符替换，支持循环
4. **动态扩展**: 自动追加行/列
5. **多工作表**: 同时填充多个 sheet
