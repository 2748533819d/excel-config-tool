# Excel Config Tool - 导出/填充模式详解

> 导出是提取的逆过程，支持多种填充模式

---

## 一、导出模式总览

### 1.1 基础模式

| 模式 | 说明 | 输入数据类型 | 典型场景 |
|------|------|-------------|----------|
| `FILL_CELL` | 填充单个单元格 | Object | 标题、汇总值 |
| `FILL_DOWN` | 向下填充列数据 | Array | 订单号列表 |
| `FILL_RIGHT` | 向右填充行数据 | Array | 表头行 |
| `FILL_BLOCK` | 填充区域矩阵 | Array<Array> | 数据矩阵 |

### 1.2 表格模式

| 模式 | 说明 | 输入数据类型 | 典型场景 |
|------|------|-------------|----------|
| `FILL_TABLE` | 填充表格（带表头） | Array<Object> | 订单明细表 |
| `APPEND_ROWS` | 追加行 | Array | 日志追加 |
| `APPEND_COLS` | 追加列 | Array | 动态列 |

### 1.3 高级模式

| 模式 | 说明 | 输入数据类型 | 典型场景 |
|------|------|-------------|----------|
| `REPLACE_AREA` | 替换区域 | Any | 重置区域 |
| `FILL_TEMPLATE` | 模板填充（占位符） | Object | 合同、发票 |
| `MULTI_SHEET_FILL` | 多工作表填充 | Object | 分月报表 |

---

## 二、基础填充模式

### 2.1 FILL_CELL - 填充单个单元格

#### 配置

```json
{
  "exports": [
    {
      "key": "title",
      "header": { "match": "报表标题" },
      "mode": "FILL_CELL"
    }
  ]
}
```

#### 数据

```json
{
  "title": "2024 年销售报表"
}
```

#### 结果

```
┌─────────────────────────┐
│  2024 年销售报表          │ ← 填充后的值
└─────────────────────────┘
```

---

### 2.2 FILL_DOWN - 向下填充列数据

#### 配置

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

#### 数据

```json
{
  "orderNos": ["ORD001", "ORD002", "ORD003", "ORD004"]
}
```

#### 结果

```
┌─────────────┐
│ 订单号      │ ← 表头
├─────────────┤
│ ORD001      │
│ ORD002      │
│ ORD003      │
│ ORD004      │
└─────────────┘
```

---

### 2.3 FILL_RIGHT - 向右填充行数据

#### 配置

```json
{
  "exports": [
    {
      "key": "headers",
      "header": { "match": "月份" },
      "mode": "FILL_RIGHT"
    }
  ]
}
```

#### 数据

```json
{
  "headers": ["1 月", "2 月", "3 月", "4 月"]
}
```

#### 结果

```
┌─────┬─────┬─────┬─────┐
│ 1 月 │ 2 月 │ 3 月 │ 4 月 │
└─────┴─────┴─────┴─────┘
```

---

### 2.4 FILL_BLOCK - 填充区域矩阵

#### 配置

```json
{
  "exports": [
    {
      "key": "matrix",
      "header": { "match": "数据区" },
      "mode": "FILL_BLOCK"
    }
  ]
}
```

#### 数据

```json
{
  "matrix": [
    [1, 2, 3],
    [4, 5, 6],
    [7, 8, 9]
  ]
}
```

#### 结果

```
┌───┬───┬───┐
│ 1 │ 2 │ 3 │
├───┼───┼───┤
│ 4 │ 5 │ 6 │
├───┼───┼───┤
│ 7 │ 8 │ 9 │
└───┴───┴───┘
```

---

## 三、表格填充模式

### 3.1 FILL_TABLE - 填充表格

#### 配置

```json
{
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
        "fontColor": "#FFFFFF"
      },
      "alternateRows": true,
      "autoWidth": true
    }
  ]
}
```

#### 数据

```json
{
  "orders": [
    { "orderNo": "ORD001", "amount": 100.00, "date": "2024-01-01" },
    { "orderNo": "ORD002", "amount": 200.00, "date": "2024-01-02" },
    { "orderNo": "ORD003", "amount": 150.00, "date": "2024-01-03" }
  ]
}
```

#### 结果

```
┌───────────────┬────────────┬──────────────┐
│   订单号      │    金额    │     日期     │ ← 表头（蓝色背景，白色粗体）
├───────────────┼────────────┼──────────────┤
│ ORD001        │   100.00   │ 2024-01-01   │ ← 隔行换色
├───────────────┼────────────┼──────────────┤
│ ORD002        │   200.00   │ 2024-01-02   │
├───────────────┼────────────┼──────────────┤
│ ORD003        │   150.00   │ 2024-01-03   │
└───────────────┴────────────┴──────────────┘
```

---

### 3.2 APPEND_ROWS - 追加行

#### 配置

```json
{
  "exports": [
    {
      "key": "logs",
      "header": { "match": "日志" },
      "mode": "APPEND_ROWS"
    }
  ]
}
```

#### 数据

```json
{
  "logs": ["日志 1", "日志 2", "日志 3"]
}
```

#### 结果

```
原始：
┌─────────────┐
│ 日志        │
├─────────────┤
│ 原有日志     │
└─────────────┘

填充后：
┌─────────────┐
│ 日志        │
├─────────────┤
│ 原有日志     │
│ 日志 1       │ ← 追加
│ 日志 2       │ ← 追加
│ 日志 3       │ ← 追加
└─────────────┘
```

---

## 四、高级填充模式

### 4.1 REPLACE_AREA - 替换区域

#### 配置

```json
{
  "exports": [
    {
      "key": "reportData",
      "position": { "areaRef": "A2:D10" },
      "mode": "REPLACE_AREA"
    }
  ]
}
```

#### 说明

先清空指定区域，再填充数据。适用于需要完全重置的场景。

---

### 4.2 FILL_TEMPLATE - 模板填充（占位符替换）

#### 配置

```json
{
  "exports": [
    {
      "key": "invoice",
      "mode": "FILL_TEMPLATE",
      "placeholder": {
        "prefix": "{{",
        "suffix": "}}"
      }
    }
  ]
}
```

#### 模板

```
发票号：{{invoiceNo}}
客户：{{customer}}
金额：{{amount}}
```

#### 数据

```json
{
  "invoice": {
    "invoiceNo": "INV-2024-001",
    "customer": "某某公司",
    "amount": 1000.00
  }
}
```

#### 结果

```
发票号：INV-2024-001
客户：某某公司
金额：1000.00
```

---

## 五、样式配置

### 5.1 StyleConfig

| 字段 | 类型 | 说明 |
|------|------|------|
| format | String | 数字格式（如 `#,##0.00`、`yyyy-MM-dd`） |
| horizontalAlign | String | 水平对齐（LEFT/CENTER/RIGHT） |
| verticalAlign | String | 垂直对齐（TOP/CENTER/BOTTOM） |
| bold | Boolean | 是否加粗 |
| italic | Boolean | 是否倾斜 |
| background | String | 背景色（十六进制，如 `#4472C4`） |
| fontColor | String | 字体色（十六进制，如 `#FFFFFF`） |
| fontSize | Integer | 字体大小 |
| borderBottom | String | 下边框样式 |
| borderTop | String | 上边框样式 |
| borderLeft | String | 左边框样式 |
| borderRight | String | 右边框样式 |

### 5.2 样式示例

```json
{
  "exports": [
    {
      "key": "amounts",
      "header": { "match": "金额" },
      "mode": "FILL_DOWN",
      "style": {
        "format": "#,##0.00",
        "horizontalAlign": "RIGHT",
        "borderBottom": "THIN"
      },
      "headerStyle": {
        "bold": true,
        "background": "#4472C4",
        "fontColor": "#FFFFFF",
        "horizontalAlign": "CENTER"
      }
    }
  ]
}
```

---

## 六、最佳实践

### 6.1 选择填充模式

```
根据数据结构选择模式：

• 单个值        → FILL_CELL
• 一维数组      → FILL_DOWN 或 FILL_RIGHT
• 二维数组      → FILL_BLOCK
• 对象数组      → FILL_TABLE
• 追加数据      → APPEND_ROWS
• 占位符模板    → FILL_TEMPLATE
```

### 6.2 推荐配置

```json
// ✅ 推荐：使用表头匹配 + 自动扩展
{
  "exports": [
    {
      "key": "orders",
      "header": { "match": "订单号" },
      "mode": "FILL_TABLE",
      "columns": [...],
      "autoWidth": true,
      "alternateRows": true
    }
  ]
}

// ❌ 不推荐：硬编码位置
{
  "exports": [
    {
      "key": "orders",
      "position": { "cellRef": "A2" },
      "mode": "FILL_DOWN",
      "maxRows": 100
    }
  ]
}
```

---

## 七、常见问题

### Q: 如何设置列宽？

**A:** 在 columns 配置中设置 width：

```json
{
  "columns": [
    {
      "key": "orderNo",
      "header": "订单号",
      "width": 15  // Excel 列宽单位
    }
  ]
}
```

或者使用 autoWidth：

```json
{
  "autoWidth": true  // 自动调整列宽
}
```

### Q: 如何设置数字格式？

**A:** 在 style 配置中设置 format：

```json
{
  "style": {
    "format": "#,##0.00"  // 千分位，两位小数
  }
}
```

### Q: 如何实现隔行换色？

**A:** 设置 alternateRows 为 true：

```json
{
  "alternateRows": true  // 隔行换色
}
```

---

## 八、参考资料

- [FillEngine 实现](../packages/core/src/main/java/com/excelconfig/export/FillEngine.java)
- [提取模式详解](./EXTRACT_MODES.md)
- [列隔离机制](./COLUMN_ISOLATION.md)
