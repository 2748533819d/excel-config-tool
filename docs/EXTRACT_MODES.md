# Excel Config Tool - 提取模式详解

> 详细说明每种提取模式的配置方式、使用场景和注意事项

---

## 一、提取模式总览

| 模式 | 说明 | 输入示例 | 输出类型 | 典型场景 |
|------|------|----------|----------|----------|
| `SINGLE` | 单个单元格 | [A1] | Object | 标题、汇总值 |
| `DOWN` | 向下提取列 | [A2,A3,A4] | Array | 订单号列表 |
| `RIGHT` | 向右提取行 | [A1,B1,C1] | Array | 表头、月份 |
| `BLOCK` | 区域矩阵 | [A1:C10] | Array<Array> | 数据矩阵 |
| `UNTIL_EMPTY` | 提取到空行 | [A2,A3...∅] | Array | 动态列表 |

---

## 二、基础提取模式

### 2.1 SINGLE - 单个单元格

#### 配置

```json
{
  "extractions": [
    {
      "key": "title",
      "header": { "match": "报表标题" },
      "mode": "SINGLE"
    }
  ]
}
```

#### Excel

```
┌─────────────────────────┐
│  2024 年销售报表          │ ← A1 (title)
└─────────────────────────┘
```

#### 结果

```json
{
  "title": "2024 年销售报表"
}
```

---

### 2.2 DOWN - 向下提取列

#### 配置

```json
{
  "extractions": [
    {
      "key": "orderNos",
      "header": { "match": "订单号" },
      "mode": "DOWN",
      "range": { "skipEmpty": true }
    }
  ]
}
```

#### Excel

```
┌─────────────┐
│ 订单号      │ ← 表头
├─────────────┤
│ ORD001      │
│ ORD002      │
│ ORD003      │
└─────────────┘
```

#### 结果

```json
{
  "orderNos": ["ORD001", "ORD002", "ORD003"]
}
```

---

### 2.3 RIGHT - 向右提取行

#### 配置

```json
{
  "extractions": [
    {
      "key": "months",
      "header": { "match": "月份" },
      "mode": "RIGHT"
    }
  ]
}
```

#### Excel

```
┌─────┬─────┬─────┬─────┐
│ 1 月 │ 2 月 │ 3 月 │ 4 月 │ ← 提取这一行
└─────┴─────┴─────┴─────┘
```

#### 结果

```json
{
  "months": ["1 月", "2 月", "3 月", "4 月"]
}
```

---

### 2.4 BLOCK - 区域矩阵

#### 配置

```json
{
  "extractions": [
    {
      "key": "matrix",
      "header": { "match": "数据区" },
      "mode": "BLOCK",
      "range": { "rows": 3, "cols": 3 }
    }
  ]
}
```

#### Excel

```
┌───┬───┬───┐
│ 1 │ 2 │ 3 │
├───┼───┼───┤
│ 4 │ 5 │ 6 │
├───┼───┼───┤
│ 7 │ 8 │ 9 │
└───┴───┴───┘
```

#### 结果

```json
{
  "matrix": [
    [1, 2, 3],
    [4, 5, 6],
    [7, 8, 9]
  ]
}
```

---

### 2.5 UNTIL_EMPTY - 提取到空行

#### 配置

```json
{
  "extractions": [
    {
      "key": "logs",
      "header": { "match": "日志" },
      "mode": "UNTIL_EMPTY"
    }
  ]
}
```

#### Excel

```
┌─────────────┐
│ 日志        │ ← 表头
├─────────────┤
│ 日志行 1     │
│ 日志行 2     │
│ 日志行 3     │
│             │ ← 空行，停止
├─────────────┤
│ 其他内容    │ ← 不提取
└─────────────┘
```

#### 结果

```json
{
  "logs": ["日志行 1", "日志行 2", "日志行 3"]
}
```

---

## 三、扩展提取模式

### 3.1 KEY_VALUE - 键值对提取

#### 配置

```json
{
  "extractions": [
    {
      "key": "config",
      "header": { "match": "配置项" },
      "mode": "KEY_VALUE"
    }
  ]
}
```

#### Excel

```
┌───────────┬───────────┐
│ 配置项    │ 值        │
├───────────┼───────────┤
│ timeout   │ 30        │
│ retries   │ 3         │
│ debug     │ true      │
└───────────┴───────────┘
```

#### 结果

```json
{
  "config": {
    "timeout": "30",
    "retries": "3",
    "debug": "true"
  }
}
```

---

### 3.2 TABLE - 表格提取

#### 配置

```json
{
  "extractions": [
    {
      "key": "orders",
      "header": { "match": "订单号" },
      "mode": "TABLE"
    }
  ]
}
```

#### Excel

```
┌───────────┬────────┬────────────┐
│ 订单号    │ 金额   │ 日期       │
├───────────┼────────┼────────────┤
│ ORD001    │ 100    │ 2024-01-01 │
│ ORD002    │ 200    │ 2024-01-02 │
│ ORD003    │ 150    │ 2024-01-03 │
└───────────┴────────┴────────────┘
```

#### 结果

```json
{
  "orders": [
    { "订单号": "ORD001", "金额": 100, "日期": "2024-01-01" },
    { "订单号": "ORD002", "金额": 200, "日期": "2024-01-02" },
    { "订单号": "ORD003", "金额": 150, "日期": "2024-01-03" }
  ]
}
```

---

### 3.3 CROSS_TAB - 交叉表提取

#### 配置

```json
{
  "extractions": [
    {
      "key": "sales",
      "header": { "match": "月份\\城市" },
      "mode": "CROSS_TAB"
    }
  ]
}
```

#### Excel

```
┌─────────┬───────┬───────┬───────┐
│ 月份\城市│ 北京  │ 上海  │ 广州  │
├─────────┼───────┼───────┼───────┤
│ 1 月     │ 100   │ 150   │ 120   │
│ 2 月     │ 110   │ 160   │ 130   │
│ 3 月     │ 120   │ 170   │ 140   │
└─────────┴───────┴───────┴───────┘
```

#### 结果

```json
{
  "sales": {
    "北京": { "1 月": 100, "2 月": 110, "3 月": 120 },
    "上海": { "1 月": 150, "2 月": 160, "3 月": 170 },
    "广州": { "1 月": 120, "2 月": 130, "3 月": 140 }
  }
}
```

---

### 3.4 GROUPED - 分组提取

#### 配置

```json
{
  "extractions": [
    {
      "key": "ordersByRegion",
      "header": { "match": "区域" },
      "mode": "GROUPED"
    }
  ]
}
```

#### Excel

```
┌─────────┬───────────┬────────┐
│ 区域    │ 订单号    │ 金额   │
├─────────┼───────────┼────────┤
│ 华东    │ ORD001    │ 100    │
│ 华东    │ ORD002    │ 200    │
│ 华北    │ ORD003    │ 150    │
│ 华北    │ ORD004    │ 180    │
│ 华南    │ ORD005    │ 120    │
└─────────┴───────────┴────────┘
```

#### 结果

```json
{
  "ordersByRegion": {
    "华东": [
      { "订单号": "ORD001", "金额": 100 },
      { "订单号": "ORD002", "金额": 200 }
    ],
    "华北": [
      { "订单号": "ORD003", "金额": 150 },
      { "订单号": "ORD004", "金额": 180 }
    ],
    "华南": [
      { "订单号": "ORD005", "金额": 120 }
    ]
  }
}
```

---

### 3.5 HIERARCHY - 层级提取

#### 配置

```json
{
  "extractions": [
    {
      "key": "categories",
      "header": { "match": "分类" },
      "mode": "HIERARCHY"
    }
  ]
}
```

#### Excel

```
┌───────────────┬───────────┐
│ 分类          │ 金额      │
├───────────────┼───────────┤
│ 电子产品      │           │ ← 父节点
│ ├─手机        │ 5000      │
│ ├─电脑        │ 8000      │
│ 服装          │           │ ← 父节点
│ ├─男装        │ 3000      │
│ └─女装        │ 4000      │
└───────────────┴───────────┘
```

#### 结果

```json
{
  "categories": [
    {
      "name": "电子产品",
      "children": [
        { "name": "手机", "amount": 5000 },
        { "name": "电脑", "amount": 8000 }
      ]
    },
    {
      "name": "服装",
      "children": [
        { "name": "男装", "amount": 3000 },
        { "name": "女装", "amount": 4000 }
      ]
    }
  ]
}
```

---

### 3.6 MERGED_CELLS - 合并单元格提取

#### 配置

```json
{
  "extractions": [
    {
      "key": "report",
      "header": { "match": "报表" },
      "mode": "MERGED_CELLS"
    }
  ]
}
```

#### Excel

```
┌───────────────────────┬───────────┐
│         报表          │ 2024 Q1   │ ← 合并单元格
├───────────────────────┼───────────┤
│  订单汇总             │           │
│  ├─华东区             │ 10000     │
│  └─华北区             │ 8000      │
└───────────────────────┴───────────┘
```

#### 结果

```json
{
  "report": {
    "title": "2024 Q1",
    "sections": [
      {
        "name": "订单汇总",
        "items": [
          { "region": "华东区", "amount": 10000 },
          { "region": "华北区", "amount": 8000 }
        ]
      }
    ]
  }
}
```

---

### 3.7 MULTI_SHEET - 多工作表提取

#### 配置

```json
{
  "extractions": [
    {
      "key": "monthlySales",
      "sheet": "*",
      "header": { "match": "销售额" },
      "mode": "MULTI_SHEET"
    }
  ]
}
```

#### Excel（多个 Sheet）

```
Sheet: 1 月
┌───────────┐
│ 销售额    │
├───────────┤
│ 100       │
└───────────┘

Sheet: 2 月
┌───────────┐
│ 销售额    │
├───────────┤
│ 150       │
└───────────┘
```

#### 结果

```json
{
  "monthlySales": {
    "1 月": [100],
    "2 月": [150]
  }
}
```

---

## 四、范围配置（RangeConfig）

### 4.1 配置项

| 字段 | 类型 | 说明 |
|------|------|------|
| skipEmpty | Boolean | 是否跳过空行 |
| startRow | Integer | 起始行（从 0 开始） |
| endRow | Integer | 结束行 |
| maxRows | Integer | 最大行数 |
| rows | Integer | 固定行数 |
| cols | Integer | 固定列数（RIGHT/BLOCK 模式） |

### 4.2 示例

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
      "key": "headers",
      "header": { "match": "月份" },
      "mode": "RIGHT",
      "range": {
        "cols": 12
      }
    }
  ]
}
```

---

## 五、解析器配置（ParserConfig）

### 5.1 配置项

| 字段 | 类型 | 说明 |
|------|------|------|
| type | String | 解析器类型（string/number/date/boolean） |
| format | String | 格式（数字格式或日期格式） |
| scale | Integer | 小数位数（number 类型） |

### 5.2 示例

```json
{
  "extractions": [
    {
      "key": "amounts",
      "header": { "match": "金额" },
      "mode": "DOWN",
      "parser": {
        "type": "number",
        "format": "#,##0.00",
        "scale": 2
      }
    },
    {
      "key": "dates",
      "header": { "match": "日期" },
      "mode": "DOWN",
      "parser": {
        "type": "date",
        "format": "yyyy-MM-dd"
      }
    }
  ]
}
```

---

## 六、最佳实践

### 6.1 模式选择

```
根据数据结构选择模式：

• 单个值         → SINGLE
• 一列数据       → DOWN
• 一行数据       → RIGHT
• 矩形区域       → BLOCK
• 不定长列表     → UNTIL_EMPTY
• 键值对         → KEY_VALUE
• 表格数据       → TABLE
• 交叉表         → CROSS_TAB
• 分组数据       → GROUPED
• 层级数据       → HIERARCHY
```

### 6.2 推荐配置

```json
// ✅ 推荐：使用表头匹配 + skipEmpty
{
  "extractions": [
    {
      "key": "orderNos",
      "header": { "match": "订单号" },
      "mode": "DOWN",
      "range": { "skipEmpty": true }
    }
  ]
}

// ❌ 不推荐：硬编码位置
{
  "extractions": [
    {
      "key": "orderNos",
      "position": { "cellRef": "A2" },
      "mode": "DOWN",
      "range": { "rows": 100 }
    }
  ]
}
```

---

## 七、常见问题

### Q: 如何处理空值？

**A:** 使用 skipEmpty 跳过空行：

```json
{
  "range": { "skipEmpty": true }
}
```

### Q: 如何限制最大行数？

**A:** 使用 maxRows：

```json
{
  "range": { "maxRows": 1000 }
}
```

### Q: 如何提取合并单元格的值？

**A:** 使用 MERGED_CELLS 模式或 BLOCK 模式自动处理合并单元格。

---

## 八、参考资料

- [ExtractEngine 实现](../packages/core/src/main/java/com/excelconfig/extract/ExtractEngine.java)
- [填充模式详解](./FILL_MODES.md)
- [表头匹配机制](./HEADER_MATCHING.md)
