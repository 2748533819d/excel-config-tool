# Excel Config Tool - 表头匹配与动态扩展

> **核心原则**：配置通过表头文字匹配定位，数据量由实际数据决定，模板空间不足时自动扩展

---

## 一、问题场景

### 场景 1：数据量超过模板预留空间

```
Excel 模板：
┌─────────────────────────┐
│ A1: 订单号 | B1: 金额   │ ← 表头
├─────────────────────────┤
│ A2: ORD001 | B2: 100   │
│ A3: ORD002 | B3: 200   │
│ ...        | ...       │
│ A8: ORD007 | B8: 700   │ ← 模板只预留了 7 行数据
│ A9: (空)               │
├─────────────────────────┤
│ A10: 客户 | B10: 订单数│ ← 下一个表
└─────────────────────────┘

实际数据：
{
  "订单号": ["ORD001", "ORD002", ... "ORD020"],  // 20 条！
  "金额": [100, 200, ... 2000],                   // 20 条！
}

问题：
- 模板 A 列只预留到 A8（7 行数据）
- 实际数据有 20 条
- 如果直接填充，会覆盖 A10 的"客户"表！

解决方案：自动下移下方内容
```

### 场景 2：表头位置不固定

```
Excel 模板 1：
┌─────────────┐
│ 标题行      │
│ A1: 订单号  │ ← 表头在 A1
├─────────────┤
│ A2: 数据... │
└─────────────┘

Excel 模板 2：
┌─────────────┐
│ 标题行 1    │
│ 标题行 2    │
│ A3: 订单号  │ ← 表头在 A3
├─────────────┤
│ A4: 数据... │
└─────────────┘

问题：表头位置不固定，不能用固定的 cellRef 配置
解决：通过表头文字匹配自动定位
```

---

## 二、解决方案

### 2.1 表头匹配机制

```java
// 配置
{
  "key": "orderNos",
  "header": { "match": "订单号" }
}

// 定位流程
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
                if (config.getMatch().equals(value)) {
                    return new Position(rowNum, cell.getColumnIndex());
                }
            }
        }
        
        throw new HeaderNotFoundException("未找到表头：" + config.getMatch());
    }
}
```

### 2.2 动态扩展机制

```
填充流程：

1. 定位表头
   配置：{ header: { match: "订单号" } }
   结果：定位到 A1

2. 检查下方空间
   - 从表头下方（A2）开始检查
   - 查找下一个配置点（如"客户"表头）
   - 计算可用行数 = 下一个配置点行号 - 表头行号 - 1

3. 计算需要行数
   需要行数 = 数据数组长度

4. 判断是否需要扩展
   if (需要行数 > 可用行数) {
       // 下移下方内容
       int shiftRows = 需要行数 - 可用行数;
       sheet.shiftRows(下一个配置点行号，最后一行，shiftRows);
   }

5. 填充数据
   从 A2 开始，填充数据
```

---

## 三、配置语法

### 3.1 基础表头匹配

```json
{
  "extractions": [
    {
      "key": "orderNos",
      "header": {
        "match": "订单号"
      },
      "mode": "DOWN"
    }
  ]
}
```

### 3.2 指定搜索范围

```json
{
  "extractions": [
    {
      "key": "orderNos",
      "header": {
        "match": "订单号",
        "inRows": [1, 10]
      },
      "mode": "DOWN"
    }
  ]
}
```

说明：`inRows: [1, 10]` 表示在第 1 行到第 10 行之间搜索表头

### 3.3 多表配置

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

---

## 四、自动扩展详解

### 4.1 单列扩展

```
原始模板：
┌─────────────┐
│ A1: 订单号   │
│ A2:         │
│ A3:         │
│ A4:         │
├─────────────┤
│ A5: 合计     │  ← 原有内容
└─────────────┘

填充 5 条数据：
["ORD001", "ORD002", "ORD003", "ORD004", "ORD005"]

结果：
┌─────────────┐
│ A1: 订单号   │
│ A2: ORD001  │
│ A3: ORD002  │
│ A4: ORD003  │
│ A5: ORD004  │
│ A6: ORD005  │
├─────────────┤
│ A7: 合计     │  ← 自动下移 2 行
└─────────────┘
```

### 4.2 多列扩展（列隔离）

```
原始模板：
┌─────────┬─────────┬─────────┐
│ A1: 订单号│ B1: 金额 │ C1: 日期 │
│ A2:      │ B2:     │ C2:     │
│ A3:      │ B3:     │ C3:     │
├─────────┴─────────┴─────────┤
│ A4: 合计                     │  ← 原有内容
└─────────────────────────────┘

填充数据：
{
  "订单号": ["ORD001", "ORD002", "ORD003", "ORD004", "ORD005"],  // 5 条
  "金额": [100, 200, 300],                                       // 3 条
  "日期": ["2024-01-01", "2024-01-02", "2024-01-03", "2024-01-04"]  // 4 条
}

结果（每列独立扩展）：
┌─────────┬─────────┬─────────┐
│ A1: 订单号│ B1: 金额 │ C1: 日期 │
│ A2: ORD001│ B2: 100 │ C2: 2024-01-01 │
│ A3: ORD002│ B3: 200 │ C3: 2024-01-02 │
│ A4: ORD003│ B4: 300 │ C4: 2024-01-03 │
│ A5: ORD004│ B5:     │ C5: 2024-01-04 │
│ A6: ORD005│         │     │
├─────────┴─────────┴─────────┤
│ A7: 合计                     │  ← 下移到第 7 行（最大偏移）
└─────────────────────────────┘

规则：
- A 列填充 5 行 → 偏移 5 行
- B 列填充 3 行 → 偏移 3 行
- C 列填充 4 行 → 偏移 4 行
- 下方内容偏移量 = max(5, 3, 4) = 5 行
```

---

## 五、边界检测

### 5.1 提取边界

```
提取时遇到以下情况停止：

1. 达到 maxRows 限制
   { "range": { "maxRows": 100 } } → 最多提取 100 行

2. 遇到空行（skipEmpty=true）
   { "range": { "skipEmpty": true } } → 遇到空行停止

3. 遇到下一个配置点
   如果配置了多个表，遇到下一个表头时停止

4. 到达 Sheet 末尾
   自然结束
```

### 5.2 填充边界

```
填充时的边界处理：

1. 有下一个配置点
   → 检查空间，不足则下移

2. 没有下一个配置点
   → 直接扩展，无限制

3. 达到 maxRows 限制
   → 截断数据，只填充 maxRows 行
```

---

## 六、最佳实践

### 6.1 推荐配置方式

```json
// ✅ 推荐：使用表头匹配，不指定固定位置
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

// ❌ 不推荐：硬编码固定位置
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

### 6.2 多表配置技巧

```json
{
  "version": "1.0",
  "extractions": [
    {
      "key": "orders",
      "header": { 
        "match": "订单号",
        "inRows": [1, 50]  // 限制搜索范围，避免匹配到下面的表
      },
      "mode": "DOWN"
    },
    {
      "key": "customers",
      "header": { 
        "match": "客户",
        "inRows": [55, 100]  // 在下方区域搜索
      },
      "mode": "DOWN"
    }
  ]
}
```

### 6.3 避免覆盖

```
如果模板中有多个表：

1. 为每个表配置 inRows 范围
2. 确保表之间有足够的间隔
3. 填充时会自动检测并下移

配置示例：
┌──────────────────────────────┐
│ 订单表 (inRows: 1-50)         │
│ ...                          │
├──────────────────────────────┤
│ 间隔行（作为缓冲区）          │
├──────────────────────────────┤
│ 客户表 (inRows: 55-100)       │
└──────────────────────────────┘
```

---

## 七、常见问题

### Q: 表头文字包含空格怎么办？

**A:** 配置中的 match 值需要精确匹配：
```json
{
  "header": { "match": "订单号 " }  // 注意空格
}
```

或者在代码中预处理表头：
```java
// 未来支持模糊匹配/正则匹配
{
  "header": { 
    "match": "订单号",
    "matchMode": "CONTAINS"  // TODO: 未来支持
  }
}
```

### Q: 如何处理合并的表头单元格？

**A:** 当前实现支持合并单元格，会读取合并区域的左上角值。

### Q: 表头在最后一行怎么办？

**A:** 需要指定 inRows 范围：
```json
{
  "header": { 
    "match": "合计",
    "inRows": [50, 100]  // 在 50-100 行搜索
  }
}
```

---

## 八、参考资料

- [HeaderLocator 实现](../packages/core/src/main/java/com/excelconfig/locator/HeaderLocator.java)
- [FillEngine 实现](../packages/core/src/main/java/com/excelconfig/export/FillEngine.java)
- [列隔离机制](./COLUMN_ISOLATION.md)
