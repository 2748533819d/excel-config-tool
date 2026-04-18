# Excel Config Tool - 列隔离与行偏移机制

> 解决多列同时填充时的数据冲突问题
> **核心原则：填充行数由数据量决定，而非配置决定**

---

## 一、问题场景

### 场景描述

```
原始模板 Excel：
┌──────────────┬──────────────┬──────────────┐
│   A 列        │   B 列        │   C 列        │
├──────────────┼──────────────┼──────────────┤
│ A1: 订单号    │ B1: 金额     │ C1: 备注     │
│ A2: 原有订单  │ B2: 原有金额 │ C2: 原有备注 │
│ A3: 预留行    │ B3: 预留行   │ C3: 预留行   │
└──────────────┴──────────────┴──────────────┘

填充配置：
- A1 向下填充 orderNos 数组（假设 5 条数据）
- B1 向下填充 amounts 数组（假设 3 条数据）
- C 列不填充

期望结果：
┌──────────────┬──────────────┬──────────────┐
│   A 列        │   B 列        │   C 列        │
├──────────────┼──────────────┼──────────────┤
│ A1: 订单号    │ B1: 金额     │ C1: 备注     │
│ A2-A6: 新订单│ B2-B4: 新金额│ C2: 原有备注 │
│ A7: 原有订单 │ B5: 原有金额 │ C3: 预留行   │
└──────────────┴──────────────┴──────────────┘

关键：
- A 列填充 5 行 → 原数据从 A7 开始
- B 列填充 3 行 → 原数据从 B5 开始
- 每列填充行数由实际数据量决定，不是配置写死
```

### 核心问题

1. **列隔离**：每列的填充独立计算，互不干扰
2. **行偏移**：原数据需要根据填充行数向下偏移
3. **样式保护**：模板原有样式、公式、合并单元格不能破坏

---

## 二、设计原则

### 2.1 列独立性原则

每列的填充范围独立计算，**填充行数由数据量决定**：

```json
{
  "exports": [
    {
      "key": "orderNos",
      "header": { "match": "订单号" },
      "mode": "FILL_DOWN"
      // 不指定 rows，填充行数 = orderNos 数组的长度
    },
    {
      "key": "amounts",
      "header": { "match": "金额" },
      "mode": "FILL_DOWN"
      // 不指定 rows，填充行数 = amounts 数组的长度
    }
  ]
}
```

**数据驱动示例**：

```json
// 如果传入的数据是：
{
  "orderNos": ["ORD001", "ORD002", "ORD003"],  // 3 条
  "amounts": [100, 200, 300, 400, 500]  // 5 条
}

// 结果：
// A 列自动填充 3 行
// B 列自动填充 5 行
```

### 2.2 行偏移原则

原数据的偏移量 = 该列的填充行数

```
对于列 A：
- A1 是表头，不偏移
- A2 开始的填充区域：填充行数 = orderNos 数组长度
- 原 A2 的数据 → 偏移到 A(2 + 填充行数)

对于列 B：
- B1 是表头，不偏移
- B2 开始的填充区域：填充行数 = amounts 数组长度
- 原 B2 的数据 → 偏移到 B(2 + 填充行数)
```

---

## 三、核心机制

### 3.1 填充行数计算

```java
public class FillEngine {
    
    /**
     * 计算每列需要填充的行数
     */
    private Map<String, Integer> calculateFillRows(
        Map<String, Object> data, 
        List<ExportConfig> exports
    ) {
        Map<String, Integer> fillRows = new HashMap<>();
        
        for (ExportConfig export : exports) {
            String key = export.getKey();
            Object value = data.get(key);
            
            if (value instanceof List) {
                fillRows.put(key, ((List<?>) value).size());
            } else {
                fillRows.put(key, 1);  // 单值填充 1 行
            }
        }
        
        return fillRows;
    }
}
```

### 3.2 列偏移计算

```java
public class FillEngine {
    
    /**
     * 计算每列的偏移量
     */
    private Map<Integer, Integer> calculateColumnOffsets(
        Map<String, Integer> fillRows,
        Map<String, ExportConfig> exportByColumn
    ) {
        Map<Integer, Integer> columnOffsets = new HashMap<>();
        
        for (Map.Entry<String, Integer> entry : fillRows.entrySet()) {
            String key = entry.getKey();
            int rows = entry.getValue();
            
            ExportConfig config = exportByColumn.get(key);
            int column = config.getHeaderColumn();
            
            // 该列的偏移量 = 填充行数
            columnOffsets.put(column, rows);
        }
        
        return columnOffsets;
    }
}
```

### 3.3 下方内容下移

```java
public class FillEngine {
    
    /**
     * 下移下方内容
     */
    private void shiftDownContent(
        Sheet sheet, 
        int startRow, 
        int maxOffset
    ) {
        if (maxOffset <= 0) {
            return;  // 不需要偏移
        }
        
        // 找到 startRow 下方的第一个非空行
        int firstDataRow = findFirstNonEmptyRow(sheet, startRow);
        if (firstDataRow < 0) {
            return;  // 下方没有内容
        }
        
        // 使用 POI 的 shiftRows 方法下移
        sheet.shiftRows(
            firstDataRow,           // 起始行
            sheet.getLastRowNum(),  // 结束行
            maxOffset               // 偏移量
        );
    }
}
```

---

## 四、完整流程

### 4.1 填充流程

```
┌─────────────────────────────────────────────────────────────────┐
│                        填充流程                                  │
├─────────────────────────────────────────────────────────────────┤
│                                                                 │
│  1. 解析配置                                                     │
│     └─→ 获取所有 ExportConfig                                   │
│                                                                 │
│  2. 定位表头                                                     │
│     └─→ 对每个 ExportConfig，通过 HeaderLocator 定位              │
│                                                                 │
│  3. 计算填充行数                                                 │
│     └─→ 对每个 key，fillRows = data.get(key).size()             │
│                                                                 │
│  4. 计算列偏移                                                   │
│     └─→ columnOffset[column] = fillRows[key]                    │
│                                                                 │
│  5. 计算最大偏移                                                 │
│     └─→ maxOffset = max(columnOffsets.values())                 │
│                                                                 │
│  6. 下移下方内容                                                 │
│     └─→ sheet.shiftRows(firstDataRow, lastRow, maxOffset)       │
│                                                                 │
│  7. 填充数据                                                     │
│     └─→ 对每个列，从表头下方开始填充                             │
│                                                                 │
└─────────────────────────────────────────────────────────────────┘
```

### 4.2 示例

```
原始模板：
┌─────────┬─────────┬─────────┐
│ A1: 订单号│ B1: 金额 │ C1: 日期 │
│ A2:      │ B2:     │ C2:     │
│ A3:      │ B3:     │ C3:     │
├─────────┴─────────┴─────────┤
│ A4: 合计                     │
└─────────────────────────────┘

数据：
{
  "orderNos": ["ORD001", "ORD002", "ORD003", "ORD004", "ORD005"],  // 5 条
  "amounts": [100, 200, 300],                                       // 3 条
  "dates": ["2024-01-01", "2024-01-02", "2024-01-03", "2024-01-04"] // 4 条
}

执行流程：

1. 定位表头 → A1, B1, C1

2. 计算填充行数：
   - A 列：5 行
   - B 列：3 行
   - C 列：4 行

3. 计算列偏移：
   - A 列偏移：5
   - B 列偏移：3
   - C 列偏移：4

4. 最大偏移：max(5, 3, 4) = 5

5. 下移下方内容：
   - "合计"行从 R4 下移到 R9

6. 填充数据：
   - A 列：A2-A6 填充 5 条订单号
   - B 列：B2-B4 填充 3 条金额
   - C 列：C2-C5 填充 4 条日期

结果：
┌─────────┬─────────┬─────────┐
│ A1: 订单号│ B1: 金额 │ C1: 日期 │
│ A2: ORD001│ B2: 100 │ C2: 2024-01-01 │
│ A3: ORD002│ B3: 200 │ C3: 2024-01-02 │
│ A4: ORD003│ B4: 300 │ C4: 2024-01-03 │
│ A5: ORD004│ B5:     │ C5: 2024-01-04 │
│ A6: ORD005│         │         │
├─────────┴─────────┴─────────┤
│ A9: 合计                     │
└─────────────────────────────┘
```

---

## 五、样式保护

### 5.1 样式复制机制

```java
public class FillEngine {
    
    /**
     * 填充时复制源行样式
     */
    private void fillWithStyle(
        Row targetRow, 
        Row sourceRow, 
        Object value,
        int cellIndex
    ) {
        Cell targetCell = targetRow.createCell(cellIndex);
        
        // 如果源行有样式，复制样式
        if (sourceRow != null) {
            Cell sourceCell = sourceRow.getCell(cellIndex);
            if (sourceCell != null && sourceCell.getCellStyle() != null) {
                targetCell.setCellStyle(sourceCell.getCellStyle());
            }
        }
        
        // 设置值
        setCellValue(targetCell, value);
    }
}
```

### 5.2 样式配置

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
        "background": "#FFFFFF"
      }
    }
  ]
}
```

---

## 六、最佳实践

### 6.1 推荐配置

```json
// ✅ 推荐：让数据量决定填充行数
{
  "exports": [
    {
      "key": "orderNos",
      "header": { "match": "订单号" },
      "mode": "FILL_DOWN"
    }
  ]
}

// ❌ 不推荐：硬编码行数
{
  "exports": [
    {
      "key": "orderNos",
      "header": { "match": "订单号" },
      "mode": "FILL_DOWN",
      "maxRows": 100  // 除非确实需要限制
    }
  ]
}
```

### 6.2 多列表格配置

```json
{
  "exports": [
    {
      "key": "orders",
      "header": { "match": "订单号" },
      "mode": "FILL_TABLE",
      "columns": [
        { "key": "orderNo", "header": "订单号" },
        { "key": "amount", "header": "金额" },
        { "key": "date", "header": "日期" }
      ],
      "alternateRows": true,
      "autoWidth": true
    }
  ]
}
```

---

## 七、常见问题

### Q: 如果某列数据为空怎么办？

**A:** 空列表不填充，该列保持原样：

```java
if (value instanceof List && ((List<?>) value).isEmpty()) {
    // 跳过空列表
    return;
}
```

### Q: 如何保证填充后公式正确？

**A:** POI 会自动更新公式引用，但如果公式引用了被覆盖的单元格，需要重新计算：

```java
// 填充完成后，重新计算公式
sheet.getDataValidationHelper();
workbook.getCreationHelper().createFormulaEvaluator()
    .evaluateAll();
```

### Q: 填充会破坏合并单元格吗？

**A:** 不会。填充引擎会检测合并区域，避免覆盖：

```java
// 检查单元格是否在合并区域内
for (CellRangeAddress mergedRegion : sheet.getMergedRegions()) {
    if (mergedRegion.isInRange(rowNum, cellIndex)) {
        // 跳过合并单元格
        return;
    }
}
```

---

## 八、参考资料

- [FillEngine 实现](../packages/core/src/main/java/com/excelconfig/export/FillEngine.java)
- [表头匹配机制](./HEADER_MATCHING.md)
- [填充模式详解](./FILL_MODES.md)
