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
4. **依赖检测**：如果 B 列依赖 A 列的公式，需要特殊处理

---

## 二、设计原则

### 1. 列独立性原则

每列的填充范围独立计算，**填充行数由数据量决定**：

```yaml
exports:
  - key: "orderNos"
    position: { cellRef: "A2" }
    mode: FILL_DOWN
    # 不指定 rows，填充行数 = orderNos 数组的长度
    
  - key: "amounts"
    position: { cellRef: "B2" }
    mode: FILL_DOWN
    # 不指定 rows，填充行数 = amounts 数组的长度
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

### 2. 行偏移原则

原数据的偏移量 = 该列的填充行数（由数据量决定）

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

**示例**（数据量不同）：
```
数据：
{
  "orderNos": ["ORD001", "ORD002", "ORD003", "ORD004", "ORD005"],  // 5 条
  "amounts": [100, 200, 300]  // 3 条
}

结果：
- A 列填充 5 行 → 原 A2 数据偏移到 A7
- B 列填充 3 行 → 原 B2 数据偏移到 B5
```

### 3. 样式继承原则

```
填充区域的样式优先级：
1. 配置中指定的 style > 模板样式
2. 未指定 style 时，继承模板对应位置的样式
3. 超出模板范围的新行，使用模板最后一行的样式（模式行）
```

---

## 三、实现机制

### 1. 列元数据收集

```java
public class ColumnMetadata {
    // 列索引
    private int columnIndex;
    
    // 填充起始行（表头下方第一行）
    private int fillStartRow;
    
    // 填充结束行
    private int fillEndRow;
    
    // 填充行数
    private int fillRows;
    
    // 原始数据起始行
    private int originalDataStartRow;
    
    // 原始数据偏移后的行
    private int shiftedDataRow;
    
    // 该列是否有原始数据
    private boolean hasOriginalData;
    
    // 该列的填充配置
    private ExportConfig exportConfig;
}
```

### 2. 填充计划计算

```java
public class FillPlan {
    // 所有涉及的列
    private Map<Integer, ColumnMetadata> columnMetadatas;
    
    // 最大填充行数（所有列中填充最多的）
    private int maxFillRows;
    
    // 需要下移的原始数据行
    private List<OriginalDataRow> rowsToShift;
    
    /**
     * 计算填充计划
     * @param workbook Excel 工作簿
     * @param exports 导出配置列表
     * @param data 实际数据（用于计算每列的填充行数）
     */
    public FillPlan calculate(Workbook workbook, List<ExportConfig> exports, Map<String, Object> data) {
        FillPlan plan = new FillPlan();
        
        for (ExportConfig export : exports) {
            if (export.getMode() == FILL_DOWN) {
                // 从实际数据中获取该列的行数
                Object columnData = data.get(export.getKey());
                int actualRows = getDataRowCount(columnData, export);
                
                ColumnMetadata meta = calculateColumnMetadata(workbook, export, actualRows);
                plan.columnMetadatas.put(meta.getColumnIndex(), meta);
                plan.maxFillRows = Math.max(plan.maxFillRows, actualRows);
            }
        }
        
        // 计算需要下移的原始数据
        plan.rowsToShift = calculateRowsToShift(workbook, plan);
        
        return plan;
    }
    
    /**
     * 根据数据类型计算行数
     */
    private int getDataRowCount(Object data, ExportConfig config) {
        if (data instanceof List) {
            int size = ((List<?>) data).size();
            return config.getMaxRows() != null ? Math.min(size, config.getMaxRows()) : size;
        } else if (data.getClass().isArray()) {
            int size = Array.getLength(data);
            return config.getMaxRows() != null ? Math.min(size, config.getMaxRows()) : size;
        } else {
            return 1;  // 单个值
        }
    }
}
```

### 3. 行偏移计算

```java
private List<OriginalDataRow> calculateRowsToShift(Workbook workbook, FillPlan plan) {
    List<OriginalDataRow> rowsToShift = new ArrayList<>();
    
    Sheet sheet = workbook.getSheetAt(0);
    
    // 找出所有受影响的原始数据行
    for (Row row : sheet) {
        int rowNum = row.getRowNum();
        
        // 检查这行是否在任意填充区域下方
        for (ColumnMetadata meta : plan.columnMetadatas.values()) {
            if (rowNum >= meta.getOriginalDataStartRow()) {
                // 这行数据需要下移
                // 下移量 = 该列的填充行数
                int shiftAmount = meta.getFillRows();
                rowsToShift.add(new OriginalDataRow(rowNum, shiftAmount));
                break;
            }
        }
    }
    
    return rowsToShift;
}
```

### 4. 安全填充策略

```java
public class SafeFillStrategy {
    
    public void fill(Workbook workbook, List<ExportConfig> exports) {
        // 步骤 1: 计算填充计划
        FillPlan plan = fillPlanCalculator.calculate(workbook, exports);
        
        // 步骤 2: 从下往上移动原始数据（避免覆盖）
        shiftOriginalData(workbook, plan);
        
        // 步骤 3: 执行各列填充
        for (ExportConfig export : exports) {
            FillStrategy strategy = getStrategy(export.getMode());
            strategy.fill(workbook, createFillContext(export, plan));
        }
        
        // 步骤 4: 应用样式
        applyStyles(workbook, exports, plan);
    }
    
    private void shiftOriginalData(Workbook workbook, FillPlan plan) {
        Sheet sheet = workbook.getSheetAt(0);
        
        // 从最后一行开始往上移，避免覆盖
        // 找出最大偏移行
        int maxRow = sheet.getLastRowNum();
        
        for (int i = maxRow; i >= 0; i--) {
            Row row = sheet.getRow(i);
            if (row == null) continue;
            
            // 检查这行是否需要移动
            Integer shiftAmount = getShiftAmountForRow(i, plan);
            if (shiftAmount != null && shiftAmount > 0) {
                // 创建新行
                Row newRow = sheet.getRow(i + shiftAmount);
                if (newRow == null) {
                    newRow = sheet.createRow(i + shiftAmount);
                }
                
                // 复制单元格（值、样式、公式）
                copyRow(row, newRow, getCellsToShift(row, plan));
            }
        }
    }
    
    private Integer getShiftAmountForRow(int rowNum, FillPlan plan) {
        Integer maxShift = null;
        
        for (ColumnMetadata meta : plan.columnMetadatas.values()) {
            if (rowNum >= meta.getOriginalDataStartRow()) {
                // 这行在该列下方，需要下移
                maxShift = Math.max(maxShift == null ? 0 : maxShift, meta.getFillRows());
            }
        }
        
        return maxShift;
    }
}
```

---

## 四、配置语法增强

### 基础列隔离配置

```yaml
exports:
  # A 列：订单号，向下填充 10 行
  - key: "orderNos"
    position:
      cellRef: "A2"
    mode: FILL_DOWN
    parser:
      type: string
    range:
      rows: 10
    
  # B 列：金额，向下填充 5 行
  - key: "amounts"
    position:
      cellRef: "B2"
    mode: FILL_DOWN
    parser:
      type: number
      format: "#,##0.00"
    range:
      rows: 5
  
  # C 列：不填充，保留原样
  # (不需要配置)
```

### 带表头的列隔离

```yaml
exports:
  # 表头（不填充，只是标识）
  - key: "headers"
    position:
      cellRef: "A1"
    mode: FILL_CELL
    parser:
      type: string
    
  # 数据列配置
  - key: "data"
    position:
      cellRef: "A2"
    mode: FILL_TABLE
    columns:
      - key: "orderNo"
        header: "订单号"
        fillMode: FILL_DOWN
        range: { rows: 10 }
      - key: "amount"
        header: "金额"
        fillMode: FILL_DOWN
        range: { rows: 5 }  # 可以和 orderNo 不同
      - key: "remark"
        header: "备注"
        fillMode: FILL_DOWN
        range: { rows: 10 }
```

### 保护区域配置

```yaml
exports:
  - key: "orderNos"
    position:
      cellRef: "A2"
    mode: FILL_DOWN
    range:
      rows: 10
    protection:
      # 保护区域，不被其他列的填充影响
      protectedArea:
        areaRef: "A2:A100"
        preserveStyles: true
        preserveFormulas: true
    
  - key: "amounts"
    position:
      cellRef: "B2"
    mode: FILL_DOWN
    range:
      rows: 5
```

---

## 五、复杂场景处理

### 场景 1：多组独立填充

```
A 列、B 列是一组（订单数据）
C 列、D 列是另一组（客户数据）
两组独立填充，互不影响

配置：
```yaml
exports:
  # 第一组：订单数据
  - key: "orders"
    position: { cellRef: "A2" }
    mode: FILL_TABLE
    range: { rows: 10 }
    columns:
      - { key: "orderNo", header: "订单号" }
      - { key: "amount", header: "金额" }
  
  # 第二组：客户数据
  - key: "customers"
    position: { cellRef: "C2" }
    mode: FILL_TABLE
    range: { rows: 5 }
    columns:
      - { key: "customerName", header: "客户" }
      - { key: "phone", header: "电话" }
```

### 场景 2：跨列公式保护

```
A 列：订单号（填充 10 行）
B 列：金额（填充 10 行）
C 列：=A2*B2（公式，需要自动填充）

配置：
```yaml
exports:
  - key: "orderData"
    position: { cellRef: "A2" }
    mode: FILL_TABLE
    range: { rows: 10 }
    columns:
      - { key: "orderNo", header: "订单号" }
      - { key: "amount", header: "金额" }
    
  # 公式列特殊处理
  - key: "totalFormula"
    position: { cellRef: "C2" }
    mode: FILL_FORMULA
    formula: "=A2*B2"
    autoFill: true  # 自动填充到所有行
```

### 场景 3：合并单元格保护

```
A 列：订单号（普通填充）
B 列+C 列：合并单元格（特殊处理）

配置：
```yaml
exports:
  - key: "orderNos"
    position: { cellRef: "A2" }
    mode: FILL_DOWN
    range: { rows: 10 }
  
  - key: "mergedData"
    position: { cellRef: "B2" }
    mode: FILL_DOWN
    range: { rows: 5 }
    mergedCells:
      enabled: true
      mergeAcross: 2  # B 列和 C 列合并
```

---

## 六、填充执行流程

```
┌─────────────────────────────────────────────────────────────────┐
│                      填充执行流程                                │
├─────────────────────────────────────────────────────────────────┤
│                                                                 │
│  1. 解析配置                                                     │
│     ↓                                                           │
│  2. 收集所有列的填充元数据                                        │
│     - 每列的填充起始位置                                          │
│     - 每列的填充行数                                              │
│     - 每列的原始数据位置                                          │
│     ↓                                                           │
│  3. 计算偏移计划                                                  │
│     - 哪些行需要下移                                              │
│     - 每行下移多少                                                │
│     ↓                                                           │
│  4. 从下往上移动原始数据（避免覆盖）                               │
│     ↓                                                           │
│  5. 执行各列填充                                                  │
│     - 按列独立填充                                                │
│     - 应用样式                                                    │
│     ↓                                                           │
│  6. 处理公式和合并单元格                                          │
│     ↓                                                           │
│  7. 应用条件格式                                                  │
│     ↓                                                           │
│  8. 完成输出                                                     │
│                                                                 │
└─────────────────────────────────────────────────────────────────┘
```

---

## 七、Java 实现骨架

### 列隔离填充器

```java
public class ColumnIsolatedFiller {
    
    private final FillPlanCalculator planCalculator;
    private final DataShifter dataShifter;
    private final List<FillStrategy> strategies;
    
    public void fill(Workbook workbook, List<ExportConfig> exports) {
        // 1. 计算填充计划
        FillPlan plan = planCalculator.calculate(workbook, exports);
        
        log.debug("填充计划：最大填充行数 = {}", plan.getMaxFillRows());
        for (ColumnMetadata meta : plan.getColumnMetadatas().values()) {
            log.debug("列 {} - 填充 {} 行，原始数据从行 {} 开始",
                meta.getColumnIndex(),
                meta.getFillRows(),
                meta.getOriginalDataStartRow());
        }
        
        // 2. 移动原始数据（从下往上）
        dataShifter.shift(workbook, plan);
        
        // 3. 执行各列填充
        for (ExportConfig export : exports) {
            FillStrategy strategy = findStrategy(export.getMode());
            ColumnFillContext context = new ColumnFillContext(
                workbook, export, plan);
            strategy.fill(context);
        }
        
        // 4. 处理跨列依赖（公式、合并单元格）
        handleCrossColumnDependencies(workbook, exports, plan);
    }
    
    private void handleCrossColumnDependencies(
        Workbook workbook,
        List<ExportConfig> exports,
        FillPlan plan) {
        
        for (ExportConfig export : exports) {
            if (export.getMode() == FILL_FORMULA) {
                FormulaFillStrategy formulaStrategy = 
                    (FormulaFillStrategy) findStrategy(FILL_FORMULA);
                formulaStrategy.fillFormula(workbook, export, plan);
            }
        }
    }
}
```

### 数据偏移器

```java
public class DataShifter {
    
    public void shift(Workbook workbook, FillPlan plan) {
        Sheet sheet = workbook.getSheetAt(0);
        
        // 收集所有需要移动的行
        Set<Integer> rowsToShift = new TreeSet<>(Collections.reverseOrder());
        for (int i = sheet.getLastRowNum(); i >= 0; i--) {
            if (shouldShiftRow(i, plan)) {
                rowsToShift.add(i);
            }
        }
        
        // 从下往上移动
        for (int rowNum : rowsToShift) {
            Row oldRow = sheet.getRow(rowNum);
            if (oldRow == null) continue;
            
            int shiftAmount = getShiftAmountForRow(rowNum, plan);
            int newRowNum = rowNum + shiftAmount;
            
            Row newRow = getOrCreateRow(sheet, newRowNum);
            copyRow(oldRow, newRow, getAffectedCells(oldRow, plan));
        }
    }
    
    private boolean shouldShiftRow(int rowNum, FillPlan plan) {
        for (ColumnMetadata meta : plan.getColumnMetadatas().values()) {
            if (rowNum >= meta.getOriginalDataStartRow()) {
                return true;
            }
        }
        return false;
    }
    
    private int getShiftAmountForRow(int rowNum, FillPlan plan) {
        int maxShift = 0;
        for (ColumnMetadata meta : plan.getColumnMetadatas().values()) {
            if (rowNum >= meta.getOriginalDataStartRow()) {
                maxShift = Math.max(maxShift, meta.getFillRows());
            }
        }
        return maxShift;
    }
}
```

---

## 八、测试用例

### 测试 1：基本列隔离

```java
@Test
public void testBasicColumnIsolation() {
    // 准备模板
    Workbook template = new XSSFWorkbook();
    Sheet sheet = template.createSheet("Test");
    
    // 创建原始数据
    sheet.createRow(0).createCell(0).setCellValue("订单号");
    sheet.getRow(0).createCell(1).setCellValue("金额");
    sheet.createRow(1).createCell(0).setCellValue("原订单 A");
    sheet.getRow(1).createCell(1).setCellValue("100");
    
    // 配置
    List<ExportConfig> exports = List.of(
        ExportConfig.builder()
            .key("orderNos")
            .position(Position.of("A2"))
            .mode(FILL_DOWN)
            .range(Range.rows(10))
            .build(),
        ExportConfig.builder()
            .key("amounts")
            .position(Position.of("B2"))
            .mode(FILL_DOWN)
            .range(Range.rows(5))
            .build()
    );
    
    // 执行填充
    ColumnIsolatedFiller filler = new ColumnIsolatedFiller();
    filler.fill(template, exports);
    
    // 验证
    Sheet result = template.getSheetAt(0);
    assertEquals("新订单 1", result.getRow(1).getCell(0).getStringCellValue());
    assertEquals("新订单 10", result.getRow(10).getCell(0).getStringCellValue());
    assertEquals("原订单 A", result.getRow(11).getCell(0).getStringCellValue());
    
    assertEquals("新金额 1", result.getRow(1).getCell(1).getStringCellValue());
    assertEquals("新金额 5", result.getRow(5).getCell(1).getStringCellValue());
    assertEquals("100", result.getRow(6).getCell(1).getStringCellValue());
}
```

---

## 九、总结

### 核心设计点

| 机制 | 说明 | 实现方式 |
|------|------|----------|
| 列隔离 | 每列独立填充，互不影响 | 按列收集元数据，独立计算填充范围 |
| 行偏移 | 原数据根据填充行数下移 | 从下往上移动，避免覆盖 |
| 样式保护 | 保留模板样式 | 复制单元格时保留样式 |
| 公式处理 | 自动填充和调整公式引用 | 填充后处理公式列 |

### 配置要点

1. **每个 FILL_DOWN 配置独立指定 range.rows**
2. **不需要特殊配置列隔离，自动处理**
3. **原始数据自动下移，无需手动指定**
4. **公式列使用 FILL_FORMULA 特殊处理**

### 限制

1. 不支持交叉填充（A 列的数据依赖 B 列的填充结果）
2. 合并单元格跨列时需要特殊配置
3. 超大偏移量（>10000 行）可能影响性能
