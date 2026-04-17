# Excel Config Tool - 表头匹配与动态扩展

> **核心原则：配置通过表头匹配定位，数据量由实际数据决定，模板空间不足时自动扩展**

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
  "客户": ["张三", "李四", "王五"]                  // 3 条
}

问题：
- 模板 A 列只预留到 A8（7 行数据）
- 实际数据有 20 条
- 如果直接填充，会覆盖 A10 的"客户"表！
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
```

---

## 二、解决方案：表头匹配 + 动态扩展

### 方案总览

```
┌─────────────────────────────────────────────────────────────────┐
│                    整体流程                                      │
├─────────────────────────────────────────────────────────────────┤
│                                                                 │
│  1. 表头匹配                                                    │
│     配置：{ header: "订单号" }                                   │
│     行为：扫描第 1 行（或指定行），找到包含"订单号"的单元格            │
│     结果：定位到 A1                                             │
│                                                                 │
│  2. 数据提取/填充                                               │
│     提取：从表头下方开始，直到遇到边界（空行/下一个表头）         │
│     填充：从表头下方开始，根据数据量动态扩展                     │
│                                                                 │
│  3. 边界处理                                                    │
│     提取：遇到空行或下一个已知表头时停止                         │
│     填充：如果下方有其他表，自动下移其他表                       │
│                                                                 │
└─────────────────────────────────────────────────────────────────┘
```

---

## 三、配置语法

### 1. 表头匹配配置

```yaml
extractions:
  # 通过表头文字匹配定位
  - key: "orderNos"
    header:
      match: "订单号"           # 精确匹配
      searchRow: 1             # 在第 1 行搜索（默认）
      partialMatch: false      # 是否部分匹配（默认 false）
    mode: DOWN
    
  - key: "amounts"
    header:
      match: "金额"
    mode: DOWN
    
  - key: "customers"
    header:
      match: "客户"
    mode: DOWN
```

### 2. 部分匹配

```yaml
extractions:
  - key: "orderNos"
    header:
      match: "订单"            # 部分匹配
      partialMatch: true       # 可以匹配"订单号"、"订单编号"等
```

### 3. 正则匹配

```yaml
extractions:
  - key: "orderNos"
    header:
      match: "订单.*号"        # 正则表达式
      regex: true
```

### 4. 多行搜索

```yaml
extractions:
  - key: "orderNos"
    header:
      match: "订单号"
      searchRows:
        start: 1
        end: 5                # 在第 1-5 行搜索表头
```

### 5. 导出配置（同样方式）

```yaml
exports:
  - key: "orderNos"
    header:
      match: "订单号"
    mode: FILL_DOWN
    
  - key: "customers"
    header:
      match: "客户"
    mode: FILL_DOWN
```

---

## 四、核心机制

### 1. 表头定位器

```java
public class HeaderLocator {
    
    /**
     * 根据表头文字定位单元格
     */
    public Position locate(Sheet sheet, HeaderConfig config) {
        int startRow = config.getSearchRow() != null ? config.getSearchRow() : 1;
        int endRow = config.getSearchRows() != null ? config.getSearchRows().getEnd() : startRow;
        
        for (int rowNum = startRow; rowNum <= endRow; rowNum++) {
            Row row = sheet.getRow(rowNum);
            if (row == null) continue;
            
            for (Cell cell : row) {
                String cellValue = getCellValueAsString(cell);
                
                if (matches(cellValue, config)) {
                    return new Position(rowNum, cell.getColumnIndex());
                }
            }
        }
        
        throw new HeaderNotFoundException(
            String.format("未找到表头 '%s' (搜索范围：行 %d-%d)", 
                config.getMatch(), startRow, endRow));
    }
    
    private boolean matches(String cellValue, HeaderConfig config) {
        if (cellValue == null) return false;
        
        if (config.isRegex()) {
            return cellValue.matches(config.getMatch());
        } else if (config.isPartialMatch()) {
            return cellValue.contains(config.getMatch());
        } else {
            return cellValue.equals(config.getMatch());
        }
    }
}
```

### 2. 边界检测器（提取）

```java
public class ExtractBoundaryDetector {
    
    /**
     * 检测提取的结束行
     */
    public int detectEndRow(Sheet sheet, Position headerPos, List<Position> knownHeaders) {
        int column = headerPos.getColumn();
        int startRow = headerPos.getRow() + 1;  // 从表头下方开始
        
        for (int rowNum = startRow; ; rowNum++) {
            Row row = sheet.getRow(rowNum);
            
            // 空行：停止
            if (row == null || isRowEmpty(row)) {
                return rowNum - 1;
            }
            
            // 检查是否是已知表头位置
            if (isKnownHeaderPosition(rowNum, column, knownHeaders)) {
                return rowNum - 1;  // 在表头之前停止
            }
            
            // 检查是否是新表头（通过样式检测）
            if (isHeaderRow(row)) {
                return rowNum - 1;  // 在新表头之前停止
            }
        }
    }
    
    private boolean isKnownHeaderPosition(int rowNum, int column, List<Position> knownHeaders) {
        return knownHeaders.stream()
            .anyMatch(p -> p.getRow() == rowNum && p.getColumn() == column);
    }
}
```

### 3. 动态扩展器（导出）

```java
public class DynamicExpander {
    
    /**
     * 导出时动态扩展空间
     */
    public void expandAndFill(
        Sheet sheet, 
        Position headerPos, 
        List<Object> data,
        List<Position> knownHeaders) {
        
        int column = headerPos.getColumn();
        int startRow = headerPos.getRow() + 1;
        
        // 检查下方是否有其他表
        Position nextHeader = findNextHeaderBelow(startRow, column, knownHeaders, sheet);
        
        if (nextHeader != null) {
            int availableRows = nextHeader.getRow() - startRow;
            int requiredRows = data.size();
            
            if (requiredRows > availableRows) {
                // 需要扩展：下移下方的表
                int rowsToShift = requiredRows - availableRows;
                shiftBelowRows(sheet, nextHeader.getRow(), rowsToShift);
            }
        }
        
        // 填充数据
        for (int i = 0; i < data.size(); i++) {
            Row row = getOrCreateRow(sheet, startRow + i);
            Cell cell = getOrCreateCell(row, column);
            fillCell(cell, data.get(i));
        }
    }
    
    /**
     * 下移指定行及其下方的所有内容
     */
    private void shiftBelowRows(Sheet sheet, int fromRow, int rowsToShift) {
        // 从最后一行开始往上移，避免覆盖
        for (int rowNum = sheet.getLastRowNum(); rowNum >= fromRow; rowNum--) {
            Row oldRow = sheet.getRow(rowNum);
            if (oldRow == null) continue;
            
            Row newRow = sheet.getRow(rowNum + rowsToShift);
            if (newRow == null) {
                newRow = sheet.createRow(rowNum + rowsToShift);
            }
            
            // 复制整行
            copyRow(oldRow, newRow);
            
            // 清空原行（可选）
            // clearRow(oldRow);
        }
    }
}
```

---

## 五、完整示例

### 示例 1：提取配置

```yaml
version: "1.0"
templateName: "多表提取"

extractions:
  # 订单表
  - key: "orderNos"
    header:
      match: "订单号"
      searchRow: 1
    mode: DOWN
    parser:
      type: string
      
  - key: "amounts"
    header:
      match: "金额"
      searchRow: 1
    mode: DOWN
    parser:
      type: number
      
  # 客户表（在订单表下方）
  - key: "customers"
    header:
      match: "客户"
    mode: DOWN
    parser:
      type: string
```

**提取行为**：
```
Excel:
A1: 订单号 | B1: 金额
A2-A8: 订单数据 (7 条)
A9: (空)
A10: 客户

结果：
- orderNos: 提取 A2-A8 (7 条)
- amounts: 提取 B2-B8 (7 条)
- customers: 提取 A11 开始（遇到空行停止，则到 sheet 末尾）
```

### 示例 2：导出配置

```yaml
version: "1.0"

exports:
  # 订单表 - 数据量可能超过模板
  - key: "orderNos"
    header:
      match: "订单号"
    mode: FILL_DOWN
    # 不指定 maxRows，根据数据量自动扩展
    
  - key: "amounts"
    header:
      match: "金额"
    mode: FILL_DOWN
    
  # 客户表 - 在订单表下方
  - key: "customers"
    header:
      match: "客户"
    mode: FILL_DOWN
```

**导出行为**：
```
模板：
A1: 订单号 | B1: 金额
A2-A8: 预留 7 行
A10: 客户

数据：
{
  "orderNos": [20 条],  // 20 条数据
  "amounts": [20 条],
  "customers": [3 条]
}

行为：
1. 定位"订单号"表头 → A1
2. 检查下方：发现 A10 有"客户"表
3. 计算：需要 20 行，可用 8 行，需要下移 12 行
4. 下移：将 A10 及下方所有内容下移 12 行 → A10 移到 A22
5. 填充：A2-A21 填充订单数据
6. 定位"客户"表头 → 现在在 A22
7. 填充：A23-A25 填充客户数据
```

---

## 六、配置优先级

### 方式 1：表头匹配（推荐）

```yaml
# ✅ 推荐：通过表头匹配定位，不依赖固定位置
extractions:
  - key: "orderNos"
    header:
      match: "订单号"
    mode: DOWN
```

### 方式 2：固定位置

```yaml
# 固定位置：适用于格式固定的模板
extractions:
  - key: "orderNos"
    position:
      cellRef: "A1"
    mode: DOWN
```

### 方式 3：混合模式

```yaml
# 混合：优先表头匹配，找不到时使用备用位置
extractions:
  - key: "orderNos"
    header:
      match: "订单号"
    position:
      cellRef: "A1"   # 备用位置
    mode: DOWN
```

---

## 七、Java 实现骨架

### 配置类

```java
@Data
@Builder
public class ExtractConfig {
    // 字段名（映射到 JSON 结果的 key）
    private String key;
    
    // 表头匹配配置（方式 1）
    private HeaderConfig header;
    
    // 固定位置配置（方式 2）
    private PositionConfig position;
    
    // 提取模式
    private ExtractMode mode;
    
    // 范围配置
    private RangeConfig range;
    
    // 解析器配置
    private ParserConfig parser;
}

@Data
@Builder
public class HeaderConfig {
    // 匹配的表头文字
    private String match;
    
    // 是否部分匹配
    private boolean partialMatch;
    
    // 是否正则匹配
    private boolean regex;
    
    // 搜索行
    private Integer searchRow;
    
    // 搜索行范围
    private RowRange searchRows;
}
```

### 提取引擎

```java
public class ExtractEngine {
    
    private final HeaderLocator headerLocator;
    private final ExtractBoundaryDetector boundaryDetector;
    
    public Map<String, Object> extract(InputStream input, ExcelConfig config) {
        Workbook workbook = WorkbookFactory.create(input);
        Sheet sheet = workbook.getSheetAt(0);
        
        // 收集所有已知表头位置
        List<Position> knownHeaders = locateAllHeaders(sheet, config);
        
        Map<String, Object> result = new HashMap<>();
        
        for (ExtractConfig extract : config.getExtractions()) {
            // 1. 定位表头
            Position headerPos;
            if (extract.getHeader() != null) {
                headerPos = headerLocator.locate(sheet, extract.getHeader());
            } else {
                headerPos = extract.getPosition().toPosition();
            }
            
            // 2. 检测边界
            int endRow = boundaryDetector.detectEndRow(
                sheet, headerPos, knownHeaders);
            
            // 3. 执行提取
            List<Object> data = extractRange(
                sheet, headerPos, endRow, extract.getParser());
            
            result.put(extract.getKey(), data);
        }
        
        return result;
    }
    
    private List<Position> locateAllHeaders(Sheet sheet, ExcelConfig config) {
        List<Position> positions = new ArrayList<>();
        
        for (ExtractConfig extract : config.getExtractions()) {
            if (extract.getHeader() != null) {
                try {
                    Position pos = headerLocator.locate(sheet, extract.getHeader());
                    positions.add(pos);
                } catch (HeaderNotFoundException e) {
                    // 忽略，使用备用位置
                }
            }
        }
        
        return positions;
    }
}
```

### 导出引擎

```java
public class ExportEngine {
    
    private final HeaderLocator headerLocator;
    private final DynamicExpander expander;
    
    public byte[] export(InputStream template, Map<String, Object> data, ExcelConfig config) {
        Workbook workbook = WorkbookFactory.create(template);
        Sheet sheet = workbook.getSheetAt(0);
        
        // 收集所有已知表头位置
        List<Position> knownHeaders = locateAllHeaders(sheet, config);
        
        // 按行号排序（从下往上处理，避免覆盖）
        knownHeaders.sort(Comparator.comparingInt(Position::getRow).reversed());
        
        for (ExportConfig export : config.getExports()) {
            // 1. 定位表头
            Position headerPos = headerLocator.locate(sheet, export.getHeader());
            
            // 2. 获取数据
            List<Object> columnData = getColumnData(data, export.getKey());
            
            // 3. 动态扩展并填充
            expander.expandAndFill(sheet, headerPos, columnData, knownHeaders);
        }
        
        ByteArrayOutputStream output = new ByteArrayOutputStream();
        workbook.write(output);
        return output.toByteArray();
    }
}
```

---

## 八、总结

### 核心原则

| 原则 | 说明 |
|------|------|
| 表头匹配定位 | 配置通过表头文字匹配，不依赖固定位置 |
| 数据量驱动 | 提取/填充的行数由实际数据决定 |
| 自动扩展 | 导出时如果空间不足，自动下移下方内容 |
| 边界保护 | 提取时遇到已知表头位置或空行停止 |

### 配置对比

```yaml
# ❌ 旧方式：固定位置，不灵活
extractions:
  - key: "orderNos"
    position: { cellRef: "A1" }
    mode: DOWN

# ✅ 新方式：表头匹配，灵活
extractions:
  - key: "orderNos"
    header: { match: "订单号" }
    mode: DOWN
```

### 解决的问题

1. ✅ **表头位置不固定** → 通过表头文字匹配定位
2. ✅ **数据量超过模板** → 自动下移下方内容，动态扩展
3. ✅ **多表共存** → 已知表头位置作为边界参考，自动协调
