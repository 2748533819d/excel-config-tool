# Excel Config Tool - 配置驱动的边界检测

> **核心原则：边界由配置决定，而不是自动检测**

---

## 一、问题场景

### 场景 1：同一列有多个提取区域

```
配置：
- A1: DOWN 模式（提取订单表）
- A10: DOWN 模式（提取客户表）

Excel 布局：
┌─────────────────────────┐
│ A1: 订单号表头          │ ← 配置点 1 (DOWN)
├─────────────────────────┤
│ A2-A8: 订单数据 (7 条)   │
│ A9: (空行)              │
├─────────────────────────┤
│ A10: 客户统计表头       │ ← 配置点 2 (DOWN)
├─────────────────────────┤
│ A11-A15: 客户数据       │
└─────────────────────────┘

期望行为：
- A1 的 DOWN 提取：A1 到 A9（下一个配置点之前）
- A10 的 DOWN 提取：A10 到 sheet 末尾（或下一个配置点）
```

### 场景 2：数据量超过预期

```
配置：
- A1: DOWN 模式
- A10: DOWN 模式

实际问题：
- A1 的订单数据有 20 条（不是 7 条）
- A10 的客户表被覆盖！

┌─────────────────────────┐
│ A1: 订单号表头          │
├─────────────────────────┤
│ A2-A21: 订单数据 (20 条) │ ← 超长了！
│ A10: 客户统计表头       │ ← 被覆盖！
│ A11-A15: 客户数据       │ ← 被覆盖！
└─────────────────────────┘
```

---

## 二、设计方案

### 设计 1：配置点即边界

**核心思想**：每个配置点定义了一个区域的起点，下一个配置点就是当前区域的终点。

```yaml
extractions:
  # 订单表
  - key: "orderNos"
    position: { cellRef: "A1" }
    mode: DOWN
    range:
      # 默认行为：提取到下一个配置点之前
      # 自动检测：下一个配置点在 A10，所以提取 A1-A9
      
  # 客户表
  - key: "customers"
    position: { cellRef: "A10" }
    mode: DOWN
    range:
      # 没有下一个配置点，提取到 sheet 末尾
```

### 设计 2：显式指定结束位置

```yaml
extractions:
  # 订单表
  - key: "orderNos"
    position: { cellRef: "A1" }
    mode: DOWN
    range:
      endPosition: { cellRef: "A9" }  # 显式指定结束位置
      
  # 客户表
  - key: "customers"
    position: { cellRef: "A10" }
    mode: DOWN
```

### 设计 3：配置点自动检测 + 安全限制

```yaml
extractions:
  # 订单表
  - key: "orderNos"
    position: { cellRef: "A1" }
    mode: DOWN
    range:
      # 自动检测下一个配置点
      stopAtNextConfigPoint: true
      
      # 安全限制：最多提取多少行
      maxRows: 100
      
  # 客户表
  - key: "customers"
    position: { cellRef: "A10" }
    mode: DOWN
    range:
      maxRows: 100
```

---

## 三、核心机制

### 1. 配置点收集

```java
public class ConfigPointCollector {
    
    /**
     * 收集所有配置点，按位置排序
     */
    public List<ConfigPoint> collect(ExcelConfig config) {
        List<ConfigPoint> points = new ArrayList<>();
        
        for (ExtractConfig extract : config.getExtractions()) {
            Position pos = extract.getPosition();
            points.add(new ConfigPoint(pos, extract));
        }
        
        // 按列分组，每列内按行排序
        Map<Integer, List<ConfigPoint>> byColumn = points.stream()
            .collect(Collectors.groupingBy(
                p -> p.getPosition().getColumn(),
                TreeMap::new,
                Collectors.toList()
            ));
        
        // 每列内按行号排序
        byColumn.forEach((col, list) -> 
            list.sort(Comparator.comparingInt(p -> p.getPosition().getRow())));
        
        return points;
    }
}
```

### 2. 提取范围计算

```java
public class ExtractRangeCalculator {
    
    public Range calculateRange(ExtractConfig config, List<ConfigPoint> allPoints) {
        Position start = config.getPosition();
        int column = start.getColumn();
        int startRow = start.getRow();
        
        // 找到同列的下一个配置点
        ConfigPoint nextPoint = findNextPointInSameColumn(column, startRow, allPoints);
        
        if (nextPoint != null) {
            // 有下一个配置点：提取到下一个配置点之前
            int endRow = nextPoint.getPosition().getRow() - 1;
            return Range.builder()
                .startRow(startRow + 1)  // 表头下方开始
                .endRow(endRow)
                .maxRows(endRow - startRow)  // 硬限制
                .build();
        } else {
            // 没有下一个配置点：使用 skipEmpty 或 maxRows
            return calculateDynamicRange(config, start);
        }
    }
    
    private ConfigPoint findNextPointInSameColumn(
        int column, 
        int startRow, 
        List<ConfigPoint> allPoints) {
        
        return allPoints.stream()
            .filter(p -> p.getPosition().getColumn() == column)
            .filter(p -> p.getPosition().getRow() > startRow)
            .min(Comparator.comparingInt(p -> p.getPosition().getRow()))
            .orElse(null);
    }
}
```

### 3. 导出数据量检查

```java
public class DataValidator {
    
    /**
     * 检查数据量是否会覆盖下一个配置点
     */
    public ValidationResult validate(ExportConfig config, Object data, List<ConfigPoint> allPoints) {
        int actualRows = getDataRowCount(data);
        Position start = config.getPosition();
        
        // 找到同列的下一个配置点
        ConfigPoint nextPoint = findNextPointInSameColumn(start.getColumn(), start.getRow(), allPoints);
        
        if (nextPoint != null) {
            int availableRows = nextPoint.getPosition().getRow() - start.getRow() - 1;
            
            if (actualRows > availableRows) {
                return ValidationResult.error(
                    String.format("数据量 (%d 行) 超过可用空间 (%d 行)，会覆盖下一个配置点 %s",
                        actualRows, availableRows, nextPoint.getPosition()));
            }
        }
        
        return ValidationResult.ok();
    }
}
```

---

## 四、配置语法

### 基础配置（自动检测边界）

```yaml
extractions:
  # 订单表 - 自动在 A10 之前停止
  - key: "orderNos"
    position: { cellRef: "A1" }
    mode: DOWN
    
  # 客户表 - 从 A10 开始
  - key: "customers"
    position: { cellRef: "A10" }
    mode: DOWN
```

### 显式指定边界

```yaml
extractions:
  # 订单表 - 显式指定提取到 A9
  - key: "orderNos"
    position: { cellRef: "A1" }
    mode: DOWN
    range:
      endPosition: { cellRef: "A9" }
      
  # 客户表
  - key: "customers"
    position: { cellRef: "A10" }
    mode: DOWN
```

### 安全限制

```yaml
extractions:
  # 订单表 - 最多提取 100 行，即使后面没有配置点
  - key: "orderNos"
    position: { cellRef: "A1" }
    mode: DOWN
    range:
      maxRows: 100
      
  # 客户表
  - key: "customers"
    position: { cellRef: "A10" }
    mode: DOWN
```

### 导出配置

```yaml
exports:
  # 订单表 - 最多填充 100 行，防止覆盖下方内容
  - key: "orderNos"
    position: { cellRef: "A1" }
    mode: FILL_DOWN
    maxRows: 100
    
  # 客户表
  - key: "customers"
    position: { cellRef: "A10" }
    mode: FILL_DOWN
```

---

## 五、完整示例

### 示例 1：同一列多个表

```yaml
version: "1.0"
templateName: "多表提取示例"

extractions:
  # 表 1：订单表 (A1-A9)
  - key: "orderTable"
    position: { cellRef: "A1" }
    mode: DOWN
    parser: { type: string }
    
  # 表 2：客户表 (A10-A19)
  - key: "customerTable"
    position: { cellRef: "A10" }
    mode: DOWN
    parser: { type: string }
    
  # 表 3：产品表 (A20-A29)
  - key: "productTable"
    position: { cellRef: "A20" }
    mode: DOWN
    parser: { type: string }
```

### 示例 2：多列配置

```yaml
version: "1.0"

extractions:
  # A 列：订单号
  - key: "orderNos"
    position: { cellRef: "A1" }
    mode: DOWN
    
  # B 列：金额
  - key: "amounts"
    position: { cellRef: "B1" }
    mode: DOWN
    
  # A 列第二个表
  - key: "summary"
    position: { cellRef: "A20" }
    mode: DOWN
    
  # B 列第二个表
  - key: "metrics"
    position: { cellRef: "B20" }
    mode: DOWN
```

### 示例 3：导出 + 安全限制

```yaml
version: "1.0"

exports:
  # 订单表 - 限制最大行数
  - key: "orderNos"
    position: { cellRef: "A1" }
    mode: FILL_DOWN
    maxRows: 100  # 防止覆盖下方内容
    
  # 汇总行 - 固定位置
  - key: "totalAmount"
    position: { cellRef: "A102" }  # 在订单表下方
    mode: FILL_CELL
```

---

## 六、实现流程

```
┌─────────────────────────────────────────────────────────────┐
│                    配置驱动的提取流程                        │
├─────────────────────────────────────────────────────────────┤
│                                                             │
│  1. 收集所有配置点                                          │
│     ↓                                                       │
│  2. 按列分组，按行排序                                       │
│     ↓                                                       │
│  3. 对每个配置点，找到同列的下一个配置点                     │
│     ↓                                                       │
│  4. 计算提取范围：当前位置 → 下一个配置点 -1                 │
│     ↓                                                       │
│  5. 执行提取                                                │
│                                                             │
└─────────────────────────────────────────────────────────────┘

┌─────────────────────────────────────────────────────────────┐
│                    配置驱动的导出流程                        │
├─────────────────────────────────────────────────────────────┤
│                                                             │
│  1. 收集所有配置点                                          │
│     ↓                                                       │
│  2. 验证数据量：检查是否超过可用空间                        │
│     ↓                                                       │
│  3. 如果有超出的风险：                                       │
│     - 发出警告                                              │
│     - 截断数据（应用 maxRows 限制）                          │
│     ↓                                                       │
│  4. 执行填充（从下往上，避免覆盖）                           │
│                                                             │
└─────────────────────────────────────────────────────────────┘
```

---

## 七、Java 实现骨架

### 配置点

```java
@Data
@Builder
public class ConfigPoint {
    // 位置
    private Position position;
    
    // 所属配置
    private ExtractConfig extractConfig;
    private ExportConfig exportConfig;
    
    // 配置类型
    private ConfigType type;
    
    public enum ConfigType {
        EXTRACT,  // 提取配置
        EXPORT    // 导出配置
    }
}
```

### 范围计算器

```java
public class RangeCalculator {
    
    public ExtractRange calculate(
        ExtractConfig config, 
        List<ConfigPoint> sortedPoints) {
        
        Position start = config.getPosition();
        int column = start.getColumn();
        int startRow = start.getRow();
        
        // 找到同列下一个配置点
        ConfigPoint nextPoint = findNextInColumn(column, startRow, sortedPoints);
        
        if (nextPoint != null) {
            // 有下一个配置点：提取范围是当前点到下一个点之前
            int endRow = nextPoint.getPosition().getRow() - 1;
            int maxRows = endRow - startRow;
            
            return ExtractRange.builder()
                .startRow(startRow + 1)  // 跳过表头
                .endRow(endRow)
                .maxRows(maxRows)
                .build();
        } else {
            // 没有下一个配置点：动态检测
            return ExtractRange.builder()
                .startRow(startRow + 1)
                .endRow(null)  // 不指定，动态检测
                .maxRows(config.getRange().getMaxRows())  // 使用配置的 maxRows
                .build();
        }
    }
    
    private ConfigPoint findNextInColumn(
        int column, 
        int currentRow, 
        List<ConfigPoint> points) {
        
        return points.stream()
            .filter(p -> p.getPosition().getColumn() == column)
            .filter(p -> p.getPosition().getRow() > currentRow)
            .findFirst()
            .orElse(null);
    }
}
```

### 数据验证器

```java
public class ExportValidator {
    
    public void validateBeforeFill(
        ExportConfig config, 
        Object data, 
        List<ConfigPoint> allPoints) {
        
        int actualRows = getDataRowCount(data);
        Position start = config.getPosition();
        
        // 找到同列下一个配置点
        ConfigPoint nextPoint = findNextInColumn(start.getColumn(), start.getRow(), allPoints);
        
        if (nextPoint != null) {
            int availableRows = nextPoint.getPosition().getRow() - start.getRow() - 1;
            
            if (actualRows > availableRows) {
                throw new ExcelConfigException(
                    String.format("数据量过大：%s 位置的数据有 %d 行，但可用空间只有 %d 行（下一个配置点在 %s）",
                        start.getCellRef(), actualRows, availableRows, 
                        nextPoint.getPosition().getCellRef()));
            }
        }
        
        // 检查 maxRows 限制
        if (config.getMaxRows() != null && actualRows > config.getMaxRows()) {
            log.warn("数据量 {} 超过 maxRows 限制 {}，将截断到 {} 行", 
                actualRows, config.getMaxRows(), config.getMaxRows());
        }
    }
}
```

---

## 八、总结

### 核心原则

1. **配置点即边界** - 每个配置点定义了一个区域的起点，下一个配置点就是当前区域的终点
2. **显式优于隐式** - 可以显式指定 `endPosition`，但默认自动检测
3. **安全第一** - `maxRows` 作为最后的保护网

### 配置建议

```yaml
# ✅ 推荐：同一列多个表，自动检测边界
extractions:
  - key: "table1"
    position: { cellRef: "A1" }
    mode: DOWN
    
  - key: "table2"
    position: { cellRef: "A10" }
    mode: DOWN

# ✅ 推荐：导出时限制最大行数
exports:
  - key: "orderNos"
    position: { cellRef: "A1" }
    mode: FILL_DOWN
    maxRows: 100  # 防止数据过多覆盖下方内容
```

### 注意事项

1. **配置点冲突检测** - 如果两个配置点太近，应该警告用户
2. **数据量验证** - 导出前验证数据量是否超过可用空间
3. **填充顺序** - 从下往上填充，避免覆盖
