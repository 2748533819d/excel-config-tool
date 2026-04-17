# Excel Config Tool - 导入导出行数动态确定机制

> **核心原则：配置不写死行数，行数由实际数据决定**

---

## 一、问题背景

### 用户痛点

```yaml
# ❌ 错误设计 - 配置时无法预知数据量
exports:
  - key: "orderNos"
    position: { cellRef: "A2" }
    mode: FILL_DOWN
    range: { rows: 10 }  # 问题：数据可能只有 5 条，也可能有 100 条

extractions:
  - key: "orderNos"
    position: { cellRef: "A2" }
    mode: DOWN
    range: { rows: 100 }  # 问题：实际数据可能只有 50 行
```

### 正确设计

```yaml
# ✅ 正确设计 - 行数由数据自动决定
exports:
  - key: "orderNos"
    position: { cellRef: "A2" }
    mode: FILL_DOWN
    # 不指定 rows，根据 orderNos 数组长度自动填充

extractions:
  - key: "orderNos"
    position: { cellRef: "A2" }
    mode: DOWN
    range:
      skipEmpty: true  # 遇到空行停止
      maxRows: 1000    # 可选：安全限制
```

---

## 二、导入（提取）模式 - 行数确定机制

### 1. skipEmpty 模式（默认推荐）

```yaml
extractions:
  - key: "orderNos"
    position: { cellRef: "A2" }
    mode: DOWN
    range:
      skipEmpty: true  # 一直读取，直到遇到空行
```

**行为**：
```
Excel 数据：
A2: ORD001
A3: ORD002
A4: ORD003
A5: (空)
A6: ORD005  ← 不会被读取

提取结果：["ORD001", "ORD002", "ORD003"]
读取行数：3 行（由数据决定）
```

### 2. fixed rows 模式（固定行数）

```yaml
extractions:
  - key: "fixedData"
    position: { cellRef: "A2" }
    mode: DOWN
    range:
      rows: 10  # 固定读取 10 行
```

**适用场景**：固定格式的表格，如月历、固定 12 个月的报表等

### 3. maxRows 限制（安全保护）

```yaml
extractions:
  - key: "largeData"
    position: { cellRef: "A2" }
    mode: DOWN
    range:
      skipEmpty: true
      maxRows: 10000  # 最多读取 10000 行，防止内存溢出
```

**适用场景**：处理可能非常大的数据文件

### 4. untilCondition 模式（条件停止）

```yaml
extractions:
  - key: "data"
    position: { cellRef: "A2" }
    mode: DOWN
    range:
      untilCondition:
        column: "C"
        contains: "总计"  # 读取到包含"总计"的行停止
```

**适用场景**：数据区域有明确的结束标记

---

## 三、导出（填充）模式 - 行数确定机制

### 1. 数据驱动模式（默认）

```yaml
exports:
  - key: "orderNos"
    position: { cellRef: "A2" }
    mode: FILL_DOWN
    # 不指定 rows，根据 orderNos 数组长度自动填充
```

**行为**：
```java
// 传入数据
{
  "orderNos": ["ORD001", "ORD002", "ORD003", "ORD004", "ORD005"]
}

// 自动填充 5 行
A2: ORD001
A3: ORD002
A4: ORD003
A5: ORD004
A6: ORD005
```

### 2. maxRows 限制（安全保护）

```yaml
exports:
  - key: "largeData"
    position: { cellRef: "A2" }
    mode: FILL_DOWN
    maxRows: 1000  # 最多填充 1000 行
```

**行为**：
```java
// 如果数据有 2000 条，只填充前 1000 行
// 如果数据只有 500 条，填充 500 行
```

### 3. 多列独立行数（列隔离）

```yaml
exports:
  - key: "orderNos"
    position: { cellRef: "A2" }
    mode: FILL_DOWN
    
  - key: "amounts"
    position: { cellRef: "B2" }
    mode: FILL_DOWN
```

**行为**：
```java
// 传入不同长度的数组
{
  "orderNos": ["ORD001", "ORD002", "ORD003", "ORD004", "ORD005"],  // 5 条
  "amounts": [100, 200, 300]  // 3 条
}

// 结果：
// A 列填充 5 行 → A2:A6
// B 列填充 3 行 → B2:B4
// 每列独立计算，互不影响
```

---

## 四、实现机制

### 1. 导入行数计算

```java
public class ExtractRangeCalculator {
    
    public int calculateRows(Sheet sheet, Position start, RangeConfig range) {
        // 1. 固定行数模式
        if (range.getRows() != null) {
            return range.getRows();
        }
        
        // 2. skipEmpty 模式
        if (range.isSkipEmpty()) {
            return countUntilEmpty(sheet, start, range.getMaxRows());
        }
        
        // 3. untilCondition 模式
        if (range.getUntilCondition() != null) {
            return countUntilCondition(sheet, start, range.getUntilCondition());
        }
        
        // 默认：读取到 sheet 末尾
        return sheet.getLastRowNum() - start.getRow();
    }
    
    private int countUntilEmpty(Sheet sheet, Position start, Integer maxRows) {
        int count = 0;
        int max = maxRows != null ? maxRows : Integer.MAX_VALUE;
        
        for (int i = start.getRow(); i < sheet.getLastRowNum() && count < max; i++) {
            Row row = sheet.getRow(i);
            if (row == null || isCellEmpty(row, start.getColumn())) {
                break;
            }
            count++;
        }
        
        return count;
    }
}
```

### 2. 导出数据行数计算

```java
public class FillRowCountCalculator {
    
    public int calculateRows(Object data, ExportConfig config) {
        // 1. 获取数据
        Object value = data.get(config.getKey());
        
        // 2. 根据数据类型计算行数
        if (value instanceof List) {
            int size = ((List<?>) value).size();
            // 应用 maxRows 限制
            if (config.getMaxRows() != null) {
                return Math.min(size, config.getMaxRows());
            }
            return size;
        } else if (value.getClass().isArray()) {
            int size = Array.getLength(value);
            if (config.getMaxRows() != null) {
                return Math.min(size, config.getMaxRows());
            }
            return size;
        } else {
            // 单个值，按 1 行处理
            return 1;
        }
    }
}
```

---

## 五、配置建议

### 导入配置最佳实践

```yaml
# ✅ 推荐：适用于大多数场景
extractions:
  - key: "orderNos"
    position: { cellRef: "A2" }
    mode: DOWN
    range:
      skipEmpty: true  # 自动检测数据结束
      maxRows: 10000   # 安全限制，防止读取过多

# ✅ 推荐：固定格式报表
extractions:
  - key: "monthHeaders"
    position: { cellRef: "B1" }
    mode: RIGHT
    range:
      cols: 12  # 固定 12 个月

# ✅ 推荐：有结束标记的数据
extractions:
  - key: "detailData"
    position: { cellRef: "A2" }
    mode: DOWN
    range:
      untilCondition:
        column: "A"
        equals: "总计行"
```

### 导出配置最佳实践

```yaml
# ✅ 推荐：数据驱动，自动计算行数
exports:
  - key: "orderNos"
    position: { cellRef: "A2" }
    mode: FILL_DOWN
    # 不指定 rows，由数据量决定

# ✅ 推荐：大量数据时限制最大行数
exports:
  - key: "reportData"
    position: { cellRef: "A2" }
    mode: FILL_DOWN
    maxRows: 1000  # 防止填充过多行导致 Excel 过大
```

---

## 六、导入导出对比

| 方面 | 导入（提取） | 导出（填充） |
|------|-------------|-------------|
| 行数来源 | Excel 中的实际数据行数 | 传入数据的数组长度 |
| 检测方式 | skipEmpty / untilCondition | 数组 size() |
| 安全限制 | maxRows | maxRows |
| 固定行数 | rows（固定格式） | 不适用 |
| 列隔离 | 每列独立读取 | 每列独立填充 |

---

## 七、总结

### 核心原则

1. **配置不写死行数** - 行数由实际数据决定
2. **skipEmpty 是默认行为** - 遇到空行/空列自动停止
3. **maxRows 是安全网** - 防止意外读取/填充过多数据
4. **列隔离** - 每列的行数独立计算

### 配置简化对比

```yaml
# ❌ 旧设计 - 复杂且不灵活
exports:
  - key: "orderNos"
    mode: FILL_DOWN
    range: { rows: 10 }  # 需要预知行数

# ✅ 新设计 - 简洁且灵活
exports:
  - key: "orderNos"
    mode: FILL_DOWN
    # 自动根据数据量填充
    maxRows: 1000  # 可选：安全限制
```
