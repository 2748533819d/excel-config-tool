# Excel Config Tool - 数据区域边界检测机制

> 解决"数据区域末尾不是空行"的复杂场景

---

## 一、问题场景

### 场景 1：数据 + 汇总行

```
┌─────────────────────────────────────────┐
│  订单明细表                              │
│  ┌───────┬────────┬────────┬─────────┐  │
│  │ 订单号 │ 金额   │ 日期   │ 客户    │  │ ← 表头 (row 1)
│  ├───────┼────────┼────────┼─────────┤  │
│  │ ORD001│ 100.00 │01-01   │ 张三    │  │ ← 数据 (row 2)
│  │ ORD002│ 200.00 │01-02   │ 李四    │  │ ← 数据 (row 3)
│  │ ORD003│ 300.00 │01-03   │ 王五    │  │ ← 数据 (row 4)
│  ├───────┴────────┴────────┴─────────┤  │
│  │ 总计：3 条订单                      │  │ ← 汇总 (row 5) ❌ 问题：不是空行！
│  └─────────────────────────────────┘    │
```

**问题**：`skipEmpty: true` 会在 row 5 停止，但 row 5 有数据（汇总行）！

### 场景 2：数据 + 新表头

```
┌─────────────────────────────────────────┐
│  订单明细表                              │
│  ┌───────┬────────┬────────┬─────────┐  │
│  │ 订单号 │ 金额   │ 日期   │ 客户    │  │ ← 表头 1 (row 1)
│  ├───────┼────────┼────────┼─────────┤  │
│  │ ORD001│ 100.00 │01-01   │ 张三    │  │ ← 数据 (row 2)
│  │ ORD002│ 200.00 │01-02   │ 李四    │  │ ← 数据 (row 3)
│  │ ORD003│ 300.00 │01-03   │ 王五    │  │ ← 数据 (row 4)
│  └───────┴────────┴────────┴─────────┘  │
│  客户统计表                              │ ← 新表头 (row 5) ❌ 问题：新表开始！
│  ┌───────┬────────┐                     │
│  │ 客户   │ 订单数  │                     │
│  └───────┴────────┘                     │
```

**问题**：新表头在 row 5，但它不是空行，如何区分？

### 场景 3：数据中间有空行

```
┌─────────────────────────────────────────┐
│  订单明细表                              │
│  ┌───────┬────────┬────────┬─────────┐  │
│  │ 订单号 │ 金额   │ 日期   │ 客户    │  │
│  ├───────┼────────┼────────┼─────────┤  │
│  │ ORD001│ 100.00 │01-01   │ 张三    │
│  │ ORD002│ 200.00 │01-02   │ 李四    │
│  │       │        │        │         │  │ ← 空行（可能是误操作）
│  │ ORD003│ 300.00 │01-03   │ 王五    │
│  │ ORD004│ 400.00 │01-04   │ 赵六    │
│  └─────────────────────────────────────┘
```

**问题**：`skipEmpty: true` 会在空行停止，但后面还有数据！

---

## 二、解决方案总览

| 方案 | 适用场景 | 配置复杂度 | 准确率 |
|------|----------|-----------|--------|
| 1. untilCondition | 有明确结束标记 | 低 | 高 |
| 2. columnPattern | 多列数据对齐 | 中 | 高 |
| 3. blockDetection | 数据块边界明显 | 中 | 中 |
| 4. headerDetection | 新表头特征明显 | 高 | 高 |
| 5. manualRange | 固定格式 | 低 | 100% |

---

## 三、方案详解

### 方案 1：untilCondition（条件停止）

**适用场景**：数据区域有明确的结束标记

#### 配置示例

```yaml
extractions:
  - key: "orderNos"
    position: { cellRef: "A2" }
    mode: DOWN
    range:
      untilCondition:
        # 停止条件：A 列包含"总计"
        column: "A"
        contains: "总计"
        
  - key: "amounts"
    position: { cellRef: "B2" }
    mode: DOWN
    range:
      untilCondition:
        # 停止条件：B 列包含"总计"
        column: "B"
        contains: "总计"
```

#### 支持的条件类型

```yaml
# 1. 包含文本
untilCondition:
  column: "A"
  contains: "总计"

# 2. 完全匹配
untilCondition:
  column: "A"
  equals: "总计："

# 3. 正则匹配
untilCondition:
  column: "A"
  matches: "^总计：.*订单.*$"

# 4. 单元格样式（如加粗、背景色）
untilCondition:
  column: "A"
  style:
    bold: true
    background: "#FFFF00"

# 5. 公式
untilCondition:
  column: "A"
  formula: "ISNUMBER(SEARCH('总计', A1))"
```

#### 多条件组合

```yaml
# AND 条件：同时满足才停止
untilCondition:
  all:
    - column: "A"
      contains: "总计"
    - column: "B"
      isEmpty: true

# OR 条件：满足任一就停止
untilCondition:
  any:
    - column: "A"
      contains: "总计"
    - column: "A"
      contains: "备注："
```

---

### 方案 2：columnPattern（列模式检测）

**适用场景**：多列数据，通过列的对齐模式判断边界

#### 原理

```
数据行特征：
- A 列有值 && B 列有值 && C 列有值 → 数据行
- A 列有值 && B 列为空 && C 列为空 → 可能是汇总行/边界

汇总行特征：
- 只有 A 列有值（"总计"）
- B 列、C 列为空或合并
```

#### 配置示例

```yaml
extractions:
  - key: "orderNos"
    position: { cellRef: "A2" }
    mode: DOWN
    range:
      columnPattern:
        # 定义数据行的模式：所有列都有值
        dataPattern:
          columns: ["A", "B", "C", "D"]
          allFilled: true  # 所有列都有值
        
        # 定义停止模式：某一列开始为空
        stopPattern:
          columns: ["A"]
          hasValue: true   # A 列有值
          adjacentEmpty: true  # 相邻列为空（可能是汇总）
```

#### 智能检测

```yaml
extractions:
  - key: "orderData"
    position: { cellRef: "A2" }
    mode: DOWN
    range:
      columnPattern:
        # 数据行：A、B、C、D 都有值
        requiredColumns: ["A", "B", "C", "D"]
        
        # 允许的连续空行数（容忍中间的空行）
        allowConsecutiveEmptyRows: 1
        
        # 停止条件：连续 2 行不符合数据模式
        stopAfterMismatchRows: 2
```

---

### 方案 3：blockDetection（数据块检测）

**适用场景**：数据形成明显的块状区域

#### 原理

```
数据块边界检测：
┌─────────┐
│ 表头    │ ← 上边界：有表头特征
├─────────┤
│ 数据行  │
│ 数据行  │ ← 数据块内部：连续的数据行
│ 数据行  │
├─────────┤
│ 汇总行  │ ← 下边界：样式/模式变化
└─────────┘
```

#### 配置示例

```yaml
extractions:
  - key: "orderData"
    position: { cellRef: "A2" }
    mode: BLOCK
    range:
      blockDetection:
        # 上边界：表头特征
        topBoundary:
          row: 1
          isHeader: true  # 自动检测表头（加粗、背景色）
        
        # 下边界：汇总行特征
        bottomBoundary:
          style:
            bold: true
            background: "#FFFF00"
          contains: ["总计", "合计", "汇总"]
        
        # 左右边界
        leftColumn: "A"
        rightColumn: "D"
```

---

### 方案 4：headerDetection（新表头检测）

**适用场景**：同一工作表有多个表格

#### 原理

```
表头特征检测：
1. 整行加粗
2. 有背景色
3. 连续的非空单元格
4. 下方有边框线
5. 文本特征（如包含"表"、"统计"、"明细"等）
```

#### 配置示例

```yaml
extractions:
  - key: "firstTable"
    position: { cellRef: "A2" }
    mode: DOWN
    range:
      stopAtNextHeader: true  # 遇到新表头停止
      
      # 表头特征
      headerPattern:
        bold: true
        hasBackgroundColor: true
        borderBottom: true
        textContains: ["表", "统计", "明细", "汇总"]
```

#### 表头检测算法

```java
public class HeaderDetector {
    
    public boolean isHeader(Row row) {
        int score = 0;
        
        // 特征 1：整行加粗
        if (isRowBold(row)) score += 30;
        
        // 特征 2：有背景色
        if (hasBackgroundColor(row)) score += 20;
        
        // 特征 3：下方有边框
        if (hasBottomBorder(row)) score += 20;
        
        // 特征 4：连续非空单元格 >= 2
        if (countConsecutiveNonEmpty(row) >= 2) score += 15;
        
        // 特征 5：文本包含表头关键词
        if (containsHeaderKeyword(row)) score += 15;
        
        return score >= 60;  // 60 分以上判定为表头
    }
}
```

---

### 方案 5：manualRange（手动指定）

**适用场景**：固定格式的模板

#### 配置示例

```yaml
extractions:
  - key: "orderNos"
    position: { cellRef: "A2" }
    mode: DOWN
    range:
      fixed:
        rows: 10  # 固定 10 行数据
        
  - key: "monthHeaders"
    position: { cellRef: "B1" }
    mode: RIGHT
    range:
      fixed:
        cols: 12  # 固定 12 列（12 个月）
```

---

## 四、推荐方案组合

### 最佳实践配置

```yaml
version: "1.0"
extractions:
  # 主数据提取：组合使用多种检测
  - key: "orderNos"
    position: { cellRef: "A2" }
    mode: DOWN
    range:
      # 首选：条件停止（如果有汇总行）
      untilCondition:
        column: "A"
        contains: ["总计", "合计"]
      
      # 备选：表头检测（如果有新表）
      stopAtNextHeader: true
      
      # 保底：最大行数限制
      maxRows: 1000
      
      # 容错：允许中间有 1 个空行
      allowConsecutiveEmptyRows: 1
```

### 智能检测流程

```
┌─────────────────────────────────────────────────────────────┐
│                    智能边界检测流程                          │
├─────────────────────────────────────────────────────────────┤
│                                                             │
│  1. 检查 untilCondition 配置                                 │
│     ↓ 有配置                                                 │
│  2. 扫描直到满足条件                                        │
│     ↓ 满足条件                                               │
│  3. 停止并返回结果                                          │
│                                                             │
│  1. 检查 stopAtNextHeader                                   │
│     ↓ 有配置                                                 │
│  2. 逐行检测是否是表头                                      │
│     ↓ 检测到表头                                             │
│  3. 停止并返回结果                                          │
│                                                             │
│  1. 检查 columnPattern                                      │
│     ↓ 有配置                                                 │
│  2. 检测数据模式是否匹配                                    │
│     ↓ 连续 2 行不匹配                                        │
│  3. 停止并返回结果                                          │
│                                                             │
│  默认：skipEmpty: true（遇到空行停止）                       │
│                                                             │
└─────────────────────────────────────────────────────────────┘
```

---

## 五、Java 实现骨架

### 边界检测器接口

```java
public interface BoundaryDetector {
    /**
     * 检测是否应该停止
     * @param row 当前行
     * @param context 检测上下文
     * @return true=停止，false=继续
     */
    boolean shouldStop(Row row, DetectionContext context);
    
    /**
     * 检测优先级（数字越小优先级越高）
     */
    int getPriority();
}
```

### 条件停止检测器

```java
public class ConditionBoundaryDetector implements BoundaryDetector {
    
    private final UntilConditionConfig config;
    
    @Override
    public boolean shouldStop(Row row, DetectionContext context) {
        Cell cell = row.getCell(config.getColumn());
        if (cell == null) return false;
        
        String value = getCellValueAsString(cell);
        
        // 包含检测
        if (config.getContains() != null) {
            for (String keyword : config.getContains()) {
                if (value.contains(keyword)) return true;
            }
        }
        
        // 完全匹配检测
        if (config.getEquals() != null) {
            if (value.equals(config.getEquals())) return true;
        }
        
        // 正则检测
        if (config.getMatches() != null) {
            if (value.matches(config.getMatches())) return true;
        }
        
        return false;
    }
    
    @Override
    public int getPriority() {
        return 1;  // 最高优先级
    }
}
```

### 表头检测器

```java
public class HeaderBoundaryDetector implements BoundaryDetector {
    
    private final HeaderPatternConfig config;
    
    @Override
    public boolean shouldStop(Row row, DetectionContext context) {
        // 检测是否是表头
        if (isHeader(row)) {
            return true;  // 遇到新表头，停止
        }
        return false;
    }
    
    private boolean isHeader(Row row) {
        int score = calculateHeaderScore(row);
        return score >= 60;
    }
    
    private int calculateHeaderScore(Row row) {
        int score = 0;
        
        // 特征 1：整行加粗（+30）
        if (isRowBold(row)) score += 30;
        
        // 特征 2：有背景色（+20）
        if (hasBackgroundColor(row)) score += 20;
        
        // 特征 3：下方有边框（+20）
        if (hasBottomBorder(row)) score += 20;
        
        // 特征 4：连续非空单元格（+15）
        if (countConsecutiveNonEmpty(row) >= 2) score += 15;
        
        // 特征 5：包含表头关键词（+15）
        if (containsHeaderKeyword(row)) score += 15;
        
        return score;
    }
    
    @Override
    public int getPriority() {
        return 2;  // 次高优先级
    }
}
```

### 智能提取引擎

```java
public class SmartExtractEngine {
    
    private final List<BoundaryDetector> detectors;
    
    public SmartExtractEngine() {
        detectors = new ArrayList<>();
        detectors.add(new ConditionBoundaryDetector());  // 优先级 1
        detectors.add(new HeaderBoundaryDetector());     // 优先级 2
        detectors.add(new PatternBoundaryDetector());    // 优先级 3
        detectors.add(new EmptyRowBoundaryDetector());   // 优先级 4（默认）
    }
    
    public List<Object> extract(Sheet sheet, ExtractConfig config, Map<String, Object> context) {
        Position start = config.getPosition().getCellRef();
        int columnIndex = start.getColumn();
        List<Object> result = new ArrayList<>();
        
        // 按优先级排序检测器
        detectors.sort(Comparator.comparingInt(BoundaryDetector::getPriority));
        
        // 选择启用的检测器
        List<BoundaryDetector> activeDetectors = selectActiveDetectors(config, detectors);
        
        for (int rowNum = start.getRow(); ; rowNum++) {
            Row row = sheet.getRow(rowNum);
            if (row == null) break;
            
            Cell cell = row.getCell(columnIndex);
            
            // 检查是否应该停止
            for (BoundaryDetector detector : activeDetectors) {
                if (detector.shouldStop(row, new DetectionContext(sheet, rowNum, config))) {
                    return result;  // 停止提取
                }
            }
            
            // 提取数据
            if (cell != null) {
                result.add(parseCell(cell, config.getParser()));
            }
        }
        
        return result;
    }
    
    private List<BoundaryDetector> selectActiveDetectors(
        ExtractConfig config, 
        List<BoundaryDetector> allDetectors) {
        
        List<BoundaryDetector> active = new ArrayList<>();
        
        // 如果配置了 untilCondition，启用条件检测器
        if (config.getRange().getUntilCondition() != null) {
            active.add(allDetectors.stream()
                .filter(d -> d instanceof ConditionBoundaryDetector)
                .findFirst()
                .orElseThrow());
        }
        
        // 如果配置了 stopAtNextHeader，启用表头检测器
        if (config.getRange().isStopAtNextHeader()) {
            active.add(allDetectors.stream()
                .filter(d -> d instanceof HeaderBoundaryDetector)
                .findFirst()
                .orElseThrow());
        }
        
        // 默认启用空行检测器
        if (active.isEmpty()) {
            active.add(allDetectors.stream()
                .filter(d -> d instanceof EmptyRowBoundaryDetector)
                .findFirst()
                .orElseThrow());
        }
        
        return active;
    }
}
```

---

## 六、配置示例大全

### 示例 1：数据 + 汇总行

```yaml
extractions:
  - key: "orderNos"
    position: { cellRef: "A2" }
    mode: DOWN
    range:
      # 检测汇总行并停止
      untilCondition:
        column: "A"
        contains: ["总计", "合计", "汇总"]
```

### 示例 2：数据 + 新表头

```yaml
extractions:
  - key: "orderNos"
    position: { cellRef: "A2" }
    mode: DOWN
    range:
      # 检测新表头并停止
      stopAtNextHeader: true
      headerPattern:
        bold: true
        hasBackgroundColor: true
```

### 示例 3：多表提取

```yaml
extractions:
  # 提取第一个表
  - key: "orderTable"
    position: { cellRef: "A2" }
    mode: DOWN
    range:
      stopAtNextHeader: true
  
  # 提取第二个表
  - key: "customerTable"
    position: { cellRef: "A10" }  # 手动指定起始位置
    mode: DOWN
    range:
      skipEmpty: true
```

### 示例 4：带容错的提取

```yaml
extractions:
  - key: "orderNos"
    position: { cellRef: "A2" }
    mode: DOWN
    range:
      # 允许中间有 1 个空行
      allowConsecutiveEmptyRows: 1
      
      # 连续 2 行不符合数据模式则停止
      stopAfterMismatchRows: 2
      
      # 最多读取 1000 行
      maxRows: 1000
```

---

## 七、总结

| 方案 | 优点 | 缺点 | 推荐度 |
|------|------|------|--------|
| untilCondition | 精确、配置简单 | 需要有明确标记 | ⭐⭐⭐⭐⭐ |
| stopAtNextHeader | 自动检测、智能 | 配置复杂 | ⭐⭐⭐⭐ |
| columnPattern | 适应性强 | 配置较复杂 | ⭐⭐⭐⭐ |
| blockDetection | 适合复杂布局 | 计算开销大 | ⭐⭐⭐ |
| manualRange | 100% 准确 | 不灵活 | ⭐⭐⭐ |

### 推荐配置策略

```yaml
# 最佳实践：多层检测 + 保底限制
extractions:
  - key: "data"
    position: { cellRef: "A2" }
    mode: DOWN
    range:
      # 第一层：条件检测（最精确）
      untilCondition:
        column: "A"
        contains: ["总计", "合计"]
      
      # 第二层：表头检测（备选）
      stopAtNextHeader: true
      
      # 第三层：容错处理
      allowConsecutiveEmptyRows: 1
      
      # 第四层：保底限制
      maxRows: 1000
```
