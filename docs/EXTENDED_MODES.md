# Excel Config Tool - 扩展提取模式设计

> 在基础模式之上的高级提取模式，满足复杂场景需求

---

## 一、基础模式回顾

| 模式 | 说明 | 输出 |
|------|------|------|
| SINGLE | 单一单元格 | Object |
| DOWN | 向下提取 | Array |
| RIGHT | 向右提取 | Array |
| BLOCK | 区域块 | Array<Array> |
| UNTIL_EMPTY | 直到空值 | Array |

---

## 二、扩展模式设计

### 6. KEY_VALUE - 键值对提取模式

#### 场景
Excel 中常见的配置表格式，A 列是 key，B 列是 value：

```
┌─────────────────┐
│  配置项  │  值   │
├─────────────────┤
│ 公司名称 │ XX 公司│
│ 地址    │ XX 市  │
│ 电话    │ 12345 │
│ 邮箱    │ a@b.c │
└─────────────────┘
```

#### 配置
```yaml
companyInfo:
  key: "companyInfo"
  position:
    areaRef: "A2:B10"
  mode: KEY_VALUE
  parser:
    keyType: string
    valueType: any
```

#### 输出
```json
{
  "companyInfo": {
    "公司名称": "XX 公司",
    "地址": "XX 市",
    "电话": "12345",
    "邮箱": "a@b.c"
  }
}
```

#### 变体：指定 key 列和 value 列
```yaml
keyValueData:
  position:
    areaRef: "A2:D20"
  mode: KEY_VALUE
  keyColumn: 0      # A 列作为 key
  valueColumns: [1, 2, 3]  # B/C/D 列作为 value
  valueType: "row"  # 每行转为对象
```

输出：
```json
{
  "keyValueData": [
    { "key": "配置 1", "values": ["值 1B", "值 1C", "值 1D"] },
    { "key": "配置 2", "values": ["值 2B", "值 2C", "值 2D"] }
  ]
}
```

---

### 7. TABLE - 表格提取模式（带表头）

#### 场景
标准的数据表格，第一行是表头，下面是数据：

```
┌────────┬────────┬────────┬────────┐
│ 订单号  │  金额   │  日期   │  客户  │ ← 表头
├────────┼────────┼────────┼────────┤
│ ORD001 │ 100.00 │01-01   │ 张三   │ ← 数据
│ ORD002 │ 200.00 │01-02   │ 李四   │
│ ORD003 │ 300.00 │01-03   │ 王五   │
└────────┴────────┴────────┴────────┘
```

#### 配置
```yaml
orders:
  key: "orders"
  position:
    areaRef: "A1:D100"
  mode: TABLE
  headerRow: 0    # 表头在第 1 行 (0-based)
  dataStartRow: 1 # 数据从第 2 行开始
  parser:
    autoType: true  # 根据表头推断类型
```

#### 输出
```json
{
  "orders": [
    { "订单号": "ORD001", "金额": 100.00, "日期": "2024-01-01", "客户": "张三" },
    { "订单号": "ORD002", "金额": 200.00, "日期": "2024-01-02", "客户": "李四" },
    { "订单号": "ORD003", "金额": 300.00, "日期": "2024-01-03", "客户": "王五" }
  ]
}
```

#### 进阶：指定字段映射
```yaml
orders:
  position:
    areaRef: "A1:D100"
  mode: TABLE
  headerRow: 0
  columns:
    - header: "订单号"
      key: "orderNo"
      type: string
    - header: "金额"
      key: "amount"
      type: number
    - header: "日期"
      key: "orderDate"
      type: date
      format: "yyyy-MM-dd"
    - header: "客户"
      key: "customerName"
      type: string
```

输出：
```json
{
  "orders": [
    { "orderNo": "ORD001", "amount": 100.00, "orderDate": "2024-01-01", "customerName": "张三" },
    ...
  ]
}
```

---

### 8. GROUPED - 分组提取模式

#### 场景
数据按某列分组，需要分组提取：

```
┌─────────────────────────────────┐
│  部门  │  姓名  │  工资  │  入职日期│
├─────────────────────────────────┤
│  技术部 │ 张三   │ 10000 │ 2023-01 │
│  技术部 │ 李四   │ 12000 │ 2023-02 │ ← 技术部一组
│  销售部 │ 王五   │ 8000  │ 2023-03 │
│  销售部 │ 赵六   │ 9000  │ 2023-04 │ ← 销售部一组
│  人事部 │ 钱七   │ 7000  │ 2023-05 │ ← 人事部一组
└─────────────────────────────────┘
```

#### 配置
```yaml
employeesByDept:
  key: "employeesByDept"
  position:
    areaRef: "A2:D100"
  mode: GROUPED
  groupBy:
    column: 0           # 按 A 列 (部门) 分组
    headerRow: 0        # 表头在第 1 行
  columns:
    - index: 1
      key: "name"
      type: string
    - index: 2
      key: "salary"
      type: number
    - index: 3
      key: "hireDate"
      type: date
```

#### 输出
```json
{
  "employeesByDept": {
    "技术部": [
      { "name": "张三", "salary": 10000, "hireDate": "2023-01-01" },
      { "name": "李四", "salary": 12000, "hireDate": "2023-02-01" }
    ],
    "销售部": [
      { "name": "王五", "salary": 8000, "hireDate": "2023-03-01" },
      { "name": "赵六", "salary": 9000, "hireDate": "2023-04-01" }
    ],
    "人事部": [
      { "name": "钱七", "salary": 7000, "hireDate": "2023-05-01" }
    ]
  }
}
```

---

### 9. HIERARCHY - 层级提取模式

#### 场景
有缩进或层级关系的数据：

```
┌─────────────────────────────────────┐
│  科目名称            │  金额        │
├─────────────────────────────────────┤
│  营业收入            │  1000000     │
│    主营业务收入      │  800000      │ ← 缩进 1 级
│      产品 A 收入      │  500000      │ ← 缩进 2 级
│      产品 B 收入      │  300000      │ ← 缩进 2 级
│    其他业务收入      │  200000      │ ← 缩进 1 级
│  营业成本            │  600000      │
│    主营业务成本      │  500000      │ ← 缩进 1 级
└─────────────────────────────────────┘
```

#### 配置
```yaml
financialHierarchy:
  key: "financialHierarchy"
  position:
    areaRef: "A2:B50"
  mode: HIERARCHY
  indent:
    column: 0         # 从 A 列检测缩进
    indentSize: 2     # 2 个空格 = 1 级缩进
  valueColumn: 1      # B 列是值
```

#### 输出
```json
{
  "financialHierarchy": {
    "name": "营业收入",
    "value": 1000000,
    "children": [
      {
        "name": "主营业务收入",
        "value": 800000,
        "children": [
          { "name": "产品 A 收入", "value": 500000 },
          { "name": "产品 B 收入", "value": 300000 }
        ]
      },
      {
        "name": "其他业务收入",
        "value": 200000
      }
    ]
  }
}
```

---

### 10. CROSS_TAB - 交叉表提取模式

#### 场景
行列都有表头的交叉表：

```
┌──────────┬──────────────────────────────┐
│  产品\月  │  1 月  │  2 月  │  3 月  │  4 月  │
├──────────┼────────┼────────┼────────┼────────┤
│  产品 A   │  100   │  120   │  150   │  130   │
│  产品 B   │  200   │  220   │  250   │  230   │
│  产品 C   │  300   │  320   │  350   │  330   │
└──────────┴────────┴────────┴────────┴────────┘
```

#### 配置
```yaml
salesCrossTab:
  key: "salesCrossTab"
  position:
    areaRef: "A1:E4"
  mode: CROSS_TAB
  rowHeader:
    column: 0         # A 列是行表头 (产品名称)
    startRow: 1       # 从第 2 行开始
  columnHeader:
    row: 0            # 第 1 行是列表头 (月份)
    startColumn: 1    # 从 B 列开始
  dataArea:
    topRow: 1         # 数据从第 2 行开始
    leftColumn: 1     # 数据从 B 列开始
```

#### 输出
```json
{
  "salesCrossTab": {
    "headers": {
      "rows": ["产品 A", "产品 B", "产品 C"],
      "columns": ["1 月", "2 月", "3 月", "4 月"]
    },
    "data": [
      [100, 120, 150, 130],
      [200, 220, 250, 230],
      [300, 320, 350, 330]
    ],
    "byRow": {
      "产品 A": { "1 月": 100, "2 月": 120, "3 月": 150, "4 月": 130 },
      "产品 B": { "1 月": 200, "2 月": 220, "3 月": 250, "4 月": 230 },
      "产品 C": { "1 月": 300, "2 月": 320, "3 月": 350, "4 月": 330 }
    }
  }
}
```

---

### 11. MERGED_CELLS - 合并单元格提取模式

#### 场景
有合并单元格的复杂表格：

```
┌─────────────────────────────────────────┐
│  部门   │  姓名  │  1 月  │  2 月  │  3 月  │
├─────────────────────────────────────────┤
│         │ 张三   │  100  │  120  │  150  │
│  技术部  │ 李四   │  200  │  220  │  250  │ ← "技术部"跨 2 行合并
│         │ 王五   │  300  │  320  │  350  │
├─────────────────────────────────────────┤
│  销售部  │ 赵六   │  400  │  420  │  450  │
│         │ 钱七   │  500  │  520  │  550  │ ← "销售部"跨 2 行合并
└─────────────────────────────────────────┘
```

#### 配置
```yaml
deptSalesWithMerged:
  key: "deptSalesWithMerged"
  position:
    areaRef: "A1:E8"
  mode: MERGED_CELLS
  mergeDirection: VERTICAL   # 垂直合并
  mergeColumn: 0             # A 列有合并
  extractMode: TABLE         # 基础模式是表格
```

#### 输出
```json
{
  "deptSalesWithMerged": [
    { "department": "技术部", "name": "张三", "sales": [100, 120, 150] },
    { "department": "技术部", "name": "李四", "sales": [200, 220, 250] },
    { "department": "技术部", "name": "王五", "sales": [300, 320, 350] },
    { "department": "销售部", "name": "赵六", "sales": [400, 420, 450] },
    { "department": "销售部", "name": "钱七", "sales": [500, 520, 550] }
  ]
}
```

**关键点**: 自动检测合并单元格，将合并值填充到所有相关行

---

### 12. PIVOT - 透视表提取模式

#### 场景
需要按条件汇总的透视表：

```
原始数据：
┌────────┬────────┬────────┐
│  地区   │  产品   │  金额   │
├────────┼────────┼────────┤
│  华东   │  产品 A │  100   │
│  华东   │  产品 B │  200   │
│  华北   │  产品 A │  150   │
│  华北   │  产品 B │  250   │
│  华南   │  产品 A │  120   │
└────────┴────────┴────────┘

需要提取为:
┌────────┬────────┬────────┬────────┐
│  地区   │  产品 A │  产品 B │  合计   │
├────────┼────────┼────────┼────────┤
│  华东   │  100   │  200   │  300   │
│  华北   │  150   │  250   │  400   │
│  华南   │  120   │   0    │  120   │
└────────┴────────┴────────┴────────┘
```

#### 配置
```yaml
salesPivot:
  key: "salesPivot"
  position:
    areaRef: "A2:C100"
  mode: PIVOT
  pivot:
    rows:
      - column: 0       # 地区作为行
        key: "region"
    columns:
      - column: 1       # 产品作为列
    values:
      - column: 2       # 金额作为值
        aggregate: SUM  # 汇总方式
```

#### 输出
```json
{
  "salesPivot": {
    "headers": ["地区", "产品 A", "产品 B", "合计"],
    "data": [
      ["华东", 100, 200, 300],
      ["华北", 150, 250, 400],
      ["华南", 120, 0, 120]
    ],
    "byRegion": {
      "华东": { "产品 A": 100, "产品 B": 200, "合计": 300 },
      "华北": { "产品 A": 150, "产品 B": 250, "合计": 400 },
      "华南": { "产品 A": 120, "产品 B": 0, "合计": 120 }
    }
  }
}
```

---

### 13. MULTI_SHEET - 多工作表提取模式

#### 场景
数据分布在多个工作表中：

```
Workbook: 2024 年销售数据.xlsx
├── Sheet: 1 月 (A1:B100)
├── Sheet: 2 月 (A1:B100)
├── Sheet: 3 月 (A1:B100)
└── ...
```

#### 配置
```yaml
monthlySales:
  key: "monthlySales"
  position:
    sheets: ["1 月", "2 月", "3 月", "4 月", "5 月", "6 月"]
    areaRef: "A2:B100"
  mode: MULTI_SHEET
  sheetKey: "month"     # 用 sheet 名作为 key
  columns:
    - index: 0
      key: "productName"
      type: string
    - index: 1
      key: "amount"
      type: number
```

#### 输出
```json
{
  "monthlySales": {
    "1 月": [
      { "productName": "产品 A", "amount": 1000 },
      ...
    ],
    "2 月": [
      { "productName": "产品 A", "amount": 1200 },
      ...
    ],
    ...
  }
}
```

#### 变体：使用通配符
```yaml
monthlySales:
  mode: MULTI_SHEET
  position:
    sheetPattern: "^[0-9]+ 月$"   # 正则匹配 sheet 名
    areaRef: "A2:B100"
```

---

### 14. FORMULA - 公式计算模式

#### 场景
需要提取公式计算结果：

```
┌─────────────────────────────────┐
│  项目   │  数值  │  公式        │
├─────────────────────────────────┤
│  销售额  │ 1000  │              │
│  成本    │ 600   │              │
│  利润    │       │ =B2-B3       │ ← 公式计算
│  利润率  │       │ =B4/B2*100   │ ← 公式计算
└─────────────────────────────────┘
```

#### 配置
```yaml
profitMetrics:
  key: "profitMetrics"
  position:
    areaRef: "A1:C5"
  mode: FORMULA
  formulas:
    - key: "profit"
      formula: "B2-B3"
      type: number
    - key: "profitMargin"
      formula: "B4/B2*100"
      type: number
      format: "0.00"
```

#### 输出
```json
{
  "profitMetrics": {
    "profit": 400,
    "profitMargin": 40.00
  }
}
```

---

### 15. CONDITIONAL - 条件过滤提取模式

#### 场景
只提取符合条件的行：

```
┌─────────────────────────────────────┐
│  订单号  │  金额   │  状态  │  日期   │
├─────────────────────────────────────┤
│ ORD001  │ 100.00  │ 已完成  │01-01   │ ← 提取
│ ORD002  │ 200.00  │ 进行中  │01-02   │
│ ORD003  │ 300.00  │ 已完成  │01-03   │ ← 提取
│ ORD004  │ 400.00  │ 已取消  │01-04   │
└─────────────────────────────────────┘
```

#### 配置
```yaml
completedOrders:
  key: "completedOrders"
  position:
    areaRef: "A2:D100"
  mode: CONDITIONAL
  condition:
    column: 2           # C 列 (状态)
    operator: EQUALS    # 等于
    value: "已完成"
  headerRow: 0
  extractMode: TABLE    # 提取为表格
```

#### 输出
```json
{
  "completedOrders": [
    { "orderNo": "ORD001", "amount": 100.00, "status": "已完成", "date": "2024-01-01" },
    { "orderNo": "ORD003", "amount": 300.00, "status": "已完成", "date": "2024-01-03" }
  ]
}
```

#### 支持的操作符
```yaml
condition:
  # 基本比较
  - operator: EQUALS       # 等于
  - operator: NOT_EQUALS   # 不等于
  - operator: GREATER_THAN # 大于
  - operator: LESS_THAN    # 小于
  
  # 模糊匹配
  - operator: CONTAINS     # 包含
  - operator: STARTS_WITH  # 开头是
  - operator: ENDS_WITH    # 结尾是
  - operator: REGEX        # 正则匹配
  
  # 范围
  - operator: BETWEEN      # 在...之间
  - operator: IN           # 在列表中
  
  # 空值
  - operator: IS_EMPTY     # 为空
  - operator: IS_NOT_EMPTY # 不为空
```

#### 多条件组合
```yaml
condition:
  logic: AND   # AND / OR
  conditions:
    - column: 2
      operator: EQUALS
      value: "已完成"
    - column: 1
      operator: GREATER_THAN
      value: 500
```

---

## 三、模式分类总览

```
┌─────────────────────────────────────────────────────────────────────────────────┐
│                           提取模式完整分类                                       │
├─────────────────────────────────────────────────────────────────────────────────┤
│                                                                                 │
│  基础模式 (5 种)                                                                  │
│  ├── SINGLE         - 单一单元格                                                 │
│  ├── DOWN           - 向下提取                                                   │
│  ├── RIGHT          - 向右提取                                                   │
│  ├── BLOCK          - 区域块                                                     │
│  └── UNTIL_EMPTY    - 直到空值                                                   │
│                                                                                 │
│  结构化模式 (3 种)                                                                │
│  ├── KEY_VALUE      - 键值对提取                                                  │
│  ├── TABLE          - 表格提取 (带表头)                                            │
│  └── CROSS_TAB      - 交叉表提取                                                  │
│                                                                                 │
│  高级模式 (4 种)                                                                  │
│  ├── GROUPED        - 分组提取                                                    │
│  ├── HIERARCHY      - 层级提取                                                    │
│  ├── MERGED_CELLS   - 合并单元格提取                                              │
│  └── MULTI_SHEET    - 多工作表提取                                                │
│                                                                                 │
│  分析模式 (3 种)                                                                  │
│  ├── PIVOT          - 透视表提取                                                  │
│  ├── FORMULA        - 公式计算提取                                                │
│  └── CONDITIONAL    - 条件过滤提取                                                │
│                                                                                 │
│  总计：15 种模式                                                                  │
│                                                                                 │
└─────────────────────────────────────────────────────────────────────────────────┘
```

---

## 四、模式选择决策树

```
开始
│
├─ 提取单个值？
│  └──→ SINGLE
│
├─ 提取列表？
│   ├── 纵向列表？
│   │   ├── 确定行数？
│   │   │   ├── 是 → DOWN
│   │   │   └── 否 → UNTIL_EMPTY
│   │   └── 有过滤条件？ → CONDITIONAL
│   │
│   └── 横向列表？
│       └──→ RIGHT
│
├─ 提取表格 (带表头)?
│   ├── 有合并单元格？ → MERGED_CELLS
│   ├── 需要分组？ → GROUPED
│   ├── 有层级缩进？ → HIERARCHY
│   └── 普通表格 → TABLE
│
├─ 提取区域块？
│   ├── 行列表头都有？ → CROSS_TAB
│   ├── 键值对格式？ → KEY_VALUE
│   └── 纯数据矩阵 → BLOCK
│
├─ 跨多个工作表？
│   └──→ MULTI_SHEET
│
├─ 需要汇总计算？
│   ├── 透视表？ → PIVOT
│   └── 公式计算？ → FORMULA
│
└── 以上都不是 → 使用基础模式组合
```

---

## 五、优先级建议

### Phase 1 (MVP - 必须实现)
| 模式 | 优先级 | 理由 |
|------|--------|------|
| SINGLE | ★★★★★ | 最基础 |
| DOWN | ★★★★★ | 最常用 |
| RIGHT | ★★★★☆ | 表头提取 |
| UNTIL_EMPTY | ★★★★☆ | 动态数据 |
| BLOCK | ★★★☆☆ | 矩阵数据 |

### Phase 2 (常用扩展)
| 模式 | 优先级 | 理由 |
|------|--------|------|
| TABLE | ★★★★★ | 标准表格场景多 |
| KEY_VALUE | ★★★★☆ | 配置表常见 |
| CONDITIONAL | ★★★★☆ | 过滤需求多 |
| MERGED_CELLS | ★★★☆☆ | 复杂表格 |

### Phase 3 (高级功能)
| 模式 | 优先级 | 理由 |
|------|--------|------|
| CROSS_TAB | ★★★☆☆ | 交叉表场景 |
| GROUPED | ★★★☆☆ | 分组场景 |
| MULTI_SHEET | ★★☆☆☆ | 多 sheet 场景 |
| HIERARCHY | ★★☆☆☆ | 层级数据 |
| PIVOT | ★★☆☆☆ | 汇总分析 |
| FORMULA | ★★☆☆☆ | 公式计算 |

---

## 六、实现建议

### 策略接口统一

```java
public interface ExtractStrategy {
    /**
     * 执行提取
     */
    Object extract(ExtractContext context);
    
    /**
     * 支持的提取模式
     */
    Set<ExtractMode> supportedModes();
}

// 每种模式一个实现类
public class SingleStrategy implements ExtractStrategy { ... }
public class DownStrategy implements ExtractStrategy { ... }
public class TableStrategy implements ExtractStrategy { ... }
public class KeyValueStrategy implements ExtractStrategy { ... }
// ...
```

### 模式可以组合使用

```yaml
# 同一个区域，用不同模式提取
cells:
  # 方式 1: 作为表格提取
  ordersTable:
    position: { areaRef: "A2:D100" }
    mode: TABLE
    
  # 方式 2: 作为 KEY_VALUE 提取 (A 列 key, B 列 value)
  ordersKeyValue:
    position: { areaRef: "A2:B100" }
    mode: KEY_VALUE
```

---

## 七、总结

### 内置 15 种模式覆盖的场景

| 类别 | 场景数 | 覆盖率 |
|------|--------|--------|
| 基础模式 | 简单数据提取 | 60% |
| 结构化模式 | 标准表格 | 25% |
| 高级模式 | 复杂表格 | 10% |
| 分析模式 | 数据分析 | 5% |

### 设计原则

1. **80/20 原则**: 优先实现 20% 最常用的模式，覆盖 80% 的场景
2. **可组合性**: 模式之间可以组合使用
3. **可扩展**: 用户可以实现自定义模式
4. **渐进式**: 从 MVP 开始，逐步增加
