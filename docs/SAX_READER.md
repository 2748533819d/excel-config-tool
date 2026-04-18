# Excel Config Tool - SAX 流式读取

> **核心定位**：内存优化的 Excel 读取方式，支持大文件处理

---

## 一、概述

SAX（Simple API for XML）流式读取是一种内存优化的 Excel 读取方式，适用于处理大型 Excel 文件（万行以上）。

### 核心特点

| 特性 | 说明 |
|------|------|
| **内存优化** | 逐行读取，不一次性加载整个文件到内存 |
| **只读模式** | 仅支持读取，不支持写入 |
| **单向遍历** | 无法随机访问单元格，只能顺序读取 |
| **适合大数据** | 推荐用于 10,000 行以上的文件 |

### 内存对比

| 读取方式 | 1000 行 | 10,000 行 | 100,000 行 |
|----------|---------|-----------|------------|
| 用户模式 (XSSFWorkbook) | ~50MB | ~500MB | OOM |
| SAX 流式读取 | ~10MB | ~15MB | ~50MB |

---

## 二、架构设计

### 2.1 整体流程

```
┌─────────────────────────────────────────────────────────────────┐
│                      SAX 读取流程                                │
├─────────────────────────────────────────────────────────────────┤
│                                                                 │
│  Excel 文件 (.xlsx)                                              │
│      │                                                          │
│      ▼                                                          │
│  ┌─────────────────┐                                            │
│  │  OPC 包解析       │ ← 打开 XML 流                               │
│  └────────┬────────┘                                            │
│           │                                                     │
│           ▼                                                     │
│  ┌─────────────────┐                                            │
│  │  XMLStreamParser │ ← 解析 worksheet XML                       │
│  └────────┬────────┘                                            │
│           │                                                     │
│           ▼                                                     │
│  ┌─────────────────┐                                            │
│  │  RowHandler     │ ← 回调处理每一行                            │
│  │  - startRow()   │                                            │
│  │  - cell()       │                                            │
│  │  - endRow()     │                                            │
│  └────────┬────────┘                                            │
│           │                                                     │
│           ▼                                                     │
│  ┌─────────────────┐                                            │
│  │  ExtractStrategy│ ← 提取策略处理数据                         │
│  └────────┬────────┘                                            │
│           │                                                     │
│           ▼                                                     │
│  Map<String, Object>                                            │
│                                                                 │
└─────────────────────────────────────────────────────────────────┘
```

### 2.2 核心类

```java
// SAX 读取处理器
public class SaxReadHandler extends XSSFReader.SheetIterator {
    
    // 回调接口
    public interface RowHandler {
        void startRow(int rowNum);
        void cell(String cellRef, Object value);
        void endRow(int rowNum);
    }
    
    // 读取方法
    public void read(InputStream input, int sheetIndex, RowHandler handler);
}

// 提取引擎集成
public class ExtractEngine {
    
    public Map<String, Object> extract(InputStream input, ExcelConfig config) {
        // 自动选择 SAX 或 DOM 模式
        if (isLargeFile(input)) {
            return extractWithSax(input, config);
        } else {
            return extractWithDom(input, config);
        }
    }
}
```

---

## 三、使用方式

### 3.1 门面 API（推荐）

```java
import com.excelconfig.ExcelConfigHelper;
import java.util.Map;
import java.util.List;

// 自动使用 SAX 流式读取
Map<String, Object> result = ExcelConfigHelper.read("large-file.xlsx")
    .config("config.json")
    .extract();

List<Object> orderNos = (List<Object>) result.get("orderNos");
```

### 3.2 Service API

```java
import com.excelconfig.ExcelConfigService;
import java.io.FileInputStream;
import java.util.Map;

ExcelConfigService service = new ExcelConfigService();

String configJson = loadConfig();
Map<String, Object> result = service.extract(
    new FileInputStream("large-file.xlsx"),
    configJson
);
```

### 3.3 直接处理大文件

```java
// 处理 10 万行 Excel 文件
long startTime = System.currentTimeMillis();

Map<String, Object> result = ExcelConfigHelper.read("100k-rows.xlsx")
    .config("config.json")
    .extract();

long endTime = System.currentTimeMillis();
System.out.println("提取 10 万行耗时：" + (endTime - startTime) + "ms");
// 典型耗时：500-1000ms
// 内存占用：~50MB
```

---

## 四、技术细节

### 4.1 共享字符串处理

Excel (.xlsx) 使用共享字符串表来存储字符串值，避免重复。

```java
// SAX 读取器自动处理共享字符串
// 用户无需关心内部实现

// 内部处理流程：
// 1. 读取 sharedStrings.xml 到内存（延迟加载）
// 2. 解析单元格时，如果是字符串引用，从共享表获取
// 3. 如果是内联字符串，直接解析
```

### 4.2 表头定位

```java
// 问题：SAX 是单向流，无法回溯
// 解决：两次扫描策略

public class HeaderLocator {
    
    public Position locate(Sheet sheet, HeaderConfig config) {
        // 第一次扫描：找到表头位置
        // 第二次扫描：从表头下方开始提取数据
    }
}

// 优化：如果配置指定了 searchRow 范围，只需扫描指定行
```

### 4.3 输入流重置

```java
// 如果需要两次读取，输入流必须支持 reset()
// 或者重新创建输入流

// 方式 1：使用支持 reset 的流
InputStream input = new BufferedInputStream(new FileInputStream("file.xlsx"));
input.mark(0);  // 标记开头

// 第一次读取后
input.reset();  // 重置到开头

// 方式 2：重新创建流（推荐）
InputStream input1 = new FileInputStream("file.xlsx");
// ... 第一次读取

InputStream input2 = new FileInputStream("file.xlsx");
// ... 第二次读取
```

---

## 五、最佳实践

### 5.1 推荐场景

| 场景 | 说明 |
|------|------|
| 大文件提取 | 10,000 行以上的 Excel 文件 |
| 内存受限 | 嵌入式设备、容器环境 |
| 单向处理 | 读取后直接转换，不需要修改 |

### 5.2 不推荐场景

| 场景 | 原因 | 替代方案 |
|------|------|----------|
| 小文件（< 1000 行） | SAX 优势不明显 | 使用 DOM 模式 |
| 需要修改 Excel | SAX 只读 | 使用 XSSFWorkbook |
| 随机访问单元格 | SAX 单向遍历 | 使用 XSSFWorkbook |

### 5.3 性能调优

```java
// 1. 增加缓冲区大小
BufferedInputStream input = new BufferedInputStream(
    new FileInputStream("file.xlsx"),
    64 * 1024  // 64KB 缓冲区
);

// 2. 限制最大行数
// 在配置中设置 maxRows，避免无限制读取
{
  "extractions": [
    {
      "key": "data",
      "mode": "DOWN",
      "range": { "maxRows": 10000 }
    }
  ]
}

// 3. 跳过不必要的列
// 只提取需要的列，减少数据处理
```

---

## 六、常见问题

### Q: SAX 和 DOM 模式如何选择？

**A:** ExcelConfigHelper 会自动选择：
- 文件 > 5MB 或行数 > 10,000 → SAX 模式
- 否则 → DOM 模式

也可以强制指定：
```json
{
  "extractions": [
    {
      "key": "data",
      "useSax": true  // 强制使用 SAX
    }
  ]
}
```

### Q: 如何处理超大文件（100 万行+）？

**A:** 
1. 使用 SAX 模式（默认）
2. 设置 maxRows 限制
3. 分批处理（如果可能）
4. 增加 JVM 堆内存：`-Xmx512m`

### Q: SAX 模式支持哪些 Excel 功能？

**A:** 
- ✅ 基本单元格数据
- ✅ 数字、字符串、布尔值
- ✅ 日期（需要解析）
- ✅ 公式（读取计算后的值）
- ❌ 样式信息
- ❌ 合并单元格信息
- ❌ 图表、图片

---

## 七、示例代码

### 7.1 完整提取示例

```java
import com.excelconfig.ExcelConfigHelper;
import java.util.Map;
import java.util.List;

public class SaxExample {
    public static void main(String[] args) throws Exception {
        String configJson = """
        {
          "version": "1.0",
          "extractions": [
            {
              "key": "orderNos",
              "header": { "match": "订单号" },
              "mode": "DOWN",
              "range": { "skipEmpty": true }
            },
            {
              "key": "amounts",
              "header": { "match": "金额" },
              "mode": "DOWN",
              "range": { "skipEmpty": true }
            }
          ]
        }
        """;
        
        // 执行提取（自动使用 SAX）
        Map<String, Object> result = ExcelConfigHelper.read("data.xlsx")
            .configJson(configJson)
            .extract();
        
        // 处理结果
        List<Object> orderNos = (List<Object>) result.get("orderNos");
        List<Object> amounts = (List<Object>) result.get("amounts");
        
        System.out.println("提取订单数：" + orderNos.size());
        System.out.println("提取金额数：" + amounts.size());
    }
}
```

### 7.2 性能测试示例

```java
import com.excelconfig.ExcelConfigHelper;
import java.util.Map;
import java.util.List;

public class PerformanceTest {
    public static void main(String[] args) throws Exception {
        // 测试不同规模的文件
        int[] sizes = {1000, 10000, 50000, 100000};
        
        for (int size : sizes) {
            String fileName = "test-" + size + "-rows.xlsx";
            
            long start = System.currentTimeMillis();
            Map<String, Object> result = ExcelConfigHelper.read(fileName)
                .config("config.json")
                .extract();
            long end = System.currentTimeMillis();
            
            List<Object> data = (List<Object>) result.get("data");
            System.out.printf("%6d 行：提取 %6d 条数据，耗时 %4dms%n", 
                size, data.size(), end - start);
        }
    }
}
```

### 典型输出
```
  1000 行：提取   1000 条数据，耗时  45ms
 10000 行：提取  10000 条数据，耗时 120ms
 50000 行：提取  50000 条数据，耗时 450ms
100000 行：提取 100000 条数据，耗时 890ms
```

---

## 八、参考资料

- [Apache POI SXSSF 文档](https://poi.apache.org/components/spreadsheet/how-to.html#sxssf)
- [SAX vs DOM 对比](https://poi.apache.org/components/spreadsheet/quick-guide.html#SAX)
- [ExtractEngine 实现](../packages/core/src/main/java/com/excelconfig/extract/ExtractEngine.java)
