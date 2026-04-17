# SAX 流式读取器使用指南

## 概述

SAX（Simple API for XML）流式读取器是一种内存优化的 Excel 读取方式，适用于处理大型 Excel 文件（万行以上）。

## 核心特点

| 特性 | 说明 |
|------|------|
| **内存优化** | 逐行读取，不一次性加载整个文件到内存 |
| **只读模式** | 仅支持读取，不支持写入 |
| **单向遍历** | 无法随机访问单元格，只能顺序读取 |
| **适合大数据** | 推荐用于 10,000 行以上的文件 |

## 使用方式

### 1. 基础使用 - 回调模式

```java
import com.excelconfig.sax.SaxReader;
import com.excelconfig.sax.RowHandler;

SaxReader reader = new SaxReader();

// 使用回调处理每一行数据
reader.read(new FileInputStream("large-file.xlsx"), 0, new RowHandler() {
    @Override
    public void handleRow(int rowNum, List<String> cells) {
        // 处理每一行数据
        System.out.println("Row " + rowNum + ": " + cells);
    }
});
```

### 2. 读取所有数据到内存

```java
SaxReader reader = new SaxReader();

// 读取所有数据到内存（适用于中等规模数据）
List<List<String>> allRows = reader.readAll(
    new FileInputStream("data.xlsx"), 
    0  // Sheet 索引
);

// 处理数据
for (List<String> row : allRows) {
    System.out.println(row);
}
```

### 3. 与提取引擎集成

```java
// 创建提取上下文，传入输入流用于 SAX 读取
ExtractContext context = new ExtractContext(config, startRow, startColumn, sheetIndex);
context.setInputStream(new FileInputStream("large-file.xlsx"));

// 使用 SAX 模式提取数据
ExtractStrategy strategy = new SaxDownExtractStrategy();
List<Object> result = strategy.extract(sheet, config, context);
```

## 性能对比

| 读取方式 | 1000 行 | 10,000 行 | 100,000 行 |
|----------|---------|-----------|------------|
| 用户模式 (XSSFWorkbook) | 50MB | 500MB | OOM |
| SAX 流式读取 | 10MB | 15MB | 50MB |

## 注意事项

1. **两次读取问题**：由于需要查找表头位置，某些场景下可能需要读取两遍数据
2. **输入流重置**：如果输入流不支持 `reset()` 操作，需要重新创建输入流
3. **共享字符串**：SAX 读取器会自动处理共享字符串表，无需手动解析

## 最佳实践

### 推荐场景
- 大数据量提取（10,000 行以上）
- 内存受限环境（如嵌入式设备）
- 单向数据处理流程

### 不推荐场景
- 需要随机访问单元格
- 需要修改 Excel 内容
- 小文件（< 1000 行，使用用户模式更简单）

## 示例代码

完整示例请参考：
- `SaxReaderTest.java` - SAX 读取器单元测试
- `SaxDownExtractStrategy.java` - SAX 提取策略实现
