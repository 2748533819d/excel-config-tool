# Excel Config Tool 使用示例

## 快速开始

### 1. 添加依赖

```xml
<dependency>
    <groupId>io.github.cynosure-tech</groupId>
    <artifactId>excel-config-core</artifactId>
    <version>1.0.1</version>
</dependency>
```

### 2. 准备配置文件

创建 `config.json`：

```json
{
  "version": "1.0",
  "templateName": "订单管理",
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
  ],
  "exports": [
    {
      "key": "orderNos",
      "header": { "match": "订单号" },
      "mode": "FILL_DOWN"
    }
  ]
}
```

---

## 示例 1：基础数据提取（门面 API）

### 场景
从 Excel 表格中提取"订单号"和"金额"列的数据。

### Excel 模板
```
| 订单号 | 金额   | 日期       |
|--------|--------|------------|
| ORD001 | 100.00 | 2024-01-01 |
| ORD002 | 200.00 | 2024-01-02 |
| ORD003 | 150.00 | 2024-01-03 |
```

### Java 代码
```java
import com.excelconfig.ExcelConfigHelper;
import java.util.Map;
import java.util.List;

public class ExtractExample {
    public static void main(String[] args) throws Exception {
        // 提取数据
        Map<String, Object> result = ExcelConfigHelper.read("template.xlsx")
            .config("config.json")
            .extract();
        
        // 获取结果
        List<Object> orderNos = (List<Object>) result.get("orderNos");
        List<Object> amounts = (List<Object>) result.get("amounts");
        
        // 输出
        System.out.println("订单号：" + orderNos);
        System.out.println("金额：" + amounts);
    }
}
```

### 输出结果
```
订单号：[ORD001, ORD002, ORD003]
金额：[100.0, 200.0, 150.0]
```

---

## 示例 2：基础数据填充（门面 API）

### 场景
将订单号列表填充到 Excel 模板中。

### Excel 模板
```
| 订单号 |
|--------|
|        |
|        |
```

### Java 代码
```java
import com.excelconfig.ExcelConfigHelper;
import java.util.Map;
import java.util.Arrays;

public class FillExample {
    public static void main(String[] args) throws Exception {
        // 准备数据
        Map<String, Object> data = Map.of(
            "orderNos", Arrays.asList("ORD001", "ORD002", "ORD003", "ORD004")
        );
        
        // 填充数据并写入文件
        ExcelConfigHelper.write("template.xlsx")
            .config("config.json")
            .data(data)
            .writeTo("output.xlsx");
        
        System.out.println("填充完成！");
    }
}
```

### 输出结果
```
| 订单号 |
|--------|
| ORD001 |
| ORD002 |
| ORD003 |
| ORD004 |
```

---

## 示例 3：从 InputStream 读取/写入 OutputStream

### 场景
在 Web 应用中，从 HTTP 请求读取 Excel，处理后直接返回给客户端。

### Java 代码
```java
import com.excelconfig.ExcelConfigHelper;
import jakarta.servlet.http.HttpServlet;
import jakarta.servlet.http.HttpServletRequest;
import jakarta.servlet.http.HttpServletResponse;
import java.io.InputStream;
import java.util.Map;

public class UploadServlet extends HttpServlet {
    
    protected void doPost(HttpServletRequest request, HttpServletResponse response) {
        try {
            // 从请求读取 Excel
            InputStream inputStream = request.getInputStream();
            
            // 提取数据
            Map<String, Object> result = ExcelConfigHelper.read(inputStream)
                .configJson(CONFIG_JSON)  // 直接使用 JSON 字符串
                .extract();
            
            // 处理数据...
            List<Object> orderNos = (List<Object>) result.get("orderNos");
            
            // 返回 JSON 响应
            response.setContentType("application/json");
            response.getWriter().write(toJson(result));
            
        } catch (Exception e) {
            throw new RuntimeException("处理失败", e);
        }
    }
}
```

---

## 示例 4：填充数据并返回字节数组

### 场景
生成 Excel 文件供用户下载。

### Java 代码
```java
import com.excelconfig.ExcelConfigHelper;
import jakarta.servlet.http.HttpServlet;
import jakarta.servlet.http.HttpServletRequest;
import jakarta.servlet.http.HttpServletResponse;
import java.util.Map;
import java.util.Arrays;

public class DownloadServlet extends HttpServlet {
    
    protected void doGet(HttpServletRequest request, HttpServletResponse response) {
        try {
            // 准备数据
            Map<String, Object> data = Map.of(
                "orderNos", Arrays.asList("ORD001", "ORD002", "ORD003"),
                "amounts", Arrays.asList(100.0, 200.0, 150.0),
                "dates", Arrays.asList("2024-01-01", "2024-01-02", "2024-01-03")
            );
            
            // 填充模板，获取字节数组
            byte[] excelBytes = ExcelConfigHelper.write("template.xlsx")
                .config("config.json")
                .data(data)
                .write();
            
            // 设置响应头
            response.setContentType(
                "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            );
            response.setHeader("Content-Disposition", 
                "attachment; filename=\"orders.xlsx\"");
            
            // 写入响应
            response.getOutputStream().write(excelBytes);
            response.getOutputStream().flush();
            
        } catch (Exception e) {
            throw new RuntimeException("生成失败", e);
        }
    }
}
```

---

## 示例 5：Service API 使用方式

### 场景
使用传统的 Service API 方式。

### Java 代码
```java
import com.excelconfig.ExcelConfigService;
import java.io.FileInputStream;
import java.util.Map;

public class ServiceApiExample {
    public static void main(String[] args) throws Exception {
        ExcelConfigService service = new ExcelConfigService();
        
        // 读取配置文件
        String configJson = java.nio.file.Files.readString(
            java.nio.file.Paths.get("config.json")
        );
        
        // 提取数据
        Map<String, Object> result = service.extract(
            new FileInputStream("template.xlsx"),
            configJson
        );
        
        System.out.println("提取结果：" + result);
        
        // 填充数据
        Map<String, Object> inputData = Map.of(
            "orderNos", java.util.Arrays.asList("ORD001", "ORD002", "ORD003")
        );
        
        byte[] excelBytes = service.fill(
            new FileInputStream("template.xlsx"),
            inputData,
            configJson
        );
        
        // 保存到文件
        java.nio.file.Files.write(
            java.nio.file.Paths.get("output.xlsx"),
            excelBytes
        );
        
        System.out.println("填充完成！");
    }
}
```

---

## 示例 6：使用配置对象而非 JSON 文件

### 场景
在代码中动态构建配置。

### Java 代码
```java
import com.excelconfig.ExcelConfigHelper;
import com.excelconfig.model.*;
import java.util.Map;
import java.util.Arrays;

public class ConfigObjectExample {
    public static void main(String[] args) throws Exception {
        // 构建配置
        ExcelConfig config = new ExcelConfig();
        config.setVersion("1.0");
        config.setTemplateName("订单管理");
        
        // 提取配置
        ExtractConfig extractConfig = new ExtractConfig();
        extractConfig.setKey("orderNos");
        
        HeaderConfig headerConfig = new HeaderConfig();
        headerConfig.setMatch("订单号");
        extractConfig.setHeader(headerConfig);
        
        extractConfig.setMode(ExtractMode.DOWN);
        
        RangeConfig rangeConfig = new RangeConfig();
        rangeConfig.setSkipEmpty(true);
        extractConfig.setRange(rangeConfig);
        
        config.setExtractions(Arrays.asList(extractConfig));
        
        // 使用配置对象
        Map<String, Object> result = ExcelConfigHelper.read("template.xlsx")
            .configObject(config)
            .extract();
        
        System.out.println("提取结果：" + result);
    }
}
```

---

## 示例 7：大数据量处理

### 场景
处理包含数千行数据的 Excel 文件。

### Java 代码
```java
import com.excelconfig.ExcelConfigHelper;
import java.util.Map;
import java.util.List;

public class LargeFileExample {
    public static void main(String[] args) throws Exception {
        long startTime = System.currentTimeMillis();
        
        // 提取大数据量文件
        Map<String, Object> result = ExcelConfigHelper.read("large-file.xlsx")
            .config("config.json")
            .extract();
        
        List<Object> orderNos = (List<Object>) result.get("orderNos");
        System.out.println("提取行数：" + orderNos.size());
        
        long extractTime = System.currentTimeMillis() - startTime;
        System.out.println("提取耗时：" + extractTime + "ms");
        
        // 填充大数据量
        Map<String, Object> inputData = Map.of(
            "orderNos", orderNos,
            "processed", java.util.Collections.nCopies(orderNos.size(), true)
        );
        
        ExcelConfigHelper.write("template.xlsx")
            .config("config.json")
            .data(inputData)
            .writeTo("output.xlsx");
        
        long fillTime = System.currentTimeMillis() - startTime;
        System.out.println("填充耗时：" + fillTime + "ms");
    }
}
```

### 性能参考
```
1000 行数据：
- 提取耗时：约 50-100ms
- 填充耗时：约 50-100ms

10000 行数据：
- 提取耗时：约 300-500ms
- 填充耗时：约 300-500ms
```

---

## 示例 8：类型转换（extractAs）

### 场景
将提取结果直接转换为 Java 对象。

### Java 代码
```java
import com.excelconfig.ExcelConfigHelper;
import java.util.List;

// 定义 DTO 类
public class OrderDTO {
    private List<String> orderNos;
    private List<Double> amounts;
    private List<String> dates;
    
    // getters and setters...
}

public class TypeConvertExample {
    public static void main(String[] args) throws Exception {
        // 直接转换为 DTO 对象
        OrderDTO order = ExcelConfigHelper.read("template.xlsx")
            .config("config.json")
            .extractAs(OrderDTO.class);
        
        // 使用对象
        for (int i = 0; i < order.getOrderNos().size(); i++) {
            System.out.printf("订单：%s, 金额：%s, 日期：%s%n",
                order.getOrderNos().get(i),
                order.getAmounts().get(i),
                order.getDates().get(i)
            );
        }
    }
}
```

---

## 配置参考

### ExtractConfig 配置项

| 字段 | 类型 | 必填 | 说明 |
|------|------|------|------|
| key | String | 是 | 数据键名 |
| header | HeaderConfig | 是 | 表头配置 |
| mode | String | 是 | 提取模式 (SINGLE/DOWN/RIGHT/BLOCK/UNTIL_EMPTY) |
| range | RangeConfig | 否 | 范围配置 |
| parser | ParserConfig | 否 | 解析器配置 |

### ExportConfig 配置项

| 字段 | 类型 | 必填 | 说明 |
|------|------|------|------|
| key | String | 是 | 数据键名 |
| header | HeaderConfig | 是 | 表头配置 |
| mode | String | 是 | 填充模式 (FILL_CELL/FILL_DOWN/FILL_TABLE) |
| columns | List<ColumnConfig> | 否 | 列配置（FILL_TABLE 模式需要） |
| headerStyle | StyleConfig | 否 | 表头样式 |
| style | StyleConfig | 否 | 单元格样式 |
| maxRows | Integer | 否 | 最大行数 |
| alternateRows | Boolean | 否 | 是否隔行换色 |
| autoWidth | Boolean | 否 | 是否自动列宽 |

### 提取模式

| 模式 | 说明 |
|------|------|
| SINGLE | 提取单个单元格 |
| DOWN | 向下提取列数据 |
| RIGHT | 向右提取行数据 |
| BLOCK | 提取区域矩阵 |
| UNTIL_EMPTY | 提取到空行停止 |

### 填充模式

| 模式 | 说明 |
|------|------|
| FILL_CELL | 填充单个单元格 |
| FILL_DOWN | 向下填充列数据 |
| FILL_RIGHT | 向右填充行数据 |
| FILL_BLOCK | 填充区域矩阵 |
| FILL_TABLE | 填充表格（带表头） |
| APPEND_ROWS | 追加行 |
| APPEND_COLS | 追加列 |

---

## 常见问题

### Q: 如何处理多个工作表？

A: 在配置中指定工作表名称或索引：

```json
{
  "extractions": [
    {
      "key": "orderNos",
      "sheet": "订单表",
      "header": { "match": "订单号" },
      "mode": "DOWN"
    }
  ]
}
```

### Q: 如何跳过空行？

A: 在 range 配置中设置 skipEmpty：

```json
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
```

### Q: 如何限制最大行数？

A: 在 range 配置中设置 maxRows：

```json
{
  "extractions": [
    {
      "key": "orderNos",
      "header": { "match": "订单号" },
      "mode": "DOWN",
      "range": { "maxRows": 1000 }
    }
  ]
}
```

### Q: 如何自定义数字格式？

A: 在 parser 配置中设置 format：

```json
{
  "extractions": [
    {
      "key": "amounts",
      "header": { "match": "金额" },
      "mode": "DOWN",
      "parser": {
        "type": "number",
        "format": "#,##0.00"
      }
    }
  ]
}
```

---

## 更多信息

- [完整配置参考](../README.md#-配置说明)
- [提取模式详解](../docs/EXTRACT_MODES.md)
- [填充模式详解](../docs/FILL_MODES.md)
- [表头匹配机制](../docs/HEADER_MATCHING.md)
