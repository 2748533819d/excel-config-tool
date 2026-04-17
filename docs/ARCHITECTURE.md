# Excel Config Tool - 系统架构设计 (v3)

> **核心定位**：前端组件库 + 后端引擎库，两者独立发布，用户自由选择组合

---

## 一、项目整体定位

```
┌─────────────────────────────────────────────────────────────────────────────────┐
│                           Excel Config Tool                                     │
│                                                                                 │
│     一个配置化的 Excel 数据提取工具，提供前端组件和后端引擎                        │
├─────────────────────────────────────────────────────────────────────────────────┤
│                                                                                 │
│   ┌───────────────────────────────┐       ┌───────────────────────────────┐     │
│   │   @excel-config/ui            │       │   @excel-config/core          │     │
│   │   前端组件库                   │       │   后端引擎库                   │     │
│   │   (npm publish)               │       │   (maven central)             │     │
│   │                               │       │                               │     │
│   │   • UniverSheet               │       │   • ConfigEngine              │     │
│   │   • ConfigDesigner            │       │   • ExtractEngine             │     │
│   │   • PreviewPanel              │       │   • FillEngine                │     │
│   │   • useSelection 等 Hooks     │       │   • SPI (Strategy/Parser/     │     │
│   │   • Utils (cellRef 等)         │       │       Handler)                │     │
│   │                               │       │   • 内置策略实现               │     │
│   │   依赖：Univer, React         │       │   依赖：Apache POI            │     │
│   └───────────────────────────────┘       └───────────────────────────────┘     │
│                                                                                 │
│   两者关系：                                                                     │
│   • 独立开发、独立版本、独立发布                                                  │
│   • 通过 JSON Schema 约定配置格式                                                 │
│   • 用户可以选择：                                                               │
│     - 只用前端组件 + 自己写后端                                                  │
│     - 只用后端引擎 + 自己写前端                                                  │
│     - 两者都用 (通过 Demo 项目)                                                   │
│                                                                                 │
└─────────────────────────────────────────────────────────────────────────────────┘
```

---

## 二、模块划分

```
excel-config-tool/
│
├── packages/
│   │
│   ├── ui-vue/                        # 前端组件包 - Vue 版本 (Phase 1)
│   │   ├── package.json               # name: @excel-config/ui-vue
│   │   ├── package.json               # name: @excel-config/ui
│   │   ├── src/
│   │   │   ├── components/
│   │   │   │   ├── UniverSheet/
│   │   │   │   │   ├── UniverSheet.vue
│   │   │   │   │   ├── UniverSheet.types.ts
│   │   │   │   │   └── index.ts
│   │   │   │   ├── ConfigDesigner/
│   │   │   │   │   ├── ConfigDesigner.vue
│   │   │   │   │   ├── ConfigDesigner.types.ts
│   │   │   │   │   └── index.ts
│   │   │   │   ├── PreviewPanel/
│   │   │   │   │   ├── PreviewPanel.vue
│   │   │   │   │   └── index.ts
│   │   │   │   └── index.ts         # 统一导出
│   │   │   ├── composables/
│   │   │   │   ├── useSelection.ts
│   │   │   │   ├── useConfig.ts
│   │   │   │   └── index.ts
│   │   │   ├── utils/
│   │   │   │   ├── cellRef.ts       # getCellRef, parseCellRef
│   │   │   │   └── index.ts
│   │   │   └── schema/
│   │   │       ├── config.schema.ts # JSON Schema (TypeScript 版)
│   │   │       └── index.ts
│   │   └── README.md
│   │
│   ├── ui-react/                      # 前端组件包 - React 版本 (Phase 2, TODO)
│   │   ├── package.json               # name: @excel-config/ui-react
│   │   └── src/
│   │       ├── components/
│   │       │   ├── UniverSheet/
│   │       │   ├── ConfigDesigner/
│   │       │   └── PreviewPanel/
│   │       ├── hooks/
│   │       │   ├── useSelection.ts
│   │       │   ├── useConfig.ts
│   │       │   └── index.ts
│   │       ├── utils/
│   │       │   └── ...
│   │       └── schema/
│   │           └── ...
│   │
│   ├── core/                        # 后端核心包
│   │   ├── pom.xml                  # artifactId: excel-config-core
│   │   └── src/main/java/
│   │       └── com/excelconfig/
│   │           ├── engine/
│   │           │   ├── ConfigEngine.java
│   │           │   ├── ExtractEngine.java
│   │           │   └── FillEngine.java
│   │           ├── config/
│   │           │   ├── ExcelConfig.java
│   │           │   ├── CellConfig.java
│   │           │   └── ExtractMode.java
│   │           ├── spi/
│   │           │   ├── ExtractStrategy.java
│   │           │   ├── CellParser.java
│   │           │   └── DataHandler.java
│   │           └── strategy/        # 内置策略实现
│   │               ├── SingleStrategy.java
│   │               └── ...
│   │
│   └── spring-boot-starter/         # Spring Boot 集成包 (可选)
│       ├── pom.xml                  # artifactId: excel-config-spring-boot-starter
│       └── src/main/java/
│           └── com/excelconfig/
│               └── autoconfigure/
│                   ├── ExcelConfigAutoConfiguration.java
│                   └── ...
│
├── examples/
│   ├── web-demo/                    # 完整 Web 示例 (展示如何用)
│   │   ├── frontend/                # React 前端
│   │   └── backend/                 # Spring Boot 后端
│   └── cli-demo/                    # CLI 工具示例
│
└── docs/
    ├── architecture.md              # 本文档
    ├── frontend-guide.md            # 前端使用指南
    ├── backend-guide.md             # 后端使用指南
    └── schema.md                    # JSON Schema 文档
```

---

## 三、前端组件库设计

### 3.1 组件架构 - Vue 版本 (Phase 1)

```
┌─────────────────────────────────────────────────────────────────────────────────┐
│                  @excel-config/ui-vue 组件架构 (Vue 3 + TypeScript)               │
├─────────────────────────────────────────────────────────────────────────────────┤
│                                                                                 │
│  ┌───────────────────────────────────────────────────────────────────────────┐  │
│  │                         核心组件 (无状态)                                  │  │
│  │                                                                           │  │
│  │  ┌─────────────────────────────────────────────────────────────────────┐  │  │
│  │  │  UniverSheet.vue                                                    │  │  │
│  │  │                                                                     │  │  │
│  │  │  Props:                                                             │  │  │
│  │  │    file?: File                      // Excel 文件                     │  │  │
│  │  │    highlightRanges?: Range[]        // 高亮区域                       │  │  │
│  │  │    readOnly?: boolean               // 只读模式                       │  │  │
│  │  │    className?: string               // 自定义样式类                    │  │  │
│  │  │                                                                     │  │  │
│  │  │  Emits:                                                             │  │  │
│  │  │    selection-change: (ranges: Range[]) => void    // 选区变化        │  │  │
│  │  │    file-loaded: () => void                        // 文件加载完成    │  │  │
│  │  │    error: (error: Error) => void                  // 错误处理        │  │  │
│  │  │                                                                     │  │  │
│  │  │  Expose Methods (通过 defineExpose):                                │  │  │
│  │  │    getSelectedRanges(): Range[]   // 获取当前选区                   │  │  │
│  │  │    highlightCells(ranges: Range[]): void  // 高亮单元格             │  │  │
│  │  │    clearHighlights(): void              // 清除高亮                 │  │  │
│  │  │    loadFile(file: File): void           // 加载文件                 │  │  │
│  │  │                                                                     │  │  │
│  │  │  特点：                                                             │  │  │
│  │  │    • 无状态组件，状态由父组件管理                                     │  │  │
│  │  │    • 不依赖任何后端 API                                             │  │  │
│  │  │    • 可以独立使用                                                   │  │  │
│  │  └─────────────────────────────────────────────────────────────────────┘  │  │
│  │                                                                           │  │
│  │  ┌─────────────────────────────────────────────────────────────────────┐  │  │
│  │  │  ConfigDesigner.vue                                                 │  │  │
│  │  │                                                                     │  │  │
│  │  │  Props:                                                             │  │  │
│  │  │    selectedRanges: Range[]          // 当前选区                      │  │  │
│  │  │    configs: FieldConfig[]           // 配置列表                      │  │  │
│  │  │    cellValues?: Record<string, any> // 选中单元格的值 (用于参考)      │  │  │
│  │  │    readOnly?: boolean               // 只读模式                      │  │  │
│  │  │                                                                     │  │  │
│  │  │  Emits:                                                             │  │  │
│  │  │    config-add: (config: FieldConfig) => void      // 添加配置       │  │  │
│  │  │    config-update: (id: string, cfg: FieldConfig) => void // 更新    │  │  │
│  │  │    config-delete: (id: string) => void            // 删除配置       │  │  │
│  │  │    save: (configs: FieldConfig[]) => void         // 保存配置       │  │  │
│  │  │                                                                     │  │  │
│  │  │  特点：                                                             │  │  │
│  │  │    • 纯 UI 组件，不负责存储                                           │  │  │
│  │  │    • 不发送 HTTP 请求                                                │  │  │
│  │  │    • 配置格式符合 JSON Schema                                       │  │  │
│  │  └─────────────────────────────────────────────────────────────────────┘  │  │
│  │                                                                           │  │
│  │  ┌─────────────────────────────────────────────────────────────────────┐  │  │
│  │  │  PreviewPanel.vue                                                   │  │  │
│  │  │                                                                     │  │  │
│  │  │  Props:                                                             │  │  │
│  │  │    data: ExtractResult              // 提取结果数据                   │  │  │
│  │  │    loading?: boolean                // 加载状态                       │  │  │
│  │  │    errors?: ValidationError[]       // 验证错误                      │  │  │
│  │  │                                                                     │  │  │
│  │  │  Emits:                                                             │  │  │
│  │  │    confirm: () => void              // 确认                          │  │  │
│  │  │    retry: () => void                // 重试                          │  │  │
│  │  │                                                                     │  │  │
│  │  │  特点：                                                             │  │  │
│  │  │    • 纯展示组件                                                     │  │  │
│  │  │    • 不负责实际导入                                                 │  │  │
│  │  └─────────────────────────────────────────────────────────────────────┘  │  │
│  └───────────────────────────────────────────────────────────────────────────┘  │
│                                                                                 │
│  ┌───────────────────────────────────────────────────────────────────────────┐  │
│  │                         Composables (组合式函数 - Vue 3)                    │  │
│  │                                                                           │  │
│  │  useSelection:        选区状态管理                                         │  │
│  │    const { ranges, setRanges, clear } = useSelection();                  │  │
│  │                                                                           │  │
│  │  useConfig:           配置 CRUD                                            │  │
│  │    const { configs, addConfig, updateConfig, deleteConfig, serialize }   │  │
│  │      = useConfig();                                                        │  │
│  │                                                                           │  │
│  │  usePreview:          预览数据管理                                         │  │
│  │    const { result, loading, error, execute } = usePreview();             │  │
│  │                                                                           │  │
│  └───────────────────────────────────────────────────────────────────────────┘  │
│                                                                                 │
│  ┌───────────────────────────────────────────────────────────────────────────┐  │
│  │                         Utils (工具函数)                                   │  │
│  │                                                                           │  │
│  │  cellRef.ts:                                                              │  │
│  │    getCellRef(row: number, col: number): string     // (0,0) => "A1"     │  │
│  │    parseCellRef(ref: string): { row, col }          // "A1" => (0,0)     │  │
│  │    parseAreaRef(areaRef: string): Range             // "A1:C10" => Range │  │
│  │                                                                           │  │
│  │  schema.ts:                                                               │  │
│  │    validateConfig(config: any): ValidationResult  // 配置验证             │  │
│  │                                                                           │  │
│  └───────────────────────────────────────────────────────────────────────────┘  │
│                                                                                 │
│  技术栈：                                                                       │
│  • Vue 3.4+ (Composition API)                                                   │
│  • TypeScript 5.x                                                               │
│  • Univer (表格引擎)                                                            │
│  • Element Plus / Ant Design Vue (UI 组件)                                      │
│  • Vite 5.x (构建工具)                                                          │
│                                                                                 │
└─────────────────────────────────────────────────────────────────────────────────┘
```

### 3.2 使用示例 (Vue)

```vue
<!-- ===== 场景 1：只用前端组件 + 自定义后端 ===== -->
<template>
  <div class="excel-config-page">
    <UniverSheet 
      ref="sheetRef"
      :file="excelFile"
      @selection-change="handleSelectionChange"
    />
    <ConfigDesigner
      :selected-ranges="ranges"
      :configs="configs"
      @config-add="addConfig"
      @save="handleSave"
    />
  </div>
</template>

<script setup lang="ts">
import { ref } from 'vue';
import { UniverSheet, ConfigDesigner } from '@excel-config/ui-vue';
import { useConfig } from '@excel-config/ui-vue/composables';

const sheetRef = ref();
const ranges = ref([]);
const { configs, addConfig, serialize } = useConfig();

const handleSelectionChange = (newRanges: Range[]) => {
  ranges.value = newRanges;
};

const handleSave = async () => {
  const configJson = serialize();
  // 发送到自己的后端
  await fetch('/my-api/config', {
    method: 'POST',
    body: JSON.stringify(configJson),
  });
};
</script>
```

---

## 四、后端引擎库设计

### 4.1 核心引擎架构

```
┌─────────────────────────────────────────────────────────────────────────────────┐
│                     @excel-config/core 引擎架构                                 │
├─────────────────────────────────────────────────────────────────────────────────┤
│                                                                                 │
│  ┌───────────────────────────────────────────────────────────────────────────┐  │
│  │                         配置模型 (Config Model)                             │  │
│  │                                                                           │  │
│  │  public class ExcelConfig {                                               │  │
│  │      private String version;        // "1.0"                              │  │
│  │      private String templateName;   // 模板名称                           │  │
│  │      private List<SheetConfig> sheets;                                    │  │
│  │      private Map<String, CellConfig> cells;                               │  │
│  │  }                                                                        │  │
│  │                                                                           │  │
│  │  public class CellConfig {                                                │  │
│  │      private String key;            // 字段名                             │  │
│  │      private Position position;     // 位置                               │  │
│  │      private ExtractMode mode;      // 提取模式                           │  │
│  │      private ExtractRange range;    // 提取范围                           │  │
│  │      private ParserConfig parser;   // 解析器配置                         │  │
│  │      private List<HandlerConfig> handlers; // 处理器链                    │  │
│  │  }                                                                        │  │
│  │                                                                           │  │
│  │  public enum ExtractMode {                                                │  │
│  │      SINGLE,      // 单一单元格                                            │  │
│  │      DOWN,        // 向下提取                                              │  │
│  │      RIGHT,       // 向右提取                                              │  │
│  │      BLOCK,       // 区域块                                                │  │
│  │      UNTIL_EMPTY  // 直到空值                                              │  │
│  │  }                                                                        │  │
│  │                                                                           │  │
│  └───────────────────────────────────────────────────────────────────────────┘  │
│                                                                                 │
│  ┌───────────────────────────────────────────────────────────────────────────┐  │
│  │                         三大引擎 (Core Engines)                             │  │
│  │                                                                           │  │
│  │  ┌─────────────────────────────────────────────────────────────────────┐  │  │
│  │  │  ConfigEngine - 配置引擎                                             │  │  │
│  │  │                                                                      │  │  │
│  │  │  public class ConfigEngine {                                        │  │  │
│  │  │      // 从 YAML 解析                                                   │  │  │
│  │  │      ExcelConfig parseYaml(InputStream input);                      │  │  │
│  │  │                                                                     │  │  │
│  │  │      // 从 JSON 解析                                                   │  │  │
│  │  │      ExcelConfig parseJson(InputStream input);                      │  │  │
│  │  │                                                                     │  │  │
│  │  │      // 从 JSON 字符串解析                                             │  │  │
│  │  │      ExcelConfig parseJson(String json);                            │  │  │
│  │  │                                                                     │  │  │
│  │  │      // 验证配置                                                     │  │  │
│  │  │      ValidationResult validate(ExcelConfig config);                 │  │  │
│  │  │                                                                     │  │  │
│  │  │      // 序列化为 JSON                                                │  │  │
│  │  │      String toJson(ExcelConfig config);                             │  │  │
│  │  │  }                                                                  │  │  │
│  │  └─────────────────────────────────────────────────────────────────────┘  │  │
│  │                                                                           │  │
│  │  ┌─────────────────────────────────────────────────────────────────────┐  │  │
│  │  │  ExtractEngine - 提取引擎                                            │  │  │
│  │  │                                                                      │  │  │
│  │  │  public class ExtractEngine {                                       │  │  │
│  │  │      // 流式读取 (内存优化)                                           │  │  │
│  │  │      <T> T readStreaming(                                           │  │  │
│  │  │          InputStream input,                                         │  │  │
│  │  │          ExcelConfig config,                                        │  │  │
│  │  │          Class<T> resultType                                        │  │  │
│  │  │      );                                                             │  │  │
│  │  │                                                                     │  │  │
│  │  │      // 从 Workbook 提取 (用户已有 Workbook 时)                        │  │  │
│  │  │      <T> T extract(                                                 │  │  │
│  │  │          Workbook workbook,                                         │  │  │
│  │  │          ExcelConfig config                                         │  │  │
│  │  │      );                                                             │  │  │
│  │  │                                                                     │  │  │
│  │  │      // 执行单个配置                                                 │  │  │
│  │  │      Object extractCell(                                            │  │  │
│  │  │          Sheet sheet,                                               │  │  │
│  │  │          CellConfig cellConfig                                      │  │  │
│  │  │      );                                                             │  │  │
│  │  │  }                                                                  │  │  │
│  │  └─────────────────────────────────────────────────────────────────────┘  │  │
│  │                                                                           │  │
│  │  ┌─────────────────────────────────────────────────────────────────────┐  │  │
│  │  │  FillEngine - 填充引擎                                               │  │  │
│  │  │                                                                      │  │  │
│  │  │  public class FillEngine {                                          │  │  │
│  │  │      // 填充模板                                                     │  │  │
│  │  │      byte[] fillTemplate(                                           │  │  │
│  │  │          InputStream template,                                      │  │  │
│  │  │          Map<String, Object> data,                                  │  │  │
│  │  │          ExcelConfig config                                         │  │  │
│  │  │      );                                                             │  │  │
│  │  │                                                                     │  │  │
│  │  │      // 生成新 Excel                                                 │  │  │
│  │  │      byte[] generate(                                               │  │  │
│  │  │          List<List<Object>> data,                                   │  │  │
│  │  │          ExcelConfig config                                         │  │  │
│  │  │      );                                                             │  │  │
│  │  │  }                                                                  │  │  │
│  │  └─────────────────────────────────────────────────────────────────────┘  │  │
│  └───────────────────────────────────────────────────────────────────────────┘  │
│                                                                                 │
│  ┌───────────────────────────────────────────────────────────────────────────┐  │
│  │                         SPI 接口 (用户可扩展)                               │  │
│  │                                                                           │  │
│  │  public interface ExtractStrategy {                                       │  │
│  │      Object extract(Sheet sheet, CellConfig config);                      │  │
│  │      Set<ExtractMode> supportedModes();                                   │  │
│  │  }                                                                        │  │
│  │                                                                           │  │
│  │  public interface CellParser<T> {                                         │  │
│  │      T parse(String rawValue, ParseContext context);                      │  │
│  │      String getName();                                                    │  │
│  │  }                                                                        │  │
│  │                                                                           │  │
│  │  public interface DataHandler<T> {                                        │  │
│  │      T handle(T data, HandlerContext context);                            │  │
│  │      int getOrder();                                                      │  │
│  │  }                                                                        │  │
│  │                                                                           │  │
│  └───────────────────────────────────────────────────────────────────────────┘  │
│                                                                                 │
│  ┌───────────────────────────────────────────────────────────────────────────┐  │
│  │                         内置策略实现                                        │  │
│  │                                                                           │  │
│  │  ExtractStrategy:                                                         │  │
│  │    • SingleCellStrategy   - 单一单元格                                     │  │
│  │    • DownListStrategy     - 向下列表                                       │  │
│  │    • RightListStrategy    - 向右列表                                       │  │
│  │    • BlockStrategy        - 区域块                                         │  │
│  │    • UntilEmptyStrategy   - 直到空值                                       │  │
│  │                                                                           │  │
│  │  CellParser:                                                              │  │
│  │    • StringParser         - 字符串                                         │  │
│  │    • NumberParser         - 数字                                           │  │
│  │    • DateParser           - 日期                                           │  │
│  │    • BooleanParser        - 布尔值                                         │  │
│  │                                                                           │  │
│  │  DataHandler:                                                             │  │
│  │    • TrimHandler          - 去空格                                         │  │
│  │    • NotEmptyHandler      - 非空验证                                       │  │
│  │    • TransformHandler     - 数据转换                                       │  │
│  │                                                                           │  │
│  └───────────────────────────────────────────────────────────────────────────┘  │
│                                                                                 │
└─────────────────────────────────────────────────────────────────────────────────┘
```

### 4.2 使用示例

```java
// ===== 场景 1：独立使用引擎 =====
import com.excelconfig.core.engine.*;
import com.excelconfig.core.config.*;
import com.excelconfig.core.spi.*;

public class MyService {
    private ExtractEngine extractEngine = new ExtractEngine();
    private ConfigEngine configEngine = new ConfigEngine();
    
    public Map<String, Object> importData(
        InputStream excelFile,
        String configJson
    ) {
        // 1. 解析配置
        ExcelConfig config = configEngine.parseJson(configJson);
        
        // 2. 验证配置
        ValidationResult result = configEngine.validate(config);
        if (!result.isValid()) {
            throw new InvalidConfigException(result.getErrors());
        }
        
        // 3. 执行提取 (SAX 流式，内存优化)
        return extractEngine.readStreaming(excelFile, config, Map.class);
    }
}

// ===== 场景 2：自定义策略 =====
public class MyCustomStrategy implements ExtractStrategy {
    @Override
    public Object extract(Sheet sheet, CellConfig config) {
        // 自定义提取逻辑
        // 例如：提取满足特定条件的单元格
    }
    
    @Override
    public Set<ExtractMode> supportedModes() {
        return Set.of(ExtractMode.CUSTOM);
    }
}

// 注册策略
StrategyRegistry registry = new StrategyRegistry();
registry.register(ExtractMode.CUSTOM, new MyCustomStrategy());

// ===== 场景 3：使用 Spring Boot Starter =====
@SpringBootApplication
public class MyApplication {
    
    @Autowired
    private ExtractEngine extractEngine;  // 自动注入
    
    @Autowired
    private ConfigEngine configEngine;    // 自动注入
    
    @PostMapping("/import")
    public ResponseEntity<?> importData(
        @RequestParam("file") MultipartFile file,
        @RequestParam("configId") String configId
    ) {
        ExcelConfig config = loadConfig(configId); // 从数据库加载
        Map<String, Object> data = extractEngine.extract(
            file.getInputStream(), 
            config
        );
        return ResponseEntity.ok(data);
    }
}
```

---

## 五、配置格式 (JSON Schema)

### 5.1 Schema 定义

```json
{
  "$schema": "http://json-schema.org/draft-07/schema#",
  "$id": "https://excel-config.tool/schema/v1",
  "title": "ExcelTemplateConfig",
  "description": "Excel 配置化提取工具的配置格式",
  "type": "object",
  "required": ["version", "templateName", "cells"],
  "properties": {
    "version": {
      "type": "string",
      "enum": ["1.0"],
      "description": "配置版本"
    },
    "templateName": {
      "type": "string",
      "description": "模板名称"
    },
    "description": {
      "type": "string",
      "description": "模板描述"
    },
    "cells": {
      "type": "object",
      "description": "单元格配置映射",
      "additionalProperties": {
        "$ref": "#/definitions/CellConfig"
      }
    }
  },
  "definitions": {
    "CellConfig": {
      "type": "object",
      "required": ["position", "mode"],
      "properties": {
        "key": { "type": "string" },
        "position": { "$ref": "#/definitions/Position" },
        "mode": { "$ref": "#/definitions/ExtractMode" },
        "range": { "$ref": "#/definitions/ExtractRange" },
        "parser": { "$ref": "#/definitions/ParserConfig" },
        "handlers": {
          "type": "array",
          "items": { "$ref": "#/definitions/HandlerConfig" }
        }
      }
    },
    "Position": {
      "type": "object",
      "properties": {
        "cellRef": { 
          "type": "string", 
          "pattern": "^[A-Z]+\\d+$",
          "description": "Excel 单元格引用，如 A1, B2"
        },
        "areaRef": {
          "type": "string",
          "pattern": "^[A-Z]+\\d+:[A-Z]+\\d+$",
          "description": "Excel 区域引用，如 A1:C10"
        },
        "headerName": {
          "type": "string",
          "description": "表头名称（自动查找该列）"
        }
      }
    },
    "ExtractMode": {
      "type": "string",
      "enum": ["SINGLE", "DOWN", "RIGHT", "BLOCK", "UNTIL_EMPTY"],
      "description": "提取模式"
    },
    "ExtractRange": {
      "type": "object",
      "properties": {
        "rows": { "type": "integer", "minimum": 1 },
        "cols": { "type": "integer", "minimum": 1 },
        "skipEmpty": { "type": "boolean" }
      }
    }
  }
}
```

### 5.2 配置示例

```json
{
  "version": "1.0",
  "templateName": "订单导入模板",
  "description": "用于导入订单数据的 Excel 模板",
  "cells": {
    "orderNo": {
      "key": "orderNo",
      "position": { "cellRef": "A2" },
      "mode": "DOWN",
      "range": { "rows": 100, "skipEmpty": true },
      "parser": { "type": "string" },
      "handlers": [
        { "type": "trim" },
        { "type": "notEmpty" }
      ]
    },
    "amount": {
      "key": "amount",
      "position": { "cellRef": "B2" },
      "mode": "DOWN",
      "range": { "rows": 100 },
      "parser": { "type": "number", "params": { "scale": 2 } }
    },
    "orderDate": {
      "key": "orderDate",
      "position": { "cellRef": "C2" },
      "mode": "DOWN",
      "range": { "rows": 100 },
      "parser": { "type": "date", "params": { "pattern": "yyyy-MM-dd" } }
    }
  }
}
```

---

## 六、前后端协作流程

```
┌─────────────────────────────────────────────────────────────────────────────────┐
│                        典型使用流程 (用户视角)                                   │
├─────────────────────────────────────────────────────────────────────────────────┤
│                                                                                 │
│  步骤 1: 配置设计阶段                                                             │
│  ┌───────────────────────────────────────────────────────────────────────────┐  │
│  │                                                                            │  │
│  │   用户打开配置界面 (使用 @excel-config/ui)                                  │  │
│  │                                                                            │  │
│  │   ┌─────────────────┐      ┌─────────────────┐                            │  │
│  │   │   UniverSheet   │      │ ConfigDesigner  │                            │  │
│  │   │                 │      │                 │                            │  │
│  │   │  [上传 Excel]   │      │  字段名：[___]  │                            │  │
│  │   │                 │      │  位置：A2       │                            │  │
│  │   │  A | B | C      │      │  模式：[DOWN▼]  │                            │  │
│  │   │  ──┼───┼───     │      │  行数：[100 ]   │                            │  │
│  │   │  订单│金额│日期 │      │                 │                            │  │
│  │   │  [选中 A2]      │      │  [+添加] [保存] │                            │  │
│  │   └─────────────────┘      └─────────────────┘                            │  │
│  │                                                                            │  │
│  │   输出：JSON 配置                                                            │  │
│  │   { "cells": { "orderNo": { "position": {"cellRef":"A2"}, ... } } }       │  │
│  │                                                                            │  │
│  └───────────────────────────────────────────────────────────────────────────┘  │
│                                                                                 │
│  步骤 2: 配置存储阶段                                                             │
│  ┌───────────────────────────────────────────────────────────────────────────┐  │
│  │                                                                            │  │
│  │   用户自己决定如何存储配置：                                                 │  │
│  │                                                                            │  │
│  │   • 存到 MySQL 数据库                                                       │  │
│  │   • 存到 Redis 缓存                                                          │  │
│  │   • 存到文件系统                                                             │  │
│  │   • 存到配置中心 (Nacos/Apollo)                                             │  │
│  │   • 硬编码在代码中                                                           │  │
│  │                                                                            │  │
│  │   本项目不限制存储方式！                                                     │  │
│  │                                                                            │  │
│  └───────────────────────────────────────────────────────────────────────────┘  │
│                                                                                 │
│  步骤 3: 数据提取阶段                                                             │
│  ┌───────────────────────────────────────────────────────────────────────────┐  │
│  │                                                                            │  │
│  │   用户调用后端引擎：                                                         │  │
│  │                                                                            │  │
│  │   Java:                                                                     │  │
│  │   ExcelConfig config = configEngine.parseJson(configJson);                │  │
│  │   Map<String, Object> result = extractEngine.readStreaming(               │  │
│  │       excelFile, config                                                    │  │
│  │   );                                                                        │  │
│  │                                                                            │  │
│  │   输出：                                                                    │  │
│  │   {                                                                        │  │
│  │     "orderNo": ["ORD001", "ORD002", ...],                                 │  │
│  │     "amount": [100.00, 200.00, ...],                                      │  │
│  │     "orderDate": ["2024-01-01", "2024-01-02", ...]                        │  │
│  │   }                                                                         │  │
│  │                                                                            │  │
│  └───────────────────────────────────────────────────────────────────────────┘  │
│                                                                                 │
│  步骤 4: 后续处理                                                                 │
│  ┌───────────────────────────────────────────────────────────────────────────┐  │
│  │                                                                            │  │
│  │   用户自己决定如何处理结果：                                                 │  │
│  │                                                                            │  │
│  │   • 转换为业务 DTO 插入数据库                                                │  │
│  │   • 调用其他服务 API                                                        │  │
│  │   • 生成报告文件                                                             │  │
│  │   • 发送通知消息                                                             │  │
│  │                                                                            │  │
│  │   本项目不限制业务逻辑！                                                     │  │
│  │                                                                            │  │
│  └───────────────────────────────────────────────────────────────────────────┘  │
│                                                                                 │
└─────────────────────────────────────────────────────────────────────────────────┘
```

---

## 七、发布计划

```
┌─────────────────────────────────────────────────────────────────────────────────┐
│                              发布计划                                            │
├─────────────────────────────────────────────────────────────────────────────────┤
│                                                                                 │
│  Phase 1: 核心引擎 + Vue 组件 (MVP)                                               │
│  ┌───────────────────────────────────────────────────────────────────────────┐  │
│  │  @excel-config/core v1.0.0          [Maven Central]                       │  │
│  │  • ConfigEngine (YAML/JSON 解析)                                            │  │
│  │  • ExtractEngine (SAX 流式读取)                                             │  │
│  │  • 内置策略 (SINGLE/DOWN/RIGHT/BLOCK/UNTIL_EMPTY)                          │  │
│  │  • 内置解析器 (String/Number/Date/Boolean)                                 │  │
│  │                                                                            │  │
│  │  @excel-config/ui-vue v1.0.0        [npm]                                  │  │
│  │  • UniverSheet 组件 (Vue 3)                                                  │  │
│  │  • ConfigDesigner 组件 (Vue 3)                                               │  │
│  │  • PreviewPanel 组件 (Vue 3)                                                 │  │
│  │  • Composables (useSelection/useConfig/usePreview)                         │  │
│  │  • Utils (cellRef/schema)                                                  │  │
│  └───────────────────────────────────────────────────────────────────────────┘  │
│                                                                                 │
│  Phase 2: Spring Boot 集成 + React 组件                                          │
│  ┌───────────────────────────────────────────────────────────────────────────┐  │
│  │  @excel-config/spring-boot-starter v1.0.0  [Maven Central]                │  │
│  │  • 自动配置                                                                 │  │
│  │  • 条件 Bean                                                                │  │
│  │  • 属性配置                                                                 │  │
│  │                                                                            │  │
│  │  @excel-config/ui-react v1.0.0      [npm]                                  │  │
│  │  • UniverSheet 组件 (React)                                                  │  │
│  │  • ConfigDesigner 组件 (React)                                               │  │
│  │  • PreviewPanel 组件 (React)                                                 │  │
│  │  • Hooks (useSelection/useConfig/usePreview)                               │  │
│  └───────────────────────────────────────────────────────────────────────────┘  │
│                                                                                 │
│  Phase 3: Demo 项目                                                              │
│  ┌───────────────────────────────────────────────────────────────────────────┐  │
│  │  examples/web-demo-vue            [GitHub Releases]                        │  │
│  │  • Vue 3 完整示例                                                             │  │
│  │  • Docker Compose 一键部署                                                  │  │
│  │                                                                            │  │
│  │  examples/web-demo-react          [GitHub Releases]                        │  │
│  │  • React 完整示例                                                            │  │
│  │  • Docker Compose 一键部署                                                  │  │
│  └───────────────────────────────────────────────────────────────────────────┘  │
│                                                                                 │
└─────────────────────────────────────────────────────────────────────────────────┘
```

---

## 八、设计原则总结

```
┌─────────────────────────────────────────────────────────────────────────────────┐
│                              设计原则                                            │
├─────────────────────────────────────────────────────────────────────────────────┤
│                                                                                 │
│  1. 组件独立                                                                     │
│     前端组件和后端引擎独立发布，用户可以自由选择                                  │
│     前端提供 Vue 和 React 两种选择 (Vue 优先)                                       │
│                                                                                 │
│  2. 零耦合                                                                       │
│     前端不依赖后端 API，后端不依赖前端配置格式                                    │
│                                                                                 │
│  3. 配置开放                                                                     │
│     JSON Schema 公开，用户可以自己实现解析器                                      │
│                                                                                 │
│  4. SPI 拓展                                                                     │
│     策略/解析器/处理器支持用户自定义扩展                                          │
│                                                                                 │
│  5. 职责单一                                                                     │
│     每个组件/引擎只做一件事，保持简单可测试                                       │
│                                                                                 │
│  6. 内存优化                                                                     │
│     SAX 流式读取，类似 EasyExcel，支持大文件                                       │
│                                                                                 │
│  7. 无侵入                                                                       │
│     不强制用户采用特定存储方式、业务逻辑                                          │
│                                                                                 │
│  8. 渐进式发布                                                                   │
│     Phase 1: Vue + Core                                                         │
│     Phase 2: React + Spring Boot Starter                                        │
│                                                                                 │
└─────────────────────────────────────────────────────────────────────────────────┘
```
