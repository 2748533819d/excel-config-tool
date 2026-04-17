# Excel Config Tool

> 📊 配置驱动的 Excel 导入导出工具 - 通过表头匹配定位，数据量驱动，自动扩展空间

---

## 🎯 核心特性

| 特性 | 说明 |
|------|------|
| **表头匹配** | 通过表头文字匹配定位，不依赖固定单元格位置 |
| **数据驱动** | 提取/填充行数由实际数据决定，不是配置写死 |
| **自动扩展** | 模板空间不足时，自动下移下方内容，不会覆盖或截断 |
| **列隔离** | 每列独立处理，互不干扰 |
| **SAX 流式读取** | 内存优化，支持大文件处理 |
| **前后端独立** | 用户可选择只使用后端引擎、只使用前端组件、或两者结合 |

---

## 🚀 快速上手

### 导入（提取）配置

```json
// config.json
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

```java
// Java 代码
Map<String, Object> result = ExcelExtractor.extract(inputStream, config);
List<String> orderNos = (List<String>) result.get("orderNos");
```

### 导出（填充）配置

```json
// config.json
{
  "exports": [
    {
      "key": "orderNos",
      "header": { "match": "订单号" },
      "mode": "FILL_DOWN"
    }
  ]
}
```

```java
// Java 代码
Map<String, Object> data = Map.of("orderNos", Arrays.asList("ORD001", "ORD002", "ORD003"));
byte[] result = ExcelExporter.fill(templateInputStream, data, config);
```

---

## 📚 设计文档

完整的设计文档位于 [`docs/`](./docs) 文件夹：

### 核心设计
| 文档 | 说明 |
|------|------|
| [FINAL_DESIGN.md](./docs/FINAL_DESIGN.md) | **最终设计方案** - 整合所有核心设计 ⭐ |
| [ARCHITECTURE.md](./docs/ARCHITECTURE.md) | 系统架构设计 |
| [FRONTEND_DESIGN.md](./docs/FRONTEND_DESIGN.md) | 前端组件设计 |

### 模式详解
| 文档 | 说明 |
|------|------|
| [EXTRACT_MODES.md](./docs/EXTRACT_MODES.md) | 5 种基础提取模式 |
| [EXTENDED_MODES.md](./docs/EXTENDED_MODES.md) | 10 种扩展提取模式 |
| [FILL_MODES.md](./docs/FILL_MODES.md) | 10 种导出/填充模式 |

### 核心机制
| 文档 | 说明 |
|------|------|
| [HEADER_MATCHING.md](./docs/HEADER_MATCHING.md) | 表头匹配与动态扩展 |
| [COLUMN_ISOLATION.md](./docs/COLUMN_ISOLATION.md) | 列隔离与行偏移 |
| [DYNAMIC_ROW_COUNT.md](./docs/DYNAMIC_ROW_COUNT.md) | 动态行数确定 |

👉 **完整文档导航**: [docs/README_DOCS.md](./docs/README_DOCS.md)

---

## 🏗️ 项目结构

```
excel-config-tool/
├── packages/
│   ├── core/                    # 核心引擎（Java）
│   ├── spring-boot-starter/     # Spring Boot 集成
│   ├── ui-vue/                  # Vue 3 组件
│   └── ui-react/                # React 组件（Phase 2）
├── docs/                        # 设计文档
└── README.md                    # 本文件
```

---

## 📦 技术栈

| 模块 | 技术 |
|------|------|
| 后端核心 | Java 21 + Apache POI 5.2.5 + Jackson 2.16.1 |
| Spring Boot 集成 | Spring Boot 3.2.1 |
| 前端核心 | Univer + Vue 3 + TypeScript (Phase 2) |

---

## 📋 实施计划

| 阶段 | 内容 | 状态 |
|------|------|------|
| Phase 1 | 核心引擎 + Vue 组件 | ✅ 核心引擎完成 |
| Phase 2 | Spring Boot 集成 + React 组件 | ✅ Spring Boot Starter 完成 |
| Phase 3 | 完善文档 + 测试 | 🔄 进行中 |

---

## 🔨 构建与测试

```bash
# 设置 Java 21
export JAVA_HOME=$(/usr/libexec/java_home -v 21)

# 构建并安装到本地 Maven 仓库
mvn clean install

# 运行测试
mvn test
```

### 测试结果

```
[INFO] Tests run: 14, Failures: 0, Errors: 0, Skipped: 0 - Core 模块
[INFO] Tests run: 3, Failures: 0, Errors: 0, Skipped: 0 - Spring Boot Starter 模块
[INFO] BUILD SUCCESS
```

---

## 📚 类似项目对比

| 项目 | 类型 | Stars | 差异化 |
|------|------|-------|--------|
| [Alibaba EasyExcel](https://github.com/alibaba/easyexcel) | 注解驱动 | 33,750+ | 流式读写，大文件处理 |
| **本项目** | 配置驱动 | - | JSON 配置，表头自动定位，列隔离 |

---

## 📄 License

MIT License
