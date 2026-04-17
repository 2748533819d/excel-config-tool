# Univer 前端方案详细设计

## 一、Univer 简介

### 1.1 什么是 Univer

**Univer** 是阿里开源的新一代在线协同表格引擎，支持 Excel 导入导出、公式、图表等功能。

- **GitHub**: https://github.com/dream-num/univer
- **官网**: https://univer.ai/
- **文档**: https://docs.univer.ai/
- **Stars**: 40,000+
- **License**: Apache 2.0

### 1.2 为什么选择 Univer

| 对比项 | Univer | Luckysheet | SheetJS |
|--------|--------|------------|---------|
| **在线编辑** | ✅ 完整支持 | ✅ 支持 | ❌ 仅解析 |
| **React/Vue集成** | ✅ 官方支持 | ⚠️ 社区方案 | ❌ |
| **单元格选择 API** | ✅ Facade API | ⚠️ 有限 | ❌ |
| **TypeScript支持** | ✅ 完整类型 | ⚠️ 部分 | ✅ |
| **活跃维护** | ✅ 2026 仍在更新 | ⚠️ 更新放缓 | ✅ |
| **插件系统** | ✅ 完善 | ⚠️ 一般 | ❌ |

### 1.3 核心能力

```
┌─────────────────────────────────────────────────────────┐
│                    Univer 能力矩阵                       │
├─────────────────────────────────────────────────────────┤
│                                                         │
│  📊 基础能力                                             │
│  - 单元格读写、样式设置                                   │
│  - 行列操作 (插入、删除、隐藏)                           │
│  - 多 Sheet 支持                                         │
│                                                         │
│  🔍 选区操作 (核心!)                                     │
│  - 获取当前选中的单元格范围                               │
│  - 监听选区变化事件                                      │
│  - 编程式设置选区                                        │
│                                                         │
│  📈 高级功能                                             │
│  - 公式计算                                              │
│  - 数据验证                                              │
│  - 条件格式                                              │
│  - 命名范围                                              │
│                                                         │
│  🔌 扩展能力                                             │
│  - 自定义插件                                            │
│  - 自定义 UI 组件                                         │
│  - 事件系统                                              │
│                                                         │
└─────────────────────────────────────────────────────────┘
```

---

## 二、Univer 核心 API - 选区与单元格操作

### 2.1 获取选区信息

```typescript
import { FUniver } from '@univerjs/core';

// 初始化后获取 Facade API
const fUniver = await FUniver.newAPI(instance);

// 获取当前工作簿
const workbook = fUniver.getActiveWorkbook();

// 获取当前工作表
const worksheet = workbook.getActiveSheet();

// 获取当前选区
const selection = worksheet.getSelection();

// 获取选区范围数组 (可能多选)
const ranges = selection?.getRangeList() || [];

// 遍历选区
ranges.forEach(range => {
  const { startRow, endRow, startColumn, endColumn } = range;
  console.log(`选区：${getCellRef(range)}`);
  
  // 获取单元格值
  for (let r = startRow; r <= endRow; r++) {
    for (let c = startColumn; c <= endColumn; c++) {
      const cellValue = worksheet.getCell(r, c)?.getValue();
      console.log(`Cell ${getColumnName(c)}${r + 1}: ${cellValue}`);
    }
  }
});
```

### 2.2 监听选区变化

```typescript
import { useEffect } from 'react';

// 在 React 组件中监听选区变化
useEffect(() => {
  const fUniver = await FUniver.newAPI(univerInstance);
  const workbook = fUniver.getActiveWorkbook();
  const worksheet = workbook.getActiveSheet();
  
  // 监听选区变化事件
  const subscription = worksheet.onSelectionChange((event) => {
    console.log('选区变化:', event);
    // event.ranges: 新的选区范围数组
    // 触发前端状态更新
    setSelectedRanges(event.ranges);
  });
  
  return () => subscription.unsubscribe();
}, []);
```

### 2.3 编程式设置选区

```typescript
// 设置单个单元格选中
worksheet.setSelection({
  startRow: 1,
  endRow: 1,
  startColumn: 0,
  endColumn: 0,
}); // 选中 A2

// 设置区域选中
worksheet.setSelection({
  startRow: 1,
  endRow: 10,
  startColumn: 0,
  endColumn: 3,
}); // 选中 A2:D11

// 多选区
worksheet.setSelections([
  { startRow: 0, endRow: 0, startColumn: 0, endColumn: 0 }, // A1
  { startRow: 5, endRow: 10, startColumn: 1, endColumn: 1 }, // B6:B11
]);
```

### 2.4 单元格操作

```typescript
// 读取单元格值
const value = worksheet.getCell(row, column)?.getValue();

// 设置单元格值
worksheet.setCellValue(row, column, '新的值');

// 设置单元格样式
worksheet.setCellStyle(row, column, {
  bg: { rgb: '#ff0000' },  // 背景色
  cl: { rgb: '#ffffff' },  // 字体颜色
  fs: 14,                   // 字体大小
  bl: 1,                    // 加粗
});

// 获取单元格样式
const style = worksheet.getCellStyle(row, column);
```

### 2.5 行列信息

```typescript
// 获取总行数
const rowCount = worksheet.getRowCount();

// 获取总列数
const colCount = worksheet.getColumnCount();

// 获取表头 (第一行)
const headers = [];
for (let c = 0; c < colCount; c++) {
  headers.push(worksheet.getCell(0, c)?.getValue());
}

// 获取列名 (0 -> 'A', 1 -> 'B', 26 -> 'AA')
const colName = getColumnName(columnIndex);
```

---

## 三、React + Univer 完整示例

### 3.1 项目初始化

```bash
# 创建项目
npm create vite@latest excel-config-frontend -- --template react-ts
cd excel-config-frontend

# 安装 Univer 核心依赖
npm install @univerjs/core @univerjs/ui @univerjs/preset-sheets-core
npm install @univerjs/icons @univerjs/design

# 安装样式
npm install -D less
```

### 3.2 Univer 初始化配置

```typescript
// src/univer/init.ts
import { Univer, UniverInstanceType } from '@univerjs/core';
import { defaultTheme } from '@univerjs/design';
import { UniverPresetSheetsCore } from '@univerjs/preset-sheets-core';
import { FUniver } from '@univerjs/core';

export function initUniver(containerId: string) {
  // 创建 Univer 实例
  const univer = new Univer({
    theme: defaultTheme,
    container: containerId,
  });
  
  // 注册核心插件
  univer.registerPlugin(new UniverPresetSheetsCore());
  
  // 创建空工作簿
  const workbook = univer.createUnit(UniverInstanceType.UNIVER_SHEET, {
    id: 'workbook1',
    name: '订单模板',
    sheetOrder: ['sheet1'],
    sheets: {
      sheet1: {
        id: 'sheet1',
        name: 'Sheet1',
        rowCount: 100,
        columnCount: 26,
      },
    },
  });
  
  return { univer, workbook };
}

export async function getUniverAPI(univer: Univer) {
  return await FUniver.newAPI(univer);
}
```

### 3.3 React 组件封装

```typescript
// src/components/UniverSheet/UniverSheet.tsx
import React, { useEffect, useRef, useState } from 'react';
import { Univer, FUniver } from '@univerjs/core';
import { initUniver } from '../../univer/init';
import './UniverSheet.less';

export interface CellRange {
  startRow: number;
  endRow: number;
  startColumn: number;
  endColumn: number;
}

export interface UniverSheetProps {
  onSelectionChange?: (ranges: CellRange[]) => void;
  onFileUpload?: (file: File) => void;
}

export const UniverSheet: React.FC<UniverSheetProps> = ({
  onSelectionChange,
  onFileUpload,
}) => {
  const containerRef = useRef<HTMLDivElement>(null);
  const univerRef = useRef<Univer | null>(null);
  const [selectedRanges, setSelectedRanges] = useState<CellRange[]>([]);
  const [cellValues, setCellValues] = useState<Record<string, any>>({});

  useEffect(() => {
    if (!containerRef.current) return;

    // 初始化 Univer
    const { univer } = initUniver(containerRef.current.id);
    univerRef.current = univer;

    // 设置选区监听
    setupSelectionListener(univer);

    return () => {
      univer?.dispose();
    };
  }, []);

  async function setupSelectionListener(univer: Univer) {
    const fUniver = await FUniver.newAPI(univer);
    const workbook = fUniver.getActiveWorkbook();
    const worksheet = workbook.getActiveSheet();

    // 监听选区变化
    worksheet.onSelectionChange((event) => {
      const ranges = event.ranges.map(r => ({
        startRow: r.startRow,
        endRow: r.endRow,
        startColumn: r.startColumn,
        endColumn: r.endColumn,
      }));
      
      setSelectedRanges(ranges);
      onSelectionChange?.(ranges);
      
      // 获取选中单元格的值
      fetchCellValues(worksheet, ranges);
    });
  }

  async function fetchCellValues(worksheet: any, ranges: CellRange[]) {
    const values: Record<string, any> = {};
    
    ranges.forEach(range => {
      for (let r = range.startRow; r <= range.endRow; r++) {
        for (let c = range.startColumn; c <= range.endColumn; c++) {
          const cellRef = getCellRef(r, c);
          const cell = worksheet.getCell(r, c);
          values[cellRef] = cell?.getValue() || '';
        }
      }
    });
    
    setCellValues(values);
  }

  // 暴露方法给父组件
  React.useImperativeHandle(ref, () => ({
    getSelectedRanges: () => selectedRanges,
    getCellValues: () => cellValues,
    loadExcelFile: async (file: File) => {
      // 实现 Excel 文件导入
    },
    highlightCells: (ranges: CellRange[], color: string) => {
      // 高亮指定单元格
    },
  }));

  return (
    <div className="univer-sheet-container">
      <div id={containerRef.current?.id || 'univer-container'} ref={containerRef} />
    </div>
  );
};
```

### 3.4 配置面板组件

```typescript
// src/components/ConfigPanel/ConfigPanel.tsx
import React, { useState } from 'react';
import { CellRange } from '../UniverSheet/UniverSheet';
import './ConfigPanel.less';

export interface FieldConfig {
  id: string;
  key: string;
  position: {
    cellRef?: string;
    areaRef?: string;
  };
  extractMode: 'SINGLE' | 'DOWN' | 'RIGHT' | 'BLOCK' | 'UNTIL_EMPTY';
  range?: {
    rows?: number;
    cols?: number;
    skipEmpty?: boolean;
  };
  type: 'STRING' | 'NUMBER' | 'DATE' | 'BOOLEAN';
  required: boolean;
}

export interface ConfigPanelProps {
  selectedRanges: CellRange[];
  cellValues: Record<string, any>;
  configs: FieldConfig[];
  onAddConfig: (config: FieldConfig) => void;
  onUpdateConfig: (id: string, updates: Partial<FieldConfig>) => void;
  onDeleteConfig: (id: string) => void;
  onSave: () => void;
}

export const ConfigPanel: React.FC<ConfigPanelProps> = ({
  selectedRanges,
  cellValues,
  configs,
  onAddConfig,
  onUpdateConfig,
  onDeleteConfig,
  onSave,
}) => {
  const [newKeyName, setNewKeyName] = useState('');

  // 从选区添加配置
  const handleAddFromSelection = () => {
    if (selectedRanges.length === 0) {
      alert('请先在表格中选择单元格');
      return;
    }

    const range = selectedRanges[0];
    const cellRef = getCellRef(range.startRow, range.startColumn);
    
    const newConfig: FieldConfig = {
      id: `field_${Date.now()}`,
      key: newKeyName || `field_${configs.length}`,
      position: { cellRef },
      extractMode: 'SINGLE',
      type: 'STRING',
      required: false,
    };

    onAddConfig(newConfig);
    setNewKeyName('');
  };

  return (
    <div className="config-panel">
      <div className="config-header">
        <h3>字段配置</h3>
      </div>

      {/* 快速添加 */}
      <div className="quick-add">
        <input
          type="text"
          placeholder="字段名 (如：orderNo)"
          value={newKeyName}
          onChange={(e) => setNewKeyName(e.target.value)}
        />
        <button onClick={handleAddFromSelection}>
          + 从选区添加
        </button>
      </div>

      {/* 配置列表 */}
      <div className="config-list">
        {configs.map((config) => (
          <div key={config.id} className="config-item">
            <div className="config-item-header">
              <span className="field-key">{config.key}</span>
              <button
                className="delete-btn"
                onClick={() => onDeleteConfig(config.id)}
              >
                ×
              </button>
            </div>

            <div className="config-item-body">
              <div className="form-row">
                <label>位置:</label>
                <input
                  type="text"
                  value={config.position.cellRef || ''}
                  readOnly
                  className="readonly-input"
                />
              </div>

              <div className="form-row">
                <label>提取模式:</label>
                <select
                  value={config.extractMode}
                  onChange={(e) =>
                    onUpdateConfig(config.id, {
                      extractMode: e.target.value as FieldConfig['extractMode'],
                    })
                  }
                >
                  <option value="SINGLE">单一单元格</option>
                  <option value="DOWN">向下列表</option>
                  <option value="RIGHT">向右列表</option>
                  <option value="BLOCK">区域块</option>
                  <option value="UNTIL_EMPTY">直到空值</option>
                </select>
              </div>

              {(config.extractMode === 'DOWN' ||
                config.extractMode === 'RIGHT' ||
                config.extractMode === 'BLOCK') && (
                <div className="form-row">
                  <label>范围:</label>
                  <div className="range-inputs">
                    <input
                      type="number"
                      placeholder="行数"
                      value={config.range?.rows || ''}
                      onChange={(e) =>
                        onUpdateConfig(config.id, {
                          range: { ...config.range, rows: parseInt(e.target.value) },
                        })
                      }
                    />
                    <input
                      type="number"
                      placeholder="列数"
                      value={config.range?.cols || ''}
                      onChange={(e) =>
                        onUpdateConfig(config.id, {
                          range: { ...config.range, cols: parseInt(e.target.value) },
                        })
                      }
                    />
                  </div>
                </div>
              )}

              <div className="form-row">
                <label>
                  <input
                    type="checkbox"
                    checked={config.required}
                    onChange={(e) =>
                      onUpdateConfig(config.id, { required: e.target.checked })
                    }
                  />
                  必填
                </label>
              </div>
            </div>
          </div>
        ))}
      </div>

      <div className="config-footer">
        <button className="save-btn" onClick={onSave}>
          保存配置
        </button>
      </div>
    </div>
  );
};
```

### 3.5 主页面整合

```typescript
// src/pages/ConfigEditor/ConfigEditor.tsx
import React, { useState, useRef } from 'react';
import { UniverSheet, CellRange } from '../../components/UniverSheet/UniverSheet';
import { ConfigPanel, FieldConfig } from '../../components/ConfigPanel/ConfigPanel';
import { uploadTemplate, saveConfig } from '../../api/excelConfig';
import './ConfigEditor.less';

export const ConfigEditor: React.FC = () => {
  const univerSheetRef = useRef<any>(null);
  const [selectedRanges, setSelectedRanges] = useState<CellRange[]>([]);
  const [cellValues, setCellValues] = useState<Record<string, any>>({});
  const [configs, setConfigs] = useState<FieldConfig[]>([]);
  const [loading, setLoading] = useState(false);

  const handleSelectionChange = (ranges: CellRange[]) => {
    setSelectedRanges(ranges);
    // cellValues 会在 UniverSheet 内部更新
  };

  const handleAddConfig = (config: FieldConfig) => {
    setConfigs([...configs, config]);
  };

  const handleUpdateConfig = (id: string, updates: Partial<FieldConfig>) => {
    setConfigs(configs.map(c => 
      c.id === id ? { ...c, ...updates } : c
    ));
  };

  const handleDeleteConfig = (id: string) => {
    setConfigs(configs.filter(c => c.id !== id));
  };

  const handleSaveConfig = async () => {
    setLoading(true);
    try {
      // 将配置转换为后端格式
      const configDTO = {
        templateName: '订单模板',
        cells: configs.reduce((acc, c) => ({ ...acc, [c.key]: c }), {}),
      };
      
      await saveConfig(configDTO);
      alert('配置保存成功!');
    } catch (error) {
      alert('保存失败');
    } finally {
      setLoading(false);
    }
  };

  const handleFileUpload = async (file: File) => {
    setLoading(true);
    try {
      const result = await uploadTemplate(file);
      // 可以用返回的结构信息初始化界面
      console.log('Excel 结构:', result);
    } catch (error) {
      console.error('上传失败', error);
    } finally {
      setLoading(false);
    }
  };

  return (
    <div className="config-editor">
      <div className="toolbar">
        <input
          type="file"
          accept=".xlsx,.xls"
          onChange={(e) => {
            const file = e.target.files?.[0];
            if (file) handleFileUpload(file);
          }}
        />
        <span className="status">
          {loading ? '处理中...' : '就绪'}
        </span>
      </div>

      <div className="main-content">
        <div className="sheet-area">
          <UniverSheet
            ref={univerSheetRef}
            onSelectionChange={handleSelectionChange}
          />
        </div>

        <div className="config-area">
          <ConfigPanel
            selectedRanges={selectedRanges}
            cellValues={cellValues}
            configs={configs}
            onAddConfig={handleAddConfig}
            onUpdateConfig={handleUpdateConfig}
            onDeleteConfig={handleDeleteConfig}
            onSave={handleSaveConfig}
          />
        </div>
      </div>
    </div>
  );
};
```

### 3.6 样式文件

```less
// src/pages/ConfigEditor/ConfigEditor.less
.config-editor {
  display: flex;
  flex-direction: column;
  height: 100vh;

  .toolbar {
    display: flex;
    justify-content: space-between;
    align-items: center;
    padding: 10px 20px;
    border-bottom: 1px solid #e0e0e0;
    background: #f5f5f5;

    .status {
      color: #666;
      font-size: 14px;
    }
  }

  .main-content {
    display: flex;
    flex: 1;
    overflow: hidden;

    .sheet-area {
      flex: 1;
      min-width: 600px;
      border-right: 1px solid #e0e0e0;
    }

    .config-area {
      width: 400px;
      overflow-y: auto;
      background: #fff;
    }
  }
}
```

---

## 四、API 接口定义

### 4.1 前端 API 封装

```typescript
// src/api/excelConfig.ts
import axios from 'axios';

const API_BASE = '/api/excel-config';

export interface ExcelStructure {
  sheets: Array<{
    name: string;
    rowCount: number;
    columnCount: number;
  }>;
  headers: string[];
  namedRanges: Array<{
    name: string;
    ref: string;
  }>;
}

export interface FieldConfigDTO {
  key: string;
  position: {
    cellRef?: string;
    areaRef?: string;
    headerName?: string;
  };
  extractMode: 'SINGLE' | 'DOWN' | 'RIGHT' | 'BLOCK' | 'UNTIL_EMPTY';
  range?: {
    rows?: number;
    cols?: number;
    skipEmpty?: boolean;
    untilCondition?: string;
  };
  parserType?: string;
  type?: 'STRING' | 'NUMBER' | 'DATE' | 'BOOLEAN';
  required?: boolean;
}

export interface ExcelTemplateConfigDTO {
  templateName: string;
  cells: Record<string, FieldConfigDTO>;
}

// 上传模板
export async function uploadTemplate(file: File): Promise<ExcelStructure> {
  const formData = new FormData();
  formData.append('file', file);
  
  const response = await axios.post(`${API_BASE}/upload`, formData, {
    headers: { 'Content-Type': 'multipart/form-data' },
  });
  
  return response.data;
}

// 保存配置
export async function saveConfig(config: ExcelTemplateConfigDTO) {
  return await axios.post(`${API_BASE}/config`, config);
}

// 获取配置
export async function getConfig(id: string) {
  const response = await axios.get(`${API_BASE}/config/${id}`);
  return response.data;
}

// 执行导入
export async function doImport(
  file: File,
  configId: string
): Promise<{ success: boolean; data: any[]; errors: any[] }> {
  const formData = new FormData();
  formData.append('file', file);
  formData.append('configId', configId);
  
  const response = await axios.post(`${API_BASE}/import`, formData);
  return response.data;
}

// 执行导出
export async function doExport(configId: string, data: any): Promise<Blob> {
  const response = await axios.post(`${API_BASE}/export`, {
    configId,
    data,
  }, {
    responseType: 'blob',
  });
  return response.data;
}
```

---

## 五、交互流程图

```
┌─────────────────────────────────────────────────────────────────┐
│                        用户操作流程                              │
└─────────────────────────────────────────────────────────────────┘

  ┌──────────────┐
  │   进入页面    │
  └──────┬───────┘
         │
         ▼
  ┌──────────────┐
  │  上传 Excel   │
  │   模板文件    │
  └──────┬───────┘
         │
         ▼
  ┌──────────────┐      ┌──────────────┐
  │  后端解析    │─────▶│  返回结构    │
  │  返回结构    │      │  信息        │
  └──────────────┘      └──────────────┘
                                │
                                ▼
  ┌──────────────┐      ┌──────────────┐
  │  Univer 渲染 │◀─────│  前端接收    │
  │  Excel 预览   │      │  并加载      │
  └──────────────┘      └──────────────┘
                                │
                                ▼
  ┌──────────────┐
  │  用户选择    │
  │  单元格区域   │
  └──────┬───────┘
         │
         ▼
  ┌──────────────┐      ┌──────────────┐
  │  监听选区    │─────▶│  显示选中    │
  │  变化事件    │      │  单元格信息  │
  └──────────────┘      └──────────────┘
                                │
                                ▼
  ┌──────────────┐
  │  输入字段名  │
  │  配置提取模式 │
  └──────┬───────┘
         │
         ▼
  ┌──────────────┐
  │  点击添加    │
  │  加入配置列表 │
  └──────┬───────┘
         │
         ▼
  ┌──────────────┐
  │  重复上述    │
  │  步骤添加更多 │
  └──────┬───────┘
         │
         ▼
  ┌──────────────┐      ┌──────────────┐
  │  点击保存    │─────▶│  POST /config│
  │  保存配置    │      │  保存配置    │
  └──────────────┘      └──────────────┘
```

---

## 六、技术栈详细

| 类别 | 技术 | 版本 | 说明 |
|------|------|------|------|
| **框架** | React | 18.x | 主流 UI 框架 |
| **语言** | TypeScript | 5.x | 类型安全 |
| **表格引擎** | Univer | latest | 阿里开源 |
| **UI 组件** | Ant Design | 5.x | 企业级组件 |
| **HTTP 客户端** | Axios | 1.x | API 请求 |
| **状态管理** | Zustand | 4.x | 轻量状态管理 |
| **构建工具** | Vite | 5.x | 快速构建 |
| **样式** | Less | 4.x | CSS 预处理器 |

---

## 七、关键实现细节

### 7.1 单元格引用转换工具

```typescript
// src/utils/cellRef.ts

/**
 * 行列号转 Excel 引用 (0, 0) => "A1"
 */
export function getCellRef(row: number, col: number): string {
  const colName = getColumnName(col);
  return `${colName}${row + 1}`;
}

/**
 * 列号转列名 0 => "A", 26 => "AA"
 */
export function getColumnName(col: number): string {
  let name = '';
  let n = col;
  
  while (n >= 0) {
    name = String.fromCharCode(65 + (n % 26)) + name;
    n = Math.floor(n / 26) - 1;
  }
  
  return name;
}

/**
 * Excel 引用转行列号 "A1" => { row: 0, col: 0 }
 */
export function parseCellRef(ref: string): { row: number; col: number } {
  const match = ref.match(/^([A-Z]+)(\d+)$/);
  if (!match) throw new Error(`Invalid cell ref: ${ref}`);
  
  const colName = match[1];
  const rowNum = parseInt(match[2]);
  
  let col = 0;
  for (let i = 0; i < colName.length; i++) {
    col = col * 26 + (colName.charCodeAt(i) - 64);
  }
  
  return {
    row: rowNum - 1,
    col: col - 1,
  };
}

/**
 * 解析区域引用 "A1:C10" => { startRow, endRow, startCol, endCol }
 */
export function parseAreaRef(areaRef: string): {
  startRow: number;
  endRow: number;
  startCol: number;
  endCol: number;
} {
  const [startRef, endRef] = areaRef.split(':');
  const start = parseCellRef(startRef);
  const end = parseCellRef(endRef);
  
  return {
    startRow: start.row,
    endRow: end.row,
    startCol: start.col,
    endCol: end.col,
  };
}
```

### 7.2 配置验证

```typescript
// src/utils/configValidator.ts
import { FieldConfig } from '../components/ConfigPanel/ConfigPanel';

export interface ValidationError {
  field: string;
  message: string;
}

export function validateConfigs(configs: FieldConfig[]): ValidationError[] {
  const errors: ValidationError[] = [];
  
  // 检查必填字段
  configs.forEach(config => {
    if (!config.key) {
      errors.push({ field: config.id, message: '字段名不能为空' });
    }
    
    if (!config.position.cellRef && !config.position.areaRef) {
      errors.push({ field: config.key, message: '位置未设置' });
    }
    
    // 检查范围配置
    if (['DOWN', 'BLOCK'].includes(config.extractMode)) {
      if (!config.range?.rows || config.range.rows <= 0) {
        errors.push({ 
          field: config.key, 
          message: '向下提取需要设置行数' 
        });
      }
    }
  });
  
  // 检查字段名重复
  const keys = configs.map(c => c.key);
  const duplicates = keys.filter((k, i) => keys.indexOf(k) !== i);
  if (duplicates.length > 0) {
    errors.push({
      field: 'duplicate',
      message: `字段名重复：${[...new Set(duplicates)].join(', ')}`,
    });
  }
  
  return errors;
}
```

---

## 八、部署与集成

### 8.1 Vite 配置

```typescript
// vite.config.ts
import { defineConfig } from 'vite';
import react from '@vitejs/plugin-react';

export default defineConfig({
  plugins: [react()],
  server: {
    port: 3000,
    proxy: {
      '/api': {
        target: 'http://localhost:8080',
        changeOrigin: true,
      },
    },
  },
});
```

### 8.2 Docker 部署

```dockerfile
# Dockerfile
FROM node:20-alpine as build

WORKDIR /app
COPY package*.json ./
RUN npm ci
COPY . .
RUN npm run build

FROM nginx:alpine
COPY --from=build /app/dist /usr/share/nginx/html
COPY nginx.conf /etc/nginx/nginx.conf

EXPOSE 80
CMD ["nginx", "-g", "daemon off;"]
```

---

## 九、开发注意事项

### 9.1 Univer 使用注意

1. **初始化时机**: 确保 DOM 元素已挂载后再初始化 Univer
2. **内存管理**: 组件卸载时调用 `univer.dispose()` 释放资源
3. **异步 API**: Facade API 需要通过 `FUniver.newAPI()` 异步获取
4. **选区监听**: 使用 `onSelectionChange` 监听，记得在 cleanup 中取消订阅

### 9.2 性能优化

1. **大文件处理**: 超过 10000 行的文件建议只预览前 1000 行
2. **配置保存**: 配置变更时做防抖处理，避免频繁请求
3. **状态同步**: 选区变化频繁时使用节流

### 9.3 用户体验

1. **加载状态**: 上传、解析、保存时显示 loading
2. **错误提示**: 配置验证失败时明确告知用户
3. **快捷键**: 支持 Ctrl+S 保存等常用快捷键
4. **撤销重做**: 配置修改支持撤销重做

---

## 十、参考资源

- **Univer 官方文档**: https://docs.univer.ai/
- **GitHub 仓库**: https://github.com/dream-num/univer
- **React Demo**: https://github.com/awesome-univer/sheets-react-demo
- **StackBlitz 在线体验**: https://stackblitz.com/github/awesome-univer/sheets-react-demo
- **NPM 包**: https://www.npmjs.com/package/@univerjs/ui
