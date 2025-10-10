# Excel 操作快速参考

## 常见场景与工具选择

### 数据输入场景
- **插入表格数据** → `insertData` (includeHeaders: true)
- **填充计算结果** → `insertData`
- **插入公式** → `insertFormula`

### 格式化场景
- **设置单个区域背景色** → `setCellColor`
- **批量设置多个区域背景色** → `setCellColorsBatch`（推荐）
- **设置字体样式** → `setCellFont`
- **批量格式化** → `formatRange`

### 数据可视化场景
- **趋势分析** → `createChart` (type: "Line")
- **数据对比** → `createChart` (type: "ColumnClustered")
- **占比分析** → `createChart` (type: "Pie")
- **相关性分析** → `createChart` (type: "XYScatter")

## 工具参数速查

### insertData
```json
{
  "address": "A1",
  "data": [["标题1", "标题2"], ["值1", "值2"]],
  "includeHeaders": true,
  "autoFitColumns": true
}
```

### setCellColorsBatch（推荐用于多区域）
```json
{
  "colorRanges": [
    { "address": "A1:A5", "backgroundColor": "red" },
    { "address": "B1:B5", "backgroundColor": "green" }
  ]
}
```

### setCellFont
```json
{
  "address": "A1:C1",
  "bold": true,
  "fontSize": 14,
  "fontColor": "blue",
  "horizontalAlignment": "Center"
}
```

### createChart
```json
{
  "type": "ColumnClustered",
  "dataRange": "A1:B10",
  "title": "销售趋势",
  "showLegend": true,
  "xAxisTitle": "月份",
  "yAxisTitle": "销售额"
}
```

### insertFormula
```json
{
  "address": "C1",
  "formula": "=SUM(A1:B1)",
  "calculateImmediately": true
}
```

## 最佳实践

1. **地址格式**
   - 使用大写字母：`"A1"` ✅，不是 `"a1"` ❌
   - 范围用冒号：`"A1:C10"` ✅

2. **颜色选择**
   - 优先使用标准颜色名：`"red"`, `"green"`, `"blue"`
   - 十六进制格式：`"#FF0000"`

3. **数据格式**
   - 二维数组：`[["行1列1", "行1列2"], ["行2列1", "行2列2"]]`
   - 包含标题时设置 `includeHeaders: true`

4. **性能优化**
   - 多个颜色设置 → 使用 `setCellColorsBatch` 而不是多次 `setCellColor`
   - 大量数据插入 → 一次性插入而不是逐行插入

5. **公式注意事项**
   - 必须以 `=` 开头
   - 使用正确的 Excel 函数名（英文）
   - 注意引用的单元格范围

## 错误处理

常见错误及解决方法：

- **"地址无效"** → 检查地址格式是否正确
- **"数据格式错误"** → 确保数据是二维数组
- **"公式错误"** → 验证公式语法，确保以 `=` 开头
- **"颜色无效"** → 使用标准颜色名或正确的十六进制格式
