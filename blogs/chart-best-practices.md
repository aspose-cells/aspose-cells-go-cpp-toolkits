# 图表最佳实践指南

本指南提供使用 toolkit 图表 API 的最佳实践，帮助您创建专业、高效的图表。

## 1. 选择合适的图表类型

### 数据比较
- **柱状图** (`ChartTypeColumn`): 比较不同类别的数值
- **条形图** (`ChartTypeBar`): 类别名称较长时水平显示
- **3D 柱状图** (`ChartTypeColumn3D`): 视觉效果更强，适合演示

### 趋势分析
- **折线图** (`ChartTypeLine`): 显示随时间变化的趋势
- **带标记折线图** (`ChartTypeLineWithMarkers`): 突出显示数据点
- **面积图** (`ChartTypeArea`): 强调数量随时间的变化

### 占比分析
- **饼图** (`ChartTypePie`): 显示各部分占总体的比例
- **环形图** (`ChartTypeDoughnut`): 类似饼图，可以显示多个数据系列
- **堆叠图** (`ChartTypeColumnStacked`): 显示各部分如何组成整体

### 相关性分析
- **散点图** (`ChartTypeScatter`): 显示两个变量之间的关系
- **气泡图** (`ChartTypeBubble`): 散点图的扩展，增加第三个维度

### 金融数据
- **K 线图** (`ChartTypeStock`): 显示股票的开盘/最高/最低/收盘价

## 2. 使用预定义模板

对于常见场景，使用预定义模板可以节省时间并确保一致性：

```go
// 专业报表
editor.AddChart(editor.ChartTypeColumn, "A1:C5", true, 5, 0, 20, 7,
    editor.WithChartPreset(editor.ProfessionalColumn),
    editor.WithChartTitle("季度销售报告"),
)

// 仪表板
editor.AddChart(editor.ChartTypeLine, "A1:C5", true, 5, 0, 20, 7,
    editor.WithChartPreset(editor.DashboardLine),
    editor.WithChartTitle("趋势分析"),
)

// 金融报表
editor.AddChart(editor.ChartTypeStock, "A1:E20", true, 5, 0, 25, 10,
    editor.WithChartPreset(editor.FinancialCandlestick),
)
```

### 行业特定模板
- `SalesPerformance`: 销售业绩报表
- `FinancialSummary`: 财务报表
- `MarketShare`: 市场份额
- `ScientificData`: 科学数据
- `SurveyResults`: 调查结果
- `BudgetVariance`: 预算差异分析
- `KPIDashboard`: KPI 仪表板

## 3. 数据范围选择

### 正确的数据布局
```
Category    Series1    Series2    Series3
Q1          100        150        200
Q2          120        180        220
Q3          140        200        240
Q4          160        220        260
```

### 使用 byColumn 参数
- `byColumn=true`: 每列是一个系列（最常用）
- `byColumn=false`: 每行是一个系列

### 显式设置分类数据
当自动检测不符合预期时，使用 `WithChartCategoryData`：

```go
editor.AddChart(editor.ChartTypeColumn, "B1:D5", true, 5, 0, 20, 7,
    editor.WithChartCategoryData("A2:A5"), // 明确指定分类列
)
```

## 4. 图表样式

### 使用 ChartStylePreset
`ChartStylePreset` 允许您打包多个设置以便重用：

```go
preset := editor.ChartStylePreset{
    ChartType:      editor.ChartTypeColumn,
    Style:          7,              // 内置样式 1-48
    Title:          "季度报告",
    ShowLegend:     true,
    LegendPosition: editor.ChartLegendBottom,
    DataRange:      "A1:C5",
    ByColumn:       true,
    Bounds:         editor.ChartBounds{
        TopRow: 5, LeftColumn: 0,
        BottomRow: 20, RightColumn: 7,
    },
}

// 在多个图表中重用
editor.AddChart(editor.ChartTypeColumn, "A1:C5", true, 5, 0, 20, 7,
    editor.WithChartPreset(preset),
)
```

### 部分预设
零值字段会被跳过，允许部分预设：

```go
// 只设置样式和标题，其他保持不变
partialPreset := editor.ChartStylePreset{
    Style: 12,
    Title: "更新标题",
}
```

## 5. 数据验证

在创建图表前验证数据：

```go
// 验证数据范围
if err := editor.ValidateChartDataRange("A1:C5", true); err != nil {
    log.Printf("数据范围无效: %v", err)
    // 使用建议的数据范围
    suggestion := editor.SuggestDataRange(10, 3, true)
    log.Printf("建议使用: %s", suggestion)
}

// 验证图表边界
if err := editor.ValidateChartBounds(5, 0, 20, 7); err != nil {
    log.Printf("图表边界无效: %v", err)
}

// 验证样式编号
if err := editor.ValidateChartStyle(7); err != nil {
    log.Printf("样式编号无效: %v", err)
}
```

## 6. 常见错误避免

### ❌ 错误：单单元格范围
```go
// 错误：单单元格无法生成图表
editor.AddChart(editor.ChartTypeColumn, "A1:A1", true, ...)
```

### ✅ 正确：多单元格范围
```go
// 正确：至少包含多个单元格
editor.AddChart(editor.ChartTypeColumn, "A1:B5", true, ...)
```

### ❌ 错误：反转边界
```go
// 错误：bottomRow < topRow
editor.AddChart(editor.ChartTypeColumn, "A1:B5", true, 20, 0, 5, 7, ...)
```

### ✅ 正确：正确顺序
```go
// 正确：topRow < bottomRow
editor.AddChart(editor.ChartTypeColumn, "A1:B5", true, 5, 0, 20, 7, ...)
```

### ❌ 错误：越界范围
```go
// 错误：超出工作表网格
editor.AddChart(editor.ChartTypeColumn, "XFE1:XFE5", true, ...)
```

### ✅ 正确：网格内范围
```go
// 正确：在工作表网格内
editor.AddChart(editor.ChartTypeColumn, "A1:E10", true, ...)
```

## 7. 性能优化

### 批量创建图表
在单次 `EditSpreadsheet` 调用中创建多个图表：

```go
// ✅ 推荐：一次调用创建多个图表
editor.EditSpreadsheet(source,
    editor.InWorksheet("Sheet1",
        editor.AddChart(editor.ChartTypeColumn, "A1:B5", true, 5, 0, 20, 7),
        editor.AddChart(editor.ChartTypePie, "A1:B5", true, 5, 8, 20, 15),
        editor.AddChart(editor.ChartTypeLine, "A1:B5", true, 22, 0, 37, 7),
    ),
)

// ❌ 避免：多次调用
for i := 0; i < 3; i++ {
    editor.EditSpreadsheet(source, ...) // 每次都会重新加载工作簿
}
```

### 使用预设重用配置
预定义常用配置，避免重复代码：

```go
// 定义一次
var reportPreset = editor.ChartStylePreset{
    Style: 7,
    ShowLegend: true,
    LegendPosition: editor.ChartLegendBottom,
}

// 多次使用
editor.AddChart(editor.ChartTypeColumn, "A1:B5", true, 5, 0, 20, 7,
    editor.WithChartPreset(reportPreset),
)
editor.AddChart(editor.ChartTypeBar, "A1:B5", true, 5, 8, 20, 15,
    editor.WithChartPreset(reportPreset),
)
```

## 8. 查询和验证图表

使用 `query` 包读取和验证图表：

```go
// 获取图表信息
info, err := query.ChartInfo(source, 0)
if err != nil {
    log.Fatal(err)
}
fmt.Printf("图表类型: %s\n", info.Type)
fmt.Printf("系列数量: %d\n", info.SeriesCount)

// 获取系列数据
series, err := query.ChartSeriesData(source, 0)
if err != nil {
    log.Fatal(err)
}
for i, s := range series {
    fmt.Printf("系列 %d: %s\n", i, s.Values)
}

// 统计图表数量
count, err := query.ChartCount(source)
if err != nil {
    log.Fatal(err)
}
fmt.Printf("工作表共有 %d 个图表\n", count)
```

## 9. 评估模式注意事项

在评估模式下：
- 工作表名称可能在加载时被破坏（约 2% 的概率）
- 使用 `WithSheetIndex` 而不是 `WithSheet` 可以避免这个问题
- 字符串读回可能偶尔出现乱码，使用 `retryStable` 模式重试

```go
// ✅ 推荐：使用索引
query.ChartInfo(source, 0, query.WithSheetIndex(0))

// ❌ 避免：使用名称（评估模式下可能不可靠）
query.ChartInfo(source, 0, query.WithSheet("Sheet1"))
```

## 10. 完整示例

```go
package main

import (
    "log"
    "github.com/aspose-cells/aspose-cells-go-cpp-toolkits/v26/datasource"
    "github.com/aspose-cells/aspose-cells-go-cpp-toolkits/v26/editor"
    "github.com/aspose-cells/aspose-cells-go-cpp-toolkits/v26/query"
)

func main() {
    // 创建工作簿
    seed, err := datasource.NewEmptyWorkbook()
    if err != nil {
        log.Fatal(err)
    }

    // 准备数据
    withData, err := editor.EditSpreadsheet(
        seed,
        editor.WithAddWorksheet("Sales"),
        editor.InWorksheet("Sales",
            editor.SetCellValue(0, 0, "Region"),
            editor.SetCellValue(0, 1, "Q1"),
            editor.SetCellValue(0, 2, "Q2"),
            editor.SetCellValue(1, 0, "North"),
            editor.SetCellValue(1, 1, 120),
            editor.SetCellValue(1, 2, 150),
            editor.SetCellValue(2, 0, "South"),
            editor.SetCellValue(2, 1, 90),
            editor.SetCellValue(2, 2, 130),
            editor.SetCellValue(3, 0, "East"),
            editor.SetCellValue(3, 1, 160),
            editor.SetCellValue(3, 2, 140),
            editor.SetCellValue(4, 0, "West"),
            editor.SetCellValue(4, 1, 70),
            editor.SetCellValue(4, 2, 110),
        ),
    )
    if err != nil {
        log.Fatal(err)
    }

    // 使用模板创建图表
    withChart, err := editor.EditSpreadsheet(
        datasource.BytesSource(withData),
        editor.InWorksheet("Sales",
            editor.AddChart(editor.ChartTypeColumn, "A1:C5", true, 5, 0, 20, 7,
                editor.WithChartPreset(editor.SalesPerformance),
                editor.WithChartTitle("季度销售报告"),
            ),
        ),
    )
    if err != nil {
        log.Fatal(err)
    }

    // 查询图表信息
    info, err := query.ChartInfo(datasource.BytesSource(withChart), 0)
    if err != nil {
        log.Fatal(err)
    }
    log.Printf("创建图表: %s, 系列数: %d", info.Type, info.SeriesCount)

    // 保存文件
    err = datasource.SaveToFile(withChart, "report.xlsx")
    if err != nil {
        log.Fatal(err)
    }
}
```

## 总结

遵循这些最佳实践，您可以：
- 选择合适的图表类型
- 使用预定义模板提高效率
- 避免常见错误
- 优化性能
- 创建专业的图表

如有疑问，请参考 `docs/editor.md` 和 `docs/query.md` 获取完整的 API 文档。
