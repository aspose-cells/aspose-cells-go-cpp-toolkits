# 设计文档:DataSource / DataSink 重构与组合函数收敛

> 状态:待评审 · 适用范围:允许改动 DataSource / DataSink 相关接口及其消费方(converter / manipulator / transfer 的公开签名)

## 1. 背景与目标

Aspose.Cells for Go via C++ Toolkits 的定位是:在 aspose-cells-go-cpp 绑定之上提供**组合封装**——一个公开函数组合多个底层操作,而不是重新暴露底层对象图。输入输出用 `DataSource`/`DataSink` 抽象、参数用函数式 Option 传入,目标是**公开 API 数量最小化 + Go 化**。

当前代码背离了这一初衷:

| # | 问题 | 现状 | 影响 |
|---|------|------|------|
| 1 | `DataSink` 闲置 | 输出靠 `*ToWriter`/`*ToFile`/`*ToZipWriter`/`*ToFolder` 变体增殖,`DataSink` 无调用方 | 接口数量膨胀,抽象名存实亡 |
| 2 | `DataSource.ByteData()` 吞错 | 读失败返回 nil;transfer 导入路径用它取数据 | 文件不可读时错误信息劣化 |
| 3 | 变体爆炸 | converter 3 + manipulator 6 + transfer 12 = **21 个**公开入口,同一能力重复三种形态 | 使用与维护成本高 |
| 4 | 位置参数群 | transfer 的 `(worksheet, beginRow, beginColumn, convertNumericData, splitter)` | 与 Option Go 化方向冲突 |

## 2. 设计原则

1. **每个能力一个组合函数**:输入 `DataSource`、参数 `Option`、输出 `DataSink`。
2. **输出方式不增殖函数,只增殖 Sink 适配器**:新增一种输出 = 新增一个几行的 Sink,不新增函数。
3. **Option 化**:所有可选/位置参数收进 `With*`;新增参数永不改签名。
4. **标准库优先**:`io.Reader`/`io.Writer` 能表达的,不发明自定义抽象。
5. **`editor` 是补充层**:细粒度 DSL,独立于组合函数,不受本次收敛影响。
6. **错误统一**:哨兵 + `%w` 保留;能报错就报错,不静默降级。

## 3. 目标接口

### 3.1 datasource — 输入

```go
// DataSource 抽象一个可读的输入。实现方只需要提供一个 Open。
type DataSource interface {
	Open() (io.ReadCloser, error)
}
```

- 删除 `ByteData()`(吞错隐患),读取统一走 `internal/aspose/cells.ReadSource`(Open → io.ReadAll → Close,错误完整传播)。
- 输入适配器(构造 DataSource 用,保留):
  ```go
  type FilePathSource string
  type BytesSource []byte
  func NewReaderSource(r io.Reader) *ReaderSource   // 输入放宽为 io.Reader(原为 io.ReadCloser)
  ```
- `FileStore` 删除(读 + 写两种能力分别由 `FilePathSource` / `FilePathSink` 承担)。

### 3.2 datasource — 输出

```go
// DataSink 抽象一个可写的输出。name 用于多输出场景(Split 每个 sheet 一个文件/条目),
// 单输出 sink 忽略 name。
type DataSink interface {
	Write(name string, data []byte) error
}
```

| Sink | 定义 | Write 语义 |
|------|------|-----------|
| `FilePathSink` | `type FilePathSink string` | `os.WriteFile(string(p), data, 0644)`,忽略 name |
| `WriterSink` | `type WriterSink struct{ w io.Writer }` + `NewWriterSink(w)` | `w.Write(data)`,忽略 name |
| `BytesSink` | `type BytesSink struct{ buf bytes.Buffer }` + `Bytes() []byte` | 累积到 buf,忽略 name |
| `FolderSink` | `type FolderSink string` | `os.WriteFile(filepath.Join(string(f), name), data, 0644)` |
| `ZipSink` | `type ZipSink struct{ zw *zip.Writer }` + `NewZipSink(zw)` | `zw.Create(name)` 后写入 |

> 跨包使用 `WriterSink` / `ZipSink` 时经构造函数构造(字段未导出)。

新增一种输出目标 = 新增一个几行的 Sink,**不新增函数**。

### 3.3 converter / manipulator / transfer — 组合函数

```go
// converter —— 1 个函数(原 3 个)
func Convert(source datasource.DataSource, opt saveoptions.SaveOption, sink datasource.DataSink) error

// manipulator —— 2 个函数(原 6 个)
func Merge(sources []datasource.DataSource, opt saveoptions.SaveOption, sink datasource.DataSink) error
func Split(source datasource.DataSource, opt saveoptions.SaveOption, sink datasource.DataSink) error

// transfer —— 6 个函数(原 12 个)
func ExportWorksheetToJson(source datasource.DataSource, sink datasource.DataSink, opts ...Option) error
func ExportRangeToJson(source datasource.DataSource, sink datasource.DataSink, opts ...Option) error
func ExportSpreadsheetToXml(source datasource.DataSource, sink datasource.DataSink, opts ...Option) error
func ImportCSV(source datasource.DataSource, data datasource.DataSource, sink datasource.DataSink, opts ...Option) error
func ImportJsonData(source datasource.DataSource, data datasource.DataSource, sink datasource.DataSink, opts ...Option) error
func ImportXMLData(source datasource.DataSource, data datasource.DataSource, sink datasource.DataSink, opts ...Option) error
```

原变体全部由「组合函数 + Sink 适配器」一行覆盖,例如:

```go
// 原 ConvertSpreadsheetToFile(a, b)          → converter.Convert(src, opt, datasource.FilePathSink(b))
// 原 ConvertToWriter(w)                      → converter.Convert(src, opt, datasource.NewWriterSink(w))
// 原 MergeSpreadsheetsToFile(files, out)     → 把 files 包成 []DataSource 后 Merge(…, FilePathSink(out))
// 原 SplitSpreadsheetToFolder(dir)           → Split(src, opt, datasource.FolderSink(dir))
// 原 SplitSpreadsheetToZipWriter(zw)         → Split(src, opt, datasource.NewZipSink(zw))
```

### 3.4 transfer — Option

```go
// 导出
func WithSheet(name string) Option        // ExportWorksheetToJson / ExportRangeToJson
func WithStartCell(ref string) Option     // 如 "A1",ExportRangeToJson
func WithEndCell(ref string) Option       // 如 "B3",ExportRangeToJson
func WithXMLMap(name string) Option       // ExportSpreadsheetToXml,如 "InventoryMap"
// 导入
func WithBeginCell(row, col int) Option   // ImportCSV / ImportJsonData / ImportXMLData
func WithConvertNumeric(b bool) Option    // ImportCSV
func WithSeparator(s string) Option       // ImportCSV
```

可选参数缺失时的默认值沿用现状(如 sheet 默认 "Sheet1"、分隔符默认 ",")。

## 4. 组合函数内部组合的底层操作

| 组合函数 | 底层操作序列 |
|---------|-------------|
| `Convert` | ReadSource → `NewWorkbook_Stream` → `opt.Apply` → `sink.Write` |
| `Merge` | `NewWorkbook` → `Worksheets.RemoveAt(0)` → 对每个 source:`NewWorkbook_Stream` + `Combine` → `SaveToStream` → `opt.Apply` → `sink.Write` |
| `Split` | 对每个 sheet:`NewWorkbook` + `CopyTheme` + defaultStyle `Copy` + `SetName` + `Copy_Worksheet` → `SaveToStream` → `opt.Apply` → `sink.Write(sheetname+"."+format, out)` |
| `ExportWorksheetToJson` | ReadSource → `NewWorkbook_Stream` → `WorksheetByName` → `GetCellAreaWithWorksheet` → json `Apply` → `sink.Write` |
| `ExportRangeToJson` | ReadSource → `NewWorkbook_Stream` → `WorksheetByName` → 由 `WithStartCell`/`WithEndCell` 建 `CellArea` → json `Apply` → `sink.Write` |
| `ExportSpreadsheetToXml` | ReadSource → `NewWorkbook_Stream` → `ExportXml(mapName)` → `sink.Write` |
| `ImportCSV` | `GetWorkbookWithDataSource` → `GetCellsWithWorksheet` → `ImportCSV_Stream` → `WorkbookToByteData` → `sink.Write` |
| `ImportJsonData` / `ImportXMLData` | 同 ImportCSV,换 `JsonUtility_ImportData` / `ImportXml_Stream` |

## 5. 兼容与迁移

**已移除**(Go 无函数重载,旧函数与新签名同名,无法共存):
- `transfer.ExportWorksheetToJson(source, worksheet string) ([]byte, error)`
- `transfer.ExportRangeToJson(source, worksheet, startCell, endCell string) ([]byte, error)`
- `transfer.ExportSpreadsheetToXml(source, mapName string) ([]byte, error)`

这三个字节版导出被 sink 化版本(`ExportWorksheetToJson(source, sink, opts...)` 等)整体取代,字节结果改用 `*datasource.BytesSink` 获取。

**保留为 Deprecated 薄包装**(签名不冲突,维持源码兼容):
- `datasource.DataSource` 删除 `ByteData()`(吞错隐患);`ReaderSource` 构造参数 `io.ReadCloser` → `io.Reader`;删除 `FileStore`。
- `datasource.DataSink.Write()` 语义从「返回 WriteCloser」改为「`Write(name, data) error`」。
- `converter.ConvertSpreadsheet` / `ConvertToWriter` / `ConvertSpreadsheetToFile` → 各自薄包装调用 `converter.Convert(source, opt, sink)`,标记 `Deprecated`。
- `manipulator.MergeSpreadsheets*` / `SplitSpreadsheet*` → 各自薄包装调用 `manipulator.Merge` / `manipulator.Split`,标记 `Deprecated`。
- `transfer.ImportCSVDataIntoSpreadsheet` / `ImportCSVFile` / `ImportJsonDataIntoSpreadsheet` / `ImportJsonFile` / `ImportXMLDataIntoSpreadsheet` / `ImportXMLFile` 及三个 `*ToFile` 导出 → 薄包装调用 sink 化函数,标记 `Deprecated`。

**不变**:
- `editor.EditSpreadsheet(source DataSource, actions ...WorkbookAction) ([]byte, error)` 签名不变(补充层,输出仍为字节;其 DataSource 用法只依赖 `Open`)。
- `saveoptions` 19 子包、`formats` 注册表、`errors` 哨兵、`internal/aspose/cells` 全部不变。
- `examples/`、`tests/`、README、docs 同步更新。

## 6. 测试策略

- `datasource_test`:新接口单元测试——各 Sink 的 Write 语义、`ReaderSource`(io.Reader)缓冲、`ReadSource` 错误传播、`FilePathSink` 自动建父目录。
- `converter/manipulator/transfer` 行为测试(新增):哨兵错误(`ErrSaveOptionNil`/`ErrUnsupportedFormat`/`ErrInvalidOutputPath`/`ErrWorksheetNotFound`/`ErrNoSources`/`ErrDataSourceNil`/`ErrDataSinkNil`)、Sink 路由(同一能力写文件/写流/取字节结果一致)、`Merge` 空源报错且合并结果含双源、`Split` 每个命名 sheet 产出独立文件/zip 条目、空 sheet 导出合法 `[]`、CSV 导入落格(导出 JSON 回读)。
- `examples/`:全部改用新签名,仍支持一次性(`./examples/run.sh`)与单个执行。
- CI 不变。

## 7. 实施计划(分阶段)

| 阶段 | 内容 | 验证 |
|------|------|------|
| A | `datasource` 重写 + `internal/aspose/cells`/editor/transfer 内部适配(去掉 ByteData) | build/vet/gofmt + 现有 tests 通过 |
| B | `converter` 收敛为 `Convert` | 新增行为测试 |
| C | `manipulator` 收敛为 `Merge`/`Split` | 新增行为测试 |
| D | `transfer` 收敛为 6 个函数 + Option | 新增行为测试 |
| E | `examples/` 4 个命令改新签名 | 一次性 + 单个运行 |
| F | README / docs / CHANGELOG 同步;旧入口在文档中以 Deprecated 章节保留 | grep 确认代码与示例无旧 API 调用(文档 Deprecated 章节除外) |

## 8. 风险与取舍

1. **Split → 字节 需显式组装 zip**:原 `SplitSpreadsheet` 一步返回 zip `[]byte`;新形态需 `zip.NewWriter(&buf)` + `ZipSink` + `Close`。这是"一个能力一个函数"的可接受代价,文档给出配方。
2. **Split → `BytesSink` 无意义**:多次 `Write` 只是拼接,不构成 zip;文档明确该组合用 `ZipSink`。
3. **单输出 Sink 的 `name` 约定**:统一传 `""`,由 Sink 忽略;避免调用方误传。
4. **保持字节式处理**:引擎 `Apply`/`SaveToStream` 输出整块 `[]byte`,本次不改内存模型。
5. **eval 模式还会追加 "Evaluation Warning" sheet**:评估版保存的工作簿可能额外多出一个名为 "Evaluation Warning" 的 sheet,导致"输出文件数 == sheet 数"类断言不稳定。行为测试改为主张「每个显式命名 sheet 产出独立输出」,不依赖精确总数;正式授权版无此噪声。
6. **加载期工作表名不可靠(eval 模式)**:引擎在 `NewWorkbook_Stream` 加载时以约 2%/次 的概率把**任意**工作表的名称改写成垃圾字节——不限于默认首表,也不存在于字节流中(探针证实:同一份字节第一次加载名称正常,第二次加载首表名可变成 `"\x00@\x12\x00"`)。后果:按名查找(`WithSheet` 默认 `"Sheet1"`)可能返回 `ErrWorksheetNotFound`;`Split` 可能以坏名产出输出文件,或对含空字节等非法文件名的坏名直接报错。对策:`WithSheet` 文档与 README 注明"请用显式命名 sheet";凡依赖名称存活的断言(行为测试)基于显式命名 sheet 构造,并用 `retryStable` 重试——每次重试重新生成输入并重新加载,直至引擎给出干净名称;真实缺陷每次重试都失败,不会被掩盖。
