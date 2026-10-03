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
func WithSheetIndex(i int) Option         // 按索引选表,对 eval 名损坏免疫,推荐
func WithStartCell(ref string) Option     // 如 "A1",ExportRangeToJson
func WithEndCell(ref string) Option       // 如 "B3",ExportRangeToJson
func WithXMLMap(name string) Option       // ExportSpreadsheetToXml,如 "InventoryMap"
// 导入
func WithBeginCell(row, col int) Option   // ImportCSV / ImportJsonData / ImportXMLData
func WithConvertNumeric(b bool) Option    // ImportCSV
func WithSeparator(s string) Option       // ImportCSV
```

可选参数缺失时的默认值:sheet 默认**第一张表(索引 0)**、分隔符默认 ","。sheet 选择(按名/按索引)与 query 共用同一 `cells.Sheet` 类型;默认索引 0 与 query 对齐,避免命名查找在 eval 模式偶尔失败。

## 4. 组合函数内部组合的底层操作

| 组合函数 | 底层操作序列 |
|---------|-------------|
| `Convert` | ReadSource → `NewWorkbook_Stream` → `opt.Apply` → `sink.Write` |
| `Merge` | `NewWorkbook` → `Worksheets.RemoveAt(0)` → 对每个 source:`NewWorkbook_Stream` + `Combine` → `SaveToStream` → `opt.Apply` → `sink.Write` |
| `Split` | 对每个 sheet:`NewWorkbook` + `CopyTheme` + defaultStyle `Copy` + `SetName` + `Copy_Worksheet` → `SaveToStream` → `opt.Apply` → `sink.Write(sheetname+"."+format, out)` |
| `ExportWorksheetToJson` | ReadSource → `NewWorkbook_Stream` → `Sheet.Resolve`(索引默认 0/按名) → json `Apply` → `sink.Write` |
| `ExportRangeToJson` | ReadSource → `NewWorkbook_Stream` → `Sheet.Resolve` → 由 `WithStartCell`/`WithEndCell` 建 `CellArea` → json `Apply` → `sink.Write` |
| `ExportSpreadsheetToXml` | ReadSource → `NewWorkbook_Stream` → `Sheet.Resolve` → `ExportXml(mapName)` → `sink.Write` |
| `ImportCSV` | `GetWorkbookWithDataSource` → `Sheet.Resolve` → `GetCells` → `ImportCSV_Stream` → `WorkbookToByteData` → `sink.Write` |
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

> **退出计划**:18 个 Deprecated 薄包装在 README「Deprecation schedule」明确将于 **v27.0.0 移除**;每个函数带 `// Deprecated:` 迁移指引。v26 系列内保持源码兼容,不再新增同形态变体。

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
6. **加载期工作表名不可靠(eval 模式)**:引擎在 `NewWorkbook_Stream` 加载时以约 2%/次 的概率把**任意**工作表的名称改写成垃圾字节——不限于默认首表,也不存在于字节流中(探针证实:同一份字节第一次加载名称正常,第二次加载首表名可变成 `"\x00@\x12\x00"`)。后果:按名查找(`WithSheet` 默认 `"Sheet1"`)可能返回 `ErrWorksheetNotFound`;`Split` 可能以坏名产出输出文件,或对含空字节等非法文件名的坏名直接报错。对策:`WithSheet` 文档与 README 注明"请用显式命名 sheet";凡依赖名称存活的断言(行为测试)基于显式命名 sheet 构造,并用 `retryStable` 重试——每次重试重新生成输入并重新加载,直至引擎给出干净名称;真实缺陷每次重试都失败,不会被掩盖。**同一机制也可能把随机单元格的值读成垃圾字节**(同一份字节可干净加载、也可返回指针垃圾),因此"加载后读值"的测试同样套 `retryStable`,`query.WithSheetIndex` 对自动化免疫。

## 9. 读取层(query)与 editor 增强(ISSUE-CELLSGO-294)

### 9.1 定位

`query` 包是 `transfer` 的读侧补充:`transfer` 把 sheet/range 序列化成格式字节(JSON/XML),`query` 把单元格读成 Go 原生类型供程序判断。二者并存——导出走 sink,读取返回 `[][]CellValue`(结构化数据无法进 sink)。不镜像底层对象模型,输出即数据。

### 9.2 query 设计

- 数据模型:`CellValue` + `CellKind`(`KindEmpty`/`KindText`/`KindInt`/`KindFloat`/`KindBool`/`KindDateTime`/`KindError`);访问器返回 `(value, ok)`,类型不匹配时 ok=false。
- 类型判别:`Cell.GetType()` 走 `CellValueType` 位枚举;数值格经 `Style.IsDateTime()` 识别日期格式(日期以序列号存储);公式格返回**计算后**值。
- 地址类型:`CellRef`(0-based)/`Area`,纯 Go 解析(`ParseCellRef`/`ParseArea`),定义在 `internal/aspose/cells`,query 以类型别名复用,供 `ReadMergedCells` 与内部助手共用。
- Option:`WithSheet`(按名)/`WithSheetIndex`(按索引,对 eval 名损坏免疫)/`WithTrimSpace`。**默认索引 0**(默认首表名恰是 eval 损坏最常命中的对象)。
- 内部收敛:`internal/aspose/cells` 新增 `GetCell`/`UsedRange`/`MergedAreas`(`Cells.GetMergedAreas` → `Area`)/`SheetNames`/`SheetByIndex`。
- 错误:新增 `ErrInvalidCellRef`/`ErrInvalidRange`;复用 `ErrDataSourceNil`/`ErrWorksheetNotFound`/`ErrInvalidSheetID`;统一 `%w` + `errors.Is`。

### 9.3 editor 增强

- `EditSpreadsheetToSink(source, sink, actions...)`:内部 `GetWorkbookWithDataSource → actions → WorkbookToByteData → sink.Write`,顺带去重 engine.go 原内联的 `GetFileFormat → FileFormatToSaveFormat → Save_SaveFormat` 段(该映射已收进 `internal/aspose/engine`,见 §11.4);`EditSpreadsheet` 降为薄包装(`*BytesSink`),返回字节语义零变化。
- `SetFormula(row, col, formula)`(WorksheetAction)+ `CalculateAll()`(WorkbookAction):绑定 `Cell.SetFormula_String` / `Workbook.CalculateFormula`。
- **日期写入修复**:`SetCellValue`/`SetValue` 写 `time.Time` 后补设日期数字格式(`yyyy-mm-dd` 或 `yyyy-mm-dd hh:mm:ss`)。根因:绑定的 `PutValue_Object(NewObject_Date)` 只写数值序列号、不设日期格式(探针:`string="45293"`、`type=2/IsNumeric`、`ObjectType_Number`),导致该格被 Excel 与 query 当作普通数字。补设格式后保存再加载返回 `type=4/IsDateTime`、`ObjectType_Date`、`GetStringValue="2024-01-02"`。

### 9.4 结构化表格读写(query.ReadRows / editor.WriteRows)

- **定位**:一张表 ↔ `[]struct` 的反射映射读写,是「组合封装、收敛 API」的最高价值一跳——一个函数顶掉上百个单元格级调用,且与 aspose 对象模型彻底解耦。
- **契约**(先定后实现):
  - `query.ReadRows[T](source, opts...) ([]T, error)`:首行作表头,`excel:"name"` tag(缺省用字段名)映射列;**大小写不敏感 + 去空白**匹配表头;`excel:"-"` 跳过字段;字段在表头缺失 → `ErrColumnNotFound`;表头多余的列忽略;空单元格留零值。支持 string/整型/无符号/float/bool/time.Time;类型不匹配或溢出 → `ErrInvalidValue`(KindInt 可入 float 字段,任意非空标量可入 string 字段取规范文本)。T 须为 struct,空表返回空非 nil 切片。
  - `editor.WriteRows[T](source, sink, rows, opts...) error`:写入 counterpart,Option 为 `WithWriteHeader(bool)`/`WithSheetIndex(i)`/`WithSheet(name)`,默认首表(索引 0)与 query/transfer 对齐;写路径复用 `SetCellValue`(含 time.Time 日期格式);写在块外单元格不动。
- **共享收敛**:反射列映射抽到新包 `internal/rows`(`Columns(reflect.Type) ([]Column, error)`),query 与 editor 两侧共用同一 tag 规则,杜绝漂移。
- **补齐**:`toObject` 补 `uint`/`uint8`/`uint32`;`CellKind` 加 `String()` 供错误消息;新增哨兵 `ErrColumnNotFound`。

### 9.5 命名区域 / 单元格批注 / 工作簿加密(P2-7)

延续「契约先行 → 实现 → 行为测试 → docs」的流程,三个能力都只暴露 Go 原生数据、不镜像 aspose 对象模型:

- **命名区域(workbook 级)**:`query.NamedRanges(source, opts...) ([]NamedRange, error)`,其中 `NamedRange{Name, RefersTo, Area}`——`Area` 由 `Name.GetRange()` 尽力解析(公式名无连续区域时置零);`query.ReadNamedRange(source, name, opts...) ([][]CellValue, error)` 经 `Range.GetWorksheet()` 解析到目标表与区域读网格,与当前 sheet 选择无关;名字缺失 → 新哨兵 `ErrNameNotFound`。写侧 `editor.DefineNamedRange(name, startRow, startCol, endRow, endCol)`(WorksheetAction,作用于被应用的表),refersTo 文本用 `CellRef.AbsoluteString()`(`$A$1`)拼出,查重更新保证幂等,反序坐标 → `ErrInvalidRange`。内部收敛到 `internal/aspose/cells`:`FindName`(**迭代** NameCollection 而非 `Get_String`——缺名返回悬垂句柄)、`SetOrAddNamedRange`。
- **单元格批注**:`editor.SetCellComment(row, col, text)` / `editor.ClearComments()`(WorksheetAction),读侧 `query.ReadCellComment(source, ref, opts...) (string, error)`(无批注 → 空串)。**关键坑**:绑定的 `Cell.GetComment()` 对无批注格返回非 nil 悬垂句柄,调用 `GetNote()` 直接崩溃(探针验证);因此读侧**绝不**用 `Cell.GetComment()`,而是迭代 `CommentCollection`(`Get_Int(i)` + `GetRow()/GetColumn()` 匹配),`CommentCollection.Get_Int_Int` 对缺批注格返回 `IsNull()==true` 可安全判别。
- **工作簿加密**:`editor.Encrypt(password)`(WorkbookAction)设置 `Settings.SetPassword` + `Workbook.SetEncryptionOptions(StrongCryptographicProvider, 128)`,保存后文件需密码才能打开。读回需加载期密码,工具包共享加载器目前未暴露(文档注明以底层 `LoadOptions` 读取)。测试在绑定层以 `LoadOptions.SetPassword` 验证:无密码加载失败、有密码加载后 `IsEncrypted()==true` 且值存活。
- **测试预算**:沿用 eval 100 次加载上限纪律——新增测试共用一次缓存的 build(`sync.Once`),并把 `TestQueryReadCellTypes` 从逐格 6 次 `ReadCell` 改为单次 `ReadWorksheet` 取网格(典型加载 6→1),为新测试腾出预算。

## 10. 引擎访问收口(`internal/aspose/engine`):串行锁 + 句柄解除武装

### 10.1 两个独立缺陷

原生引擎有两处会让进程直接崩溃,二者机制不同,修法也不同:

1. **引擎不是线程安全的**。两个 goroutine 同时进入会卡死或崩溃——探针:4 个 goroutine 并发加载,每次运行都挂起;加载与"finalizer 释放对象"重叠,保存路径内访问违例(15 次全套件运行中 4 次)。
2. **绑定的 finalizer 是破坏性的,不只是浪费**。`Delete_Worksheet` 真的析构 worksheet,`Delete_Cell` 真的析构 cell。工具链丢弃的包装器一旦不可达,下一次 GC 就会销毁引擎仍然持有指针的对象,其后的调用就走在已释放的对象上。探针:在"创建-丢弃句柄"循环中,`Delete_Worksheet`、`Delete_Cell`、`Delete_Charts` 三类各自在 20000 次迭代内让引擎出错;把同样的句柄解除武装,探针从 5 次中 1 次通过变成 5 次全部通过。

这两点合起来解释了此前的现象:锁只解决了第 1 点(全套件崩溃率约 27% → 6%),第 2 点**锁无法触及**——finalizer 走的是 Go 的 finalizer goroutine,不会去拿引擎锁。

### 10.2 规则

- **进入引擎即持锁**。所有可能到达引擎的公开入口点用 `engine.WithEngine(fn)` 或 `engine.LockEngine()` / `defer engine.UnlockEngine()` 包住全程。锁**按 goroutine 可重入**:工具链的 API 是组合式的(`manipulator.Merge` 先做引擎工作,再把结果交给 `saveoptions` 的 `Apply`,后者自己也会拿锁;`Split` 对每张表如此;`Convert` 本身就是一次 `Apply` 调用),普通互斥锁会在第一次嵌套调用时死锁。
- **引擎拥有的句柄在取得的那一刻解除武装**,写法是 `engine.Derive(receiver.GetX())`。`engine.OpenWorkbook` / `NewWorkbook` / `OpenWorkbookFile` / `OpenWorkbookWithOptions` 已内建解除武装。派生句柄(工作簿的 worksheet / cells / cell / style / chart / series / collection)一律经由 `Derive`,GC 因此永远不会通过 finalizer 进引擎。
- **工作簿由工具链拥有**:用 `engine.CloseWorkbook` 释放,**每本只调一次**。`DeleteWorkbook` 释放后会把句柄内部的指针置 nil,第二次调用会带着 nil 指针进入引擎的分配器;破坏是静默的,很晚才以一次无关调用中的访问违例显现。宁可泄漏一本工作簿,也不要猜第二次释放。
- **调用方自己拥有的对象不解除武装**:保存选项、加载选项、`NewObject_*` 值对象不属于任何工作簿,finalizer 是它们唯一的释放途径,解除武装等于泄漏它们。(例外:`internal/aspose/cells/xmlmap.go` 的 `XmlLoadOptions` 在交给引擎后即解除武装,因为那份所有权已经转移。)

### 10.3 验证

修完后同一测试二进制、同一负载:全套件 200 次运行 **0 次访问违例**(基线 12/200)。作为对照,`GOGC=off` 下基线同样 0 次崩溃——即 GC 是必要条件,这一点是定位第 2 个缺陷的关键实验。

残留的失败是另一回事:引擎在加载/读回时以小概率把**字符串读成垃圾字节**(同一机制此前已记录于 §8.6 的工作表名损坏),表现为断言失败而非崩溃,与句柄生命周期无关。

## 11. 绑定类型不外泄(`internal/aspose`):枚举、颜色与非枚举类型收口

### 11.1 问题

`saveoptions` 的 16 个格式包里有 75 个 `With*` 选项直接以 `asposecells.*` 枚举作参数(`csv.WithEncoding(asposecells.EncodingType)` 这类),`Config` 字段也存 `*asposecells.EncodingType`。这违反 §2「不镜像底层对象模型」和 §10.2「调用方自己拥有的对象不解除武装」所服务的那条线:调用方为了配置一次保存,必须先 `import` 绑定,于是引擎升级就成了他们的破坏性变更。另有一处反向错误:`saveoptions/image` 已经是字符串,但 `toImageType` 对无法识别的格式**静默返回 `ImageType_Unknown`**——拼错格式会得到一个"引擎顺手挑的格式"的文件,错误离原因很远,甚至根本不暴露。

### 11.2 规则

- **公开签名只出现工具链自己的类型**。枚举一律换成 `string` 名字,内部再解析成引擎枚举。
- **词汇表就是引擎的**。名字取常量后缀的 lower camel case(`EncodingType_UTF8` → `"utf8"`,`EmfRenderSetting_EmfPlusPrefer` → `"emfPlusPrefer"`),文档引用这套正规拼写。
- **匹配规则跟随 `editor`**:大小写与标点不敏感,`"UTF-8"` / `"utf8"` / `"utf-8"` 是同一个名字。工具链内只有一条匹配规则,不引入第二条。
- **不认识的 name 一律报错**(新哨兵 `errors.ErrInvalidEnumValue`),不回落默认值——这是 §2 原则 6 的直接推论。`ImageType_Unknown` 是"没有识别出格式"的哨兵而不是一种格式,因此**不进表**;它不进表,才是拼错格式能被报出来的前提。
- **解析发生在 `Apply` 里,不在 `Option` 闭包里**。`Option` 是 `func(*Config)`,没有返回错误的通道;`Apply` 是第一个能报错的点。解析先于打开工作簿,所以坏名字不会产出任何字节。
- **新包必须是叶子**。表不能放在 `saveoptions` 里:`saveoptions → internal/aspose/cells → formats → saveoptions` 已经是环,所以名字表落在只依赖 `errors` 与绑定的 `internal/aspose/enums`。

### 11.3 覆盖与验证

`internal/aspose/enums` 为 35 个枚举、125 个成员各存一张「规范化名字 → 引擎值」表,配一个同名解析函数。`tests/enums_test.go` 把这 125 个成员逐一对着引擎常量钉住,并检查大小写/标点不敏感规则;另有两条端到端测试:8 个非法名字必须在 save option 处失败且不产出字节,合法名字必须真的到达引擎并改变输出(CVS 文本、`%PDF-` 头、JPEG 的 `FF D8`)——断言落在渲染结果上,因为「选项被静默忽略」正是这次重构可能引入的失败形态。

写这条测试时抓到一个真实缺陷:表的键起初用的是正规 camelCase 拼写,而查表走的是规范化后的小写形式,于是所有多词名字(`"displayString"`、`"crossHideRight"`、`"pdfA1b"`)都解析不出来。键必须是规范化形式。

### 11.4 非枚举泄漏

有 8 个**不是枚举**的绑定类型同样出现在 `saveoptions` 的公开签名中,字符串表达不了它们,各自需要先有一个工具链自己的类型:

| 绑定类型 | 需要的东西 | 现状 |
| --- | --- | --- |
| `CellArea` | A1 区间字符串 | `json.WithExportArea` 已收口;其余 6 处停用 |
| `SheetSet` | 工作表名列表 | 停用,待补 |
| `PdfSecurityOptions` | 工具链自己的选项结构体 | 停用,待补 |
| `PdfBookmarkEntry` | 同上 | 停用,待补 |
| `RenderingWatermark` | 同上 | 停用,待补 |
| `SqlScriptColumnTypeMap` | 同上 | 停用,待补 |
| `CustomRenderSettings` | 回调类型,建议不纳入 | 停用 |
| `DrawObjectEventHandler` | 同上 | 停用 |

**「停用」= 整段注释掉,原文保留。** 这 28 个 `With*` 全库无调用方,收口形状又尚未确定,所以它们的声明被包进块注释,上面留一段说明:为什么不能这样暴露、恢复时该换成工具链自己的值。`Config` 字段与 `Apply` 里的接线原样保留,所以恢复只需删掉注释包裹、把参数类型换掉。(`editor` 的 6 个动作类型仍以 `*asposecells.*` 作参数——那是 §11.5 末尾说的同一类既有行为,收口要单独一轮。)

**`CellArea` 是第一个收口的**,形状就是 A1 区间字符串:

- `json.WithExportArea` 现在收 `"A1:C3"`(`"B2"` 视为单格),`Apply` 里解析成引擎的 `CellArea`;`transfer.ExportRangeToJson` 把解析好的四角渲染回 `"A1:C3"` 再交给它,签名里不再有绑定类型。
- 解析落在新叶子包 `internal/aspose/refs`。不能放 `internal/aspose/cells`:save option 导它就会成环(`saveoptions → internal/aspose/cells → formats → saveoptions`)。`cells` 现在把这份地址词汇表(re-export 给 `query` 的 `CellRef`/`Area`)重新导出,`ParseArea` 等仍是同一份实现。
- `refs.ParseAreaWithinGrid` 补上了原先缺的**网格边界校验**。`ParseCellRef` 是纯语法解析:任意行号、任意长度的列字母都收,而 14 个字母的列会溢出成负数。引擎两样都不查——越界的区间它照收,然后**匹配不到任何东西**——所以不校验的 `"XFE1"` 是个静默的空结果,不是错误。这正是 §2 原则 6「能报错就报错,不静默降级」。

**`formats.FileFormatToSaveFormat` 一并转为内部。** 它两侧都是引擎枚举,公开出去就是把 `formats` 的调用方绑到绑定上;实际调用方只有 `internal/aspose/cells.WorkbookToByteData` 一处,所以整张表移到 `internal/aspose/engine`(与 `OpenWorkbook`/`CloseWorkbook` 同层,都是保存路径上的引擎管道)。`formats` 因此**不再 import 绑定**,只剩扩展名注册表。

### 11.5 颜色(`internal/aspose/color`)

`Color` 不是枚举,收口方式却一样:公开签名收 `interface{}`,工具链内部转成引擎的 `Color`。`color.Resolve` 是这件事唯一的发生地,`editor`(`WithFontColor` / `WithBackgroundColor`,以及图表与条件格式的颜色入口)与 5 个 save option 的 `WithGridlineColor` 共用它——全库只有一条颜色规则,所以「怎么写一个颜色」在任何入口都是同一个答案。

- **表必须自己带,不能问引擎**。颜色名走工具链自带的 140 项构造器表(`editor/color_names.go` 迁到 `internal/aspose/color/color_names.go`),绝不交给引擎的 `Color_FromName`:它对不认识的名字**抛出未捕获的 C++ 异常**,穿过 cgo 直接杀进程,Go 侧接不住。
- **Go 颜色要「去预乘」**。`color.Color` 只提供 alpha 预乘后的通道,而文档颜色是直乘 alpha,`Resolve` 因此把 alpha 除回去:`color.NRGBA{R:0xFF, A:0x80}` 必须是「全强度红、半透明」,不是「半强度红、半透明」。除回去的结果可能超出通道范围——`color.RGBA{R:0xFF, A:0x80}` 在预乘语义下不合法,却正是「半透明红」的通俗写法——超出时夹到满值,而不是回绕成一个无关的颜色。
- **十六进制按文档的 `RRGGBBAA` 读**。引擎的 `Color_FromHex` 把 8 位串读成 `AARRGGBB`(Java/Android 顺序),而工具链文档承诺的是 `RRGGBBAA`(CSS 顺序)。这个不一致意味着 `"#33669980"` 在引擎里是「33/255 不透明的青」,不是「80/255 不透明的青」——`editor` 的 `WithFontColor` 一直带着这个缺陷。现在把 alpha 移到前面再交给引擎,文档与实现一致。
- **所有权是 `Resolve` 的第二个返回值**。引擎的 `Color` **没有 finalizer**(`Color_FromArgb` / `Color_FromHex` 只分配),所以工具链建的色只能靠显式 `DeleteColor` 释放,而调用方交进来的 `*asposecells.Color` 绝不能由工具链释放——释放它就是让调用方握着一个已死对象。save option 在 `Apply` 里 `defer DeleteColor`,即保存之后才释放,而不是 `SetGridlineColor` 返回后就释放:引擎是当场复制值还是留到保存时再读,绑定的文档没有交代,放进 `Apply` 的 defer 对两种情形都成立。
- **`editor` 的色仍不释放**。动作闭包把 `Color` 直接交给引擎 setter,而 `defer` 会在 setter 之后立刻执行——那时往往离保存还很远,而引擎是否还持有那个指针无从确认,所以这里只能维持既有行为(即每次设置颜色泄漏一个引擎色对象)。要收口需要先有「引擎确实复制了值」的证据。

`tests/color_test.go` 把这张词汇表钉住:每种写法的 ARGB 精确值(包括预乘回来的半透明与越界夹取)、`Resolve` 的所有权标志、非法输入必须是 `ErrInvalidColor`,以及 5 个 gridline 颜色入口各自能存出文件、连着存两次不炸——后两条断言故意弱(页面网格线的颜色读不回来),但它们覆盖的是一次段错误就会带走整个测试二进制的那段代码:双重释放或释放后用。
