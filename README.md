# ExcelTools v2.1

> **一句话**：Excel-DNA 插件，11 个通用数据处理工具，拷贝即用。
>
> **One-liner**: An Excel-DNA add-in with 11 general-purpose data-processing tools, copy-and-run.

---

## 构建 / Build

```bash
cd source/ExcelTools
dotnet build -c Release
```

构建产物在 `bin/Release/`：
Build artifacts are placed in `bin/Release/`:

- `ExcelTools.xll` — 发给客户就这一个文件 / the only file you ship to a client
- `ExcelTools.dna` — 配置文件（Ribbon 定义）/ config file (Ribbon definition)
- `ExcelTools.dll` — 核心程序集 / core assembly

---

## 部署到客户机 / Deploy to a client machine

1. 把 `ExcelTools.xll` + `ExcelTools.dna` + `ExcelTools.dll` 拷到一个目录
   Copy `ExcelTools.xll` + `ExcelTools.dna` + `ExcelTools.dll` into one folder.
2. 双击 `.xll`（或在 Excel 里 文件→选项→加载项→Excel 加载项→浏览→选 .xll）
   Double-click the `.xll` (or in Excel: File → Options → Add-ins → Excel Add-ins → Browse → pick the .xll).
3. Ribbon 出现「Excel 工具」选项卡
   The "Excel Tools" tab appears on the Ribbon.
4. 不需要管理员权限、不需要 COM 注册、不需要证书
   No admin rights, no COM registration, no code-signing certificate required.

---

## 11 个工具 / The 11 tools

### 数据处理 / Data processing
| 工具 Tool | 功能 Function |
| --- | --- |
| 拆分工作表 / Split sheet | 按某列值将总表拆成多个 Sheet / Split a sheet into multiple sheets by the values in a column |
| 条件筛选 / Filter | 多条件筛选到新表（逐行匹配 / SQL 两种模式）/ Filter rows into a new sheet (row-by-row or SQL mode) |
| 按列排序 / Sort | 按指定列排序，自动备份原始数据到 Original 表 / Sort by a chosen column, auto-backup originals to an "Original" sheet |
| 删除匹配行 / Delete matched rows | 删除与另一表匹配的行，全部参数化 / Delete rows that match another sheet; fully parameterized |

### 比对 / Compare
| 工具 Tool | 功能 Function |
| --- | --- |
| 两表比对 / Compare two sheets | 比较两个工作表的数据差异 / Find data differences between two sheets |
| 差异比对 / Diff compare | 按匹配列联两表，找比对列不一致的行 / Join two sheets by a key column, flag mismatches |
| 累计去重计数 / Cumulative distinct count | 截至每行统计某列累计不重复值数 / Count cumulative distinct values up to each row |

### 统计 / Statistics
| 工具 Tool | 功能 Function |
| --- | --- |
| 分组统计 / Group statistics | 按分类字段生成交叉汇总表（分类×月份+合计）/ Cross-tab summary by category (category × month + totals) |
| 权限矩阵 / Permission matrix | 从数据生成角色×功能权限矩阵，★标记 / Build a role × function permission matrix, marked with ★ |
| 文件名列表 / List file names | 递归提取目录下所有文件名到当前表 / Recursively list all file names under a folder into the active sheet |

### 匹配 / Match
| 工具 Tool | 功能 Function |
| --- | --- |
| 多字段匹配 / Multi-field match | 按多字段联合匹配两表，提取Key值填充目标表 / Join two sheets on multiple fields, fill a key into the target |

---

## 设计守则 / Design principles

1. **所有参数用户输入** — 列名、表名全部参数化，不硬编码
   **All parameters are user-input** — column and sheet names are parameterized, never hardcoded.
2. **动态类型操作** — `dynamic app = (dynamic)ExcelDnaUtil.Application` 代替 VSTO
   **Dynamic COM access** — uses `dynamic app = (dynamic)ExcelDnaUtil.Application` instead of VSTO.
3. **OleDb 读取** — 大数据量用 SQL 查询，不走逐格 COM 调用
   **OleDb reads** — large data goes through SQL queries, not cell-by-cell COM calls.
4. **进度窗体** — TopMost + 可设进度和消息
   **Progress form** — TopMost window with settable progress and message.
5. **备份优先** — 排序等破坏性操作先备份原始数据
   **Backup first** — destructive ops (e.g. sort) back up originals beforehand.
6. **输出到新表** — 不碰原数据，结果写入新工作表
   **Output to a new sheet** — source data is never mutated; results go to a new sheet.

---

## 目录结构 / Project layout

```
ExcelTools/
├── ExcelTools.csproj        # 项目配置 / project config
├── ExcelTools.dna           # Ribbon XML（11 个按钮）/ Ribbon XML (11 buttons)
├── ExcelTools.slnx          # 解决方案 / solution
├── RibbonCommands.cs        # Ribbon 事件入口 / Ribbon event entry
├── Common/
│   └── MyTools.cs           # 公共工具函数 / shared helpers
├── Forms/
│   ├── ProgressForm.cs      # 进度窗体 / progress form
│   ├── FrmSeperate.cs       # 拆分工作表 / split sheet
│   ├── FrmFilter.cs         # 条件筛选 / filter
│   ├── FrmSort.cs           # 按列排序 / sort
│   ├── FrmDelete.cs         # 删除匹配行 / delete matched rows
│   ├── FrmCompare.cs        # 两表比对 / compare two sheets
│   ├── FrmTimeCompare.cs    # 差异比对 / diff compare
│   ├── FrmKeyMatch.cs       # 多字段匹配 / multi-field match
│   ├── FrmCountDistinct.cs  # 累计去重计数 / cumulative distinct count
│   ├── FrmGroupStat.cs      # 分组统计 / group statistics
│   ├── FrmPermMatrix.cs     # 权限矩阵 / permission matrix
│   └── FrmGetFileNames.cs   # 文件名列表 / list file names
└── README.md
```

---

## 依赖 / Dependencies

- .NET 6.0 Windows
- ExcelDna.Integration 1.9.0
- ExcelDna.AddIn 1.9.0
- System.Data.OleDb 7.0.0
- Microsoft ADO.NET OLE DB Provider for Microsoft ACE 12.0

---

## 许可 / License

本项目采用 MIT 许可。源代码不含任何第三方公司标识、项目代号或专有业务逻辑。
This project is licensed under the MIT License. The source code contains no third-party company identifiers, project codes, or proprietary business logic.
