# ExcelTools · Excel 数据处理插件 · Excel Data Tools Add-in

[中文](#中文) | [English](#english)

---

<a id="中文"></a>
# 中文

> **一句话**：Excel-DNA 插件，11 个通用数据处理工具，拷贝即用。

## 一、构建

```bash
cd source/ExcelTools
dotnet build -c Release
```

构建产物在 `bin/Release/`：
- `ExcelTools.xll` — 发给客户就这一个文件
- `ExcelTools.dna` — 配置文件（Ribbon 定义）
- `ExcelTools.dll` — 核心程序集

## 二、部署到客户机

1. 把 `ExcelTools.xll` + `ExcelTools.dna` + `ExcelTools.dll` 拷到一个目录
2. 双击 `.xll`（或在 Excel 里 文件→选项→加载项→Excel 加载项→浏览→选 .xll）
3. Ribbon 出现「Excel 工具」选项卡
4. 不需要管理员权限、不需要 COM 注册、不需要证书

## 三、11 个工具

### 数据处理
| 工具 | 功能 |
| --- | --- |
| 拆分工作表 | 按某列值将总表拆成多个 Sheet |
| 条件筛选 | 多条件筛选到新表（逐行匹配 / SQL 两种模式） |
| 按列排序 | 按指定列排序，自动备份原始数据到 Original 表 |
| 删除匹配行 | 删除与另一表匹配的行，全部参数化 |

### 比对
| 工具 | 功能 |
| --- | --- |
| 两表比对 | 比较两个工作表的数据差异 |
| 差异比对 | 按匹配列联两表，找比对列不一致的行 |
| 累计去重计数 | 截至每行统计某列累计不重复值数 |

### 统计
| 工具 | 功能 |
| --- | --- |
| 分组统计 | 按分类字段生成交叉汇总表（分类×月份+合计） |
| 权限矩阵 | 从数据生成角色×功能权限矩阵，★标记 |
| 文件名列表 | 递归提取目录下所有文件名到当前表 |

### 匹配
| 工具 | 功能 |
| --- | --- |
| 多字段匹配 | 按多字段联合匹配两表，提取 Key 值填充目标表 |

## 四、设计守则

1. **所有参数用户输入** — 列名、表名全部参数化，不硬编码
2. **动态类型操作** — `dynamic app = (dynamic)ExcelDnaUtil.Application` 代替 VSTO
3. **OleDb 读取** — 大数据量用 SQL 查询，不走逐格 COM 调用
4. **进度窗体** — TopMost + 可设进度和消息
5. **备份优先** — 排序等破坏性操作先备份原始数据
6. **输出到新表** — 不碰原数据，结果写入新工作表

## 五、目录结构

```
ExcelTools/
├── ExcelTools.csproj        # 项目配置
├── ExcelTools.dna           # Ribbon XML（11 个按钮）
├── ExcelTools.slnx          # 解决方案
├── RibbonCommands.cs        # Ribbon 事件入口
├── Common/
│   └── MyTools.cs           # 公共工具函数
├── Forms/
│   ├── ProgressForm.cs      # 进度窗体
│   ├── FrmSeperate.cs       # 拆分工作表
│   ├── FrmFilter.cs         # 条件筛选
│   ├── FrmSort.cs           # 按列排序
│   ├── FrmDelete.cs         # 删除匹配行
│   ├── FrmCompare.cs        # 两表比对
│   ├── FrmTimeCompare.cs    # 差异比对
│   ├── FrmKeyMatch.cs       # 多字段匹配
│   ├── FrmCountDistinct.cs  # 累计去重计数
│   ├── FrmGroupStat.cs      # 分组统计
│   ├── FrmPermMatrix.cs     # 权限矩阵
│   └── FrmGetFileNames.cs   # 文件名列表
└── README.md
```

## 六、依赖

- .NET 6.0 Windows
- ExcelDna.Integration 1.9.0
- ExcelDna.AddIn 1.9.0
- System.Data.OleDb 7.0.0
- Microsoft ADO.NET OLE DB Provider for Microsoft ACE 12.0

## 许可

本项目采用 MIT 许可。源代码不含任何第三方公司标识、项目代号或专有业务逻辑。

---

<a id="english"></a>
# English

> **One-liner**: An Excel-DNA add-in with 11 general-purpose data-processing tools, copy-and-run.

## 1. Build

```bash
cd source/ExcelTools
dotnet build -c Release
```

Build artifacts are placed in `bin/Release/`:
- `ExcelTools.xll` — the only file you ship to a client
- `ExcelTools.dna` — config file (Ribbon definition)
- `ExcelTools.dll` — core assembly

## 2. Deploy to a client machine

1. Copy `ExcelTools.xll` + `ExcelTools.dna` + `ExcelTools.dll` into one folder.
2. Double-click the `.xll` (or in Excel: File → Options → Add-ins → Excel Add-ins → Browse → pick the .xll).
3. The "Excel Tools" tab appears on the Ribbon.
4. No admin rights, no COM registration, no code-signing certificate required.

## 3. The 11 tools

### Data processing
| Tool | Function |
| --- | --- |
| Split sheet | Split a sheet into multiple sheets by the values in a column |
| Filter | Filter rows into a new sheet (row-by-row or SQL mode) |
| Sort | Sort by a chosen column, auto-backup originals to an "Original" sheet |
| Delete matched rows | Delete rows that match another sheet; fully parameterized |

### Compare
| Tool | Function |
| --- | --- |
| Compare two sheets | Find data differences between two sheets |
| Diff compare | Join two sheets by a key column, flag mismatches |
| Cumulative distinct count | Count cumulative distinct values up to each row |

### Statistics
| Tool | Function |
| --- | --- |
| Group statistics | Cross-tab summary by category (category × month + totals) |
| Permission matrix | Build a role × function permission matrix, marked with ★ |
| List file names | Recursively list all file names under a folder into the active sheet |

### Match
| Tool | Function |
| --- | --- |
| Multi-field match | Join two sheets on multiple fields, fill a key into the target |

## 4. Design principles

1. **All parameters are user-input** — column and sheet names are parameterized, never hardcoded.
2. **Dynamic COM access** — uses `dynamic app = (dynamic)ExcelDnaUtil.Application` instead of VSTO.
3. **OleDb reads** — large data goes through SQL queries, not cell-by-cell COM calls.
4. **Progress form** — TopMost window with settable progress and message.
5. **Backup first** — destructive ops (e.g. sort) back up originals beforehand.
6. **Output to a new sheet** — source data is never mutated; results go to a new sheet.

## 5. Project layout

```
ExcelTools/
├── ExcelTools.csproj        # project config
├── ExcelTools.dna           # Ribbon XML (11 buttons)
├── ExcelTools.slnx          # solution
├── RibbonCommands.cs        # Ribbon event entry
├── Common/
│   └── MyTools.cs           # shared helpers
├── Forms/
│   ├── ProgressForm.cs      # progress form
│   ├── FrmSeperate.cs       # split sheet
│   ├── FrmFilter.cs         # filter
│   ├── FrmSort.cs           # sort
│   ├── FrmDelete.cs         # delete matched rows
│   ├── FrmCompare.cs        # compare two sheets
│   ├── FrmTimeCompare.cs    # diff compare
│   ├── FrmKeyMatch.cs       # multi-field match
│   ├── FrmCountDistinct.cs  # cumulative distinct count
│   ├── FrmGroupStat.cs      # group statistics
│   ├── FrmPermMatrix.cs     # permission matrix
│   └── FrmGetFileNames.cs   # list file names
└── README.md
```

## 6. Dependencies

- .NET 6.0 Windows
- ExcelDna.Integration 1.9.0
- ExcelDna.AddIn 1.9.0
- System.Data.OleDb 7.0.0
- Microsoft ADO.NET OLE DB Provider for Microsoft ACE 12.0

## License

This project is licensed under the MIT License. The source code contains no third-party company identifiers, project codes, or proprietary business logic.
