# ExcelTools · Excel 数据处理插件 · Excel Data Tools Add-in

[中文](#中文) | [English](#english)

---

<a id="中文"></a>
# 中文

> **一句话**：Excel-DNA 插件，11 个通用数据处理工具——把"手工搬半天"的活压成一键。

## 一、构建

```bash
cd source/ExcelTools
dotnet build -c Release
```

构建产物在 `bin/Release/net10.0-windows/`：

| 文件 | 说明 |
| --- | --- |
| `ExcelTools64.xll` | ★ **64 位 Excel 用这个** |
| `ExcelTools.xll` | 32 位 Excel 用（名字里没有 32，但实测 PE 头是 x86） |
| `ExcelTools64.dna` / `ExcelTools.dna` | 配置，与上面两个 xll 一一对应 |
| `ExcelTools.dll` | 核心程序集 |
| `ExcelTools.deps.json` / `ExcelTools.runtimeconfig.json` | .NET 10 必需，缺了加载失败 |
| `runtimes/win/lib/net10.0/System.Data.OleDb.dll` | OleDb 真正的实现，缺了「SQL 查询」功能失效 |
| `publish/*-packed.xll` | Excel-DNA 单文件打包版，**本项目的 .NET 10 组合下不可用**，见下 |

## 二、部署到客户机

> **给客户装，直接用一键安装包**：`installer/` 目录下的 Inno Setup 脚本编译出来的
> `ExcelTools-Setup-*.exe`。客户双击装完就有，下面 1–7 步全都是它自动做的
> （判断 32/64 位、写信任位置、注册加载项、清禁用项、支持干净卸载）。
> 手工部署只在排查问题或客户不让装 exe 时才需要走。

1. **先确认位数**：Excel 是 64 位（默认）→ 用 `ExcelTools64.xll`；32 位 Excel → 用 `ExcelTools.xll`。**位数拿错一定报「该加载项不是有效加载项」**，这是最常见的失败原因。
2. **整目录拷，别只拷三个文件**：把 `ExcelTools64.xll` + `ExcelTools64.dna` + `ExcelTools.dll` + `ExcelTools.deps.json` + `ExcelTools.runtimeconfig.json` + `runtimes/` 一起拷到一个目录。.NET 10 与 .NET Framework 不同，少了 `deps.json`/`runtimeconfig.json` 加载会失败，少了 `runtimes/` 则 OleDb 拿不到 Windows 专用实现。
3. **不要用 `publish/` 下的 `-packed.xll`**。那是 Excel-DNA 的 LZMA 单文件打包，在 .NET 10 上存在已知缺陷（解压后的平台专用程序集解析不到，官方 issue #679），本项目的 `System.Data.OleDb` 正好属于这一类。开发期也建议只用 unpacked 版。
4. **目标机需要 .NET 10 Desktop Runtime**（`Microsoft.WindowsDesktop.App 10.0.x`）。没装就不启动，且提示不一定明显。
5. **失败过一次要检查禁用项**：Excel 会把加载失败的 xll 自动加进「文件→选项→加载项→管理:禁用项目」，之后双击静默无反应。先去那里启用回来，再重试。
6. 双击 `.xll`（或 文件→选项→加载项→Excel 加载项→浏览→选 .xll），Ribbon 出现「Excel 工具」选项卡。
7. 不需要管理员权限、不需要 COM 注册、不需要证书。

### 打包安装包（推荐交付方式）

脚本在 `installer\ExcelTools.iss`，两个变体：

| 变体 | 命令 | 体积 | 适用 |
| --- | --- | --- | --- |
| Lite | `ISCC.exe ExcelTools.iss` | ~2.9 MB | 客户机已装 .NET 10 桌面运行时 |
| Full | `ISCC.exe /DBundleRuntime ExcelTools.iss` | ~62.9 MB | 客户机可能没装 .NET 10，缺了自动装（弹一次 UAC） |

产物在 `installer\output\`。Full 包内嵌 `installer\runtime\windowsdesktop-runtime-10-x64.exe`（微软官方签名，10.0.13）；该目录不进 git。

**为什么不走 self-contained**：Excel-DNA 明确不支持（[issue #811](https://github.com/excel-dna/exceldna/issues/811)，作者原话 *"the add-in can't work in self-contained mode"*）。.NET Core 运行时没法自带进 Excel 进程。所以"客户不用管运行时"只能靠安装包内嵌运行时安装器实现。

**★ `.dna` 必须带 `xmlns="http://schemas.excel-dna.net/addin/2020/07/dnalibrary"`**：缺了它，Excel-DNA 按旧格式解析，`RuntimeVersion` 被忽略并回落到 .NET 6 —— 表现为加载项能起来、但程序集加载失败，且不报明显错误。升级目标框架时最容易踩这个坑。

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

- .NET 10.0 Windows
- ExcelDna.Integration 1.9.0
- ExcelDna.AddIn 1.9.0
- System.Data.OleDb 7.0.0
- Microsoft ADO.NET OLE DB Provider for Microsoft ACE 12.0

## 许可

本项目采用 MIT 许可。源代码不含任何第三方公司标识、项目代号或专有业务逻辑。

---

<a id="english"></a>
# English

> **One-liner**: An Excel-DNA add-in with 11 general-purpose data-processing tools — it turns half a day of manual work into one click.

## 1. Build

```bash
cd source/ExcelTools
dotnet build -c Release
```

Build artifacts are placed in `bin/Release/net10.0-windows/`:

| File | Purpose |
| --- | --- |
| `ExcelTools64.xll` | ★ **use this one for 64-bit Excel** |
| `ExcelTools.xll` | for 32-bit Excel (no "32" in the name, but the PE header is x86) |
| `ExcelTools64.dna` / `ExcelTools.dna` | config, one per xll above |
| `ExcelTools.dll` | core assembly |
| `ExcelTools.deps.json` / `ExcelTools.runtimeconfig.json` | required by .NET 10; missing them breaks loading |
| `runtimes/win/lib/net10.0/System.Data.OleDb.dll` | the real OleDb implementation; without it the SQL filter mode fails |
| `publish/*-packed.xll` | Excel-DNA single-file packing — **does not work for this project on .NET 10**, see below |

## 2. Deploy to a client machine

> **For client installs, use the one-click installer**: compile `installer/ExcelTools.iss`
> with Inno Setup 6 to get `ExcelTools-Setup-*.exe`. A single double-click does everything
> below (bitness detection, trusted location, add-in registration, disabled-item cleanup,
> clean uninstall). The manual steps remain for troubleshooting or locked-down clients.

1. **Match the bitness first**: 64-bit Excel (the default) → `ExcelTools64.xll`; 32-bit Excel → `ExcelTools.xll`. Picking the wrong one always yields "The file you selected is not a valid Office add-in" — the single most common failure.
2. **Copy the whole folder, not three files**: `ExcelTools64.xll` + `ExcelTools64.dna` + `ExcelTools.dll` + `ExcelTools.deps.json` + `ExcelTools.runtimeconfig.json` + `runtimes/`. Unlike .NET Framework, .NET 10 fails to load without `deps.json` / `runtimeconfig.json`, and falls back to a non-Windows OleDb stub without `runtimes/`.
3. **Do not ship `publish/*-packed.xll`.** That is Excel-DNA's LZMA single-file packing, which has a known defect on .NET 10 (platform-specific assemblies under `runtimes/` are not resolved after unpacking — Excel-DNA issue #679). This project's `System.Data.OleDb` is exactly such a case. Use the unpacked build, also during development.
4. **The target machine needs the .NET 10 Desktop Runtime** (`Microsoft.WindowsDesktop.App 10.0.x`). Without it the add-in silently fails to start.
5. **After one failed load, check Disabled Items**: Excel automatically adds a failing `.xll` to *File → Options → Add-ins → Manage: Disabled Items*, after which double-clicking does nothing. Re-enable it there before retrying.
6. Double-click the `.xll` (or in Excel: File → Options → Add-ins → Excel Add-ins → Browse → pick the .xll). The "Excel Tools" tab appears on the Ribbon.
7. No admin rights, no COM registration, no code-signing certificate required.

### Packaging an installer (recommended delivery)

The Inno Setup script lives in `installer\ExcelTools.iss` with two variants:

| Variant | Command | Size | Use when |
| --- | --- | --- | --- |
| Lite | `ISCC.exe ExcelTools.iss` | ~2.9 MB | the client already has the .NET 10 Desktop Runtime |
| Full | `ISCC.exe /DBundleRuntime ExcelTools.iss` | ~62.9 MB | the client may not — the runtime is bundled and installed silently (one UAC prompt) |

Output lands in `installer\output\`. The Full build embeds `installer\runtime\windowsdesktop-runtime-10-x64.exe` (Microsoft-signed, 10.0.13); that folder is git-ignored.

**Why not self-contained**: Excel-DNA does not support it ([issue #811](https://github.com/excel-dna/exceldna/issues/811) — *"the add-in can't work in self-contained mode"*). A .NET Core runtime cannot be brought into an existing Excel process. Bundling the runtime installer is the supported equivalent.

**★ The `.dna` file must carry `xmlns="http://schemas.excel-dna.net/addin/2020/07/dnalibrary"`.** Without it Excel-DNA parses the file as the legacy format, ignores `RuntimeVersion` and silently falls back to .NET 6 — the add-in starts but the assembly fails to load, with no clear error. This is the trap to watch for when retargeting the framework.

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

- .NET 10.0 Windows
- ExcelDna.Integration 1.9.0
- ExcelDna.AddIn 1.9.0
- System.Data.OleDb 7.0.0
- Microsoft ADO.NET OLE DB Provider for Microsoft ACE 12.0

## License

This project is licensed under the MIT License. The source code contains no third-party company identifiers, project codes, or proprietary business logic.
