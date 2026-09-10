# ExcelNavigatorPane

ExcelNavigatorPane 是一个基于 VSTO、WinForms 与 Excel Interop 的 Excel 导航栏加载项。当前范围只覆盖 Kutools Navigation Pane 中的“工作簿和表”模块：浏览、切换和管理已打开的工作簿与工作表，不包含资源库、名称管理器、列导航或跨工作簿查找替换。

## 功能

### 导航布局

- 工作簿区默认预留五行，超过五个时自动增高；达到窗格可用高度后才在列表内滚动，并为工作表工具栏和至少三行工作表保留空间（窗格高度足够时）
- 拖动两区之间的分隔线可手动调整高度；当前窗格会保留手动高度，刷新不会覆盖，双击分隔线恢复自动高度
- 工作表搜索独占一行，输入框随窗格宽度调整
- 新建工作簿保留快捷入口，排序、保存等操作集中在工作簿菜单
- 统一文字、行高与线性图标；隐藏和受保护状态常显，普通行操作在悬停或键盘聚焦时显示
- 工作簿使用装订册图标，工作表使用网格图标；工具栏与行内按钮共用圆角线条图标，普通操作悬停为绿色，关闭按钮为红色。当前工作簿的关闭入口也保持显示。


### 多窗口

- 每个 `Excel.Window` 拥有独立的 Navigation TaskPane
- 每个窗格绑定自己的 `Window` 与 `Workbook`
- 同一工作簿通过“视图 → 新建窗口”打开多个窗口时，每个窗格按各自的 `Window.ActiveSheet` 高亮
- 工作簿级临时隐藏状态在同一工作簿的多个窗口之间共享

### 工作簿

- 列出当前 Excel 实例中打开的工作簿并激活所选项
- 按默认打开顺序、A→Z 或 Z→A 排列；再次点击当前排序菜单项恢复默认顺序
- 新建和打开工作簿
- 保存已保存工作簿；未保存工作簿使用 Excel 原生 Save As 对话框
- 通过 `SaveAs` 重命名文件，保留扩展名并检查非法名称与已打开同名工作簿
- 关闭明确选中的工作簿，并保留 Excel 原生的保存/取消提示

### 工作表

- 列出所属工作簿中的 Worksheet、筛选名称并高亮当前窗口的活动表
- 默认列出隐藏表，标题栏按钮显示下一步操作“仅看可见表”；点击后列表和名称搜索仅显示可见工作表，按钮变为“查看全部表”。此开关仅过滤当前窗格，不改变 Sheet 的隐藏状态，底部数量仍按工作簿实际状态统计。
- “取消隐藏 / 恢复隐藏”会改变 Excel 中普通工作表的隐藏状态；这组按钮与列表筛选按钮的悬停提示均随状态切换，说明点击后的作用。
- 显示工作表总数、可见数与隐藏数
- 单击工作表名称切换工作表
- 单击眼睛图标切换普通 Visible/Hidden 状态
- `xlSheetVeryHidden` 只显示状态，不由导航栏修改
- 单击锁图标调用 Excel 原生保护/取消保护命令
- 临时显示当前工作簿中全部普通 Hidden 工作表，再次点击优先按 `CodeName` 恢复；若 Excel 未提供 CodeName，则在该工作簿生命周期内按工作表 COM 身份恢复
- 每个窗口独立记录最近使用的两张工作表，可连续往返切换
- 右键可见工作表时调用 Excel 原生工作表标签菜单（`Ply`）；隐藏表只显示适用的导航栏命令

### 公式编辑时导航

- 输入未完成公式后，单击工作簿行正文或可见工作表行正文，可继续切换并点击单元格生成引用；公式由用户按 Enter 提交或 Esc 取消。
- 工作表通过所属窗口的 Excel 原生标签切换；工作簿通过缓存的精确窗口 HWND 切换。同一工作簿仍只列一行：本工作簿使用当前窗格绑定的窗口，其他工作簿使用上次刷新时的第一扇可见窗口。
- 编辑期间保留缓存列表；切表成功后更新高亮，结束编辑后刷新实际状态，包括 Esc 返回原表的情况。
- 原生标签缺失、禁用、归属不明确或窗口失效时，提示错误并停止；恢复仍属于本次点击的列表选择，保留未完成公式。不会发送 Enter/Esc、模拟点击原生标签或退回 COM 激活。
- 保存、重命名、保护、隐藏、最近两表、右键命令、手动刷新与排序仍要求 Excel 已结束编辑且可执行命令。编辑时宿主可能直接屏蔽这些入口。

## 行为边界

- 隐藏和恢复操作不会修改 `xlSheetVeryHidden`
- 工作簿结构保护开启时不会修改工作表可见性
- 隐藏/恢复前会保证工作簿至少保留一张可见 Sheet
- 临时显示期间新增的隐藏表不进入旧快照；已删除的快照工作表会安全跳过
- SaveAs 或导航栏重命名不会改变该工作簿的临时隐藏状态
- 关闭工作簿时如果用户取消，临时隐藏状态不会被提前清除
- 用户主动执行命令失败时显示简洁错误；Excel 被动事件刷新失败只写入 Debug 输出，避免连续弹窗

## 技术栈

- .NET Framework 4.8
- VSTO Excel Add-in
- WinForms
- Microsoft.Office.Interop.Excel
- Microsoft.Office.Tools / CustomTaskPane
- .NET Framework 自带的 UIAutomationClient / UIAutomationTypes（无新增 NuGet 包）

关键文件：

- `ThisAddIn.cs`：Excel 事件、窗口级 TaskPane 生命周期和工作簿级共享状态
- `NavigationPaneControl.cs`：导航 UI、窗口上下文、命令与列表刷新
- `ExcelNavigatorPane.csproj`：VSTO、Excel 宿主与调试配置

## 运行环境

- Windows 和 Microsoft Excel 桌面版
- 支持 VSTO 的 Visual Studio
- .NET Framework 4.8 Developer Pack
- VSTO Runtime 与 Excel COM 组件

## 构建与调试

可直接用 Visual Studio 打开：

- `ExcelNavigatorPane.csproj`
- `ExcelNavigatorPane.slnx`

项目的调试宿主是 Excel，启动参数已配置为 `/x`，因此 F5 会启动一个新的 Excel 实例，而不是附加到当前正在使用的实例。

在 Visual Studio Developer PowerShell 中可执行：

```powershell
msbuild .\ExcelNavigatorPane.csproj /t:Rebuild /p:Configuration=Debug /p:Platform=AnyCPU
```

### 内部试用安装包

运行 `./installer/Build-Installer.ps1` 可发布 Release/ClickOnce 并生成 `dist/ExcelNavigator-Setup-1.0.0.2.exe`，接收方只需这个 EXE；同时生成可供上传的完整发布 ZIP。安装时保存并关闭 Excel，缺少 .NET Framework 4.8 或 VSTO Runtime 时联网安装依赖。详细步骤见生成目录中的 `安装说明.txt`。

脚本默认在当前用户证书库中创建或复用内部代码签名证书，不写入受信任根或受信任发布者，不更改项目的开发签名配置。正式证书可用 `-CertificateThumbprint` 指定，构建工具位置可用 `-MSBuild` 指定。输出目录 `dist/` 为生成物。

单文件启动器将安装源解包到 `%LOCALAPPDATA%/ExcelNavigatorPane/Installer`，再启动微软生成的安装程序；插件安装和卸载由 VSTO/ClickOnce 处理。安装器面向 Windows 10/11，其他 Excel 版本及 x86 仍需实机验证。

打包脚本会运行仅解包检查，逐文件比对发布内容的 SHA-256，不在开发机执行安装或修改 Excel 注册。当前已通过这些检查；无开发环境电脑上的安装、加载和卸载为 **NOT_RUN**。内部自签名证书可能触发发布者信任提示或被公司策略阻止，正式分发需使用组织认可的签名方案。

### 更新功能

- 工作簿右上角菜单提供“检查更新”，异步查询版本，最多等待 10 秒；“关于导航栏”显示当前发布版本。
- 未配置更新地址时明确提示，且不发出网络请求。手动检查只读取清单版本，不下载或执行程序；网页、错误产品清单、非 HTTPS 地址和临时签名 URL 不作为更新来源。
- 发布时传入 `-UpdateBaseUrl`（固定 HTTPS 目录），VSTO 清单启用每次加载时检查更新；安装程序从该在线地址安装，Office 后续也沿用这个源。插件不会在用户工作期间关闭或重启 Excel。手动发现新版后提示保存工作、关闭所有 Excel 窗口并重新打开。
- 未传地址时，EXE 保持本地安装模式，自动更新关闭。不能把当前 OneView 首页当作文件更新源。

```powershell
# 当前可生成未配置地址的包
.\installer\Build-Installer.ps1 -Version 1.0.0.2

# 生成带正式更新地址的发布包
.\installer\Build-Installer.ps1 -Version 1.0.0.2 -UpdateBaseUrl 'https://oneview.jiarui.net.cn/excel-navigation/'
```

发布后将 `ExcelNavigator-Publish-版本号.zip` 解压上传到配置目录：先上传 `Application Files` 中的新版本文件，最后覆盖根目录 `ExcelNavigatorPane.vsto`；保留旧版目录，不修改已签名的清单或程序集。每次发布递增版本号、沿用同一签名证书及固定地址。入口清单应避免长期缓存，并设置 `.vsto` 类型为 `application/x-ms-vsto`。

旧的离线来源安装用户需要先安装一次带在线地址且版本号更高的包；已安装加载项的来源迁移仍需独立电脑验收，不保证仅替换 EXE 即完成迁移。

更新检查的解析/地址校验见 `tests/UpdateCheck.cs`；发布 ZIP 中的版本、原生更新标记、签名存在及安装源检查见 `tests/UpdatePublishCheck.ps1`。这些检查不等于真实升级验收：从旧版安装 → 发布新版 → 重启 Excel 加载新版、断网启动及证书不受信任时的行为均为 **NOT_RUN**，待本版安装后在独立环境执行。

### 发布流程

RustFS 发布脚本复用同一批构建产物，默认仅预览；验证通过后加 `--apply` 上传并逐文件校验匿名下载，最后更新入口清单。版本文件已存在但内容不同会拒绝覆盖：

```powershell
python installer/Publish-RustFS.py --version 1.0.0.2 --directory dist/releases/1.0.0.2
python installer/Publish-RustFS.py --version 1.0.0.2 --directory dist/releases/1.0.0.2 --apply
```

后续每次发布均包含以下步骤，更新日志统一维护在 [CHANGELOG.md](CHANGELOG.md)：

1. 确定递增的四段版本号，将“未发布”内容整理为该版本和发布日期，写明用户可见变化、验证结果及已知限制。
2. 完成相关构建和检查，提交本次发布源码与更新日志并推送 GitHub。排除凭据、私钥、生成物和无关本地修改。
3. 从该提交构建安装包，显式指定版本、正式 `UpdateBaseUrl` 和沿用的签名证书；验证产物。创建指向该提交的 `v版本号` 标签并推送。
4. 创建 GitHub Release 草稿，说明取自该版本更新日志；附上 EXE、EXE 的 SHA-256 文件、完整发布 ZIP 和安装说明。后续 RustFS 上传复用同一批产物。
5. 将发布 ZIP 解压后的目录结构上传至 RustFS 的 `excel-navigation` 桶内固定更新目录。先上传新版本 `Application Files` 并校验下载内容，再上传根目录安装资源，最后覆盖 `ExcelNavigatorPane.vsto`。保留旧版目录，不使用删除同步。入口清单避免长期缓存；签名文件保持原样。
6. 从用户使用的 HTTPS 地址验证匿名下载、清单版本及文件完整性，再发布 Release。首次接入先发布内部试用预发布版，明确未完成的实机验收；独立环境的安装及旧版升级验收通过后再发布稳定版。报告提交、标签、Release 链接、下载地址和实际验证结果。

正式更新目录为 `https://oneview.jiarui.net.cn/excel-navigation/`，S3 Endpoint 为同域名根地址，凭据保存在 Git 忽略的本机 `.env`。2026-09-11 已通过真实上传、覆盖、签名/匿名下载及权限检查：`python tests/RustFSUploadCheck.py`（需本机已安装 boto3、python-dotenv、requests；每次仅新增一组小型诊断对象并保留）。详情见 [RustFS 接入说明](installer/RustFS接入说明.md)。1.0.0.2 为首个在线更新试用版；真实 Excel 升级验收尚未执行；任一步失败应保留已完成步骤的记录，修复后继续，不重复发布同版本的不同产物。

## 手工验收清单

### 加载与窗口隔离

- F5 启动新的 `/x` Excel 实例，加载项成功加载且左侧出现 Navigation TaskPane
- 打开两个不同工作簿；切换窗口时，各窗格的工作簿高亮和工作表列表不串扰
- 在同一工作簿中执行“视图 → 新建窗口”，两个窗口分别激活不同工作表；各窗格高亮自己的活动表

### 工作簿命令

- 新建、打开、首次保存、普通保存均由明确命令完成
- 重命名保留原扩展名；空名称、路径片段、非法字符和已打开同名工作簿得到可理解提示
- 关闭含未保存更改的工作簿时，Excel 原生保存/取消提示正常；取消后窗格仍可使用
- A→Z、Z→A 和恢复默认打开顺序均正确

### 公式编辑导航

- BookA/Sheet1 的 A1 输入 `=`，点击导航栏 Sheet2，再点击 B2；公式应保持编辑并显示 `=Sheet2!B2`。分别验证 Esc 后 A1 不提交、Enter 后引用正确。
- 重复上述操作，将目标改为 BookB 行及其 B2；检查实际目标窗口与 `[BookB.xlsx]Sheet1!$B$2` 引用，分别取消和提交。
- A、B 各开两个窗口；各窗格切表仅影响绑定窗口，跨工作簿命中刷新时缓存的可见窗口。
- 在编辑前滚动导航列表，选取原生标签条以外的工作表；关闭 Excel 的“显示工作表标签”后重试，确认只提示一次错误，原选择恢复且公式仍可继续编辑。
- 编辑时尝试命令入口，排序状态与工作簿内容不得改变；结束编辑后重试普通命令。

2026-09-08 已在独立 `/x` 实例、Excel **16.0.20326.20132 x64** 中验证跨表/跨簿引用、提交/取消、多窗口、条带外 Sheet24、标签关闭失败与编辑结束后的高亮恢复。其他 Office 版本需复跑本节，构建通过不代表 UIA 兼容。

### 工作表命令

- 筛选、活动项高亮与总数/可见数/隐藏数正确
- “仅看可见表 / 查看全部表”按钮文字随状态切换，并与名称搜索共同生效；关闭后 Hidden/VeryHidden 均不列出，重新选中后恢复，实际隐藏状态及底部计数不变
- 普通 Hidden 可通过眼睛图标显示；Visible 可在满足至少一张可见 Sheet 时隐藏
- VeryHidden 不被修改；工作簿结构保护时可见性操作被阻止
- 锁图标打开 Excel 原生保护/取消保护流程，取消不会导致导航栏异常
- 右键可见工作表打开该表的 Excel 原生工作表标签菜单

### 批量隐藏恢复

- 准备 Visible、Hidden、VeryHidden 工作表；“取消隐藏”只显示普通 Hidden
- “恢复隐藏”仅恢复本次快照中的工作表，VeryHidden 始终不变
- 同一工作簿的两个窗口同步显示“取消隐藏/恢复隐藏”状态
- 临时显示期间 SaveAs 或重命名后仍可恢复
- 删除快照内工作表或取消关闭工作簿后，恢复操作不会丢失其他有效状态

### 最近两张表

- 按 A→B→C 激活后，连续点击“最近两表” 可在 B 与 C 之间往返
- 同一工作簿的不同窗口拥有各自历史
- 工作表重命名后历史仍有效；目标被删除或隐藏时显示可理解提示

## 本版验证

- Debug/AnyCPU Rebuild：0 错误、0 警告。
- 布局检查：280/320/400 宽度、480/900 高度、1/2/5/20 个工作簿，共 24 组通过。
- 新布局在独立 `/x` Excel 中验证搜索筛选、未提交的 `=Sheet24!B2` 引用及 Esc 返回原表、跨工作簿精确窗口切换。
- 2026-09-08：用户确认交互验证无问题，授权提交本版。完整交互回归的最终确认来自用户；不将其记作自动化测试结果。

## 已知限制

- 编辑态切表依赖 Office UIA 提供程序暴露 `XLDESK → ExcelGrid → ExcelBookTabControl → SheetTab` 及 `SelectionItemPattern`；关闭原生标签或提供程序结构不同会停止切换，不使用猜测性兜底。
- Excel 会把编辑态窗格点击交给原生宿主。加载项只在所属 UI 线程转接两张导航列表正文的实际左键点击；不会整体启用被 Excel 禁用的命令。
- 当前没有自动化测试项目；窗口、对话框和 COM 事件行为需要在真实 Excel 中验收
- 导航列表只覆盖 Worksheet，不列出图表工作表等其他 Sheet 类型
- 工作簿重命名本质上是 Excel `SaveAs`，覆盖确认、格式兼容和路径长度错误由 Excel 处理
- “关于导航栏”显示打包时的发布版本；本地开发构建默认不启用自动更新
- 当前只实现“工作簿和表”模块，不承诺与 Kutools 其他模块或视觉资产兼容
