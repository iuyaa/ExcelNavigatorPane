# ExcelNavigatorPane

ExcelNavigatorPane 是一个基于 VSTO、WinForms 与 Excel Interop 的 Excel 导航栏加载项。当前范围只覆盖 Kutools Navigation Pane 中的“工作簿和表”模块：浏览、切换和管理已打开的工作簿与工作表，不包含资源库、名称管理器、列导航或跨工作簿查找替换。

## 功能

### 版本、更新与功能说明

- 导航栏底部同一行显示工作表数量与当前版本；“功能说明”保留在工作簿区域右键菜单中，指南随插件内置，离线可读。
- 初始化后延迟 15 秒检查更新，此后同一加载进程每六小时检查一次；多个窗格共享请求和结果。自动检查不弹窗、不访问 Office 对象，失败保持安静，已有新版提示不会因临时失败消失。
- 发现新版在当前版本号后显示红点，点击版本号或红点阅读更新内容；只有确认“下载并安装”才下载。菜单“检查更新”可立即手动重试。
- 打包从 `CHANGELOG.md` 提取截至目标版本的日志，通过 `latest.xml` 的 `notes` 与 `notesSignature` 属性分发和校验，原安装包签名协议保持兼容。日志按版本分段，可用“新增功能、体验优化、问题修复”三级标题分类，无内容的分类不添加；缺失或无效的说明不影响独立校验安装包。

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
- 通过 `SaveAs` 重命名文件，固定保留原扩展名与文件格式；名称中的点号按正文保留，末尾已带原扩展名时不重复追加，并检查非法名称与已打开同名工作簿
- 保存或重命名成功后自动刷新工作簿列表；Excel 忙碌时延迟重试，恢复可操作后显示新名称
- 关闭明确选中的工作簿，并保留 Excel 原生的保存/取消提示

### 工作表

- 列出所属工作簿中的 Worksheet、筛选名称并高亮当前窗口的活动表
- 默认列出隐藏表，标题栏按钮显示下一步操作“仅看可见表”；点击后列表和名称搜索仅显示可见工作表，按钮变为“查看全部表”。此开关仅过滤当前窗格，不改变 Sheet 的隐藏状态，底部数量仍按工作簿实际状态统计。
- “取消隐藏 / 恢复隐藏”会改变 Excel 中普通工作表的隐藏状态；这组按钮与列表筛选按钮的悬停提示均随状态切换，说明点击后的作用。
- 显示工作表总数、可见数与隐藏数
- 单击工作表名称切换工作表
- 单击眼睛图标切换普通 Visible/Hidden 状态
- 从列表取消单张表的隐藏后立即切换到该表；眼睛按钮与隐藏表右键菜单行为一致，批量取消隐藏不自动切表。
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

## WPS 兼容修复（1.0.3.0）

用户已确认本机 WPS 表格曾能自动加载此插件；1.0.3.0 修复 64 位 Excel 与 32 位 WPS 共存时新版注册缺失。用户反馈 WPS 切换与重复窗格问题初测已解决。本版针对该环境增加工作簿 COM 切换、同一主窗口复用导航栏和关闭后的清理；不同主窗口仍分别管理导航栏，Excel 的每窗口独立行为保留。

WPS 就绪状态使用宿主的 Ready 属性，编辑或忙碌时拒绝切换；Excel 原有公式编辑导航不变。WPS 公式编辑导航尚未适配，真实 WPS/Excel 点击回归为 **NOT_RUN**。安装器按 Windows 位数选择 MSI；64 位 Windows 同时注册 32/64 位加载项，只有 WPS 的电脑安装尚未适配。升级前必须保存并关闭 Excel 与 WPS 表格。

`tests/WpsCompatibilityCheck.cs` 使用测试进程自己的隐藏窗口和 COM 接口替身，覆盖窗格复用、创建重入、跨主窗口隔离、关闭后的清理、两种宿主取消隐藏的激活顺序，以及忙碌时禁止 COM 切换；已纳入 `tests/UpdatePublishCheck.ps1`。

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
- .NET Framework 4.8
- VSTO Runtime 与 Excel COM 组件

## 构建与调试

开发机需要支持 VSTO 的 Visual Studio 与 .NET Framework 4.8 Developer Pack；安装包接收方不需要这些开发工具。

可直接用 Visual Studio 打开：

- `ExcelNavigatorPane.csproj`
- `ExcelNavigatorPane.slnx`

项目的调试宿主是 Excel，启动参数已配置为 `/x`，因此 F5 会启动一个新的 Excel 实例，而不是附加到当前正在使用的实例。

在 Visual Studio Developer PowerShell 中可执行：

```powershell
msbuild .\ExcelNavigatorPane.csproj /t:Rebuild /p:Configuration=Debug /p:Platform=AnyCPU
```

### EXE 安装包（内含 MSI）

以下命令生成 `dist/releases/1.0.16.0/渠道/ExcelNavigator-Setup-1.0.16.0.exe`。用户只需一个 EXE，无需 Visual Studio。EXE 检查桌面 Excel 已安装后，按 Windows 位数选择内置 MSI；64 位包同时注册 32/64 位宿主，共享 AnyCPU 插件；缺少 .NET Framework 4.8 或 VSTO Runtime 时由微软引导程序下载依赖。

```powershell
# 一次性准备 WiX 构建工具；用户电脑不需要 WiX
dotnet tool install wix --version 4.0.6 --tool-path work/tools/wix
./installer/Build-Installer.ps1 -Version 1.0.16.0 -UpdateChannel oneview
./installer/Build-Installer.ps1 -Version 1.0.16.0 -UpdateChannel github
```

安装前保存工作并关闭所有 Excel 与 WPS 表格。MSI 需要管理员授权，将加载项安装到对应的 Program Files 目录，并写入 HKLM 加载项注册（64 位 Windows 同时写入 32/64 位视图），使用 `|vstolocal` 从本地加载。正常 Windows/VSTO 策略下，这条安装方式使用 Program Files 的信任机制，不需要用户导入自签名根证书。EXE 仍可能显示未知发布者或 Windows 安全提示，公司策略也可能限制安装。

从 1.0.0.2 及更早版本迁移时，先在“程序和功能”卸载旧 ClickOnce 版本，再运行新 EXE。开发机的 HKCU 开发版注册也会覆盖 HKLM 注册，必须先取消开发版注册；EXE 检测到旧注册会停止并说明原因，不自动删除注册或源文件。不要直接删除开发项目目录。其他 Windows 用户已有的 ClickOnce 注册需分别处理。

MSI 版后续升级直接运行新版 EXE，Windows Installer 负责替换旧版；拒绝降级，升级失败时通过安装事务恢复旧版。控制面板中使用“Excel Navigator”卸载。安装源按版本保存在 `%LOCALAPPDATA%/ExcelNavigatorPane/Installer/版本号`，供修复使用，安装期间不要清理。

安装器先准备运行环境，再等待 MSI 安装结束；仅成功时显示“Excel Navigator 安装成功”，需要重启时明确提示。取消或失败不会显示成功提示。保存并关闭 Excel/WPS 的提醒窗口标题统一为“Excel Navigator”。

脚本复用当前用户证书库中的现有发布证书；可显式传入 `-CertificateThumbprint`。更新签名的公钥已固定在已安装插件中，换用不同密钥会导致旧客户端拒绝新清单，不能隐式轮换。打包不导入 Root/TrustedPublisher，不修改开发签名配置；`BuildInstaller=true` 跳过 VSTO 默认的开发注册步骤。

### 更新功能

- 工作簿菜单“检查更新”读取固定 HTTPS `latest.xml`，最多等待 10 秒；“关于导航栏”显示安装版本。
- 清单包含版本、相对 EXE 路径、SHA-256 和 RSA-SHA256 签名。插件使用内置公钥验证签名，拒绝篡改、外站路径、非 HTTPS、重定向和错误产品。
- 发现新版后询问是否下载，确认后自动保存到系统临时目录下的 `ExcelNavigatorPane\Updates`（通过 `Path.GetTempPath()` 获取，通常为 `%TEMP%\ExcelNavigatorPane\Updates`），不再弹出另存为窗口；下载最长 2 分钟、最多 128 MiB。文件哈希通过后才替换同版本目标文件，失败会清理临时文件并保留原文件。
- 用户确认下载后，安装包校验通过便自动打开。安装器检测到 Excel 或 WPS 表格运行时，提示先保存工作、关闭所有窗口，再点击“重试”继续；点击“取消”退出安装。自动打开失败时提示文件路径供手动运行，不强制退出或重启 Excel/WPS；不在启动时自动检查。
- 开发构建未配置更新地址时不发起网络请求。正式打包默认使用 `https://oneview.jiarui.net.cn/excel-navigation/`，可用 `-UpdateBaseUrl` 指定其他固定 HTTPS 目录。
- 旧 ClickOnce 根入口和旧文件继续保留；本版不会通过旧 `.vsto` 入口迁移用户，需要首次手工安装 MSI 版。

本地检查：

```powershell
./tests/UpdatePublishCheck.ps1 -Directory dist/releases/1.0.16.0/oneview -Version 1.0.16.0 -UpdateChannel oneview
./tests/UpdatePublishCheck.ps1 -Directory dist/releases/1.0.16.0/github -Version 1.0.16.0 -UpdateChannel github
python tests/RustFSPublishCheck.py
```

这些检查只解包和读取 MSI 表/依赖注册，不执行安装。覆盖两种位数、Program Files/HKLM 注册、升级事务和降级条件、负载哈希、更新清单签名及篡改拒绝、发布顺序和冲突保护。无开发环境电脑上的首次安装、Excel 加载、真实升级/回滚/卸载为 **NOT_RUN**；真实 Excel 验证必须使用 `/x` 新建独立进程。

### 发布流程

更新渠道由 [installer/UpdateChannels.json](installer/UpdateChannels.json) 配置，通过构建参数 `-UpdateChannel oneview|github` 选择，默认 `oneview`；可用 `-ChannelsFile` 指定其他配置文件。地址不会写死在 C# 中，也不需要给客户端任何 GitHub Token 或 S3 凭据。

- `oneview`：读取配置中的 HTTPS `latest.xml`，从对象存储下载；保留旧版兼容。
- `github`：读取公开仓库 `releases.atom`，选择其中最高版本（包含预发布），再下载对应标签的 `latest.xml` 和 EXE；仅允许 GitHub 和其 Release 资源域名的 HTTPS 重定向。私有仓库不能用于匿名客户端更新。
- 安装包保留构建时选定的渠道。1.0.9 及更早的包均继续走 OneView，即便最初是从 GitHub 下载的；要转为 GitHub 渠道，安装一次 GitHub 渠道的新版本。不要用同版本安装包切换渠道，MSI 可能拒绝同版本替换。

1.0.16.0 为内部预发布版本，完整真实升级验收尚未完成，普通修改不自动发布。用户要求“发布”时完成以下全部步骤，更新日志维护在 [CHANGELOG.md](CHANGELOG.md)：

1. 确定递增版本号并整理更新日志，同步 `Properties/UpdateSettings.xml` 和 `Properties/AssemblyInfo.cs` 的 `AssemblyFileVersion`，避免 MSI 因 DLL 文件版本未增长而跳过替换。采用 `主.次.构建.0`，最后一段必须为 0；MSI 只比较前三段（最大分别为 255、255、65535）。下一版例如 1.0.17.0。
2. 完成相关检查，提交源码与日志并推送 GitHub；排除凭据、私钥、生成物和无关本地修改。
3. 从该提交分别用 `-UpdateChannel oneview`、`-UpdateChannel github` 构建，沿用同一发布证书。两个渠道版本号及源码相同，但嵌入的渠道配置不同，EXE 哈希及签名清单也不同，不得混用。对各自产物运行 `tests/UpdatePublishCheck.ps1 -Directory dist/releases/版本号/渠道 -Version 版本号 -UpdateChannel 渠道`，创建并推送指向同一提交的 `v版本号` 标签。
4. 创建 GitHub Release 草稿，上传 **github 目录**内的 EXE、SHA-256、x86/x64 MSI、安装说明、`latest.xml` 和 `update-public-key.xml`。用户默认下载 EXE；MSI 供具备运行环境的 IT 部署使用，必须选择 Windows 对应位数，先关闭 Excel、卸载旧 ClickOnce。
5. 将 **oneview 目录**的产物上传 RustFS：先上传并校验 `releases/版本号/` 内的七个文件，最后更新根 `latest.xml`。保留全部旧版本和旧 ClickOnce 对象，不进行删除同步；同版本内容不同会拒绝覆盖。发布端需要 boto3、python-dotenv、requests、cryptography。
6. 验证 RustFS 匿名下载及 GitHub 草稿附件哈希后发布 Release，再验证公开 GitHub 订阅源、签名清单、EXE 下载；两条渠道均须完成实网更新检查和下载哈希验证。未完成独立电脑安装和升级验收时使用预发布。报告提交、标签、Release 链接、下载地址和验证结果；失败说明已完成与待补步骤。

```powershell
# 默认只检查本地产物并预览，不访问 RustFS
python installer/Publish-RustFS.py --version 1.0.16.0 --directory dist/releases/1.0.16.0/oneview
# 用户要求发布时执行，逐文件上传并校验匿名下载，入口最后更新
python installer/Publish-RustFS.py --version 1.0.16.0 --directory dist/releases/1.0.16.0/oneview --apply
```

正式更新目录为 `https://oneview.jiarui.net.cn/excel-navigation/`，S3 Endpoint 为同域名根地址，凭据保存在 Git 忽略的本机 `.env`。路由和上传账号已在 1.0.0.2 发布时验证；迁移不改变桶、账号或域名，新的入口在本版正式发布时上传。详情见 [RustFS 接入说明](installer/RustFS接入说明.md)。

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

1.0.4.0 更新请求使用独立的 TLS 1.2 设置，不修改 Excel/WPS 全局网络配置。线上 latest.xml 返回 404 时明确提示更新清单尚未提供。已通过独立 .NET 进程的旧 TLS 默认值实网回归（tests/UpdateTransportCheck.cs）；真实安装升级和完整下载新版流程仍待验收。

1.0.5.0 导航栏工作表名称前显示 Sheet 标签原色的小色条；无色不显示，白色/浅色保留边框，活动行绿色标记独立保留。标签颜色在列表刷新或切换时同步，手动改色后可点击工作表区刷新按钮。颜色读取失败不会阻止导航。颜色转换及选中/悬停/隐藏行渲染检查通过（280/320/400 宽度），Excel/WPS 实机颜色验证 NOT_RUN。

1.0.6.0 安装器只阻止与 HKLM 安装路径冲突的用户级注册；指向同一本地清单的注册允许升级。日志位于 %LOCALAPPDATA%\ExcelNavigatorPane\Installer\logs，记录当前 EXE 版本/路径及检查时的两种注册视图。使用 EXE 的 --check-only 参数仅检查注册并写日志，不安装或更改注册表。反复旧注册提示的启动瞬间原因仍待日志确认，不能以事后注册为空声称问题已解决。

1.0.7.0 根据安装日志确认开发注册指向 bin/Debug。安装器对本插件 bin/Debug 或 bin/Release 本地注册提供显式确认迁移：保存注册值、类型及子键到 HKCU\Software\ExcelNavigatorPane\RegistrationBackups 后移除原开发注册，源文件不变。其他注册冲突保持拦截。--check-only 只读。隔离注册表子树回归通过；真实迁移安装尚待用户验证。

1.0.8.0：工作表搜索框有内容时显示清除按钮，清空关键词后保留隐藏表筛选并将焦点留在搜索框。工作簿空白区/标题区右键显示新建、打开、排序、刷新、检查更新、关于；右键具体工作簿再显示保存、重命名、关闭，操作绑定点中的工作簿，右键本身不激活工作簿。三个点入口保留并复用同一菜单，支持菜单键及 Shift+F10。离线目标绑定/空白/空列表/入口共用回归、筛选和三宽度渲染通过；真实 Excel/WPS 点击验证 NOT_RUN。
