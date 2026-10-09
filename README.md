# WordUtilityAddIn

`WordUtilityAddIn` 是一个使用 C#、VSTO 和 .NET Framework 4.8 编写的 Microsoft Word 实用工具加载项。第一版提供批量扫描和修改现有拼音指南（OOXML `w:ruby`）格式的能力，不生成或替换拼音读音。

## 第一版功能

- Ribbon 选项卡：`实用工具`
- Ribbon 分组：`拼音指南`
- 处理范围：全文、当前选区
- 扫描并显示注音数量、当前拼音字号、正文基准字号和上移参数
- 拼音字号以磅输入，以半磅单位写入 OOXML
- 拼音字体留空时保持原字体；填写时更新拼音 run 的字体
- 自动上移：每个 ruby 使用 `hpsBaseText + 2`
- 手动上移：可直接填写半磅单位值
- 全文修改前保存文档并创建同目录时间戳备份
- 修改后校验注音数量、读音指纹、字号、字体和上移参数

核心处理同时更新：

- `w:rubyPr/w:hps`
- `w:rt` 内每个 `w:r/w:rPr/w:sz` 和 `w:szCs`
- `w:rubyPr/w:hpsRaise`
- 自动模式下补齐或保持 `w:rubyPr/w:hpsBaseText`

## 解决方案结构

| 项目 | 职责 |
| --- | --- |
| `WordUtilityAddIn` | VSTO 生命周期、Ribbon 和 Word 文档事务 |
| `WordUtilityAddIn.Core` | 不依赖 Word 的 OOXML 扫描、修改与校验 |
| `WordUtilityAddIn.Verification` | 对测试副本执行离线端到端验证 |

`OoxmlRubyProcessor` 遍历 `/word/` 下所有 XML 内容部件，因此可处理正文、页眉、页脚、脚注、尾注、批注等部件中的 `w:ruby`。包内宏、媒体、关系和其他未修改部件会原样保留。

## 开发环境

- Windows 10/11
- Microsoft Word/Office 2024 64 位
- Visual Studio Community 2026
- `Office/SharePoint 开发` 工作负载
- `.NET 桌面开发` 工作负载
- .NET Framework 4.8
- Visual Studio Tools for Office Runtime

## 构建

在 Visual Studio 中打开 `WordUtilityAddIn.sln`，选择 `Debug | Any CPU` 后构建解决方案。

命令行构建：

```powershell
& 'C:\Program Files\Microsoft Visual Studio\18\Community\MSBuild\Current\Bin\MSBuild.exe' `
  .\WordUtilityAddIn.sln /t:Rebuild /p:Configuration=Debug /m
```

项目包含一个仅供本机开发的临时自签名证书 `WordUtilityAddIn_TemporaryKey.pfx`，用于生成 VSTO/ClickOnce 清单。正式发布前必须替换为组织认可的代码签名证书，并保护其私钥；不要将生产证书和密码提交到源码仓库。

## 调试

1. 用 Visual Studio 打开解决方案，将 `WordUtilityAddIn` 设为启动项目。
2. 按 `F5`。VSTO 项目系统会注册 Debug 清单并启动 Word。
3. 打开文档，在 `实用工具 > 拼音指南` 中先执行扫描。
4. 使用测试文档时先创建副本，不要直接处理原件。

如果选项卡未出现：

1. 在 Word 的 `文件 > 选项 > 加载项` 中检查 `WordUtilityAddIn` 是否位于禁用项目。
2. 检查 `HKCU\Software\Microsoft\Office\Word\Addins\WordUtilityAddIn` 的 `LoadBehavior` 是否为 `3`。
3. 关闭所有 Word 进程，在 Visual Studio 中执行 `清理`、`重新生成` 后再次按 `F5`。
4. 确认 Word 与 Visual Studio 使用同一 Windows 用户和权限级别。

## 使用方式

### 全文

1. 先将文档保存为 `docx`、`docm`、`dotx` 或 `dotm`。
2. 范围选择 `全文`，点击 `扫描`。
3. 设置拼音字号、可选字体和上移模式，点击 `应用`。
4. 加载项保存文档，并在原文档同目录创建 `<原名>.backup-yyyyMMdd-HHmmss.<扩展名>`。
5. 加载项在工作副本中修改并完整校验，然后关闭当前文档、原子替换原文件并重新打开。

处理失败时，工作副本不会替换原文件。备份已在修改前生成，可用于人工恢复。

### 当前选区

第一版使用 `Range.WordOpenXML + InsertXML`，并有以下限制：

- 仅支持主文档正文中的非空选区。
- 页眉、页脚、脚注、文本框等内容请使用全文模式。
- 选区替换处于 Word 自定义撤销记录中；插入或校验失败时会调用撤销。
- 复杂选区（跨节、内容控件边界、受保护区域）可能被 Word 拒绝，遇到此情况应改用全文模式。

## 离线验证

验证工具总是先复制源文档，再修改指定副本：

```powershell
.\WordUtilityAddIn.Verification\bin\Debug\WordUtilityAddIn.Verification.exe `
  'D:\Desktop\21天读书班阅读书籍-注音.docx' `
  '.\TestOutput\21天读书班阅读书籍-注音-9磅-验证副本.docx' `
  9
```

本机验证结果（2026-07-10）：

- 实际 `w:ruby` 数量：3761
- 修改数量：3761
- 拼音字号：9 磅，OOXML `18`
- 正文基准字号：28 磅，OOXML `56`
- 自动上移：OOXML `58`
- 注音数量和全部拼音读音指纹：通过
- 桌面源文档仅只读扫描，未写入

任务背景中的“约 3765 个”是估算值；程序以包内实际 `w:ruby` 节点数 3761 为准。

## 安装和发布

开发调试由 Visual Studio 自动注册 VSTO 清单。面向其他电脑发布时：

1. 将临时证书替换为可信代码签名证书。
2. 在项目 `发布` 设置中选择 ClickOnce 发布目录或部署地址。
3. 确保目标电脑安装 .NET Framework 4.8 和 VSTO Runtime。
4. 发布生成的安装程序和应用清单，不要只复制 DLL。
5. 在受管环境中通过受信任发布者证书或企业部署策略安装。

正式部署前应分别在 `docx`、`docm`、`dotx`、`dotm` 样本上回归，重点确认 VBA、媒体、页眉页脚、脚注、尾注和批注保持完整。

## 后续扩展

新增工具时保持以下边界：

- OOXML 领域逻辑放入 `WordUtilityAddIn.Core/Services`，不引用 Office Interop。
- Word 文档保存、撤销和生命周期操作放入 `WordUtilityAddIn/Services`。
- Ribbon 回调只做输入校验、调用服务和状态展示。
- 为每个高风险文档修改提供工作副本、修改后校验和明确的恢复路径。

## 项目记录

从本地历史项目迁移，源码和资源按原文件编码保留。本次迁移没有验证编译或运行结果。

开发记录：AI 参与过本项目的开发。

## 公开仓库的本地签名设置

仓库不包含 `WordUtilityAddIn_TemporaryKey.pfx`。首次构建前，在 Visual Studio 的加载项项目属性「签名」页创建本机测试证书，让项目更新对应证书文件名和指纹，再构建。生成的私钥文件只保存在本机，已由 `.gitignore` 排除。
