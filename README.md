# DeepSeek 文档助手

一个运行在 Microsoft 365 **Word 桌面版**侧边栏的 Office 加载项：在编辑文档的同时与 DeepSeek 对话，并把回复直接写入当前文档。Windows 已使用；Mac 可按下文侧载验证，目前尚未经过 Mac 实机测试。

## 主要功能

- **侧边栏对话**：无需切换窗口，在 Word 内与 DeepSeek 连续对话。
- **直连 DeepSeek API**：浏览器直接调用 DeepSeek 官方接口，无需本地服务，电脑上无任何常驻进程。
- **模型实时拉取**：优先从 DeepSeek API 获取可用模型；拉取失败时提供 `deepseek-flash`（V4.1 Flash）和 `deepseek-v4-pro` 作为备用选项，也可填写自定义模型。旧名称 `deepseek-v4-flash` 目前是临时兼容别名。
- **思考过程展示**：模型的思考过程以可折叠区域实时展示，可随时展开或收起。
- **思考强度调节**：可在设置中选择低 / 标准 / 最高三档思考强度（对应 DeepSeek 的 `low` / `high` / `max`），默认标准。
- **流式回复**：回复逐字实时显示，生成过程一目了然。
- **自动应用更改**：回复完成后自动把正文写入文档——有选中内容时替换选中文字，否则插入光标处，无需手动点击。
- **只插入正文**：回复中的正文用代码块包裹，写入文档时只写入代码块内的内容，修改说明不会混入文档。
- **撤销 / 接受 / 回退**：最新一条回复应用后按钮显示“撤销”，点击可撤销并切换为“接受”重新应用；更早的已应用回复显示“回退”，点击并确认后，可撤销该回复及其之后的所有修改、并删除其后的对话记录。撤销、接受与回退均基于文档存档点整文还原，不受光标位置影响，无需使用 Word 的 Ctrl+Z。
- **附带文档上下文**：可勾选附带当前选中内容或全文，让模型基于文档实际内容回答，不会凭空编造。
- **历史对话**：自动保存对话记录，支持在侧边栏展开列表切换、恢复历史对话。
- **快捷发送**：回车直接发送，Shift+回车换行。
- **隐私友好**：API Key 保存在加载项 WebView 的 `localStorage` 中，请求从加载项页面直接发送给 DeepSeek 官方 API，不经过项目自建中转服务。请勿在共享设备上保存 API Key；本地存储会随 Office 缓存清理而丢失。

## 使用环境

- Microsoft 365 桌面版 Word（Windows）；Mac 版提供侧载验证步骤，尚未经过 Mac 实机测试。无需安装 Node.js 或运行任何本地服务。

### 在 Mac 上侧载并验证

1. 下载本仓库的 `manifest.xml`。使用 Microsoft 365 桌面版 Word；不要运行仓库中的 Windows 专用 `.bat` / `.ps1` 安装和刷新脚本。
2. 在 Finder 中按 `Command` + `Shift` + `G`，前往 `~/Library/Containers/com.microsoft.Word/Data/Documents/wef`；若 `wef` 文件夹不存在，先创建它。将 `manifest.xml` 复制到该文件夹。
3. 启动 Word（若已打开则先退出并重新打开），打开任意文档，在“开始”>“加载项”中选择“DeepSeek 文档助手”。不同版本的菜单文字可能略有差异。
4. 在侧边栏设置中填入 DeepSeek API Key，检查模型列表是否加载、对话能否流式返回；在测试文档中分别验证插入、替换选区和撤销。API Key 与对话历史使用该加载项 WebView 的本地存储，Windows 和 Mac 间不会自动同步。

Mac 的侧载路径和步骤参照 [Microsoft 官方文档](https://learn.microsoft.com/en-us/office/dev/add-ins/testing/sideload-an-office-add-in-on-mac)。本项目尚无 Mac 实机测试结果；若出现侧载、网络请求或存储问题，请附上 Word/macOS 版本和具体错误反馈。

## 获取源码

```bash
curl -L -O --ssl-no-revoke https://github.com/TAP-APIA/word-deepseek-assistant/archive/refs/heads/main.zip
```

（Windows 下如提示“证书吊销检查失败”，`--ssl-no-revoke` 可跳过该检查。）
