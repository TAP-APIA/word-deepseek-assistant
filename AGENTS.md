# AGENTS.md

## 修改后自动推送与刷新

每次完成对加载项代码的修改并验证后，按以下顺序自动执行：

1. **先推送到 GitHub**：提交并推送仓库 `TAP-APIA/word-deepseek-assistant` 的 main 分支，让 GitHub Pages 部署最新版。
2. **再刷新加载项**：推送后运行刷新脚本，让 Word 加载部署后的最新版：

```powershell
powershell -NoProfile -ExecutionPolicy Bypass -File refresh-addin.ps1 -Auto
```

- 本地工作目录不是 git 仓库，推送使用临时克隆目录，提交前用文件对比确认本地改动已全部同步（允许仅 LF/CRLF 差异，全局 autocrlf=true 会规范化）。
- 提交信息用中文简要描述本次改动，一条提交包含当前批次的全部改动。
- 跳过本地配置：不提交 `.codex/` 目录下的文件。
- 推送后核对远程 HEAD 为最新提交；GitHub Pages 部署通常需几分钟，若立即刷新可能拉到旧版，应等部署完成（或检查部署状态）后再刷新。
- `-Auto` 模式会免交互自动关闭 Word、备份并清除 WEF 缓存、重新打开 Word；这是 `refresh-addin.bat` 的非交互版本（bat 用于手动双击运行，会询问确认）。
- 若用户明确要求手动验证或 Word 中有未保存文档，先提示用户再执行。
