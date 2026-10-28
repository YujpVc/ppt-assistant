# PPT 助手

DeepSeek Harness 的独立 PPT 助手预设：需求分析、大纲、可编辑 PowerPoint 生成和修改、结构质检、可选的联网配图与渲染审查。安装只插入 `ppt-assistant` 预设，不修改标准模式、模型路由、权限或现有会话。

## 安装

在 DSH 的插件管理页面添加 `https://github.com/YujpVc/ppt-assistant/releases/download/v1.0.0/ppt-assistant-1.0.0.tgz`，然后新建会话并选择 **PPT 助手**。命令行部署使用 `dsh plugin --profile web add https://github.com/YujpVc/ppt-assistant/releases/download/v1.0.0/ppt-assistant-1.0.0.tgz`。

包名为 `@rx01334/ppt-assistant`，当前通过 GitHub Release 分发，尚未发布到 npm。社区市场收录申请见 [PR #10056](https://github.com/awesome-dsh-plugin/awesome-dsh-plugin/pull/10056)，审核合并前无法在市场目录搜索到。

此包保留原有 PPT 助手提示词和工具组合，同时附带 Python 生成引擎和通用主题资源；不复制用户配置、API 密钥、机构模板、机构 Logo、历史会话或生成的文件。提示词沿用原来的引擎发现规则：`deepseek-harness/ppt-assistant`、`~/ppt-assistant` 或安装后的 `node_modules/@rx01334/ppt-assistant`。

## 运行环境

使用 Python 3.10+，在引擎目录运行 `python3 -m pip install -r requirements.txt`。这些依赖不会在插件安装时自动执行。`run_pptfast.sh` 可选使用已安装的 `@liustack/pptfast@0.20.0`，缺少时使用内置 `build_deck.py`；本包不自动下载 pptfast。

渲染需要 LibreOffice 和 `pdftoppm`，扫描文档 OCR 需要 Tesseract。联网取图可通过 `PEXELS_API_KEY` 设置个人密钥，视觉审查需要 `DEEPSEEK_API_KEY` 和可用的视觉模型权限。旧提示词提到的 dsh-office、mindmap 等补充插件需另行安装，本包不自动安装或宣称提供这些插件的工具。

## 使用

```bash
python3 build_deck.py spec.json output.pptx
python3 quality_check.py output.pptx
python3 build_deck.py template-spec.json output.pptx --template template.pptx
python3 quality_check.py output.pptx --template template.pptx
```

内置默认生成器支持原生可编辑文本、表格与图表。模板、Logo 由使用者显式提供，本包提供通用默认模板；旧版 `generate_ppt.py` 可用 `--template` 指定用户模板，不包含本机机构素材。生成质量取决于输入素材和逐页检查；没有运行视觉检查时不得声称已通过。

插件不包含模型服务或账号，遵循会话已有权限。安装需要 DSH `0.2.1-rc.3` 或兼容的 `0.2.x`，已在 Linux Desktop 0.2.1-rc.3 的插件管理入口验证安装；激活验证结果以发行说明为准。

引擎与预设提取自本机 DeepSeek Harness PPT 助手，保留 MIT 许可及上游署名。
