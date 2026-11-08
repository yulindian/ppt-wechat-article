# 课件公众号助手

将教学 PPT/PDF 整理成固定排版的公众号文章：标题、封面、一段教学导语、完整课件预览和 Word 导入稿。可作为 Codex 等智能体的 Skill 使用，也可直接运行 Python 脚本。

开头、结尾两张引导图由使用者替换，其余排版保持固定。随包图片为无品牌示例占位图，不含作者个人账号、二维码或课件产品。

## 效果与示例

仓库自带一个三页示例「学会整理书包」，完整输入位于 `examples/source/`，内容配置位于 `examples/metadata.json`，已生成的交付结果位于 `examples/result/`。

![固定排版示例](examples/article-preview.png)

示例只用于演示流程。文章中的资料清单来自配置中真实存在的文件，不自动添加不存在的发言稿、教案或视频。

## 安装与快速运行

Python 3.10+。在仓库目录打开终端：

```bash
python -m pip install -r requirements.txt
python scripts/build_article.py --source examples/source --metadata examples/metadata.json --output output/demo
```

Windows 中也可将 `python` 换成 `py -3`。每次生成使用新的或空的输出文件夹；已有成品不会被覆盖。

作为 Skill 安装：将整个仓库文件夹命名为 `ppt-wechat-article`，放到自己的 `~/.codex/skills/` 下，再让 Codex 使用 `$ppt-wechat-article`。其他智能体按其 Skill 安装规则放置；后台操作需要该智能体具备浏览器或电脑控制工具。

## 换成自己的课件和引导图

1. 准备最终 PDF、对应 PPTX 和实际配套资料。已有 PDF 时直接使用；只有 PPTX 时可在 Windows PowerPoint 下自动导出 PDF，其他系统请先自行导出。
2. 复制 `examples/metadata.json`，根据实际课件写标题、摘要、单段导语和资料清单。字段说明见 [配置规则](references/metadata-schema.md)。脚本不调用大模型；使用 Skill 时由智能体阅读课件后写配置。
3. 指定自己的两张引导图，PNG/JPEG 均可。也可直接替换 `assets/opening.png` 与 `assets/closing.png`。

```bash
python scripts/build_article.py --source "自己的课件文件夹" --metadata "自己的内容.json" --opening-image "开头图.png" --closing-image "结尾图.png" --output "output/我的文章"
```

多个 PDF 时用 `--pdf "最终版.pdf"` 指定。封面默认裁取第一页；若裁切损伤标题，可先按第一页内容重排，再用 `--cover "封面.png"` 指定约 2.35:1 的封面。

## 固定排版

开头图 → 名片留白 → 空一行 → 可编辑的“课件预览”组件 → 空一行 → 首行缩进的教学导语与加粗资料清单 → 全部课件页图 → “以上是最终PPT的完整预览，共X页” → 名片留白 → 结尾图。

资源清单与“下面按照上课顺序完整展示课件。”位于同一个段落。页数根据实际 PDF 自动计算。Markdown 用于 Obsidian 存储与编辑，Word 导入稿承载固定排版；Obsidian 主题可能改变 Markdown 预览效果。

## 如何进入公众号草稿箱

推荐在已登录的公众号后台选择“文档导入”，上传生成的 `公众号文章导入稿.docx`，再填正式标题、摘要和封面，点击“保存为草稿”。导入可能把文件名作为标题，务必重新填写。保存后刷新或重新打开，核对图片数、页序、字重、缩进与封面裁切。详见 [草稿导入与验收](references/wechat-draft-import.md)。

此路径不使用 AppSecret，但需要账号本身能正常登录、创建和保存草稿。Skill 可在用户要求时通过电脑控制执行这些操作；单独运行生成脚本只会生成本地文件，不会自动登录、上传或发表文章。

附带的 Windows `.cmd` 复制入口仅供辅助使用。一次实际测试中，插件复制的本地图片和嵌入图片在保存时被过滤；同篇 DOCX 导入的图片在保存刷新后保留。因此不把“编辑区能显示图片”当作草稿验收通过。

## 交付文件

| 文件 | 用途 |
| --- | --- |
| 公众号标题.txt | 后台标题栏 |
| 公众号封面图.png | 后台封面栏 |
| 公众号文章导入稿.docx | 后台文档导入 |
| 文章正文.md | Obsidian 编辑与存储 |
| pages/ | 全部课件 JPG，按页序编号 |
| 首尾引导图与发布信息.txt | 正文素材与后台元数据 |
| 封面方形裁切检查.png | 检查方形封面标题安全区 |
| 生成检查.txt | 本地生成检查与待人工验收项 |
| 一键复制到公众号.cmd | Windows 辅助复制入口 |

要放入 Obsidian，将整个生成目录放进自己的库，保持 Markdown 与图片的相对位置。不需要安装公众号插件才能存储这些文件。

## 许可与边界

本项目自编代码、文档和合成示例采用 MIT 许可；第三方依赖遵循各自许可证，PyMuPDF 的许可尤其需要使用者在再分发或商业集成时核对。自己的课件、品牌图与字体需要自行取得使用权。仓库不含账号凭据、登录会话和原始个人课件。

PDF 路径在本机 Windows 测试；macOS/Linux 路径未实机验收。公众号后台可能改变导入规则，实际保存并重新打开后的结果为准。手机预览需要单独检查。
