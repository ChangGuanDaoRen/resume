# 个人简历 · Resume

以 **Word 文档（`简历.docx`）为唯一内容源** 的个人简历主页，发布于 GitHub Pages。

访问地址：<https://changguandaoren.github.io/resume/>

## 如何更新简历（日常只需这一步）

**上传（替换）新的 `简历.docx` 并提交** —— 完事。

GitHub Actions 会自动：

1. 读取新的 `简历.docx`；
2. 重新生成 `index.html`，与 Word 内容**一比一对应**：表格结构、合并单元格、
   底纹颜色、加粗、字号、字色、对齐方式全部跟随 Word 文件；
3. 把新的 `index.html` 提交回仓库，GitHub Pages 随即自动更新。

全程无需手工编辑任何 HTML。

> 手动触发：仓库 **Actions** 页 → 「Build resume page」→ **Run workflow**。

## 页面右下角的「保存为 PDF」按钮

点击按钮**直接下载 `resume.pdf`**（浏览器立即开始下载，文件名
`个人简历-赵浩.pdf`），不再调起浏览器打印界面。这份 PDF 由 weasyprint
从页面精确渲染生成，因此：

- 颜色 100% 保留（模块蓝底、标签灰底、字色）；
- **没有**浏览器打印页眉页脚（网址、日期、标题都不会出现）；
- A4 纸张与页边距与页面设置完全一致。

`resume.pdf` 由 `build_html.py` 在每次构建时自动重新生成，与
`index.html` 始终同源同步。

## 本地重新生成（可选）

```bash
pip install python-docx weasyprint
python build_html.py
```

脚本自带一致性校验：Word 中每一段文字都必须按原顺序出现在 HTML 中，
否则以非零退出码报错。装了 weasyprint 时还会同时重新生成 `resume.pdf`
（Windows 需先安装 pango，如 `pacman -S mingw-w64-ucrt-x86_64-pango`）。

## 目录结构

```
简历.docx                            # 唯一内容源（只编辑这个文件）
index.html                           # 自动生成的简历页面（勿手工编辑）
resume.pdf                           # 自动生成的下载版 PDF（勿手工编辑）
build_html.py                        # Word → HTML/PDF 生成脚本
.github/workflows/build-resume.yml   # 自动构建工作流
```

## 一次性配置（GitHub Pages）

仓库 **Settings → Pages → Build and deployment**：

- Source：**Deploy from a branch**
- Branch：**main** ／ 目录 **(root)**
- Save

之后每次更新 `简历.docx`，页面都会自动更新。
