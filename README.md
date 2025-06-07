# md2word

`md2word` 是一个用户友好的桌面应用程序，旨在简化 Markdown 文件到 Microsoft Word (`.docx`) 格式的转换过程。它提供了一个直观的图形用户界面 (GUI)，让用户可以轻松地选择 Markdown 文件并将其转换为高质量的 Word 文档，同时保留原始 Markdown 的格式和结构。

## 特性

*   **直观的 GUI**：通过简单的点击操作即可完成文件转换。
*   **Markdown 到 Word 转换**：将 `.md` 文件转换为 `.docx` 格式。
*   **保留格式**：在转换过程中尽可能保留 Markdown 的标题、列表、代码块、链接、图片等格式。
*   **跨平台**：基于 PyQt5 开发，理论上支持 Windows、macOS 和 Linux。

## 安装

在运行此应用程序之前，您需要安装所有必要的依赖项。建议使用 `pip` 和 `requirements.txt` 文件进行安装。

1.  **克隆仓库** (如果尚未克隆):
    ```bash
    git clone https://github.com/zzzj131/md2word.git
    cd md2word
    ```


2.  **安装依赖**:
    ```bash
    pip install -r requirements.txt
    ```

## 使用方法

安装完所有依赖后，您可以通过运行 `main.py` 文件来启动应用程序。

```bash
python main.py
```

启动应用程序后：
1.  点击“选择 Markdown 文件”按钮，选择您要转换的 `.md` 文件。
2.  点击“选择输出路径”按钮，选择转换后的 `.docx` 文件的保存位置和文件名。
3.  点击“开始转换”按钮，程序将开始转换过程。
4.  转换完成后，您将在指定的输出路径找到生成的 Word 文档。

## 依赖

本项目依赖以下 Python 库：

*   `beautifulsoup4`
*   `PyQt5`
*   `PyQt5-Qt5`
*   `PyQt5_sip`
*   `PyQtWebEngine`
*   `PyQtWebEngine-Qt5`
*   `python-docx`
*   `soupsieve`
*   `typing_extensions`

## 贡献

欢迎贡献！如果您有任何改进建议或发现 bug，请随时提交 issue 或 pull request。

## 许可证

本项目采用 [LICENSE](LICENSE) 许可证。
