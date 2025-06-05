import sys
import os
import json
from PyQt5.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout,
    QPushButton, QLineEdit, QLabel, QFileDialog, QTabWidget,
    QFormLayout, QSpinBox, QCheckBox, QColorDialog, QDoubleSpinBox,
    QComboBox, QMessageBox
)
from PyQt5.QtGui import QColor, QFontDatabase, QTextDocument, QTextCursor, QDesktopServices, QFont
from PyQt5.QtWebEngineWidgets import QWebEngineView # Keep for potential fallback or other uses
from PyQt5.QtCore import Qt, QUrl, QSizeF, QMarginsF, QDateTime

from converter import MarkdownToWordConverter

# New: WordPreviewWidget for simulating Word document appearance
class WordPreviewWidget(QWidget):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.web_view = QWebEngineView() # Use QWebEngineView internally for rendering HTML
        
        # Simulate A4 page size (210mm x 297mm) at 96 DPI (standard for web)
        # 1 inch = 2.54 cm = 96 pixels
        # A4 width: 21.0 cm = 21.0 / 2.54 * 96 = 793.7 pixels
        # A4 height: 29.7 cm = 29.7 / 2.54 * 96 = 1122.5 pixels
        self.page_width_px = 794
        self.page_height_px = 1123

        # Default Word margins (top/bottom 2.54cm, left/right 3.17cm)
        # 2.54 cm = 96 pixels
        # 3.17 cm = 120 pixels
        self.margin_top_px = 96
        self.margin_bottom_px = 96
        self.margin_left_px = 120
        self.margin_right_px = 120

        self.init_ui()

    def init_ui(self):
        layout = QVBoxLayout(self)
        layout.setContentsMargins(0, 0, 0, 0) # No margins for the main layout

        # Create a scroll area for the page
        scroll_area = QWidget()
        scroll_layout = QVBoxLayout(scroll_area)
        scroll_layout.setAlignment(Qt.AlignHCenter | Qt.AlignTop) # Center the page horizontally
        scroll_layout.setContentsMargins(0, 0, 0, 0)

        # Page container widget
        self.page_container = QWidget()
        self.page_container.setFixedSize(self.page_width_px, self.page_height_px)
        self.page_container.setStyleSheet(f"""
            background-color: white;
            border: 1px solid #ccc;
            margin: 20px; /* Visual spacing around the page */
        """)
        
        page_content_layout = QVBoxLayout(self.page_container)
        page_content_layout.setContentsMargins(
            self.margin_left_px, self.margin_top_px,
            self.margin_right_px, self.margin_bottom_px
        )
        page_content_layout.addWidget(self.web_view)
        
        scroll_layout.addWidget(self.page_container)
        layout.addWidget(scroll_area) # Add the scroll area to the main layout

    def set_content(self, html_content, styles, base_url=None):
        # Generate dynamic CSS based on the provided styles
        dynamic_css = ""
        for style_name, style_data in styles.items():
            css_rules = []
            
            # Font properties
            if 'font_name' in style_data and style_data['font_name']:
                css_rules.append(f"font-family: '{style_data['font_name']}';")
            if 'font_size' in style_data:
                # Convert pt to px for web rendering (1pt = 1.333px approx)
                font_size_px = style_data['font_size'] * (4/3) # 1pt = 4/3px
                if style_name == 'inline_code' and 'font_size_ratio' in style_data:
                    # For inline code, font size is relative to paragraph font size
                    # This requires knowing the base paragraph font size, which is complex in pure CSS
                    # For simplicity, we'll apply the ratio to a default base or assume it's handled by converter
                    # For now, just use the font_size directly if present, or apply ratio to a default
                    pass # Handled by converter's HTML output, or default CSS below
                else:
                    css_rules.append(f"font-size: {font_size_px}px;")
            if style_data.get('bold'):
                css_rules.append("font-weight: bold;")
            if style_data.get('italic'):
                css_rules.append("font-style: italic;")
            if 'color_rgb' in style_data:
                r, g, b = style_data['color_rgb']
                css_rules.append(f"color: rgb({r}, {g}, {b});")

            # Specific styles
            if style_name == 'paragraph':
                if 'line_spacing' in style_data:
                    css_rules.append(f"line-height: {style_data['line_spacing']};")
                if 'first_line_indent_cm' in style_data:
                    # Convert cm to px (1cm = 37.795px approx at 96 DPI)
                    indent_px = style_data['first_line_indent_cm'] * 37.795
                    css_rules.append(f"text-indent: {indent_px}px;")
            
            if style_name == 'code_block':
                if 'background_color' in style_data:
                    css_rules.append(f"background-color: #{style_data['background_color']};")
                if 'line_spacing' in style_data:
                    css_rules.append(f"line-height: {style_data['line_spacing']};")
                css_rules.append("white-space: pre-wrap;") # Ensure code blocks wrap
                css_rules.append("word-wrap: break-word;")

            if style_name == 'inline_code':
                # Inline code font size ratio is tricky with global CSS.
                # It's better handled by the markdown-to-html conversion itself.
                # For now, we'll just apply the general font settings.
                pass

            # Space before/after for headings and paragraph
            if style_name.startswith('H') or style_name == 'paragraph':
                if 'space_before_pt' in style_data:
                    space_before_px = style_data['space_before_pt'] * (4/3)
                    css_rules.append(f"margin-top: {space_before_px}px;")
                if 'space_after_pt' in style_data:
                    space_after_px = style_data['space_after_pt'] * (4/3)
                    css_rules.append(f"margin-bottom: {space_after_px}px;")

            # Map style names to HTML tags/classes
            selector = ""
            if style_name == 'paragraph':
                selector = "p"
            elif style_name.startswith('H'):
                selector = style_name.lower() # H1 -> h1, H2 -> h2 etc.
            elif style_name == 'code_block':
                selector = "pre" # Markdown code blocks usually map to <pre>
            elif style_name == 'inline_code':
                selector = "code" # Markdown inline code usually maps to <code>
            elif style_name == 'bold':
                selector = "strong, b"
            elif style_name == 'italic':
                selector = "em, i"
            elif style_name == 'link':
                selector = "a"
            elif style_name == 'image':
                selector = "img"
            elif style_name == 'list_item':
                selector = "li"
            elif style_name == 'blockquote':
                selector = "blockquote"
            elif style_name == 'table':
                selector = "table"
            elif style_name == 'table_header':
                selector = "th"
            elif style_name == 'table_cell':
                selector = "td"
            
            if selector and css_rules:
                dynamic_css += f"{selector} {{\n    " + "\n    ".join(css_rules) + "\n}\n"

        # Default body styles (can be overridden by specific styles)
        body_style_data = styles.get('paragraph', {})
        body_font_name = body_style_data.get('font_name', '宋体, SimSun, serif')
        body_font_size_px = body_style_data.get('font_size', 12) * (4/3)
        body_line_height = body_style_data.get('line_spacing', 1.5)
        body_color_rgb = body_style_data.get('color_rgb', (0, 0, 0))
        body_color = f"rgb({body_color_rgb[0]}, {body_color_rgb[1]}, {body_color_rgb[2]})"

        full_html = f"""
        <!DOCTYPE html>
        <html>
        <head>
            <meta charset="utf-8">
            <style>
                body {{
                    font-family: '{body_font_name}';
                    font-size: {body_font_size_px}px;
                    line-height: {body_line_height};
                    color: {body_color};
                    margin: 0; /* Margins handled by page_content_layout */
                    padding: 0;
                }}
                /* Default styles for elements not explicitly configured */
                h1, h2, h3, h4, h5, h6 {{
                    margin-top: 0.5em;
                    margin-bottom: 0.5em;
                }}
                p {{
                    margin-top: 0.2em;
                    margin-bottom: 0.2em;
                }}
                pre {{
                    border: 1px solid #ddd;
                    padding: 10px;
                    font-family: 'Consolas', 'Courier New', monospace;
                    font-size: 0.9em;
                }}
                code {{
                    font-family: 'Consolas', 'Courier New', monospace;
                    font-size: 0.9em;
                    background-color: #eee;
                    padding: 2px 4px;
                    border-radius: 3px;
                }}
                img {{
                    max-width: 100%;
                    height: auto;
                }}
                table {{
                    border-collapse: collapse;
                    width: 100%;
                }}
                th, td {{
                    border: 1px solid #ddd;
                    padding: 8px;
                    text-align: left;
                }}
                th {{
                    background-color: #f2f2f2;
                }}
                /* List specific styles for better rendering */
                ul, ol {{
                    margin: 0;
                    padding-left: 20px; /* Base indentation for lists */
                }}
                ul ul, ol ol, ul ol, ol ul {{
                    padding-left: 20px; /* Additional indentation for nested lists */
                }}
                li {{
                    margin-bottom: 5px;
                    list-style-position: inside;
                }}
                {dynamic_css}
            </style>
        </head>
        <body>
            {html_content}
        </body>
        </html>
        """
        if base_url:
            self.web_view.setHtml(full_html, baseUrl=base_url)
        else:
            self.web_view.setHtml(full_html)

class StyleConfigWidget(QWidget):
    def __init__(self, style_name, style_data, parent=None):
        super().__init__(parent)
        self.style_name = style_name
        self.style_data = style_data
        self.init_ui()

    def init_ui(self):
        layout = QFormLayout()
        layout.setContentsMargins(10, 10, 10, 10) # Add some padding
        layout.setVerticalSpacing(10) # Add vertical spacing between rows

        # Font Name
        self.font_name_input = QComboBox()
        # Populate with system fonts
        fonts = QFontDatabase().families()
        for font in sorted(fonts):
            self.font_name_input.addItem(font)
        self.font_name_input.setEditable(True) # Allow typing custom font names
        self.font_name_input.setCurrentText(self.style_data.get('font_name', ''))
        layout.addRow("字体名称:", self.font_name_input)

        # Font Size
        self.font_size_input = QSpinBox()
        self.font_size_input.setRange(1, 72)
        self.font_size_input.setValue(self.style_data.get('font_size', 12))
        layout.addRow("字号:", self.font_size_input)

        # Bold
        self.bold_checkbox = QCheckBox("粗体")
        self.bold_checkbox.setChecked(self.style_data.get('bold', False))
        layout.addRow(self.bold_checkbox)

        # Italic
        self.italic_checkbox = QCheckBox("斜体")
        self.italic_checkbox.setChecked(self.style_data.get('italic', False))
        layout.addRow(self.italic_checkbox)

        # Color
        self.color_button = QPushButton("选择颜色")
        self.color_display = QLabel()
        self.color_button.clicked.connect(self._pick_color)
        self._set_color_display(self.style_data.get('color_rgb', (0, 0, 0)))
        color_layout = QHBoxLayout()
        color_layout.addWidget(self.color_button)
        color_layout.addWidget(self.color_display)
        layout.addRow("颜色:", color_layout)

        # Specific paragraph/code block settings
        if self.style_name == 'paragraph':
            self.line_spacing_input = QDoubleSpinBox()
            self.line_spacing_input.setRange(0.5, 5.0)
            self.line_spacing_input.setSingleStep(0.1)
            self.line_spacing_input.setValue(self.style_data.get('line_spacing', 1.5))
            layout.addRow("行距:", self.line_spacing_input)

            self.first_line_indent_input = QDoubleSpinBox()
            self.first_line_indent_input.setRange(0.0, 10.0)
            self.first_line_indent_input.setSingleStep(0.1)
            self.first_line_indent_input.setValue(self.style_data.get('first_line_indent_cm', 0.74))
            layout.addRow("首行缩进 (cm):", self.first_line_indent_input)
        
        if self.style_name == 'code_block':
            self.bg_color_button = QPushButton("选择背景颜色")
            self.bg_color_display = QLabel()
            self.bg_color_button.clicked.connect(self._pick_bg_color)
            self._set_bg_color_display(self.style_data.get('background_color', "F0F0F0"))
            bg_color_layout = QHBoxLayout()
            bg_color_layout.addWidget(self.bg_color_button)
            bg_color_layout.addWidget(self.bg_color_display)
            layout.addRow("背景颜色:", bg_color_layout)
            
            self.line_spacing_input = QDoubleSpinBox()
            self.line_spacing_input.setRange(0.5, 5.0)
            self.line_spacing_input.setSingleStep(0.1)
            self.line_spacing_input.setValue(self.style_data.get('line_spacing', 1.0))
            layout.addRow("行距:", self.line_spacing_input)

        if self.style_name == 'inline_code':
            self.font_size_ratio_input = QDoubleSpinBox()
            self.font_size_ratio_input.setRange(0.1, 2.0)
            self.font_size_ratio_input.setSingleStep(0.05)
            self.font_size_ratio_input.setValue(self.style_data.get('font_size_ratio', 0.9))
            layout.addRow("字号比例:", self.font_size_ratio_input)
            # Inline code also has color, which is handled by the general color picker

        # Space before/after for headings and paragraph
        if self.style_name.startswith('H') or self.style_name == 'paragraph':
            self.space_before_input = QSpinBox()
            self.space_before_input.setRange(0, 100)
            self.space_before_input.setValue(self.style_data.get('space_before_pt', 0))
            layout.addRow("段前间距 (pt):", self.space_before_input)

            self.space_after_input = QSpinBox()
            self.space_after_input.setRange(0, 100)
            self.space_after_input.setValue(self.style_data.get('space_after_pt', 0))
            layout.addRow("段后间距 (pt):", self.space_after_input)


        self.setLayout(layout)

    def _pick_color(self):
        initial_color = QColor(*self.style_data.get('color_rgb', (0, 0, 0)))
        color = QColorDialog.getColor(initial_color, self)
        if color.isValid():
            self.style_data['color_rgb'] = (color.red(), color.green(), color.blue())
            self._set_color_display(self.style_data['color_rgb'])

    def _set_color_display(self, rgb_tuple):
        self.color_display.setStyleSheet(f"background-color: rgb({rgb_tuple[0]}, {rgb_tuple[1]}, {rgb_tuple[2]}); border: 1px solid black;")
        self.color_display.setText(f"({rgb_tuple[0]}, {rgb_tuple[1]}, {rgb_tuple[2]})")
        self.color_display.setFixedSize(100, 20)

    def _pick_bg_color(self):
        initial_hex = self.style_data.get('background_color', "F0F0F0")
        initial_color = QColor(f"#{initial_hex}")
        color = QColorDialog.getColor(initial_color, self)
        if color.isValid():
            self.style_data['background_color'] = color.name()[1:].upper() # Get hex without #
            self._set_bg_color_display(self.style_data['background_color'])

    def _set_bg_color_display(self, hex_string):
        self.bg_color_display.setStyleSheet(f"background-color: #{hex_string}; border: 1px solid black;")
        self.bg_color_display.setText(f"#{hex_string}")
        self.bg_color_display.setFixedSize(100, 20)

    def get_current_style_data(self):
        data = {
            'font_name': self.font_name_input.currentText(),
            'font_size': self.font_size_input.value(),
            'bold': self.bold_checkbox.isChecked(),
            'italic': self.italic_checkbox.isChecked(),
            'color_rgb': self.style_data['color_rgb'] # Already updated by color picker
        }
        if self.style_name == 'paragraph':
            data['line_spacing'] = self.line_spacing_input.value()
            data['first_line_indent_cm'] = self.first_line_indent_input.value()
        if self.style_name == 'code_block':
            data['background_color'] = self.style_data['background_color'] # Already updated
            data['line_spacing'] = self.line_spacing_input.value()
        if self.style_name == 'inline_code':
            data['font_size_ratio'] = self.font_size_ratio_input.value()
        
        if self.style_name.startswith('H') or self.style_name == 'paragraph':
            data['space_before_pt'] = self.space_before_input.value()
            data['space_after_pt'] = self.space_after_input.value()

        return data


class MarkdownToWordGUI(QMainWindow):
    def __init__(self):
        super().__init__()
        self.converter = MarkdownToWordConverter()
        self.converter.load_styles() # Load styles on startup
        self.current_styles = self.converter.get_styles()
        self.set_application_font()  # 设置应用程序字体
        self.init_ui()
    
    def set_application_font(self):
        """设置整个应用程序的字体为微软雅黑"""
        font = QFont("Microsoft YaHei", 9)  # 字体名称和大小
        font.setStyleHint(QFont.SansSerif)  # 设置字体类型提示
        QApplication.instance().setFont(font)

    def init_ui(self):
        self.setWindowTitle("Markdown to Word Converter")
        self.setGeometry(100, 100, 1200, 800) # Increased window size

        central_widget = QWidget()
        self.setCentralWidget(central_widget)
        main_layout = QHBoxLayout(central_widget)

        # Left Panel: File Selection and Styles
        left_panel = QVBoxLayout()
        main_layout.addLayout(left_panel, 1) # Take 1/3 of space

        # File Selection
        file_group_box = QVBoxLayout()
        file_group_box.addWidget(QLabel("<h3>文件设置</h3>"))

        md_layout = QHBoxLayout()
        self.md_path_input = QLineEdit()
        self.md_path_input.setPlaceholderText("选择Markdown文件...")
        self.md_path_button = QPushButton("浏览Markdown")
        self.md_path_button.clicked.connect(self._browse_md_file)
        md_layout.addWidget(self.md_path_input)
        md_layout.addWidget(self.md_path_button)
        file_group_box.addLayout(md_layout)

        docx_layout = QHBoxLayout()
        self.docx_path_input = QLineEdit()
        self.docx_path_input.setPlaceholderText("保存为Word文件...")
        self.docx_path_button = QPushButton("保存为Word")
        self.docx_path_button.clicked.connect(self._browse_docx_file)
        docx_layout.addWidget(self.docx_path_input)
        docx_layout.addWidget(self.docx_path_button)
        file_group_box.addLayout(docx_layout)
        left_panel.addLayout(file_group_box)

        # Style Configuration
        style_group_box = QVBoxLayout()
        style_group_box.addWidget(QLabel("<h3>样式配置</h3>"))
        
        self.style_tabs = QTabWidget()
        self.style_widgets = {} # To store references to StyleConfigWidget instances

        for style_name, style_data in self.current_styles.items():
            widget = StyleConfigWidget(style_name, style_data)
            self.style_tabs.addTab(widget, style_name)
            self.style_widgets[style_name] = widget
        
        style_group_box.addWidget(self.style_tabs)

        style_buttons_layout = QHBoxLayout()
        self.save_styles_button = QPushButton("保存当前样式")
        self.save_styles_button.clicked.connect(self._save_styles)
        self.load_preset_button = QPushButton("导入预设")
        self.load_preset_button.clicked.connect(self._import_styles)
        self.save_preset_button = QPushButton("导出预设")
        self.save_preset_button.clicked.connect(self._export_styles)
        self.reset_styles_button = QPushButton("重置为默认")
        self.reset_styles_button.clicked.connect(self._reset_styles)
        
        style_buttons_layout.addWidget(self.save_styles_button)
        style_buttons_layout.addWidget(self.load_preset_button)
        style_buttons_layout.addWidget(self.save_preset_button)
        style_buttons_layout.addWidget(self.reset_styles_button)
        style_group_box.addLayout(style_buttons_layout)

        left_panel.addLayout(style_group_box)
        left_panel.addStretch(1) # Push content to top

        # Right Panel: Actions and Preview
        right_panel = QVBoxLayout()
        main_layout.addLayout(right_panel, 2) # Take 2/3 of space

        # Action Buttons
        action_buttons_layout = QHBoxLayout()
        self.convert_button = QPushButton("转换到Word")
        self.convert_button.clicked.connect(self._convert_markdown)
        self.preview_button = QPushButton("预览HTML")
        self.preview_button.clicked.connect(self._preview_markdown)
        action_buttons_layout.addWidget(self.convert_button)
        action_buttons_layout.addWidget(self.preview_button)
        right_panel.addLayout(action_buttons_layout)

        # Preview Area
        right_panel.addWidget(QLabel("<h3>Word样式预览</h3>"))
        self.word_preview = WordPreviewWidget()
        right_panel.addWidget(self.word_preview)

    def _browse_md_file(self):
        file_path, _ = QFileDialog.getOpenFileName(self, "选择Markdown文件", "", "Markdown Files (*.md *.markdown);;All Files (*)")
        if file_path:
            self.md_path_input.setText(file_path)
            # Suggest default output path
            dir_name, base_name = os.path.split(file_path)
            name_without_ext = os.path.splitext(base_name)[0]
            self.docx_path_input.setText(os.path.join(dir_name, f"{name_without_ext}.docx"))
            self._preview_markdown() # Auto-preview when file is selected

    def _browse_docx_file(self):
        file_path, _ = QFileDialog.getSaveFileName(self, "保存Word文件为", self.docx_path_input.text(), "Word Documents (*.docx);;All Files (*)")
        if file_path:
            if not file_path.lower().endswith('.docx'):
                file_path += '.docx'
            self.docx_path_input.setText(file_path)

    def _update_current_styles_from_gui(self):
        for style_name, widget in self.style_widgets.items():
            self.current_styles[style_name] = widget.get_current_style_data()
        self.converter.set_styles(self.current_styles) # Update converter's styles

    def _save_styles(self):
        self._update_current_styles_from_gui()
        self.converter.save_styles()
        QMessageBox.information(self, "样式保存", "样式已成功保存！")

    def _reset_styles(self):
        reply = QMessageBox.question(self, "重置样式", "确定要将所有样式重置为默认值吗？",
                                     QMessageBox.Yes | QMessageBox.No, QMessageBox.No)
        if reply == QMessageBox.Yes:
            self.converter = MarkdownToWordConverter() # Re-initialize to get default styles
            self.current_styles = self.converter.get_styles()
            # Re-populate style tabs
            self.style_tabs.clear()
            self.style_widgets = {}
            for style_name, style_data in self.current_styles.items():
                widget = StyleConfigWidget(style_name, style_data)
                self.style_tabs.addTab(widget, style_name)
                self.style_widgets[style_name] = widget
            QMessageBox.information(self, "样式重置", "样式已重置为默认值。")

    def _convert_markdown(self):
        md_path = self.md_path_input.text()
        docx_path = self.docx_path_input.text()

        if not md_path:
            QMessageBox.warning(self, "输入错误", "请选择一个Markdown文件。")
            return
        if not docx_path:
            QMessageBox.warning(self, "输入错误", "请指定一个Word输出文件路径。")
            return
        
        self._update_current_styles_from_gui() # Ensure latest styles are used

        try:
            self.converter.markdown_to_docx(md_path, docx_path)
            QMessageBox.information(self, "转换成功", f"'{md_path}' 已成功转换为 '{docx_path}'。")
        except Exception as e:
            QMessageBox.critical(self, "转换失败", f"转换过程中发生错误: {e}")

    def _preview_markdown(self):
        md_path = self.md_path_input.text()
        if not md_path:
            # QMessageBox.warning(self, "输入错误", "请选择一个Markdown文件进行预览。")
            self.word_preview.set_content("") # Clear preview if no file selected
            return
        
        if not os.path.exists(md_path):
            QMessageBox.warning(self, "文件不存在", f"Markdown文件 '{md_path}' 不存在。")
            self.word_preview.set_content("") # Clear preview if file not found
            return

        try:
            with open(md_path, 'r', encoding='utf-8') as f:
                md_content = f.read()
            
            self._update_current_styles_from_gui() # Ensure latest styles are used for preview
            html_preview_content = self.converter.markdown_to_html(md_content)
            self.word_preview.set_content(html_preview_content, self.current_styles, base_url=QUrl.fromLocalFile(os.path.abspath(md_path) + "/"))
        except Exception as e:
            QMessageBox.critical(self, "预览失败", f"生成预览时发生错误: {e}")

    def _import_styles(self):
        file_path, _ = QFileDialog.getOpenFileName(self, "导入样式预设", "presets/", "JSON Files (*.json);;All Files (*)")
        if file_path:
            try:
                with open(file_path, 'r', encoding='utf-8') as f:
                    imported_styles = json.load(f)
                
                # Update current_styles and GUI
                self.current_styles.update(imported_styles) # Merge imported styles
                self.converter.set_styles(self.current_styles) # Update converter

                # Re-populate style tabs with updated data
                self.style_tabs.clear()
                self.style_widgets = {}
                for style_name, style_data in self.current_styles.items():
                    widget = StyleConfigWidget(style_name, style_data)
                    self.style_tabs.addTab(widget, style_name)
                    self.style_widgets[style_name] = widget
                
                QMessageBox.information(self, "导入成功", f"样式预设已从 '{os.path.basename(file_path)}' 导入。")
                self._preview_markdown() # Refresh preview after importing styles

            except Exception as e:
                QMessageBox.critical(self, "导入失败", f"导入样式预设时发生错误: {e}")

    def _export_styles(self):
        # Ensure presets directory exists
        presets_dir = "presets"
        os.makedirs(presets_dir, exist_ok=True)

        # Suggest a filename based on current timestamp
        timestamp = QDateTime.currentDateTime().toString("yyyyMMdd_hhmmss")
        default_filename = f"styles_preset_{timestamp}.json"
        default_path = os.path.join(presets_dir, default_filename)

        file_path, _ = QFileDialog.getSaveFileName(self, "导出样式预设", default_path, "JSON Files (*.json);;All Files (*)")
        if file_path:
            try:
                self._update_current_styles_from_gui() # Get latest styles from GUI
                with open(file_path, 'w', encoding='utf-8') as f:
                    json.dump(self.current_styles, f, indent=4, ensure_ascii=False)
                QMessageBox.information(self, "导出成功", f"当前样式已成功导出到 '{os.path.basename(file_path)}'。")
            except Exception as e:
                QMessageBox.critical(self, "导出失败", f"导出样式预设时发生错误: {e}")


if __name__ == '__main__':
    app = QApplication(sys.argv)
    # Ensure QApplication is initialized before QFontDatabase
    # Check if PyQtWebEngine is available
    try:
        from PyQt5.QtWebEngineWidgets import QWebEngineView
    except ImportError:
        QMessageBox.critical(None, "错误", "PyQtWebEngine模块未找到。请确保已安装：\npip install PyQtWebEngine")
        sys.exit(1)

    main_window = MarkdownToWordGUI()
    main_window.show()
    sys.exit(app.exec_())
