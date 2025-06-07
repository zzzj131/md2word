import sys
import os
import json
from PyQt5.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout,
    QPushButton, QLineEdit, QLabel, QFileDialog, QTabWidget,
    QFormLayout, QSpinBox, QCheckBox, QColorDialog, QDoubleSpinBox,
    QComboBox, QMessageBox, QScrollArea
)
from PyQt5.QtGui import QColor, QFontDatabase, QTextDocument, QTextCursor, QDesktopServices, QFont
from PyQt5.QtWebEngineWidgets import QWebEngineView, QWebEngineSettings, QWebEngineScript, QWebEnginePage, QWebEngineProfile
from PyQt5.QtCore import Qt, QUrl, QSizeF, QMarginsF, QDateTime

from converter import MarkdownToWordConverter

# New: WordPreviewWidget for simulating Word document appearance
class WordPreviewWidget(QWidget):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.web_view = QWebEngineView() # Use QWebEngineView internally for rendering HTML
        self.web_view.settings().setAttribute(QWebEngineSettings.LocalContentCanAccessFileUrls, True)
        self.web_view.settings().setAttribute(QWebEngineSettings.AutoLoadImages, True) # 确保自动加载图片
        
        self.web_view.loadFinished.connect(self._on_load_finished)
        
        self.page_width_px = 794  # A4 width in pixels at 96 DPI (21cm)
        self.page_height_px = 1123 # A4 height in pixels at 96 DPI (29.7cm)

        self.margin_top_px = 96    # 1 inch margin in pixels at 96 DPI
        self.margin_bottom_px = 96
        self.margin_left_px = 120  # 1.25 inch margin
        self.margin_right_px = 120

        self.init_ui()

    def init_ui(self):
        main_layout = QVBoxLayout(self)
        main_layout.setContentsMargins(0, 0, 0, 0)

        # ScrollArea to contain the page, allowing scrolling if content is too large
        scroll_area_container = QScrollArea()
        scroll_area_container.setWidgetResizable(True)
        scroll_area_container.setHorizontalScrollBarPolicy(Qt.ScrollBarAsNeeded)
        scroll_area_container.setVerticalScrollBarPolicy(Qt.ScrollBarAsNeeded)
        scroll_area_container.setStyleSheet("QScrollArea { background-color: #E0E0E0; border: none; }") # Match background

        # Widget to center the page within the scroll area
        centering_widget = QWidget()
        centering_layout = QVBoxLayout(centering_widget)
        centering_layout.setAlignment(Qt.AlignHCenter | Qt.AlignTop)
        centering_layout.setContentsMargins(20, 20, 20, 20) # Visual padding around the page

        self.page_container = QWidget()
        self.page_container.setFixedSize(self.page_width_px, self.page_height_px)
        self.page_container.setStyleSheet(f"""
            QWidget {{
                background-color: white;
                border: 1px solid #CCCCCC;
            }}
        """)
        
        page_content_layout = QVBoxLayout(self.page_container)
        page_content_layout.setContentsMargins(
            self.margin_left_px, self.margin_top_px,
            self.margin_right_px, self.margin_bottom_px
        )
        page_content_layout.addWidget(self.web_view)
        
        centering_layout.addWidget(self.page_container)
        scroll_area_container.setWidget(centering_widget)
        main_layout.addWidget(scroll_area_container)


    def set_content(self, html_content, styles, base_url=None): # Added base_url parameter
        dynamic_css = ""
        for style_name, style_data in styles.items():
            css_rules = []
            
            if 'font_name' in style_data and style_data['font_name']:
                css_rules.append(f"font-family: '{style_data['font_name']}';")
            if 'font_size' in style_data:
                font_size_px = style_data['font_size'] * (4/3) 
                if style_name == 'inline_code' and 'font_size_ratio' in style_data:
                    # For inline code, font size is relative to paragraph font size
                    # This needs context, so we'll apply a general size here or expect it from converter
                    # For now, let converter handle specific inline code size within its generated HTML
                    # or apply ratio to default paragraph size if available
                    para_font_size = styles.get('paragraph', {}).get('font_size', 12) * (4/3)
                    css_rules.append(f"font-size: {para_font_size * style_data['font_size_ratio']}px;")
                else:
                    css_rules.append(f"font-size: {font_size_px}px;")

            if style_data.get('bold'):
                css_rules.append("font-weight: bold;")
            if style_data.get('italic'):
                css_rules.append("font-style: italic;")
            if 'color_rgb' in style_data:
                r, g, b = style_data['color_rgb']
                css_rules.append(f"color: rgb({r}, {g}, {b});")

            if style_name == 'paragraph':
                if 'line_spacing' in style_data:
                    css_rules.append(f"line-height: {style_data['line_spacing']};")
                if 'first_line_indent_cm' in style_data:
                    indent_px = style_data['first_line_indent_cm'] * 37.7952 # 1cm = 37.7952px at 96DPI
                    css_rules.append(f"text-indent: {indent_px}px;")
            
            if style_name == 'code_block':
                if 'background_color' in style_data:
                    css_rules.append(f"background-color: #{style_data['background_color']};")
                if 'line_spacing' in style_data:
                    css_rules.append(f"line-height: {style_data['line_spacing']};")
                css_rules.append("white-space: pre-wrap;") 
                css_rules.append("word-wrap: break-word;")
                css_rules.append("padding: 10px;"); # Add padding for code blocks
                css_rules.append("border: 1px solid #ddd;"); # Add border for code blocks


            if style_name.startswith('H') or style_name == 'paragraph':
                if 'space_before_pt' in style_data:
                    space_before_px = style_data['space_before_pt'] * (4/3)
                    css_rules.append(f"margin-top: {space_before_px}px;")
                if 'space_after_pt' in style_data:
                    space_after_px = style_data['space_after_pt'] * (4/3)
                    css_rules.append(f"margin-bottom: {space_after_px}px;")

            selector = ""
            if style_name == 'paragraph':
                selector = "p"
            elif style_name.startswith('H'):
                selector = style_name.lower() 
            elif style_name == 'code_block':
                selector = "pre" 
            elif style_name == 'inline_code':
                selector = "code" 
            elif style_name == 'bold': # These are less common as top-level style keys but can be for defaults
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

        body_style_data = styles.get('paragraph', {}) # Use paragraph as base for body
        body_font_name = body_style_data.get('font_name', '宋体, SimSun, serif')
        body_font_size_px = body_style_data.get('font_size', 12) * (4/3)
        body_line_height = body_style_data.get('line_spacing', 1.5)
        body_color_rgb = body_style_data.get('color_rgb', (0, 0, 0))
        body_color = f"rgb({body_color_rgb[0]}, {body_color_rgb[1]}, {body_color_rgb[2]})"

        # Default styles for elements not explicitly configured via style_tabs but common in Markdown
        default_styles = f"""
            h1, h2, h3, h4, h5, h6 {{
                /* Default heading margins if not overridden by specific H styles */
                margin-top: {1.0 * body_font_size_px}px; /* e.g., 1em */
                margin-bottom: {0.5 * body_font_size_px}px; /* e.g., 0.5em */
            }}
            p {{
                /* Default paragraph margins if not overridden */
                margin-top: {0.2 * body_font_size_px}px;
                margin-bottom: {0.2 * body_font_size_px}px;
            }}
            pre {{ /* Default for code blocks if 'code_block' style is minimal */
                border: 1px solid #ddd;
                padding: 10px;
                font-family: 'Consolas', 'Courier New', monospace; /* Monospace default */
                font-size: {0.9 * body_font_size_px}px; /* 0.9em of body font */
                background-color: #f9f9f9; /* Light gray background */
            }}
            code {{ /* Default for inline code if 'inline_code' style is minimal */
                font-family: 'Consolas', 'Courier New', monospace;
                font-size: {0.9 * body_font_size_px}px; /* 0.9em of body font, can be overridden by specific inline_code style */
                background-color: #f0f0f0;
                padding: 2px 4px;
                border-radius: 3px;
            }}
            img {{
                max-width: 100%;
                height: auto;
                display: block; /* Helps with margins if needed */
                margin-top: 5px;
                margin-bottom: 5px;
            }}
            table {{
                border-collapse: collapse;
                width: 100%;
                margin-top: 10px;
                margin-bottom: 10px;
            }}
            th, td {{
                border: 1px solid #ddd;
                padding: 8px;
                text-align: left;
            }}
            th {{
                background-color: #f2f2f2;
                font-weight: bold;
            }}
            ul, ol {{
                margin-top: 5px;
                margin-bottom: 5px;
                padding-left: 40px; /* Standard indentation for lists */
            }}
            ul ul, ol ol, ul ol, ol ul {{
                padding-left: 20px; /* Additional indentation for nested lists */
            }}
            li {{
                margin-bottom: 5px;
                /* list-style-position: inside; /* Consider 'outside' for better alignment with wrapped text */
            }}
            blockquote {{
                border-left: 4px solid #ccc;
                margin-left: 0;
                padding-left: 15px;
                color: #555;
            }}
            hr {{
                border: 0;
                border-top: 1px solid #ccc;
                margin: 20px 0;
            }}
        """

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
                    margin: 0; 
                    padding: 0;
                    word-wrap: break-word; /* Ensure long words wrap */
                }}
                {default_styles}
                {dynamic_css}
            </style>
        </head>
        <body>
            {html_content}
        </body>
        </html>
        """
        if base_url and base_url.isValid():
            # print(f"WordPreviewWidget: Setting HTML with baseUrl: {base_url.toString()}")
            self.web_view.setHtml(full_html, baseUrl=base_url)
        else:
            # print("WordPreviewWidget: Setting HTML without explicit baseUrl.")
            self.web_view.setHtml(full_html)

    def _on_load_finished(self, ok):
        if ok:
            # print("WebEngineView: 页面加载完成。")
            pass
        else:
            print("WebEngineView: 页面加载失败。 Check for resource loading errors in console if possible.")


class StyleConfigWidget(QWidget):
    def __init__(self, style_name, style_data, parent=None):
        super().__init__(parent)
        self.style_name = style_name
        # Ensure style_data is a mutable copy for this widget instance
        self.style_data = style_data.copy() if style_data else {}
        self.init_ui()

    def init_ui(self):
        layout = QFormLayout()
        layout.setContentsMargins(10, 10, 10, 10) 
        layout.setVerticalSpacing(10) 

        self.font_name_input = QComboBox()
        fonts = QFontDatabase().families()
        for font in sorted(fonts):
            self.font_name_input.addItem(font)
        self.font_name_input.setEditable(True) 
        self.font_name_input.setCurrentText(self.style_data.get('font_name', ''))
        layout.addRow("字体名称:", self.font_name_input)

        self.font_size_input = QSpinBox()
        self.font_size_input.setRange(1, 72)
        self.font_size_input.setValue(self.style_data.get('font_size', 12))
        layout.addRow("字号 (pt):", self.font_size_input)

        self.bold_checkbox = QCheckBox("粗体")
        self.bold_checkbox.setChecked(self.style_data.get('bold', False))
        layout.addRow(self.bold_checkbox)

        self.italic_checkbox = QCheckBox("斜体")
        self.italic_checkbox.setChecked(self.style_data.get('italic', False))
        layout.addRow(self.italic_checkbox)

        self.color_button = QPushButton("选择颜色")
        self.color_display = QLabel()
        self.color_button.clicked.connect(self._pick_color)
        self._set_color_display(self.style_data.get('color_rgb', (0, 0, 0)))
        color_layout = QHBoxLayout()
        color_layout.addWidget(self.color_button)
        color_layout.addWidget(self.color_display)
        layout.addRow("颜色:", color_layout)

        if self.style_name == 'paragraph':
            self.line_spacing_input = QDoubleSpinBox()
            self.line_spacing_input.setRange(0.5, 5.0)
            self.line_spacing_input.setSingleStep(0.1)
            self.line_spacing_input.setValue(self.style_data.get('line_spacing', 1.5))
            layout.addRow("行距:", self.line_spacing_input)

            self.first_line_indent_input = QDoubleSpinBox()
            self.first_line_indent_input.setRange(0.0, 10.0) # cm
            self.first_line_indent_input.setSingleStep(0.1)
            self.first_line_indent_input.setSuffix(" cm")
            self.first_line_indent_input.setValue(self.style_data.get('first_line_indent_cm', 0.0)) # Default to 0 for no indent
            layout.addRow("首行缩进:", self.first_line_indent_input)
        
        if self.style_name == 'code_block':
            self.bg_color_button = QPushButton("选择背景颜色")
            self.bg_color_display = QLabel()
            self.bg_color_button.clicked.connect(self._pick_bg_color)
            self._set_bg_color_display(self.style_data.get('background_color', "F0F0F0"))
            bg_color_layout = QHBoxLayout()
            bg_color_layout.addWidget(self.bg_color_button)
            bg_color_layout.addWidget(self.bg_color_display)
            layout.addRow("背景颜色:", bg_color_layout)
            
            self.line_spacing_input = QDoubleSpinBox() # Specific line spacing for code blocks
            self.line_spacing_input.setRange(0.5, 5.0)
            self.line_spacing_input.setSingleStep(0.1)
            self.line_spacing_input.setValue(self.style_data.get('line_spacing', 1.0))
            layout.addRow("行距:", self.line_spacing_input)

        if self.style_name == 'inline_code':
            self.font_size_ratio_input = QDoubleSpinBox()
            self.font_size_ratio_input.setRange(0.1, 2.0)
            self.font_size_ratio_input.setSingleStep(0.05)
            self.font_size_ratio_input.setToolTip("相对于段落字号的比例")
            self.font_size_ratio_input.setValue(self.style_data.get('font_size_ratio', 0.9))
            layout.addRow("字号比例:", self.font_size_ratio_input)
            # Note: Inline code font size is handled by WordPreviewWidget's CSS directly based on this ratio and paragraph font.
            # The 'font_size' field for 'inline_code' in style_data is not directly used by QSpinBox here.

        if self.style_name.startswith('H') or self.style_name == 'paragraph':
            self.space_before_input = QSpinBox()
            self.space_before_input.setRange(0, 100)
            self.space_before_input.setSuffix(" pt")
            self.space_before_input.setValue(self.style_data.get('space_before_pt', 0))
            layout.addRow("段前间距:", self.space_before_input)

            self.space_after_input = QSpinBox()
            self.space_after_input.setRange(0, 100)
            self.space_after_input.setSuffix(" pt")
            self.space_after_input.setValue(self.style_data.get('space_after_pt', 6)) # Default 6pt for paragraph after
            layout.addRow("段后间距:", self.space_after_input)

        self.setLayout(layout)

    def _pick_color(self):
        # Ensure color_rgb exists in style_data before picking
        current_rgb = self.style_data.get('color_rgb', (0, 0, 0))
        if isinstance(current_rgb, list): # Convert if it's a list from JSON
            current_rgb = tuple(current_rgb)

        initial_color = QColor(*current_rgb)
        color = QColorDialog.getColor(initial_color, self)
        if color.isValid():
            self.style_data['color_rgb'] = (color.red(), color.green(), color.blue())
            self._set_color_display(self.style_data['color_rgb'])

    def _set_color_display(self, rgb_tuple):
        if not (isinstance(rgb_tuple, (tuple, list)) and len(rgb_tuple) == 3):
            rgb_tuple = (0,0,0) # Default to black if invalid
        self.color_display.setStyleSheet(f"background-color: rgb({rgb_tuple[0]}, {rgb_tuple[1]}, {rgb_tuple[2]}); border: 1px solid black;")
        self.color_display.setText(f" R:{rgb_tuple[0]} G:{rgb_tuple[1]} B:{rgb_tuple[2]} ")
        self.color_display.setFixedSize(120, 20)


    def _pick_bg_color(self):
        initial_hex = self.style_data.get('background_color', "F0F0F0")
        initial_color = QColor(f"#{initial_hex}")
        color = QColorDialog.getColor(initial_color, self)
        if color.isValid():
            self.style_data['background_color'] = color.name()[1:].upper() 
            self._set_bg_color_display(self.style_data['background_color'])

    def _set_bg_color_display(self, hex_string):
        self.bg_color_display.setStyleSheet(f"background-color: #{hex_string}; border: 1px solid black;")
        self.bg_color_display.setText(f" #{hex_string} ")
        self.bg_color_display.setFixedSize(100, 20)

    def get_current_style_data(self):
        data = {
            'font_name': self.font_name_input.currentText(),
            'font_size': self.font_size_input.value(),
            'bold': self.bold_checkbox.isChecked(),
            'italic': self.italic_checkbox.isChecked(),
            'color_rgb': self.style_data.get('color_rgb', (0,0,0)) # Get from internal, updated by picker
        }
        if 'background_color' in self.style_data : # For code_block, from internal
             data['background_color'] = self.style_data['background_color']

        if self.style_name == 'paragraph':
            data['line_spacing'] = self.line_spacing_input.value()
            data['first_line_indent_cm'] = self.first_line_indent_input.value()
        if self.style_name == 'code_block':
            # background_color is already in self.style_data and added above
            data['line_spacing'] = self.line_spacing_input.value()
        if self.style_name == 'inline_code':
            data['font_size_ratio'] = self.font_size_ratio_input.value()
            # font_size for inline_code is not set here, it's relative.
            # color_rgb is already handled.
        
        if self.style_name.startswith('H') or self.style_name == 'paragraph':
            data['space_before_pt'] = self.space_before_input.value()
            data['space_after_pt'] = self.space_after_input.value()

        return data


class MarkdownToWordGUI(QMainWindow):
    def __init__(self):
        super().__init__()
        self.converter = MarkdownToWordConverter()
        self.converter.load_styles() # Load styles from file or use defaults
        # Make a deep copy of styles for GUI to modify independently until saved
        self.current_styles = {k: v.copy() for k, v in self.converter.get_styles().items()}
        self.set_application_font()
        self.init_ui()
    
    def set_application_font(self):
        font = QFont("Microsoft YaHei", 9) 
        font.setStyleHint(QFont.SansSerif) 
        QApplication.instance().setFont(font)

    def init_ui(self):
        self.setWindowTitle("Markdown to Word Converter")
        self.setGeometry(100, 100, 1300, 850) # Slightly larger window

        central_widget = QWidget()
        self.setCentralWidget(central_widget)
        main_layout = QHBoxLayout(central_widget)

        left_panel_widget = QWidget()
        left_panel_layout = QVBoxLayout(left_panel_widget)
        main_layout.addWidget(left_panel_widget, 1) 

        # File Selection GroupBox
        file_selection_group = QWidget()
        file_group_box = QVBoxLayout(file_selection_group)
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
        left_panel_layout.addWidget(file_selection_group)

        # Style Configuration GroupBox
        style_config_group = QWidget()
        style_group_box = QVBoxLayout(style_config_group) # Changed from left_panel to style_group_box
        style_group_box.addWidget(QLabel("<h3>样式配置</h3>"))
        
        self.style_tabs = QTabWidget()
        self.style_widgets = {} 

        # Populate style tabs using self.current_styles
        for style_name, style_data in self.current_styles.items():
            widget = StyleConfigWidget(style_name, style_data) # Pass a copy of style_data
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

        left_panel_layout.addWidget(style_config_group)
        left_panel_layout.addStretch(1) 

        # Right Panel: Actions and Preview
        right_panel_widget = QWidget()
        right_panel_layout = QVBoxLayout(right_panel_widget)
        main_layout.addWidget(right_panel_widget, 2)

        action_buttons_layout = QHBoxLayout()
        self.convert_button = QPushButton("转换到Word")
        self.convert_button.clicked.connect(self._convert_markdown)
        self.preview_button = QPushButton("预览HTML") # Changed from "预览Word"
        self.preview_button.clicked.connect(self._preview_markdown) # Connects to preview HTML
        action_buttons_layout.addWidget(self.convert_button)
        action_buttons_layout.addWidget(self.preview_button)
        right_panel_layout.addLayout(action_buttons_layout)

        right_panel_layout.addWidget(QLabel("<h3>样式预览</h3>")) # Changed from "Word样式预览"
        self.word_preview = WordPreviewWidget()
        right_panel_layout.addWidget(self.word_preview)

    def _browse_md_file(self):
        file_path, _ = QFileDialog.getOpenFileName(self, "选择Markdown文件", "", "Markdown Files (*.md *.markdown);;All Files (*)")
        if file_path:
            self.md_path_input.setText(file_path)
            dir_name, base_name = os.path.split(file_path)
            name_without_ext = os.path.splitext(base_name)[0]
            suggested_docx_path = os.path.join(dir_name, f"{name_without_ext}.docx")
            # Only set docx_path_input if it's empty or user hasn't manually set it to something different
            if not self.docx_path_input.text() or \
               self.docx_path_input.text() == os.path.join(os.path.dirname(self.docx_path_input.text()), f"{os.path.splitext(os.path.basename(self.md_path_input.text()))[0]}.docx"):
                self.docx_path_input.setText(suggested_docx_path)
            self._preview_markdown() 

    def _browse_docx_file(self):
        current_path = self.docx_path_input.text()
        if not current_path and self.md_path_input.text(): # If empty, suggest based on MD file
            dir_name, base_name = os.path.split(self.md_path_input.text())
            name_without_ext = os.path.splitext(base_name)[0]
            current_path = os.path.join(dir_name, f"{name_without_ext}.docx")

        file_path, _ = QFileDialog.getSaveFileName(self, "保存Word文件为", current_path, "Word Documents (*.docx);;All Files (*)")
        if file_path:
            if not file_path.lower().endswith('.docx'):
                file_path += '.docx'
            self.docx_path_input.setText(file_path)

    def _update_current_styles_from_gui(self):
        """Updates self.current_styles from the GUI widgets."""
        for style_name, widget in self.style_widgets.items():
            self.current_styles[style_name] = widget.get_current_style_data()
        # self.converter.set_styles(self.current_styles) # Update converter's styles only when converting/saving

    def _save_styles(self):
        """Saves the current styles (from GUI) to the converter and then to file."""
        self._update_current_styles_from_gui()
        self.converter.set_styles(self.current_styles.copy()) # Pass a copy to converter
        self.converter.save_styles() # This saves what's in converter.styles_config
        QMessageBox.information(self, "样式保存", "样式已成功保存到默认配置文件！")

    def _reset_styles(self):
        reply = QMessageBox.question(self, "重置样式", "确定要将所有样式重置为默认值吗？\n这将从头加载默认样式。",
                                     QMessageBox.Yes | QMessageBox.No, QMessageBox.No)
        if reply == QMessageBox.Yes:
            # Re-initialize a temporary converter to get fresh default styles
            temp_converter = MarkdownToWordConverter() # Gets default styles
            default_styles = temp_converter.get_styles()
            
            # Update self.current_styles with deep copies of default styles
            self.current_styles = {k: v.copy() for k, v in default_styles.items()}
            
            # Re-populate style tabs with these new default styles
            self.style_tabs.clear()
            self.style_widgets = {}
            for style_name, style_data in self.current_styles.items():
                widget = StyleConfigWidget(style_name, style_data) # Pass the copy
                self.style_tabs.addTab(widget, style_name)
                self.style_widgets[style_name] = widget
            
            # Also update the main converter instance if desired, or just rely on _update_current_styles_from_gui before conversion
            # self.converter.set_styles(self.current_styles.copy()) 
            
            QMessageBox.information(self, "样式重置", "样式已重置为默认值。")
            self._preview_markdown() # Refresh preview with new default styles

    def _convert_markdown(self):
        md_path = self.md_path_input.text()
        docx_path = self.docx_path_input.text()

        if not md_path or not os.path.exists(md_path):
            QMessageBox.warning(self, "输入错误", "请选择一个有效的Markdown文件。")
            return
        if not docx_path:
            QMessageBox.warning(self, "输入错误", "请指定一个Word输出文件路径。")
            return
        
        self._update_current_styles_from_gui() 
        self.converter.set_styles(self.current_styles.copy()) # Ensure converter has the latest styles

        try:
            self.converter.markdown_to_docx(md_path, docx_path)
            QMessageBox.information(self, "转换成功", f"'{os.path.basename(md_path)}' 已成功转换为 '{os.path.basename(docx_path)}'。")
        except Exception as e:
            QMessageBox.critical(self, "转换失败", f"转换过程中发生错误: {e}\n查看控制台获取详细信息。")
            import traceback
            traceback.print_exc()


    def _preview_markdown(self):
        md_path = self.md_path_input.text()
        if not md_path:
            self.word_preview.set_content("<!-- 请选择一个Markdown文件进行预览 -->", self.current_styles.copy()) 
            return
        
        if not os.path.exists(md_path):
            QMessageBox.warning(self, "文件不存在", f"Markdown文件 '{md_path}' 不存在。")
            self.word_preview.set_content(f"<!-- Markdown文件未找到: {md_path} -->", self.current_styles.copy())
            return

        try:
            with open(md_path, 'r', encoding='utf-8') as f:
                md_content = f.read()
            
            self._update_current_styles_from_gui() # Ensure self.current_styles is up-to-date from UI
            
            md_dir = os.path.dirname(os.path.abspath(md_path))
            
            # For QUrl.fromLocalFile, directory paths should ideally end with a separator.
            md_dir_for_url = md_dir
            if not md_dir_for_url.endswith(os.path.sep):
                md_dir_for_url += os.path.sep
            
            base_url_for_preview = QUrl.fromLocalFile(md_dir_for_url)
            if not base_url_for_preview.isValid():
                print(f"警告: 生成的baseUrl无效: {md_dir_for_url}")
            
            # The converter needs the md_dir to resolve relative paths in Markdown for image processing.
            # The converter also needs the current styles for its HTML generation logic.
            temp_converter_for_html = MarkdownToWordConverter(styles_config=self.current_styles.copy())
            html_preview_content = temp_converter_for_html.markdown_to_html(md_content, md_dir)
            
            # Pass a copy of current_styles to set_content
            self.word_preview.set_content(html_preview_content, self.current_styles.copy(), base_url=base_url_for_preview)
        except Exception as e:
            QMessageBox.critical(self, "预览失败", f"生成预览时发生错误: {e}\n查看控制台获取详细信息。")
            import traceback
            traceback.print_exc()


    def _import_styles(self):
        # Ensure presets directory exists or default to current directory
        presets_dir = "presets"
        if not os.path.exists(presets_dir):
            presets_dir = "" # Default to current dir if presets doesn't exist

        file_path, _ = QFileDialog.getOpenFileName(self, "导入样式预设", presets_dir, "JSON Files (*.json);;All Files (*)")
        if file_path:
            try:
                with open(file_path, 'r', encoding='utf-8') as f:
                    imported_styles_raw = json.load(f)
                
                # Create a new current_styles dictionary, starting with defaults and updating with imported
                temp_converter = MarkdownToWordConverter() # For default structure
                default_styles_template = temp_converter.get_styles()
                
                new_styles_from_import = {k: v.copy() for k, v in default_styles_template.items()}

                for style_name, imported_data in imported_styles_raw.items():
                    if style_name in new_styles_from_import:
                        # Ensure imported data is a dictionary and make a copy
                        if isinstance(imported_data, dict):
                            new_styles_from_import[style_name].update(imported_data.copy())
                        else:
                            print(f"警告: 导入的样式 '{style_name}' 数据格式不正确，已跳过。")
                    else:
                        # If style_name from import is not in defaults, add it (if it's a dict)
                        if isinstance(imported_data, dict):
                             new_styles_from_import[style_name] = imported_data.copy()

                self.current_styles = new_styles_from_import
                
                # Update GUI tabs
                self.style_tabs.clear()
                self.style_widgets = {}
                for style_name, style_data in self.current_styles.items():
                    # Convert color_rgb back to tuple if it's a list from JSON
                    if 'color_rgb' in style_data and isinstance(style_data['color_rgb'], list):
                        style_data['color_rgb'] = tuple(style_data['color_rgb'])
                    widget = StyleConfigWidget(style_name, style_data) # Pass copy
                    self.style_tabs.addTab(widget, style_name)
                    self.style_widgets[style_name] = widget
                
                QMessageBox.information(self, "导入成功", f"样式预设已从 '{os.path.basename(file_path)}' 导入。")
                self._preview_markdown() 

            except Exception as e:
                QMessageBox.critical(self, "导入失败", f"导入样式预设时发生错误: {e}\n查看控制台获取详细信息。")
                import traceback
                traceback.print_exc()


    def _export_styles(self):
        presets_dir = "presets"
        os.makedirs(presets_dir, exist_ok=True)

        timestamp = QDateTime.currentDateTime().toString("yyyyMMdd_hhmmss")
        default_filename = f"styles_preset_{timestamp}.json"
        default_path = os.path.join(presets_dir, default_filename)

        file_path, _ = QFileDialog.getSaveFileName(self, "导出样式预设", default_path, "JSON Files (*.json);;All Files (*)")
        if file_path:
            try:
                self._update_current_styles_from_gui() 
                
                # Prepare styles for JSON: ensure color_rgb is a list
                styles_to_export = {}
                for k, v_dict in self.current_styles.items():
                    s_copy = v_dict.copy()
                    if 'color_rgb' in s_copy and isinstance(s_copy['color_rgb'], tuple):
                        s_copy['color_rgb'] = list(s_copy['color_rgb'])
                    styles_to_export[k] = s_copy

                with open(file_path, 'w', encoding='utf-8') as f:
                    json.dump(styles_to_export, f, indent=4, ensure_ascii=False)
                QMessageBox.information(self, "导出成功", f"当前样式已成功导出到 '{os.path.basename(file_path)}'。")
            except Exception as e:
                QMessageBox.critical(self, "导出失败", f"导出样式预设时发生错误: {e}\n查看控制台获取详细信息。")
                import traceback
                traceback.print_exc()


if __name__ == '__main__':
    if hasattr(Qt, 'AA_EnableHighDpiScaling'):
        QApplication.setAttribute(Qt.AA_EnableHighDpiScaling, True)
    if hasattr(Qt, 'AA_UseHighDpiPixmaps'):
        QApplication.setAttribute(Qt.AA_UseHighDpiPixmaps, True)
        
    app = QApplication(sys.argv)

    try:
        # It's good practice to initialize the default profile if using WebEngine extensively
        # QWebEngineProfile.defaultProfile() 
        pass # QWebEngineView creation will handle necessary initializations.
    except ImportError:
        QMessageBox.critical(None, "模块错误", "PyQtWebEngine模块未找到。\n请确保已安装：pip install PyQtWebEngine")
        sys.exit(1)
    except Exception as e: 
        QMessageBox.critical(None, "QtWebEngine 初始化错误", f"加载QtWebEngine时发生错误: {e}\n请检查您的Qt安装和环境变量。")
        sys.exit(1)

    main_window = MarkdownToWordGUI()
    main_window.show()
    sys.exit(app.exec_())