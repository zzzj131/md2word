import markdown
from bs4 import BeautifulSoup
from docx import Document
from docx.shared import Pt, RGBColor, Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn, nsdecls
from docx.oxml import parse_xml
from docx.oxml.shared import OxmlElement
import os
import re
import json

class MarkdownToWordConverter:
    DEFAULT_STYLES_CONFIG = {
        'H1': {
            'font_name': '黑体',
            'font_size': 24,
            'bold': True,
            'italic': False,
            'color_rgb': (0, 0, 0),
            'space_before_pt': 12,
            'space_after_pt': 6,
        },
        'H2': {
            'font_name': '黑体',
            'font_size': 20,
            'bold': True,
            'italic': False,
            'color_rgb': (0, 0, 0),
            'space_before_pt': 10,
            'space_after_pt': 5,
        },
        'H3': {
            'font_name': '黑体',
            'font_size': 18,
            'bold': True,
            'italic': False,
            'color_rgb': (0, 0, 0),
            'space_before_pt': 8,
            'space_after_pt': 4,
        },
        'H4': {
            'font_name': '宋体',
            'font_size': 16,
            'bold': True,
            'italic': False,
            'color_rgb': (50, 50, 50),
            'space_before_pt': 6,
            'space_after_pt': 3,
        },
        'H5': {
            'font_name': '宋体',
            'font_size': 14,
            'bold': True,
            'italic': False,
            'color_rgb': (70, 70, 70),
            'space_before_pt': 5,
            'space_after_pt': 2,
        },
        'H6': {
            'font_name': '宋体',
            'font_size': 12,
            'bold': True,
            'italic': False,
            'color_rgb': (90, 90, 90),
            'space_before_pt': 4,
            'space_after_pt': 2,
        },
        'paragraph': {
            'font_name': '宋体',
            'font_size': 12,
            'bold': False,
            'italic': False,
            'color_rgb': (0, 0, 0),
            'line_spacing': 1.5,
            'first_line_indent_cm': 0.74,
            'space_after_pt': 6,
        },
        'code_block': {
            'font_name': 'Courier New',
            'font_size': 10,
            'bold': False,
            'italic': False,
            'color_rgb': (0, 0, 0),
            'background_color': "F0F0F0",
            'line_spacing': 1.0,
        },
        'inline_code': { # New style for inline code
            'font_name': 'Courier New',
            'font_size_ratio': 0.9, # Relative to paragraph font size
            'color_rgb': (50, 50, 50)
        }
    }

    def __init__(self, styles_config=None):
        self.styles_config = styles_config if styles_config is not None else self.DEFAULT_STYLES_CONFIG.copy()

    def _apply_font_style(self, run, style_config):
        font = run.font
        font.name = style_config.get('font_name', '宋体')
        font.element.rPr.rFonts.set(qn('w:eastAsia'), style_config.get('font_name', '宋体'))
        font.size = Pt(style_config.get('font_size', 12))
        font.bold = style_config.get('bold', False)
        font.italic = style_config.get('italic', False)
        if 'color_rgb' in style_config:
            font.color.rgb = RGBColor(*style_config['color_rgb'])

    def _apply_paragraph_format(self, paragraph, style_config):
        p_format = paragraph.paragraph_format
        if 'space_before_pt' in style_config:
            p_format.space_before = Pt(style_config['space_before_pt'])
        if 'space_after_pt' in style_config:
            p_format.space_after = Pt(style_config['space_after_pt'])
        if 'line_spacing' in style_config:
            p_format.line_spacing = style_config['line_spacing']
        if 'first_line_indent_cm' in style_config:
            p_format.first_line_indent = Inches(style_config['first_line_indent_cm'] / 2.54)
        if 'alignment' in style_config:
            alignment_map = {
                'LEFT': WD_ALIGN_PARAGRAPH.LEFT,
                'CENTER': WD_ALIGN_PARAGRAPH.CENTER,
                'RIGHT': WD_ALIGN_PARAGRAPH.RIGHT,
                'JUSTIFY': WD_ALIGN_PARAGRAPH.JUSTIFY
            }
            p_format.alignment = alignment_map.get(style_config['alignment'].upper(), WD_ALIGN_PARAGRAPH.LEFT)

    def _set_cell_background(self, cell, hex_color_string):
        if hex_color_string:
            # Ensure hex_color_string is a valid 6-digit hex code (e.g., "RRGGBB")
            # python-docx shading expects a hex string without '#'
            if hex_color_string.startswith('#'):
                hex_color_string = hex_color_string[1:]
            
            # Set the background color using the shading property
            # The fill property expects a hex color string (e.g., "FF0000" for red)
            # The val property specifies the shading pattern, "clear" means solid fill
            # Use OxmlElement to correctly build the XML element
            from docx.oxml.shared import OxmlElement
            shd = OxmlElement('w:shd')
            shd.set(qn('w:val'), 'clear')
            shd.set(qn('w:fill'), hex_color_string)
            cell._tc.get_or_add_tcPr().append(shd)

    def markdown_to_docx(self, md_file_path, docx_file_path):
        if not os.path.exists(md_file_path):
            print(f"错误: Markdown文件 '{md_file_path}' 不存在。")
            return

        md_dir = os.path.dirname(os.path.abspath(md_file_path))

        with open(md_file_path, 'r', encoding='utf-8') as f:
            md_content = f.read()

        html_content = markdown.markdown(md_content, extensions=['fenced_code', 'tables', 'sane_lists'])
        soup = BeautifulSoup(html_content, 'html.parser')
        doc = Document()

        default_font_name = self.styles_config.get('paragraph', {}).get('font_name', '宋体')
        default_font_size = self.styles_config.get('paragraph', {}).get('font_size', 12)
        
        style = doc.styles['Normal']
        style.font.name = default_font_name
        style.element.rPr.rFonts.set(qn('w:eastAsia'), default_font_name)
        style.font.size = Pt(default_font_size)

        for element in soup.find_all(True, recursive=False):
            tag_name = element.name

            if re.match(r'h[1-6]', tag_name):
                level = int(tag_name[1])
                style_key = tag_name.upper()
                heading_style = self.styles_config.get(style_key, self.styles_config['paragraph'])

                p = doc.add_heading(element.get_text(), level=level)
                if p.runs:
                    self._apply_font_style(p.runs[0], heading_style)
                self._apply_paragraph_format(p, heading_style)

            elif tag_name == 'p':
                para_style = self.styles_config['paragraph']
                p = doc.add_paragraph()
                
                for child in element.children:
                    if child.name == 'img':
                        img_src = child.get('src')
                        alt_text = child.get('alt', '')
                        
                        if not os.path.isabs(img_src):
                            img_path = os.path.join(md_dir, img_src)
                        else:
                            img_path = img_src
                        
                        if os.path.exists(img_path):
                            try:
                                p.add_run().add_picture(img_path)
                            except Exception as e:
                                print(f"警告: 无法插入图片 '{img_path}': {e}")
                                p.add_run(f"[图片无法加载: {alt_text or img_src}]")
                        else:
                            print(f"警告: 图片文件 '{img_path}' (源: {img_src}) 未找到。")
                            p.add_run(f"[图片未找到: {alt_text or img_src}]")

                    elif child.name == 'strong' or child.name == 'b':
                        run = p.add_run(child.get_text())
                        self._apply_font_style(run, para_style)
                        run.bold = True
                    elif child.name == 'em' or child.name == 'i':
                        run = p.add_run(child.get_text())
                        self._apply_font_style(run, para_style)
                        run.italic = True
                    elif child.name == 'code':
                        run = p.add_run(child.get_text())
                        inline_code_style = self.styles_config.get('inline_code', {})
                        code_font_name = inline_code_style.get('font_name', 'Courier New')
                        code_font_size_ratio = inline_code_style.get('font_size_ratio', 0.9)
                        code_color_rgb = inline_code_style.get('color_rgb', (50, 50, 50))
                        
                        self._apply_font_style(run, {
                            'font_name': code_font_name,
                            'font_size': para_style.get('font_size', 12) * code_font_size_ratio,
                            'color_rgb': code_color_rgb
                        })

                    elif child.name is None:
                        run = p.add_run(str(child))
                        self._apply_font_style(run, para_style)
                
                self._apply_paragraph_format(p, para_style)

            elif tag_name == 'pre':
                code_text = element.find('code').get_text() if element.find('code') else element.get_text()
                code_style = self.styles_config['code_block']

                table = doc.add_table(rows=1, cols=1)
                table.autofit = False
                table.columns[0].width = Inches(6)

                cell = table.cell(0, 0)
                
                if 'background_color' in code_style:
                    self._set_cell_background(cell, code_style['background_color'])

                if cell.paragraphs:
                    p = cell.paragraphs[0]
                    p.clear()
                else:
                    p = cell.add_paragraph()
                
                for line in code_text.splitlines():
                    run = p.add_run(line + '\n')
                    self._apply_font_style(run, code_style)
                
                if p.runs and p.runs[-1].text.endswith('\n'):
                     p.runs[-1].text = p.runs[-1].text[:-1]

                self._apply_paragraph_format(p, code_style)
                self._apply_paragraph_format(table.rows[0].cells[0].paragraphs[0], {'space_after_pt': self.styles_config.get('paragraph',{}).get('space_after_pt', 6)})

            elif tag_name == 'ul' or tag_name == 'ol':
                list_items = element.find_all('li', recursive=False)
                for item in list_items:
                    list_style_name = 'ListBullet' if tag_name == 'ul' else 'ListNumber'
                    p = doc.add_paragraph(style=list_style_name)
                    
                    para_style = self.styles_config['paragraph']
                    for child in item.children:
                        if child.name == 'strong' or child.name == 'b':
                            run = p.add_run(child.get_text())
                            self._apply_font_style(run, para_style)
                            run.bold = True
                        elif child.name == 'em' or child.name == 'i':
                            run = p.add_run(child.get_text())
                            self._apply_font_style(run, para_style)
                            run.italic = True
                        elif child.name == 'code':
                            run = p.add_run(child.get_text())
                            inline_code_style = self.styles_config.get('inline_code', {})
                            code_font_name = inline_code_style.get('font_name', 'Courier New')
                            code_font_size_ratio = inline_code_style.get('font_size_ratio', 0.9)
                            code_color_rgb = inline_code_style.get('color_rgb', (50, 50, 50))
                            self._apply_font_style(run, {
                                'font_name': code_font_name,
                                'font_size': para_style.get('font_size', 12) * code_font_size_ratio,
                                'color_rgb': code_color_rgb
                            })
                        elif child.name is None:
                            run = p.add_run(str(child))
                            self._apply_font_style(run, para_style)
                    
                    self._apply_paragraph_format(p, para_style)
                    p.paragraph_format.first_line_indent = None

            elif tag_name == 'table':
                header_row = element.find('thead').find('tr') if element.find('thead') else None
                body_rows = element.find('tbody').find_all('tr') if element.find('tbody') else element.find_all('tr')

                if not body_rows and header_row:
                     body_rows = []

                if not header_row and not body_rows:
                    continue

                num_cols = 0
                if header_row:
                    num_cols = len(header_row.find_all(['th', 'td']))
                elif body_rows:
                    num_cols = len(body_rows[0].find_all(['th', 'td']))
                
                if num_cols == 0: continue

                word_table = doc.add_table(rows=0, cols=num_cols)
                word_table.style = 'TableGrid'

                if header_row:
                    cells = header_row.find_all(['th', 'td'])
                    row = word_table.add_row().cells
                    for i, cell_content in enumerate(cells):
                        if i < num_cols:
                            p = row[i].paragraphs[0]
                            p.clear()
                            run = p.add_run(cell_content.get_text())
                            self._apply_font_style(run, self.styles_config.get('paragraph', {}))
                            run.bold = True
                            self._apply_paragraph_format(p, {'alignment': 'CENTER'})

                for body_row_html in body_rows:
                    cells = body_row_html.find_all('td')
                    row = word_table.add_row().cells
                    for i, cell_content in enumerate(cells):
                        if i < num_cols:
                            p = row[i].paragraphs[0]
                            p.clear()
                            run = p.add_run(cell_content.get_text())
                            self._apply_font_style(run, self.styles_config.get('paragraph', {}))
            
            elif tag_name == 'hr':
                # A simpler way to add a horizontal line using a paragraph with a bottom border
                # This avoids complex VML XML and namespace issues.
                p = doc.add_paragraph()
                p_format = p.paragraph_format
                p_format.space_before = Pt(6)
                p_format.space_after = Pt(6)
                
                # Add a bottom border to the paragraph
                pPr = p._element.get_or_add_pPr()
                pBdr = OxmlElement('w:pBdr')
                bottom = OxmlElement('w:bottom')
                bottom.set(qn('w:val'), 'single')
                bottom.set(qn('w:sz'), '6') # 1/2 pt line
                bottom.set(qn('w:space'), '1')
                bottom.set(qn('w:color'), 'auto')
                pBdr.append(bottom)
                pPr.append(pBdr)

        doc.save(docx_file_path)
        print(f"成功将 '{md_file_path}' 转换为 '{docx_file_path}'")

    def load_styles(self, file_path="styles_config.json"):
        if os.path.exists(file_path):
            try:
                with open(file_path, 'r', encoding='utf-8') as f:
                    loaded_config = json.load(f)
                    # Convert list (from JSON) back to tuple for RGBColor
                    for key, value in loaded_config.items():
                        if 'color_rgb' in value and isinstance(value['color_rgb'], list):
                            value['color_rgb'] = tuple(value['color_rgb'])
                    self.styles_config.update(loaded_config)
                print(f"样式已从 '{file_path}' 加载。")
            except json.JSONDecodeError as e:
                print(f"错误: 无法解析样式配置文件 '{file_path}': {e}")
            except Exception as e:
                print(f"加载样式时发生未知错误: {e}")
        else:
            print(f"样式配置文件 '{file_path}' 不存在，使用默认样式。")

    def save_styles(self, file_path="styles_config.json"):
        try:
            # Convert tuples (RGBColor) to lists for JSON serialization
            serializable_config = self.styles_config.copy()
            for key, value in serializable_config.items():
                if 'color_rgb' in value and isinstance(value['color_rgb'], tuple):
                    value['color_rgb'] = list(value['color_rgb'])
            
            with open(file_path, 'w', encoding='utf-8') as f:
                json.dump(serializable_config, f, ensure_ascii=False, indent=4)
            print(f"样式已保存到 '{file_path}'。")
        except Exception as e:
            print(f"保存样式时发生错误: {e}")

    def get_styles(self):
        return self.styles_config

    def set_styles(self, new_styles):
        self.styles_config = new_styles
        # Ensure RGB colors are tuples if they come as lists from GUI
        for key, value in self.styles_config.items():
            if 'color_rgb' in value and isinstance(value['color_rgb'], list):
                value['color_rgb'] = tuple(value['color_rgb'])

    def markdown_to_html(self, md_content):
        """
        将Markdown内容转换为HTML，并尝试注入CSS样式以模拟Word样式。
        """
        html_content = markdown.markdown(md_content, extensions=['fenced_code', 'tables', 'sane_lists'])
        soup = BeautifulSoup(html_content, 'html.parser')

        # 动态生成CSS
        css_styles = []
        for style_key, config in self.styles_config.items():
            if style_key.startswith('H'):
                tag = style_key.lower()
                css = f"{tag} {{"
                if 'font_name' in config: css += f"font-family: '{config['font_name']}';"
                if 'font_size' in config: css += f"font-size: {config['font_size']}pt;"
                if 'bold' in config and config['bold']: css += "font-weight: bold;"
                if 'italic' in config and config['italic']: css += "font-style: italic;"
                if 'color_rgb' in config: css += f"color: rgb({config['color_rgb'][0]}, {config['color_rgb'][1]}, {config['color_rgb'][2]});"
                if 'space_before_pt' in config: css += f"margin-top: {config['space_before_pt']}pt;"
                if 'space_after_pt' in config: css += f"margin-bottom: {config['space_after_pt']}pt;"
                css += "}"
                css_styles.append(css)
            elif style_key == 'paragraph':
                css = "p {"
                if 'font_name' in config: css += f"font-family: '{config['font_name']}';"
                if 'font_size' in config: css += f"font-size: {config['font_size']}pt;"
                if 'bold' in config and config['bold']: css += "font-weight: bold;"
                if 'italic' in config and config['italic']: css += "font-style: italic;"
                if 'color_rgb' in config: css += f"color: rgb({config['color_rgb'][0]}, {config['color_rgb'][1]}, {config['color_rgb'][2]});"
                if 'line_spacing' in config: css += f"line-height: {config['line_spacing']};"
                if 'first_line_indent_cm' in config: css += f"text-indent: {config['first_line_indent_cm']}cm;"
                if 'space_after_pt' in config: css += f"margin-bottom: {config['space_after_pt']}pt;"
                css += "}"
                css_styles.append(css)
            elif style_key == 'code_block':
                css = "pre, code.block {" # Target pre for block code
                if 'font_name' in config: css += f"font-family: '{config['font_name']}';"
                if 'font_size' in config: css += f"font-size: {config['font_size']}pt;"
                if 'bold' in config and config['bold']: css += "font-weight: bold;"
                if 'italic' in config and config['italic']: css += "font-style: italic;"
                if 'color_rgb' in config: css += f"color: rgb({config['color_rgb'][0]}, {config['color_rgb'][1]}, {config['color_rgb'][2]});"
                if 'background_color' in config: css += f"background-color: #{config['background_color']};"
                if 'line_spacing' in config: css += f"line-height: {config['line_spacing']};"
                css += "padding: 10px; border: 1px solid #ccc; overflow: auto;}"
                css_styles.append(css)
            elif style_key == 'inline_code':
                css = "p code, li code {" # Target inline code within p and li
                if 'font_name' in config: css += f"font-family: '{config['font_name']}';"
                # Calculate font size based on paragraph's font size and ratio
                para_font_size = self.styles_config.get('paragraph', {}).get('font_size', 12)
                font_size_ratio = config.get('font_size_ratio', 0.9)
                css += f"font-size: {para_font_size * font_size_ratio}pt;"
                if 'color_rgb' in config: css += f"color: rgb({config['color_rgb'][0]}, {config['color_rgb'][1]}, {config['color_rgb'][2]});"
                css += "background-color: #f0f0f0; padding: 2px 4px; border-radius: 3px;}"
                css_styles.append(css)

        # Add basic list styling for preview
        css_styles.append("ul, ol { margin-left: 20px; }")
        css_styles.append("li { margin-bottom: 5px; }")
        css_styles.append("table { border-collapse: collapse; width: 100%; }")
        css_styles.append("th, td { border: 1px solid #ddd; padding: 8px; text-align: left; }")
        css_styles.append("th { background-color: #f2f2f2; text-align: center; font-weight: bold; }")
        css_styles.append("img { max-width: 100%; height: auto; }") # Ensure images fit in preview

        full_html = """
        <!DOCTYPE html>
        <html>
        <head>
            <meta charset="utf-8">
            <title>Markdown Preview</title>
            <style>
                body {{ font-family: '{}'; font-size: {}pt; line-height: {}; }}
                {}
            </style>
        </head>
        <body>
            {}
        </body>
        </html>
        """.format(
            self.styles_config.get('paragraph', {}).get('font_name', '宋体'),
            self.styles_config.get('paragraph', {}).get('font_size', 12),
            self.styles_config.get('paragraph', {}).get('line_spacing', 1.5),
            "\n".join(css_styles),
            html_content
        )
        return full_html
