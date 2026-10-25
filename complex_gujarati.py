import os
import urllib.request
import math
import pathlib
import json
import re
import openpyxl
import io # New import for in-memory buffer
from reportlab.pdfgen import canvas
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.lib.pagesizes import A4
from reportlab.lib.utils import ImageReader
from pypdf import PdfReader, PdfWriter, PageObject # New imports for PDF merging


# Load configuration from JSON
CONFIG_FILE = pathlib.Path(__file__).parent / "config.json"
try:
    with open(CONFIG_FILE, "r", encoding="utf-8") as f:
        config = json.load(f)
except Exception as e:
    print(f"Error loading config file '{CONFIG_FILE}': {e}")
    print("Create a config.json with keys: font_filename, font_url, template_pdf_path, output_dir, excel_path, pdf_filename_col, text_fields")
    exit()

# Pythonic variable names (lower_snake_case)
font_filename = config.get("font_filename", "NotoSansGujarati-Bold.ttf")
font_url = config.get("font_url", "https://github.com/googlefonts/noto-fonts/raw/main/hinted/ttf/NotoSansGujarati/NotoSansGujarati-Regular.ttf")
template_pdf_path = config.get("template_pdf_path")
output_dir = config.get("output_dir", r"output")
excel_path = config.get("excel_path", r"data.xlsx")
# Single 0-based column index to use for PDF filename (preferred if set)
# Treat the value in `config.json` as 0-based (0 -> first column)
pdf_filename_col = config.get("pdf_filename_col")

# ensure output dir exists
os.makedirs(output_dir, exist_ok=True)

# `start_row` in config is 0-based (0 => first Excel row).
# Default to 1 so behavior matches previous `min_row=2` (skip header row).
start_row = config.get("start_row", 1)
try:
    start_row = int(start_row)
    if start_row < 0:
        start_row = 0
except Exception:
    start_row = 1

# `end_row` in config is 0-based (0 => first Excel row).
# If not provided or set to -1, iterate to the last row (current behavior).
end_row = config.get("end_row", -1)
try:
    end_row = int(end_row)
except Exception:
    end_row = -1

# Which text columns to render and where; list of objects with keys:
#  - col (1-based Excel column index in the `row` tuple)
#  - page, x, y, font_size, centered
text_fields = config.get("text_fields", [
    {"col": 2, "page": 0, "x": 100, "y": 200, "font_size": 25, "centered": False},
    {"col": 3, "page": 1, "x": 100, "y": 50, "font_size": 25, "centered": False},
])


# --- Import Rendering Engines ---
try:
    from PIL import Image, ImageDraw
    import uharfbuzz as hb
    import freetype
except ImportError as e:
    print("Missing dependencies!")
    print(f"Error: {e}")
    print("Please run: pip install uharfbuzz freetype-py Pillow reportlab pypdf")
    exit()


def download_font(font_name, url):
    """Downloads a font file if it doesn't already exist locally."""
    if not os.path.exists(font_name):
        print(f"Downloading {font_name}...")
        try:
            urllib.request.urlretrieve(url, font_name)
            print("Download complete.")
        except Exception as e:
            print(f"Error downloading font: {e}")
            return False
    return True


class TextRenderer:
    """
    A helper class to shape and render complex Indic scripts using 
    HarfBuzz (shaping) and FreeType (rasterization).
    """
    def __init__(self, font_path, font_size):
        self.font_path = font_path
        self.font_size = font_size
        self.scale_factor = 64  # FreeType 26.6 fixed point coordinates

        # 1. Initialize FreeType
        self.ft_face = freetype.Face(font_path)
        self.ft_face.set_char_size(font_size * self.scale_factor)

        # 2. Initialize HarfBuzz
        blob = hb.Blob.from_file_path(font_path)
        face = hb.Face(blob)
        self.hb_font = hb.Font(face)
        # Match HarfBuzz scale to FreeType size
        self.hb_font.scale = (font_size * self.scale_factor, font_size * self.scale_factor)

    def render_to_image(self, text, text_color=(255, 0, 0, 255)):
        """Shapes text and returns a PIL Image."""
        # --- A. Shaping with HarfBuzz ---
        buf = hb.Buffer()
        buf.add_str(text)
        buf.guess_segment_properties()
        hb.shape(self.hb_font, buf)

        infos = buf.glyph_infos
        positions = buf.glyph_positions

        # --- B. Calculate Dimensions ---
        x_cursor = 0
        y_cursor = 0
        
        metrics = self.ft_face.size
        ascender = metrics.ascender / self.scale_factor
        descender = metrics.descender / self.scale_factor
        height = int(math.ceil(ascender - descender))
        
        total_width = 0
        for pos in positions:
            total_width += pos.x_advance
        
        width = int(math.ceil(total_width / self.scale_factor)) + 10 # Add padding
        
        img = Image.new("RGBA", (width, height + 10), (255, 255, 255, 0))
        #img = Image.new("RGBA", (width, height + 10), (255, 0, 0, 0))
        
        # --- C. Rasterize Glyphs with FreeType ---
        pen_x = 0
        baseline = int(ascender) + 5 # Add small top padding

        for info, pos in zip(infos, positions):
            gid = info.codepoint
            
            self.ft_face.load_glyph(gid, freetype.FT_LOAD_RENDER)
            bitmap = self.ft_face.glyph.bitmap
            
            x_offset = pos.x_offset / self.scale_factor
            y_offset = pos.y_offset / self.scale_factor
            
            draw_x = int(pen_x + x_offset + self.ft_face.glyph.bitmap_left)
            draw_y = int(baseline - (y_offset + self.ft_face.glyph.bitmap_top))

            if bitmap.width > 0 and bitmap.rows > 0:
                glyph_mask = Image.frombytes("L", (bitmap.width, bitmap.rows), bytes(bitmap.buffer))
                glyph_color = Image.new("RGBA", (bitmap.width, bitmap.rows), text_color)
                img.paste(glyph_color, (draw_x, draw_y), mask=glyph_mask)

            pen_x += pos.x_advance / self.scale_factor
            
        return img

class GujaratiPDFWriter:
    # Font and URL are loaded from `config.json` as module-level vars


    def __init__(self, template_pdf_path : str, output_pdf_name : str):
        self.template_pdf_path = template_pdf_path
        self.template_pdf_reader = PdfReader(self.template_pdf_path)
        #self.template_pdf_pages = self.template_pdf_reader.pages
        self.template_pdf_pages = [p for p in self.template_pdf_reader.pages]

        self.output_pdf = os.path.join(output_dir, f"{output_pdf_name}.pdf")

        if not download_font(font_filename, font_url):
            return

        # Check for input file existence
        if not os.path.exists(self.template_pdf_path):
            print(f"Error: Input PDF not found at '{self.template_pdf_path}'. Please ensure it exists.")
            return


    
    def write_text(self, text, page_number, x, y, font_size=25, centered=False):
        # Generate the Overlay PDF (Gujarati Content) in memory ---
        packet = io.BytesIO()

        # Page from the template PDF
        base_page = self.template_pdf_pages[page_number]
        page_width = base_page.mediabox.width
        page_height = base_page.mediabox.height

        # Create the overlay canvas using the existing page size
        c = canvas.Canvas(packet, pagesize=(page_width, page_height))

        # Setup Helper Rendering Function
        def draw_complex_text(text, size, x, y, centered=False):
            try:
                renderer = TextRenderer(font_filename, size)
                img = renderer.render_to_image(text)
                
                # PDF dimensions (PIL image size is already in points/pixels for a 1:1 map)
                pdf_w = img.width 
                pdf_h = img.height
                
                draw_x = x - (pdf_w / 2) if centered else x
                draw_y = y 
                
                c.drawImage(ImageReader(img), draw_x, draw_y, width=pdf_w, height=pdf_h, mask='auto')
            except Exception as e:
                print(f"Error drawing text '{text}': {e}")

        draw_complex_text(text, font_size, x, y, centered=centered)

        # Finish generating the overlay PDF
        c.save()
        packet.seek(0)

        # Merge the Overlay with the Existing PDF
        
        overlay_pdf = PdfReader(packet)
        overlay_page = overlay_pdf.pages[0]
        
        base_page.merge_page(overlay_page)
        self.template_pdf_pages[page_number] = base_page



    def save(self):
        output_pdf_writer = PdfWriter()
        for p in self.template_pdf_pages:
            output_pdf_writer.add_page(p)

        with open(self.output_pdf, "wb") as f:
            output_pdf_writer.write(f)


################


def main():
    if not os.path.exists(excel_path) :
        print(f"Excel file {excel_path} not found.")
        return

    def _sanitize_filename(s: str) -> str:
        # Replace common Windows illegal filename chars with underscore
        if s is None:
            return ""
        s = str(s).strip()
        # Replace path separators and other reserved chars
        s = re.sub(r"[\\/:*?\"<>|]", "_", s)
        # Collapse whitespace to single underscores
        s = re.sub(r"\s+", "_", s)
        return s


    def _build_output_name(row_tuple) -> str:
        # If a single pdf_output_col is set (1-based), prefer that
        try:
            if pdf_filename_col is not None:
                # pdf_filename_col is 0-based (0 = first column)
                idx = int(pdf_filename_col)
                if 0 <= idx < len(row_tuple):
                    val = row_tuple[idx]
                    if val is not None:
                        return _sanitize_filename(val)
        except Exception:
            # ignore and fall back
            pass

        # No multi-field filename support: fallback to first column (row[0])
        if len(row_tuple) > 0 and row_tuple[0] is not None:
            return _sanitize_filename(row_tuple[0])
        return "output"


    w = openpyxl.load_workbook(excel_path)
    sheet = w.active
    # openpyxl uses 1-based row indices; convert 0-based `start_row`/`end_row` to 1-based
    min_row = int(start_row) + 1
    max_row = None
    if end_row is not None and end_row >= 0:
        # convert 0-based end_row to openpyxl 1-based max_row
        max_row = int(end_row) + 1

    if max_row is None:
        rows_iter = sheet.iter_rows(min_row=min_row, values_only=True)
    else:
        rows_iter = sheet.iter_rows(min_row=min_row, max_row=max_row, values_only=True)

    for row in rows_iter:
        pdf_name = _build_output_name(row)
        print(pdf_name)
        guj_pdf = GujaratiPDFWriter(template_pdf_path, pdf_name)

        # Iterate configured text fields and render any non-empty values
        for field in text_fields:
            col = field.get("col")
            if col is None:
                continue
            idx = int(col)# - 1
            if idx < 0 or idx >= len(row):
                continue
            val = row[idx]
            if val is None:
                continue

            page = int(field.get("page", 0))
            x = field.get("x", 0)
            y = field.get("y", 0)
            font_size = field.get("font_size", 25)
            centered = bool(field.get("centered", False))

            guj_pdf.write_text(str(val), page, x, y, font_size=font_size, centered=centered)

        guj_pdf.save()

    w.close()
    

if __name__ == "__main__":
    main()