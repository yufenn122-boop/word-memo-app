import os
import io
import zipfile
import tempfile
from dataclasses import dataclass
from typing import List, Optional, Tuple

import streamlit as st
from PIL import Image, ImageDraw, ImageFont
from docx import Document
from docx.shared import Length


# =========================================================
# 固定高清画布：3:4
# =========================================================

PAGE_W = 2160
PAGE_H = 2880

BG_IMAGE_PATH = "memo_bg.png"
BG_COLOR = (255, 255, 255)

TEXT_COLOR = "#111111"
NOTE_YELLOW = "#d6b839"

# 内容区位置：按 2160×2880 调
LEFT = 145
RIGHT = 145
CONTENT_TOP = 410
BOTTOM = 120

CONTENT_W = PAGE_W - LEFT - RIGHT

# 标题专用区域：更像参考图一，不让标题铺太满
TITLE_LEFT = LEFT
TITLE_RIGHT = 300
TITLE_W = PAGE_W - TITLE_LEFT - TITLE_RIGHT

# Word pt 转高清图 px 的倍率
# 正文如果整体太大，改 3.0；太小，改 3.5
PT_TO_PX_SCALE = 3.25


# =========================================================
# 默认模板字号 / 间距
# 标题类走固定模板；正文尽量读 Word
# =========================================================

DEFAULT_FONT_PX = {
    "title": 86,
    "h1": 72,
    "h2": 66,
    "body": 58,
    "blank": 58,
}

DEFAULT_SPACE_AFTER = {
    "title": 34,
    "h1": 24,
    "h2": 20,
    "body": 22,
    "blank": 30,
}

DEFAULT_SPACE_BEFORE = {
    "title": 8,
    "h1": 34,
    "h2": 26,
    "body": 0,
    "blank": 0,
}

DEFAULT_LINE_RATIO = {
    "title": 1.10,
    "h1": 1.28,
    "h2": 1.32,
    "body": 1.52,
    "blank": 1.0,
}

HIGHLIGHT_COLORS = {
    "yellow": "#f5df6d",
    "blue": "#bceeff",
    "pink": "#ff76d2",
    "green": "#8ee38d",
}


# =========================================================
# 数据结构
# =========================================================

@dataclass
class Chunk:
    text: str
    bold: bool = False
    underline: bool = False
    highlight: Optional[str] = None
    font_size_px: Optional[int] = None


@dataclass
class ParagraphBlock:
    chunks: List[Chunk]
    kind: str = "body"
    line_height_px: Optional[int] = None
    space_before_px: int = 0
    space_after_px: int = 0


# =========================================================
# 字体：本地优先微软雅黑；云端优先 Noto CJK
# =========================================================

FONT_CACHE = {}


def find_font_path(bold=False) -> Optional[str]:
    """
    Windows 本地：优先微软雅黑。
    Streamlit Cloud/Linux：优先 Noto Sans CJK。
    Mac：尝试 PingFang。
    """

    # 允许你自己放字体文件到项目 fonts 文件夹
    project_regular = [
        "fonts/msyh.ttc",
        "fonts/msyh.ttf",
        "fonts/NotoSansCJK-Regular.ttc",
        "fonts/NotoSansSC-Regular.otf",
        "fonts/PingFang.ttc",
        "fonts/PingFangSC-Regular.ttf",
    ]

    project_bold = [
        "fonts/msyhbd.ttc",
        "fonts/msyhbd.ttf",
        "fonts/NotoSansCJK-Bold.ttc",
        "fonts/NotoSansSC-Bold.otf",
        "fonts/PingFang.ttc",
        "fonts/PingFangSC-Semibold.ttf",
    ]

    if os.name == "nt":
        system_regular = [
            r"C:\Windows\Fonts\msyh.ttc",
            r"C:\Windows\Fonts\msyh.ttf",
            r"C:\Windows\Fonts\simhei.ttf",
        ]

        system_bold = [
            r"C:\Windows\Fonts\msyhbd.ttc",
            r"C:\Windows\Fonts\msyhbd.ttf",
            r"C:\Windows\Fonts\msyh.ttc",
            r"C:\Windows\Fonts\simhei.ttf",
        ]
    else:
        system_regular = [
            "/usr/share/fonts/opentype/noto/NotoSansCJK-Regular.ttc",
            "/usr/share/fonts/opentype/noto/NotoSansCJK-Regular.otf",
            "/usr/share/fonts/truetype/noto/NotoSansCJK-Regular.ttc",
            "/usr/share/fonts/truetype/noto/NotoSansCJK-Regular.otf",
            "/usr/share/fonts/opentype/noto/NotoSansCJKsc-Regular.otf",
            "/usr/share/fonts/truetype/arphic/uming.ttc",
            "/System/Library/Fonts/PingFang.ttc",
            "/System/Library/Fonts/STHeiti Light.ttc",
        ]

        system_bold = [
            "/usr/share/fonts/opentype/noto/NotoSansCJK-Bold.ttc",
            "/usr/share/fonts/opentype/noto/NotoSansCJK-Bold.otf",
            "/usr/share/fonts/truetype/noto/NotoSansCJK-Bold.ttc",
            "/usr/share/fonts/truetype/noto/NotoSansCJK-Bold.otf",
            "/usr/share/fonts/opentype/noto/NotoSansCJKsc-Bold.otf",
            "/usr/share/fonts/opentype/noto/NotoSansCJK-Regular.ttc",
            "/usr/share/fonts/truetype/arphic/uming.ttc",
            "/System/Library/Fonts/PingFang.ttc",
            "/System/Library/Fonts/STHeiti Medium.ttc",
        ]

    candidates = project_bold + system_bold if bold else project_regular + system_regular

    for path in candidates:
        if os.path.exists(path):
            return path

    return None


def get_font(size: int, bold: bool = False):
    key = (size, bold)

    if key in FONT_CACHE:
        return FONT_CACHE[key]

    path = find_font_path(bold)

    if not path:
        raise RuntimeError(
            "没有找到中文字体。"
            "如果是 Streamlit Cloud，请确认仓库里有 packages.txt，内容包含 fonts-noto-cjk 和 fontconfig。"
            "如果是本地 Windows，请确认系统有微软雅黑。"
        )

    font = ImageFont.truetype(path, size)
    FONT_CACHE[key] = font
    return font


# =========================================================
# Word 格式读取
# =========================================================

def get_para_kind(p, index: int) -> str:
    style_name = ""

    if p.style and p.style.name:
        style_name = p.style.name.lower()

    if "title" in style_name or style_name == "标题":
        return "title"

    if "heading 1" in style_name or "标题 1" in style_name:
        return "title"

    if "heading 2" in style_name or "标题 2" in style_name:
        return "h1"

    if "heading 3" in style_name or "标题 3" in style_name:
        return "h2"

    text = p.text.strip()

    # 第一段较短时自动当大标题
    if index == 0 and text and len(text) <= 40:
        return "title"

    return "body"


def length_to_px(value, default=None) -> Optional[int]:
    if value is None:
        return default

    try:
        if isinstance(value, Length):
            pt = value.pt
            return int(pt * PT_TO_PX_SCALE)

        if hasattr(value, "pt"):
            return int(value.pt * PT_TO_PX_SCALE)
    except Exception:
        pass

    return default


def get_effective_para_format(p):
    pf = p.paragraph_format
    sf = p.style.paragraph_format if p.style is not None else None

    line_spacing = pf.line_spacing
    space_before = pf.space_before
    space_after = pf.space_after

    if line_spacing is None and sf is not None:
        line_spacing = sf.line_spacing

    if space_before is None and sf is not None:
        space_before = sf.space_before

    if space_after is None and sf is not None:
        space_after = sf.space_after

    return line_spacing, space_before, space_after


def calc_line_height_px(p, kind: str, font_px: int) -> int:
    # 标题 / 小标题统一按模板，不被 Word 原始行距带跑
    if kind in ["title", "h1", "h2"]:
        return int(font_px * DEFAULT_LINE_RATIO.get(kind, 1.2))

    line_spacing, _, _ = get_effective_para_format(p)

    if line_spacing is None:
        return int(font_px * DEFAULT_LINE_RATIO.get(kind, 1.52))

    # Word 倍数行距，例如 1.0 / 1.15 / 1.5 / 2.0
    if isinstance(line_spacing, float):
        return int(font_px * line_spacing)

    # Word 固定值行距
    px = length_to_px(line_spacing)
    if px:
        return px

    return int(font_px * DEFAULT_LINE_RATIO.get(kind, 1.52))


def calc_space_before_after_px(p, kind: str) -> Tuple[int, int]:
    # 标题 / 小标题统一走模板，保证像参考图
    if kind in ["title", "h1", "h2"]:
        before_px = DEFAULT_SPACE_BEFORE.get(kind, 0)
        after_px = DEFAULT_SPACE_AFTER.get(kind, 20)
        return int(before_px), int(after_px)

    _, space_before, space_after = get_effective_para_format(p)

    before_px = length_to_px(space_before)
    after_px = length_to_px(space_after)

    if before_px is None:
        before_px = DEFAULT_SPACE_BEFORE.get(kind, 0)

    if after_px is None:
        after_px = DEFAULT_SPACE_AFTER.get(kind, 22)

    return int(before_px), int(after_px)


def get_effective_run_font_size_px(run, p, kind: str) -> int:
    # 标题 / 小标题不跟 Word 原字号走，直接用模板字号
    if kind in ["title", "h1", "h2"]:
        return DEFAULT_FONT_PX[kind]

    size = run.font.size

    if size is not None:
        try:
            return max(20, int(size.pt * PT_TO_PX_SCALE))
        except Exception:
            pass

    try:
        if p.style and p.style.font and p.style.font.size:
            return max(20, int(p.style.font.size.pt * PT_TO_PX_SCALE))
    except Exception:
        pass

    return DEFAULT_FONT_PX.get(kind, DEFAULT_FONT_PX["body"])


def map_highlight(color) -> Optional[str]:
    if color is None:
        return None

    s = str(color).upper()

    if "YELLOW" in s:
        return "yellow"

    if "TURQUOISE" in s or "BLUE" in s:
        return "blue"

    if "PINK" in s or "RED" in s:
        return "pink"

    if "GREEN" in s or "BRIGHT_GREEN" in s:
        return "green"

    return "yellow"


def get_list_prefix(p, counters: dict) -> str:
    style_name = ""

    if p.style and p.style.name:
        style_name = p.style.name.lower()

    if "bullet" in style_name or "项目符号" in style_name:
        return "• "

    if "number" in style_name or "编号" in style_name:
        key = style_name
        counters[key] = counters.get(key, 0) + 1
        return f"{counters[key]}. "

    try:
        num_pr = p._p.pPr.numPr

        if num_pr is not None:
            num_id = str(num_pr.numId.val) if num_pr.numId is not None else "0"
            ilvl = str(num_pr.ilvl.val) if num_pr.ilvl is not None else "0"
            key = (num_id, ilvl)
            counters[key] = counters.get(key, 0) + 1
            return f"{counters[key]}. "
    except Exception:
        pass

    return ""


def parse_docx(file_path: str) -> List[ParagraphBlock]:
    doc = Document(file_path)
    blocks: List[ParagraphBlock] = []

    counters = {}
    non_empty_index = 0

    for p in doc.paragraphs:
        raw_text = p.text.replace("\t", "    ")

        if not raw_text.strip():
            blocks.append(
                ParagraphBlock(
                    chunks=[],
                    kind="blank",
                    line_height_px=DEFAULT_FONT_PX["blank"],
                    space_before_px=0,
                    space_after_px=DEFAULT_SPACE_AFTER["blank"],
                )
            )
            continue

        kind = get_para_kind(p, non_empty_index)
        non_empty_index += 1

        default_font_px = DEFAULT_FONT_PX.get(kind, DEFAULT_FONT_PX["body"])
        para_default_bold = kind in ["title", "h1", "h2"]

        chunks: List[Chunk] = []

        prefix = get_list_prefix(p, counters)
        if prefix:
            chunks.append(
                Chunk(
                    text=prefix,
                    bold=True,
                    underline=False,
                    highlight=None,
                    font_size_px=default_font_px,
                )
            )

        if p.runs:
            for run in p.runs:
                text = run.text.replace("\t", "    ")

                if not text:
                    continue

                bold = bool(run.bold) if run.bold is not None else para_default_bold
                underline = bool(run.underline)
                highlight = map_highlight(run.font.highlight_color)

                font_px = get_effective_run_font_size_px(run, p, kind)

                chunks.append(
                    Chunk(
                        text=text,
                        bold=bold,
                        underline=underline,
                        highlight=highlight,
                        font_size_px=font_px,
                    )
                )
        else:
            chunks.append(
                Chunk(
                    text=raw_text,
                    bold=para_default_bold,
                    underline=False,
                    highlight=None,
                    font_size_px=default_font_px,
                )
            )

        max_font_px = max(
            [c.font_size_px or default_font_px for c in chunks],
            default=default_font_px,
        )

        line_height_px = calc_line_height_px(p, kind, max_font_px)
        space_before_px, space_after_px = calc_space_before_after_px(p, kind)

        blocks.append(
            ParagraphBlock(
                chunks=chunks,
                kind=kind,
                line_height_px=line_height_px,
                space_before_px=space_before_px,
                space_after_px=space_after_px,
            )
        )

    return blocks


# =========================================================
# 文本测量与换行
# =========================================================

_MEASURE_IMG = Image.new("RGB", (10, 10), "white")
_MEASURE_DRAW = ImageDraw.Draw(_MEASURE_IMG)


def text_width(text: str, font) -> int:
    if not text:
        return 0

    return int(_MEASURE_DRAW.textlength(text, font=font))


def is_ascii_word_char(ch: str) -> bool:
    return ch.isascii() and (ch.isalnum() or ch in "-_'/.")


def tokenize_text(text: str) -> List[str]:
    tokens = []
    buf = ""

    for ch in text:
        if ch == "\n":
            if buf:
                tokens.append(buf)
                buf = ""
            tokens.append("\n")
            continue

        if is_ascii_word_char(ch):
            buf += ch
        else:
            if buf:
                tokens.append(buf)
                buf = ""

            tokens.append(ch)

    if buf:
        tokens.append(buf)

    return tokens


def split_chunk_by_token(chunk: Chunk) -> List[Chunk]:
    result = []

    for token in tokenize_text(chunk.text):
        result.append(
            Chunk(
                text=token,
                bold=chunk.bold,
                underline=chunk.underline,
                highlight=chunk.highlight,
                font_size_px=chunk.font_size_px,
            )
        )

    return result


def wrap_chunks(chunks: List[Chunk], max_width: int) -> List[List[Chunk]]:
    lines: List[List[Chunk]] = []
    current: List[Chunk] = []
    current_w = 0

    def flush():
        nonlocal current, current_w

        if current:
            lines.append(current)

        current = []
        current_w = 0

    for chunk in chunks:
        token_chunks = split_chunk_by_token(chunk)

        for tk in token_chunks:
            token = tk.text

            if token == "\n":
                flush()
                continue

            if token == "":
                continue

            if not current and token.isspace():
                continue

            font_size = tk.font_size_px or DEFAULT_FONT_PX["body"]
            font = get_font(font_size, tk.bold)
            w = text_width(token, font)

            if current and current_w + w > max_width:
                flush()

                if token.isspace():
                    continue

            if w > max_width:
                for ch in token:
                    font = get_font(font_size, tk.bold)
                    cw = text_width(ch, font)

                    if current and current_w + cw > max_width:
                        flush()

                    current.append(
                        Chunk(
                            text=ch,
                            bold=tk.bold,
                            underline=tk.underline,
                            highlight=tk.highlight,
                            font_size_px=tk.font_size_px,
                        )
                    )
                    current_w += cw
            else:
                current.append(tk)
                current_w += w

    flush()

    return lines


# =========================================================
# 背景图：优先用 memo_bg.png；没有就自动画高清背景
# =========================================================

def draw_default_notes_background() -> Image.Image:
    img = Image.new("RGB", (PAGE_W, PAGE_H), BG_COLOR)
    draw = ImageDraw.Draw(img)

    font = get_font(72, False)
    color = NOTE_YELLOW

    # 左侧返回箭头
    x = 96
    y = 185
    w = 8
    draw.line((x + 38, y - 44, x, y), fill=color, width=w)
    draw.line((x, y, x + 38, y + 44), fill=color, width=w)

    # 备忘录
    draw.text((150, 136), "备忘录", font=font, fill=color)

    # 分享图标
    sx = PAGE_W - 560
    sy = 150
    icon_w = 8

    draw.rounded_rectangle(
        (sx - 34, sy + 28, sx + 48, sy + 122),
        radius=14,
        outline=color,
        width=icon_w,
    )

    # 顶部开口
    draw.line((sx - 8, sy + 28, sx + 22, sy + 28), fill=BG_COLOR, width=18)

    # 箭头
    draw.line((sx + 8, sy - 34, sx + 8, sy + 70), fill=color, width=icon_w)
    draw.line((sx + 8, sy - 34, sx - 22, sy - 2), fill=color, width=icon_w)
    draw.line((sx + 8, sy - 34, sx + 38, sy - 2), fill=color, width=icon_w)

    # 更多图标
    mx = PAGE_W - 310
    my = 186
    r = 58
    draw.ellipse((mx - r, my - r, mx + r, my + r), outline=color, width=icon_w)

    dot_r = 8
    for dx in [-24, 0, 24]:
        draw.ellipse(
            (mx + dx - dot_r, my - dot_r, mx + dx + dot_r, my + dot_r),
            fill=color,
        )

    return img


def new_page() -> Tuple[Image.Image, ImageDraw.ImageDraw]:
    if os.path.exists(BG_IMAGE_PATH):
        bg = Image.open(BG_IMAGE_PATH).convert("RGB")
        bg = bg.resize((PAGE_W, PAGE_H), Image.LANCZOS)
        img = bg.copy()
    else:
        img = draw_default_notes_background()

    return img, ImageDraw.Draw(img)


# =========================================================
# 绘制文字：先画底线 / 高亮，再画字
# =========================================================

def get_line_max_font_size(line: List[Chunk]) -> int:
    return max(
        [c.font_size_px or DEFAULT_FONT_PX["body"] for c in line],
        default=DEFAULT_FONT_PX["body"],
    )


def draw_line(draw: ImageDraw.ImageDraw, line: List[Chunk], x: int, y: int):
    cursor_x = x

    # 第一遍：画高亮和下划线，保证置底
    for chunk in line:
        if not chunk.text:
            continue

        font_size = chunk.font_size_px or DEFAULT_FONT_PX["body"]
        font = get_font(font_size, chunk.bold)
        w = text_width(chunk.text, font)

        if chunk.highlight:
            color = HIGHLIGHT_COLORS.get(chunk.highlight, HIGHLIGHT_COLORS["yellow"])

            if chunk.highlight == "blue":
                rect = (
                    cursor_x - 18,
                    y + int(font_size * 0.04),
                    cursor_x + w + 18,
                    y + int(font_size * 1.18),
                )
                draw.rounded_rectangle(rect, radius=36, fill=color)

            elif chunk.highlight == "yellow":
                rect = (
                    cursor_x - 10,
                    y + int(font_size * 0.56),
                    cursor_x + w + 10,
                    y + int(font_size * 1.13),
                )
                draw.rounded_rectangle(rect, radius=18, fill=color)

            elif chunk.highlight == "green":
                rect = (
                    cursor_x - 10,
                    y + int(font_size * 0.68),
                    cursor_x + w + 10,
                    y + int(font_size * 1.17),
                )
                draw.rounded_rectangle(rect, radius=10, fill=color)

            else:
                rect = (
                    cursor_x - 10,
                    y + int(font_size * 0.62),
                    cursor_x + w + 10,
                    y + int(font_size * 1.14),
                )
                draw.rounded_rectangle(rect, radius=14, fill=color)

        if chunk.underline:
            underline_y = y + int(font_size * 1.04)
            draw.line(
                (cursor_x, underline_y, cursor_x + w, underline_y),
                fill=HIGHLIGHT_COLORS["pink"],
                width=18,
            )

        cursor_x += w

    # 第二遍：画文字
    cursor_x = x

    for chunk in line:
        if not chunk.text:
            continue

        font_size = chunk.font_size_px or DEFAULT_FONT_PX["body"]
        font = get_font(font_size, chunk.bold)
        w = text_width(chunk.text, font)

        draw.text(
            (cursor_x, y),
            chunk.text,
            font=font,
            fill=TEXT_COLOR,
        )

        cursor_x += w


# =========================================================
# 渲染分页
# =========================================================

def render_pages(blocks: List[ParagraphBlock]) -> List[Image.Image]:
    pages: List[Image.Image] = []

    img, draw = new_page()
    y = CONTENT_TOP

    for block in blocks:
        if block.kind == "blank":
            y += block.space_after_px
            continue

        if block.kind == "title":
            y += max(block.space_before_px, 24)
        else:
            y += block.space_before_px

        # 标题用较窄宽度，避免标题拉太开
        if block.kind == "title":
            lines = wrap_chunks(block.chunks, TITLE_W)
            draw_x = TITLE_LEFT
        else:
            lines = wrap_chunks(block.chunks, CONTENT_W)
            draw_x = LEFT

        line_height = block.line_height_px or int(DEFAULT_FONT_PX["body"] * DEFAULT_LINE_RATIO["body"])

        for line in lines:
            line_max_font = get_line_max_font_size(line)
            effective_line_height = max(line_height, int(line_max_font * 1.18))

            if y + effective_line_height > PAGE_H - BOTTOM:
                pages.append(img)
                img, draw = new_page()

                if block.kind == "title":
                    y = CONTENT_TOP + max(block.space_before_px, 24)
                else:
                    y = CONTENT_TOP + block.space_before_px

            draw_line(draw, line, draw_x, y)
            y += effective_line_height

        y += block.space_after_px

        if y > PAGE_H - BOTTOM:
            pages.append(img)
            img, draw = new_page()
            y = CONTENT_TOP

    pages.append(img)

    return pages


# =========================================================
# 导出
# =========================================================

def image_to_bytes(img: Image.Image) -> bytes:
    buf = io.BytesIO()
    img.save(buf, format="PNG", dpi=(300, 300))
    return buf.getvalue()


def make_long_image(pages: List[Image.Image]) -> Image.Image:
    w = pages[0].width
    h = pages[0].height * len(pages)

    long_img = Image.new("RGB", (w, h), BG_COLOR)

    for i, page in enumerate(pages):
        long_img.paste(page, (0, i * page.height))

    return long_img


def make_zip(pages: List[Image.Image], long_img: Image.Image) -> bytes:
    zip_buf = io.BytesIO()

    with zipfile.ZipFile(zip_buf, "w", zipfile.ZIP_DEFLATED) as z:
        for i, page in enumerate(pages, start=1):
            z.writestr(f"memo_page_{i:02d}.png", image_to_bytes(page))

        z.writestr("memo_long_image.png", image_to_bytes(long_img))

    return zip_buf.getvalue()


# =========================================================
# Streamlit 页面
# =========================================================

st.set_page_config(
    page_title="Word 转高清备忘录长图",
    page_icon="📝",
    layout="centered",
)

st.title("Word 转高清备忘录长图")
st.caption("固定 2160×2880，3:4；标题走备忘录模板，正文读取 Word 行距、段前、段后。")

uploaded_file = st.file_uploader("上传 Word 文件（.docx）", type=["docx"])

st.markdown(
    """
**Word 设置规则：**

- `标题 / 标题1`：转成备忘录大标题版式。
- `标题2`：转成一级小标题。
- `标题3`：转成二级小标题。
- 正文：尽量读取 Word 的字号、行距、段前、段后。
- 直接按 `Enter`：新段落，程序会读取段前 / 段后。
- 按 `Shift + Enter`：同段换行，只走普通行距。
- 加粗：保留加粗。
- 下划线：转成粉色荧光线。
- 黄色高亮：转成黄色荧光笔线。
- 蓝色高亮：转成蓝色圆角底块。
- 绿色高亮：转成绿色荧光线。
- 本地 Windows 默认使用微软雅黑；网页部署默认使用 Noto CJK 中文字体。
"""
)

if os.path.exists(BG_IMAGE_PATH):
    st.info("已检测到 memo_bg.png，会使用这张图作为背景。")
else:
    st.warning("没有检测到 memo_bg.png，会自动绘制一个高清备忘录背景。")

if uploaded_file:
    if st.button("生成高清图片", type="primary"):
        with tempfile.NamedTemporaryFile(delete=False, suffix=".docx") as tmp:
            tmp.write(uploaded_file.getvalue())
            tmp_path = tmp.name

        try:
            blocks = parse_docx(tmp_path)
            pages = render_pages(blocks)
            long_img = make_long_image(pages)
            zip_bytes = make_zip(pages, long_img)

            st.success(f"生成完成：共 {len(pages)} 页，单页尺寸：{PAGE_W}×{PAGE_H}")

            st.download_button(
                label="下载全部分页 PNG（ZIP）",
                data=zip_bytes,
                file_name="memo_pages_hd.zip",
                mime="application/zip",
            )

            st.download_button(
                label="下载完整长图 PNG",
                data=image_to_bytes(long_img),
                file_name="memo_long_hd.png",
                mime="image/png",
            )

            st.subheader("预览")
            for i, page in enumerate(pages, start=1):
                st.image(page, caption=f"第 {i} 页", width=360)

        finally:
            try:
                os.remove(tmp_path)
            except Exception:
                pass
