"""In-memory model of a Word document after style resolution."""
from __future__ import annotations

from dataclasses import dataclass, field

from ..theme import Color


@dataclass(frozen=True)
class RunStyle:
    font_ascii: str = "Times New Roman"
    font_hansi: str = "Times New Roman"
    font_ea: str = "Times New Roman"
    font_cs: str = "Times New Roman"
    size: float = 10.0
    size_cs: float = 10.0
    bold: bool = False
    italic: bool = False
    bold_cs: bool = False
    italic_cs: bool = False
    underline: str | None = None
    u_color: Color | None = None
    strike: bool = False
    dstrike: bool = False
    color: Color | None = None
    highlight: Color | None = None
    shading: Color | None = None
    vert: str | None = None
    caps: bool = False
    small_caps: bool = False
    spacing: float = 0.0
    position: float = 0.0
    scale: float = 100.0
    hidden: bool = False
    hint: str | None = None
    border: object = None
    kern: float = 0.0  # kerning applies from this font size (pt); 0 = off
    math_script: bool = False  # sub/superscript inside an equation: does not size the line


@dataclass
class Text:
    text: str
    style: RunStyle
    link: str | None = None


@dataclass
class Tab:
    style: RunStyle
    link: str | None = None


@dataclass
class Break:
    kind: str  # line | page | column
    style: RunStyle


@dataclass
class Picture:
    data: bytes | None
    width: float
    height: float
    style: RunStyle
    link: str | None = None
    crop: tuple | None = None
    # extra layout space around the picture (effectExtent) l, t, r, b
    extent: tuple = (0.0, 0.0, 0.0, 0.0)
    shape: "Shape | None" = None


@dataclass
class Field:
    kind: str  # PAGE | NUMPAGES | SECTIONPAGES
    style: RunStyle
    fmt: str | None = None
    fallback: str = "1"
    picture: str | None = None  # numeric picture switch (\# "00")


@dataclass
class MathItem:
    """An Office Math equation (m:oMath) laid out as one box."""
    element: object
    style: RunStyle
    display: bool = False


@dataclass
class Separator:
    """Footnote separator line."""
    style: RunStyle


@dataclass
class Bookmark:
    name: str


@dataclass
class FootnoteRef:
    id: str
    text: str
    style: RunStyle
    kind: str = "footnote"


@dataclass
class Shape:
    """A DrawingML shape (optionally a text box)."""
    width: float
    height: float
    fill: Color | None = None
    line: Color | None = None
    line_width: float = 0.75
    geom: str = "rect"
    blocks: list = field(default_factory=list)
    insets: tuple = (7.2, 3.6, 7.2, 3.6)  # l, t, r, b
    v_anchor: str = "t"
    data: bytes | None = None  # picture fill
    rotation: float = 0.0
    autofit: bool = False  # grow to fit the text (spAutoFit)
    children: list = field(default_factory=list)  # group members: (dx, dy, Shape|Picture)
    vert: str | None = None  # text direction of a text box: vert | vert270
    flip_x: bool = False
    flip_y: bool = False
    dash: tuple | None = None
    start_arrow: str | bool = False  # VML arrow kind: block | oval | ...
    end_arrow: str | bool = False
    points: list = field(default_factory=list)  # geom "path": [(x, y)] in pt
    closed: bool = False


@dataclass
class Floating:
    """An anchored (floating) drawing object."""
    width: float
    height: float
    h_rel: str
    h_align: str | None
    h_off: float
    v_rel: str
    v_align: str | None
    v_off: float
    wrap: str  # none | square | tight | through | topAndBottom
    behind: bool
    dist: tuple  # t, b, l, r
    picture: Picture | None = None
    shape: Shape | None = None
    z: int = 0
    h_pct: float | None = None  # wp14 percentage offsets (fraction of the reference)
    v_pct: float | None = None
    size_pct: dict = field(default_factory=dict)  # "w"/"h" -> (relativeFrom, fraction)
    dropcap: int = 0  # drop cap frame: number of text lines it spans (sized at layout)


@dataclass
class Paragraph:
    ppr: dict
    items: list
    mark: RunStyle
    style_id: str | None = None
    label: tuple | None = None  # (text, RunStyle, level) for numbered paragraphs
    sect: "Section | None" = None
    in_table: bool = False
    mark_hidden: bool = False  # hidden paragraph mark: runs on into the next paragraph


@dataclass
class Cell:
    tcpr: dict
    blocks: list
    col: int = 0
    span: int = 1
    vmerge: str | None = None  # None | restart | continue
    row_span: int = 1


@dataclass
class Row:
    trpr: dict
    cells: list


@dataclass
class Table:
    tblpr: dict
    grid: list
    rows: list
    compat_legacy: bool = False
    float: dict | None = None  # tblpPr positioning for floating tables


@dataclass
class Section:
    page_w: float = 612.0
    page_h: float = 792.0
    margin_top: float = 72.0
    margin_bottom: float = 72.0
    margin_left: float = 72.0
    margin_right: float = 72.0
    header_dist: float = 36.0
    footer_dist: float = 36.0
    gutter: float = 0.0
    cols: int = 1
    col_space: float = 36.0
    col_widths: list | None = None
    col_sep: bool = False
    title_pg: bool = False
    start: str = "nextPage"
    headers: dict = field(default_factory=dict)  # kind -> blocks
    footers: dict = field(default_factory=dict)
    pg_start: int | None = None
    pg_fmt: str = "decimal"
    v_align: str = "top"
    blocks: list = field(default_factory=list)
    borders: dict = field(default_factory=dict)
