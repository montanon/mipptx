from __future__ import annotations

from enum import Enum

from pptx.enum.text import MSO_AUTO_SIZE as AS
from pptx.enum.text import MSO_VERTICAL_ANCHOR as VA
from pptx.enum.text import PP_PARAGRAPH_ALIGNMENT as PPA


class ParagraphAlignment(str, Enum):
    left = "left"
    center = "center"
    right = "right"
    justify = "justify"
    distributed = "distributed"


class VerticalAnchor(str, Enum):
    top = "top"
    middle = "middle"
    bottom = "bottom"


class AutoSize(str, Enum):
    none = "none"
    text_to_fit_shape = "text_to_fit_shape"
    shape_to_fit_text = "shape_to_fit_text"


class Underline(str, Enum):
    none = "none"
    single = "single"


class ShapeKind(str, Enum):
    text_box = "text_box"
    picture = "picture"
    table = "table"
    chart = "chart"
    unknown = "unknown"


# -- helpers to map to/from python-pptx enums (optional: best-effort)
def to_alignment_enum(pp_align) -> ParagraphAlignment | None:
    mapping = {
        PPA.LEFT: ParagraphAlignment.left,
        PPA.CENTER: ParagraphAlignment.center,
        PPA.RIGHT: ParagraphAlignment.right,
        PPA.JUSTIFY: ParagraphAlignment.justify,
        PPA.DISTRIBUTE: ParagraphAlignment.distributed,
    }
    return mapping.get(pp_align)


def from_alignment_enum(alignment: ParagraphAlignment | None):
    if alignment is None:
        return None

    reverse = {
        ParagraphAlignment.left: PPA.LEFT,
        ParagraphAlignment.center: PPA.CENTER,
        ParagraphAlignment.right: PPA.RIGHT,
        ParagraphAlignment.justify: PPA.JUSTIFY,
        ParagraphAlignment.distributed: PPA.DISTRIBUTE,
    }
    return reverse[alignment]


def to_vertical_anchor(pp_anchor) -> VerticalAnchor | None:
    mapping = {
        VA.TOP: VerticalAnchor.top,
        VA.MIDDLE: VerticalAnchor.middle,
        VA.BOTTOM: VerticalAnchor.bottom,
    }
    return mapping.get(pp_anchor)


def from_vertical_anchor(anchor: VerticalAnchor | None):
    if anchor is None:
        return None
    reverse = {
        VerticalAnchor.top: VA.TOP,
        VerticalAnchor.middle: VA.MIDDLE,
        VerticalAnchor.bottom: VA.BOTTOM,
    }
    return reverse[anchor]


def to_auto_size(pp_autosize) -> AutoSize | None:
    mapping = {
        None: AutoSize.none,
        AS.NONE: AutoSize.none,
        AS.TEXT_TO_FIT_SHAPE: AutoSize.text_to_fit_shape,
        AS.SHAPE_TO_FIT_TEXT: AutoSize.shape_to_fit_text,
    }
    return mapping.get(pp_autosize)


def from_auto_size(auto_size: AutoSize | None):
    if auto_size is None:
        return None

    reverse = {
        AutoSize.none: AS.NONE,
        AutoSize.text_to_fit_shape: AS.TEXT_TO_FIT_SHAPE,
        AutoSize.shape_to_fit_text: AS.SHAPE_TO_FIT_TEXT,
    }
    return reverse[auto_size]
