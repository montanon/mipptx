from __future__ import annotations

from typing import Annotated, List, Literal, Optional, Union

from pydantic import ConfigDict, Field, field_validator, model_validator

from .base import JsonModel
from .charts import ChartModel
from .enums import ShapeKind
from .text import TextFrameModel
from .utils import emu_to_pt, pt_to_emu


class BaseShapeModel(JsonModel):
    model_config = ConfigDict(extra="forbid")

    type: Literal[ShapeKind.unknown] = Field(
        ShapeKind.unknown, description="Discriminator for union"
    )
    name: Optional[str] = None
    # positions/sizes in points for JSON-friendliness
    left_pt: float = Field(...)
    top_pt: float = Field(...)
    width_pt: float = Field(..., ge=0)
    height_pt: float = Field(..., ge=0)
    rotation: Optional[float] = Field(0, ge=-360.0, le=360.0)

    @classmethod
    def from_pptx(cls, shape) -> "BaseShapeModel":
        left = emu_to_pt(shape.left) or 0.0
        top = emu_to_pt(shape.top) or 0.0
        width = emu_to_pt(shape.width) or 0.0
        height = emu_to_pt(shape.height) or 0.0
        base_kwargs = dict(
            name=getattr(shape, "name", None),
            left_pt=left,
            top_pt=top,
            width_pt=width,
            height_pt=height,
            rotation=getattr(shape, "rotation", 0.0) or 0.0,
        )

        # text box
        if getattr(shape, "has_text_frame", False):
            tf = TextFrameModel.from_pptx(shape.text_frame)
            return TextBoxModel(type=ShapeKind.text_box, text_frame=tf, **base_kwargs)

        # chart
        if getattr(shape, "has_chart", False):
            try:
                return ChartModel.from_pptx(shape)
            except Exception:
                pass

        # table
        if getattr(shape, "has_table", False):
            # extract rows/columns via this module's TableModel helper
            tbl_model = TableModel.from_pptx(shape.table)
            return TableModel(
                type=ShapeKind.table,
                **base_kwargs,
                rows=tbl_model.rows,
                col_widths_pt=tbl_model.col_widths_pt,
                first_row=tbl_model.first_row,
                last_row=tbl_model.last_row,
                first_col=tbl_model.first_col,
                last_col=tbl_model.last_col,
                horz_banding=tbl_model.horz_banding,
                vert_banding=tbl_model.vert_banding,
                row_heights_pt=tbl_model.row_heights_pt,
            )  # type: ignore[arg-type]

        # picture
        if shape.__class__.__name__ == "Picture":
            # python-pptx doesn't expose source path reliably once embedded; emit as unknown
            return PictureModel(type=ShapeKind.picture, image_path=None, **base_kwargs)

        # fallback
        return BaseShapeModel(**base_kwargs)

    def apply_geometry(self, shp) -> None:
        shp.left = pt_to_emu(self.left_pt)
        shp.top = pt_to_emu(self.top_pt)
        shp.width = pt_to_emu(self.width_pt)
        shp.height = pt_to_emu(self.height_pt)
        if self.rotation is not None:
            try:
                shp.rotation = float(self.rotation)
            except Exception:
                pass


class TextBoxModel(BaseShapeModel):
    type: Literal[ShapeKind.text_box] = ShapeKind.text_box
    text_frame: TextFrameModel = Field(default_factory=TextFrameModel)

    def apply_to_slide(self, slide) -> None:
        # add text box and apply content
        left = pt_to_emu(self.left_pt)
        top = pt_to_emu(self.top_pt)
        width = pt_to_emu(self.width_pt)
        height = pt_to_emu(self.height_pt)
        tb = slide.shapes.add_textbox(left, top, width, height)
        if self.name:
            tb.name = self.name
        self.text_frame.to_pptx(tb.text_frame)
        if self.rotation:
            tb.rotation = float(self.rotation)


class PictureModel(BaseShapeModel):
    type: Literal[ShapeKind.picture] = ShapeKind.picture
    image_path: Optional[str] = Field(
        None,
        description="Path to image file to embed. When created from an existing PPTX the path is unknown and set to None.",
    )

    def apply_to_slide(self, slide) -> None:
        if not self.image_path:
            # cannot rehydrate without a source; skip
            return
        left = pt_to_emu(self.left_pt)
        top = pt_to_emu(self.top_pt)
        width = pt_to_emu(self.width_pt)
        height = pt_to_emu(self.height_pt)
        pic = slide.shapes.add_picture(
            self.image_path, left, top, width=width, height=height
        )
        if self.name:
            pic.name = self.name
        if self.rotation:
            pic.rotation = float(self.rotation)


class TableModel(BaseShapeModel):
    type: Literal[ShapeKind.table] = ShapeKind.table
    rows: List[List[TextFrameModel | str]] = Field(
        default_factory=list,
        description="2D grid of cell content (TextFrameModel or str)",
    )
    col_widths_pt: Optional[List[float]] = None
    # Table-level style flags aligning with python-pptx Table API
    first_row: Optional[bool] = None
    last_row: Optional[bool] = None
    first_col: Optional[bool] = None
    last_col: Optional[bool] = None
    horz_banding: Optional[bool] = None
    vert_banding: Optional[bool] = None
    # Optional explicit row heights (in points) matching `rows` length
    row_heights_pt: Optional[List[float]] = None

    @classmethod
    def from_pptx(cls, tbl) -> "TableModel":
        rows: List[List[TextFrameModel | str]] = []
        for r in tbl.rows:
            row: List[TextFrameModel | str] = []
            for c in r.cells:
                row.append(TextFrameModel.from_pptx(c.text_frame))
            rows.append(row)
        col_widths_pt = [emu_to_pt(col.width) or 0.0 for col in tbl.columns]
        # optional flags if available on table
        first_row = getattr(tbl, "first_row", None)
        last_row = getattr(tbl, "last_row", None)
        first_col = getattr(tbl, "first_col", None)
        last_col = getattr(tbl, "last_col", None)
        horz_banding = getattr(tbl, "horz_banding", None)
        vert_banding = getattr(tbl, "vert_banding", None)
        # row heights
        try:
            row_heights_pt = [
                emu_to_pt(r.height) or 0.0 for r in getattr(tbl, "rows", [])
            ]
        except Exception:
            row_heights_pt = None
        # geometry is applied by parent factory
        return cls(
            rows=rows,
            col_widths_pt=col_widths_pt,
            first_row=first_row,
            last_row=last_row,
            first_col=first_col,
            last_col=last_col,
            horz_banding=horz_banding,
            vert_banding=vert_banding,
            row_heights_pt=row_heights_pt,
            left_pt=0,
            top_pt=0,
            width_pt=0,
            height_pt=0,
        )

    def apply_to_slide(self, slide) -> None:
        if not self.rows:
            return
        n_rows = len(self.rows)
        n_cols = max(len(r) for r in self.rows)
        left = pt_to_emu(self.left_pt)
        top = pt_to_emu(self.top_pt)
        width = pt_to_emu(self.width_pt)
        height = pt_to_emu(self.height_pt)
        shape = slide.shapes.add_table(n_rows, n_cols, left, top, width, height)
        if self.name:
            shape.name = self.name
        tbl = shape.table
        # column widths
        if self.col_widths_pt and len(self.col_widths_pt) == n_cols:
            for idx, w in enumerate(self.col_widths_pt):
                tbl.columns[idx].width = pt_to_emu(w)
        # table flags
        for attr in (
            "first_row",
            "last_row",
            "first_col",
            "last_col",
            "horz_banding",
            "vert_banding",
        ):
            val = getattr(self, attr, None)
            if val is not None and hasattr(tbl, attr):
                try:
                    setattr(tbl, attr, bool(val))
                except Exception:
                    pass
        # row heights
        if self.row_heights_pt and len(self.row_heights_pt) == n_rows:
            try:
                for ridx, hpt in enumerate(self.row_heights_pt):
                    tbl.rows[ridx].height = pt_to_emu(hpt)
            except Exception:
                pass
        # content
        for r_idx, row in enumerate(self.rows):
            for c_idx, cell_content in enumerate(row):
                cell = tbl.cell(r_idx, c_idx)
                tf = cell.text_frame
                if isinstance(cell_content, str):
                    tf.text = cell_content
                elif isinstance(cell_content, TextFrameModel):
                    cell_content.to_pptx(tf)
                else:
                    # If a plain dict sneaks in, coerce to model
                    TextFrameModel.model_validate(cell_content).to_pptx(tf)
        if self.rotation:
            shape.rotation = float(self.rotation)


# Use a discriminated union for reliable JSON (de)serialization by `type`
ShapeModel = Annotated[
    Union[TextBoxModel, PictureModel, TableModel, ChartModel, BaseShapeModel],
    Field(discriminator="type"),
]
