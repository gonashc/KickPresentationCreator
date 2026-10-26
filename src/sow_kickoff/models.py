from __future__ import annotations

from typing import Literal

from pydantic import BaseModel, Field


class SourceDocument(BaseModel):
    text: str
    paragraph_count: int
    table_count: int
    truncated: bool = False


class TemplateTarget(BaseModel):
    kind: Literal["shape", "table_cell"]
    shape_id: int
    shape_name: str
    current_text: str
    placeholder_type: str
    row_index: int = -1
    column_index: int = -1


class SlideInventory(BaseModel):
    slide_number: int
    title: str
    targets: list[TemplateTarget]


class TemplateInventory(BaseModel):
    slide_count: int
    editable_target_count: int
    slides: list[SlideInventory]


class SlideEdit(BaseModel):
    slide_number: int = Field(description="One-based slide number from the supplied inventory")
    target_kind: Literal["shape", "table_cell"]
    shape_id: int
    row_index: int = Field(description="Zero-based table row, or -1 for a shape")
    column_index: int = Field(description="Zero-based table column, or -1 for a shape")
    text: str = Field(description="Replacement text. Use an empty string when evidence is missing.")
    source_references: list[str] = Field(description="SoW paragraph/table references supporting the text")


class DeckPlan(BaseModel):
    executive_summary: str
    edits: list[SlideEdit]
    warnings: list[str]


class RenderResult(BaseModel):
    presentation_bytes: bytes
    applied_count: int
    skipped_count: int
    applied_updates: list[dict]
    skipped_updates: list[dict]

