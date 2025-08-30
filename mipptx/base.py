from __future__ import annotations

from pydantic import BaseModel as _PydanticBaseModel, ConfigDict


class JsonModel(_PydanticBaseModel):
    """Base model with JSON helpers.

    Provides `to_json()` and `from_json()` thin wrappers around Pydantic v2
    `model_dump_json` and `model_validate_json`.
    """

    model_config = ConfigDict()

    def to_json(self, **kwargs) -> str:
        return self.model_dump_json(**kwargs)

    @classmethod
    def from_json(cls, data: str):
        return cls.model_validate_json(data)

