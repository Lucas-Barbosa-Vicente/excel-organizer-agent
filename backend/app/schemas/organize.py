from typing import List, Optional, Dict, Literal
from pydantic import BaseModel


class SortRule(BaseModel):
    column: str
    direction: Literal["asc", "desc"] = "asc"


class ColorRule(BaseModel):
    column: str
    operator: Literal[
        "equals", "not_equals", "starts_with", "ends_with",
        "contains", "greater_than", "less_than", "is_empty", "is_not_empty"
    ]
    value: Optional[str] = None
    color: Literal["yellow", "green", "red", "blue", "orange", "gray", "pink"]


class OrganizeParameters(BaseModel):
    sort_by: Optional[List[SortRule]] = None
    remove_duplicates: Optional[bool] = False
    duplicate_columns: Optional[List[str]] = None
    keep_duplicate: Optional[Literal["first", "last"]] = "first"
    standardize_text: Optional[Dict[str, Literal["upper", "lower", "capitalize", "strip"]]] = None
    color_rules: Optional[List[ColorRule]] = None
    split_by_category: Optional[str] = None
    keep_original_sheet: Optional[bool] = True
    create_summary_sheet: Optional[bool] = False
    natural_language_instruction: Optional[str] = None
    force_override: Optional[bool] = False


class OrganizeRequest(BaseModel):
    parameters: Optional[OrganizeParameters] = None
    profile_id: Optional[int] = None


class OrganizeResponse(BaseModel):
    status: str
    message: str
    transformations_applied: List[str]
    rows_before: int
    rows_after: int
    sheets_created: List[str]
    download_token: Optional[str] = None
    requires_confirmation: bool = False
    existing_customizations: List[str] = []
