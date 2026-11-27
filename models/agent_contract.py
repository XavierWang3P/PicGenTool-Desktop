from __future__ import annotations

SCHEMA_VERSION = 1

ERROR_EXIT_CODES = {
    "INVALID_REQUEST": 2,
    "IMAGE_NOT_FOUND": 2,
    "UNSUPPORTED_FORMAT": 2,
    "IMAGE_READ_FAILED": 2,
    "TOO_MANY_IMAGES": 2,
    "TEMPLATE_MISSING": 3,
    "OUTPUT_EXISTS": 4,
    "OUTPUT_NOT_WRITABLE": 4,
    "CANCELLED": 5,
    "INTERNAL_ERROR": 1,
}

REQUEST_SCHEMA = {
    "$schema": "https://json-schema.org/draft/2020-12/schema",
    "title": "PicGenTool generation request",
    "type": "object",
    "additionalProperties": False,
    "required": [
        "schema_version",
        "title",
        "location",
        "activity_date",
        "images",
        "output",
    ],
    "properties": {
        "schema_version": {"const": SCHEMA_VERSION},
        "title": {"type": "string", "minLength": 1},
        "location": {"type": "string", "minLength": 1},
        "activity_date": {"type": "string", "format": "date"},
        "images": {
            "type": "array",
            "items": {"type": "string", "minLength": 1},
            "minItems": 1,
            "maxItems": 30,
        },
        "output": {"type": "string", "minLength": 1},
        "options": {
            "type": "object",
            "additionalProperties": False,
            "properties": {
                "if_exists": {"enum": ["fail", "skip", "overwrite"]},
                "on_image_error": {"const": "fail"},
                "open": {"const": False},
                "use_settings": {"const": False},
            },
        },
    },
}


def request_schema() -> dict:
    """Return an independent JSON-compatible copy of the public schema."""
    import copy

    return copy.deepcopy(REQUEST_SCHEMA)
