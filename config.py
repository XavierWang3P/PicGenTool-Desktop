APP_TITLE = "活动照片文档生成"
APP_VERSION = "3.1.1"
MAX_IMAGES = 30
MARGINS = (20, 18, 20, 18)
SPACING = 12
LOG_LEVEL = "INFO"
LOG_FILE_NAME = "app.log"
LOG_MAX_BYTES = 5 * 1024 * 1024
LOG_BACKUP_COUNT = 3
DEFAULT_ACTIVITY_LOCATION = ""

SUPPORTED_FORMATS = {".png", ".jpg", ".jpeg", ".gif", ".bmp", ".heic", ".heif"}
IMAGE_QUALITY = 60
IMAGE_OPTIMIZE = False
MAX_IMAGE_WORKERS = 4
TARGET_WIDTH_CM = 6.4
TARGET_HEIGHT_CM = 4.8
DPI = 240

APP_STYLES = {
    "label": {
        "font-size": "14px",
        "font-weight": "600",
        "color": "#2F3437",
    },
    "line_edit": {
        "padding": "9px 12px",
        "border": "1px solid #D7DDE5",
        "border-radius": "10px",
        "background": "white",
        "font-size": "14px",
        "selection-background-color": "#0F6FFF",
    },
    "line_edit_focus": {
        "border": "2px solid #0F6FFF",
        "padding": "8px 11px",
    },
    "line_edit_disabled": {
        "border-color": "#E3E8EF",
        "background": "#F5F7FA",
        "color": "#98A2AE",
    },
    "button": {
        "border-radius": "12px",
        "font-size": "14px",
        "font-weight": "600",
        "border": "1px solid #D7DDE5",
        "background": "white",
        "height": "44px",
        "min-width": "96px",
        "padding": "0 16px",
    },
    "button_hover": {
        "background": "#F6F9FC",
        "border-color": "#BFC9D6",
    },
    "button_focus": {
        "border": "2px solid #0F6FFF",
        "padding": "0 15px",
    },
    "button_disabled": {
        "background": "#F1F4F8",
        "border-color": "#E3E8EF",
        "color": "#6B7785",
    },
    "generate_button": {
        "background": "#0F6FFF",
        "color": "white",
        "border": "1px solid #0F6FFF",
    },
    "generate_button_hover": {
        "background": "#0B5DDB",
    },
    "generate_button_focus": {
        "border": "2px solid #083E99",
        "padding": "0 15px",
    },
    "generate_button_disabled": {
        "background": "#B7C8E8",
        "border-color": "#B7C8E8",
        "color": "#31415C",
    },
    "clear_button": {
        "background": "#F3F5F7",
        "color": "#58636F",
        "border": "1px solid #D7DDE5",
    },
    "cancel_button": {
        "background": "#FFF4F2",
        "color": "#B3421E",
        "border": "1px solid #F0C4B8",
    },
    "checkbox": {
        "font-size": "14px",
        "color": "#2F3437",
        "spacing": "8px",
    },
}
