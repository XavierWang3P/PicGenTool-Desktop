# 常量配置
MAX_IMAGES = 30
MARGINS = (12, 12, 12, 12)
SPACING = 6

# 图片处理配置
SUPPORTED_FORMATS = {'.png', '.jpg', '.jpeg', '.gif', '.bmp'}
IMAGE_QUALITY = 60  # 图片压缩质量
TARGET_WIDTH_CM = 6.4  # 目标图片宽度（厘米）
TARGET_HEIGHT_CM = 4.8  # 目标图片高度（厘米）
DPI = 240  # 图片分辨率

# UI样式配置
UI_STYLES = {
    'label': {
        'font-size': '13px',
        'color': '#333'
    },
    'line_edit': {
        'padding': '4px',
        'border': '1px solid #DDD',
        'border-radius': '4px',
        'background': 'white',
        'font-size': '13px',
        'height': '20px'
    },
    'line_edit_focus': {
        'border-color': '#06F'
    },
    'button': {
        'border-radius': '4px',
        'font-size': '13px',
        'border': '1px solid #DDD',
        'background': 'white',
        'height': '42px',
        'width': '100px'
    },
    'button_hover': {
        'background': '#F5F5F5',
        'border-color': '#CCC'
    },
    'generate_button': {
        'background': '#06F',
        'color': 'white'
    },
    'generate_button_hover': {
        'background': '#05C'
    },
    'clear_button': {
        'background': '#F5F5F5',
        'color': '#666'
    },
    'checkbox': {
        'font-size': '13px',
        'color': '#333'
    },
    'calendar': {
        'background': 'white',
        'tool_button': {
            'color': '#333',
            'padding': '4px'
        },
        'menu': {
            'width': '150px',
            'left': '20px',
            'color': '#333'
        },
        'spin_box': {
            'width': '60px',
            'font-size': '13px'
        },
        'table_view': {
            'selection-background-color': '#06F'
        }
    }
}