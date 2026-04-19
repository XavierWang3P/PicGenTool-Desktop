from __future__ import annotations

import sys

from utils.logging_utils import setup_logging


def main() -> int:
    logger = setup_logging()
    try:
        from ui.main_window import MainWindow, create_application

        app = create_application()
        window = MainWindow()
        window.show()
        return app.exec()
    except Exception as exc:
        logger.critical("应用崩溃: %s", exc, exc_info=True)
        raise


if __name__ == "__main__":
    sys.exit(main())
