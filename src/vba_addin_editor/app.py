"""Application bootstrap: logging setup and entry (plan 20)."""

from __future__ import annotations

import logging
import sys
from logging.handlers import RotatingFileHandler

from vba_addin_editor.platform.paths import log_dir
from vba_addin_editor.version import APP_NAME, PYOPENVBA_PIN, VERSION


def setup_logging() -> logging.Logger:
    log_dir().mkdir(parents=True, exist_ok=True)
    logger = logging.getLogger("vba_addin_editor")
    if logger.handlers:
        return logger
    handler = RotatingFileHandler(
        log_dir() / "vba_addin_editor.log",
        maxBytes=2 * 1024 * 1024,
        backupCount=4,
        encoding="utf-8",
    )
    handler.setFormatter(
        logging.Formatter("%(asctime)s %(levelname)s %(name)s %(message)s")
    )
    logger.addHandler(handler)
    logger.setLevel(logging.INFO)
    logger.info("%s %s starting; pyopenvba pin=%s; python=%s",
                APP_NAME, VERSION, PYOPENVBA_PIN, sys.version.split()[0])
    return logger


def main() -> int:
    setup_logging()
    from vba_addin_editor.__main__ import main as cli_main

    return cli_main()


if __name__ == "__main__":
    sys.exit(main())
