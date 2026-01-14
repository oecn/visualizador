from __future__ import annotations

import sys
from pathlib import Path


def bootstrap(argv: list[str] | None = None) -> int:
    project_root = Path(__file__).resolve().parent.parent
    if str(project_root) not in sys.path:
        sys.path.insert(0, str(project_root))

    try:
        from ventas_app.main import main as run_app
    except ImportError:
        print("No se pudo importar ventas_app. Ejecuta instalar_appnueva.ps1 para crear el entorno.")
        return 1

    return run_app(argv)


if __name__ == "__main__":
    sys.exit(bootstrap(sys.argv[1:]))
