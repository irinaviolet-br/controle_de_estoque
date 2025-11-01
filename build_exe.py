"""Helper script to bundle the application as a Windows executable.

The script converts the provided PNG icon to the ICO format expected by
PyInstaller, then invokes PyInstaller with the right arguments so the
resulting executable ships with the Tkinter interface, stock spreadsheets
and the public logo.
"""
from __future__ import annotations

import os
import shutil
import subprocess
import sys
from pathlib import Path
from typing import Iterable, Tuple, TYPE_CHECKING

if TYPE_CHECKING:
    from PyInstaller import __main__ as PyInstallerMain  # type: ignore
else:  # pragma: no cover - used only for typing
    PyInstallerMain = object  # type: ignore

BASE_DIR = Path(__file__).resolve().parent
APP_SCRIPT = BASE_DIR / "main.py"
RUNTIME_LOGO = BASE_DIR / "LOGO.png"
ICON_SOURCE = BASE_DIR / "LOGO Estoque.png"
ICON_OUTPUT = BASE_DIR / "build_assets" / "LOGO_Estoque.ico"
EXECUTABLE_NAME = "Controle de Estoque"


def require_windows() -> None:
    """Abort if the script is executed on a non-Windows platform."""
    if os.name != "nt":
        raise SystemExit(
            "Esta etapa precisa ser executada no Windows para gerar um .exe."
        )


def ensure_dependencies() -> tuple[str | None, PyInstallerMain | None]:
    """Ensure optional build-time dependencies are available."""

    missing: list[str] = []

    try:
        import PIL  # type: ignore  # noqa: F401
    except ModuleNotFoundError:
        missing.append("pillow")

    try:
        import PyInstaller.__main__ as pyinstaller_main  # type: ignore
    except ModuleNotFoundError:
        pyinstaller_main = None
    else:
        # keep reference type-friendly for callers
        pyinstaller_main = pyinstaller_main

    pyinstaller_cli = shutil.which("pyinstaller")

    if pyinstaller_cli is None and pyinstaller_main is None:
        missing.append("pyinstaller")

    if missing:
        raise SystemExit(
            "Dependências ausentes: " + ", ".join(missing)
            + "\nInstale-as com 'pip install "
            + " ".join(missing)
            + "'."
        )

    return pyinstaller_cli, pyinstaller_main


def convert_icon(source: Path, destination: Path) -> None:
    """Convert the PNG icon provided by the user to ICO format."""
    from PIL import Image

    destination.parent.mkdir(parents=True, exist_ok=True)

    with Image.open(source) as image:
        if image.mode not in ("RGBA", "RGB"):
            image = image.convert("RGBA")
        sizes: Iterable[Tuple[int, int]] = [
            (16, 16),
            (32, 32),
            (48, 48),
            (64, 64),
            (128, 128),
            (256, 256),
        ]
        image.save(destination, format="ICO", sizes=list(sizes))


def run_pyinstaller(
    icon_path: Path,
    pyinstaller_cli: str | None,
    pyinstaller_main: PyInstallerMain | None,
) -> None:
    """Invoke PyInstaller with the required parameters."""
    data_sep = ";" if os.name == "nt" else ":"
    cmd_args = [
        "--noconfirm",
        "--clean",
        "--windowed",
        "--name",
        EXECUTABLE_NAME,
        f"--icon={icon_path}",
        "--hidden-import=msoffcrypto",
        f"--add-data={RUNTIME_LOGO}{data_sep}.",
        str(APP_SCRIPT),
    ]

    if pyinstaller_cli is not None:
        cmd = ["pyinstaller", *cmd_args]
        print("Executando:", " ".join(cmd))
        subprocess.run(cmd, check=True)
    else:
        assert pyinstaller_main is not None
        print("Executando via PyInstaller.__main__")
        pyinstaller_main.run(cmd_args)


def main() -> None:
    require_windows()
    pyinstaller_cli, pyinstaller_main = ensure_dependencies()

    if not APP_SCRIPT.exists():
        raise SystemExit("main.py não foi encontrado.")

    if not RUNTIME_LOGO.exists():
        raise SystemExit("LOGO.png não foi encontrado na pasta do aplicativo.")

    if not ICON_SOURCE.exists():
        raise SystemExit(
            "LOGO Estoque.png não foi encontrado. Coloque o arquivo ao lado do script."
        )

    convert_icon(ICON_SOURCE, ICON_OUTPUT)
    run_pyinstaller(ICON_OUTPUT, pyinstaller_cli, pyinstaller_main)

    dist_path = BASE_DIR / "dist" / f"{EXECUTABLE_NAME}.exe"
    if dist_path.exists():
        print("Executável gerado em:", dist_path)
    else:
        print(
            "A compilação terminou, mas o executável não foi encontrado em 'dist/'. "
            "Verifique a saída do PyInstaller para mais detalhes."
        )


if __name__ == "__main__":
    main()
