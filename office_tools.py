# office_tools.py
import subprocess, shutil, tempfile, os
from pathlib import Path

def _soffice_path() -> str:
    # 1) If bundled portable copy exists, use it
    portable = Path(__file__).with_name("bin") / "LibreOffice" / "program" / "soffice.exe"
    if portable.exists():
        return str(portable)
    # 2) Else look in PATH
    exe = shutil.which("soffice") or shutil.which("soffice.exe")
    if not exe:
        raise RuntimeError("LibreOffice (soffice) not found. Install it or bundle a portable copy.")
    return exe

def convert_with_soffice(src: str, target_ext: str) -> str:
    """
    Convert a document to another format using LibreOffice headless.
    Example: DOCX -> PDF  or  PDF -> DOCX
    """
    soffice = _soffice_path()
    src_path = Path(src).resolve()
    out_dir = Path(tempfile.mkdtemp())
    cmd = [
        soffice, "--headless", "--convert-to", target_ext, "--outdir", str(out_dir), str(src_path)
    ]
    result = subprocess.run(cmd, stdout=subprocess.PIPE, stderr=subprocess.PIPE, text=True)
    if result.returncode != 0:
        raise RuntimeError(f"LibreOffice failed:\n{result.stderr.strip() or result.stdout.strip()}")

    # LibreOffice keeps the same base filename, only extension changes
    output_file = out_dir / f"{src_path.stem}.{target_ext}"
    if not output_file.exists():
        raise RuntimeError(f"Conversion failed – output not produced.\n"
                        f"Tried to create: {output_file}")

    # Move beside source (or wherever caller wants later)
    final_path = src_path.with_suffix(f".{target_ext}")
    output_file.replace(final_path)
    return str(final_path)
