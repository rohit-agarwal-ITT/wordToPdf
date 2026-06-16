"""Run LibreOffice without showing a console window (Windows)."""
import os
import platform
import subprocess
import time
from typing import Callable, List, Optional

WINDOWS_SOFFICE_DIR = r'C:\Program Files\LibreOffice\program'


def get_soffice_path() -> str:
    if platform.system() == 'Windows':
        for name in ('soffice.com', 'soffice.exe'):
            path = os.path.join(WINDOWS_SOFFICE_DIR, name)
            if os.path.exists(path):
                return path
        return os.path.join(WINDOWS_SOFFICE_DIR, 'soffice.com')
    return 'soffice'


def _subprocess_kwargs() -> dict:
    if platform.system() != 'Windows':
        return {}
    startupinfo = subprocess.STARTUPINFO()
    startupinfo.dwFlags |= subprocess.STARTF_USESHOWWINDOW
    startupinfo.wShowWindow = subprocess.SW_HIDE
    return {
        'creationflags': subprocess.CREATE_NO_WINDOW,
        'startupinfo': startupinfo,
    }


def run_soffice(
    args: List[str],
    timeout: Optional[int] = None,
    check: bool = False,
) -> subprocess.CompletedProcess:
    cmd = [get_soffice_path()] + args
    return subprocess.run(
        cmd,
        capture_output=True,
        timeout=timeout,
        check=check,
        **_subprocess_kwargs(),
    )


def convert_docx_files_to_pdf(
    docx_paths: List[str],
    output_dir: str,
    timeout: int = 300,
    should_cancel: Optional[Callable[[], bool]] = None,
) -> None:
    """Convert Word files to PDF in headless mode without a visible terminal."""
    if should_cancel and should_cancel():
        raise RuntimeError('Conversion cancelled by user.')

    args = [
        '--headless',
        '--invisible',
        '--norestore',
        '--nologo',
        '--convert-to',
        'pdf',
        '--outdir',
        output_dir,
    ] + docx_paths
    cmd = [get_soffice_path()] + args
    proc = subprocess.Popen(
        cmd,
        stdout=subprocess.PIPE,
        stderr=subprocess.PIPE,
        **_subprocess_kwargs(),
    )
    deadline = time.time() + timeout
    try:
        while proc.poll() is None:
            if should_cancel and should_cancel():
                proc.terminate()
                try:
                    proc.wait(timeout=5)
                except subprocess.TimeoutExpired:
                    proc.kill()
                    proc.wait()
                raise RuntimeError('Conversion cancelled by user.')
            if time.time() > deadline:
                proc.kill()
                proc.wait()
                raise subprocess.TimeoutExpired(cmd, timeout)
            time.sleep(0.2)
    except Exception:
        if proc.poll() is None:
            proc.kill()
            proc.wait()
        raise

    if proc.returncode != 0:
        stderr = (proc.stderr.read() or b'').decode(errors='replace')
        raise subprocess.CalledProcessError(proc.returncode, cmd, stderr=stderr)
