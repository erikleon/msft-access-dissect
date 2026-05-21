"""
COMSession: manages the Access.Application COM object lifecycle.

Responsibilities:
  - Preflight checks (bitness, ACE driver, Access COM registration)
  - Opening and closing the Access application safely
  - Providing references to CurrentDb(), VBE, and CurrentProject
  - Ensuring Application.Quit() always runs, even on failure

Critical design: COMSession is injected into all extractors.
Extractors never call win32com.client.Dispatch() themselves.
This enables unit testing with mock COM objects (no Access required).
"""

from __future__ import annotations

import os
import platform
import struct
import subprocess
import sys
import winreg
from pathlib import Path
from typing import Any

from access_dissect.utils.errors import (
    AccessNotInstalledError,
    BitnessError,
    COMSessionError,
    DatabaseOpenError,
)

# Late import — only available on Windows with pywin32 installed
try:
    import pywintypes
    import win32com.client
    import win32process

    _PYWIN32_AVAILABLE = True
except ImportError:
    _PYWIN32_AVAILABLE = False


# ---------------------------------------------------------------------------
# Preflight checks (run before any COM interaction)
# ---------------------------------------------------------------------------


def _python_is_64bit() -> bool:
    return struct.calcsize("P") == 8


def _check_pywin32() -> None:
    if not _PYWIN32_AVAILABLE:
        raise AccessNotInstalledError(
            "pywin32 is not installed. Run: pip install pywin32\n"
            "pywin32 is required for COM automation on Windows."
        )


def _find_ace_driver_path() -> str | None:
    """Read the ACE ODBC driver path from the Windows registry."""
    driver_keys = [
        r"SOFTWARE\ODBC\ODBCINST.INI\Microsoft Access Driver (*.mdb, *.accdb)",
        r"SOFTWARE\WOW6432Node\ODBC\ODBCINST.INI\Microsoft Access Driver (*.mdb, *.accdb)",
    ]
    for subkey in driver_keys:
        try:
            with winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, subkey) as key:
                driver_path, _ = winreg.QueryValueEx(key, "Driver")
                return str(driver_path)
        except OSError:
            continue
    return None


def _check_bitness() -> None:
    """
    Verify that the Python interpreter bitness matches the ACE driver bitness.

    If they don't match, pyodbc.connect() will fail with IM014.
    win32com may or may not work (COM routing can bridge bitness in some scenarios,
    but it's unreliable).
    """
    driver_path = _find_ace_driver_path()
    if driver_path is None:
        # ACE driver not found — we'll warn at connection time
        return

    python_64 = _python_is_64bit()
    # Check if the driver DLL lives in a 32-bit location
    driver_lower = driver_path.lower()
    driver_32 = "syswow64" in driver_lower or "program files (x86)" in driver_lower
    driver_64 = not driver_32

    if python_64 and driver_32:
        raise BitnessError(
            "Bitness mismatch: 64-bit Python is trying to use the 32-bit ACE driver.\n\n"
            f"  Python: 64-bit\n"
            f"  ACE driver: 32-bit ({driver_path})\n\n"
            "Fix options:\n"
            "  Option A (recommended): Install the 64-bit ACE redistributable:\n"
            "    https://www.microsoft.com/en-us/download/details.aspx?id=54920\n"
            "  Option B: Use 32-bit Python (not recommended for new projects)\n\n"
            "Note: If you have 32-bit Microsoft Office installed, you cannot install\n"
            "the 64-bit ACE redistributable alongside it. Consider using Office 365\n"
            "(64-bit) or the standalone ACE redistributable on a machine without Office."
        )
    if not python_64 and driver_64:
        raise BitnessError(
            "Bitness mismatch: 32-bit Python is trying to use the 64-bit ACE driver.\n\n"
            f"  Python: 32-bit\n"
            f"  ACE driver: 64-bit ({driver_path})\n\n"
            "Fix: Use 64-bit Python."
        )


def _check_access_com_registered() -> None:
    """Check that Access.Application is registered in COM."""
    access_keys = [
        r"SOFTWARE\Classes\Access.Application",
        r"SOFTWARE\Classes\Access.Application.16",
        r"SOFTWARE\Classes\Access.Application.15",
        r"SOFTWARE\WOW6432Node\Classes\Access.Application",
    ]
    for subkey in access_keys:
        try:
            with winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, subkey):
                return  # found
        except OSError:
            continue

    raise AccessNotInstalledError(
        "Microsoft Access is not installed or not registered in COM.\n\n"
        "access-dissect requires Microsoft Access (any version from 2010+) or\n"
        "the free Microsoft Access Runtime:\n"
        "  https://support.microsoft.com/en-us/office/download-and-install-microsoft-access-runtime\n\n"
        "Note: The ACE ODBC redistributable alone is not sufficient for full extraction.\n"
        "Full extraction (forms, reports, VBA) requires the Access application."
    )


def run_preflight_checks() -> None:
    """
    Run all startup checks before attempting COM interaction.
    Raises a descriptive error if any check fails.
    Call this once at the start of the `extract` command.
    """
    if platform.system() != "Windows":
        raise AccessNotInstalledError(
            "access-dissect extraction requires Windows.\n"
            "The render and analyze commands can run on any platform using\n"
            "a catalog.json file created on Windows."
        )
    _check_pywin32()
    _check_bitness()
    _check_access_com_registered()


# ---------------------------------------------------------------------------
# COMSession
# ---------------------------------------------------------------------------


class COMSession:
    """
    Manages the Access.Application COM object for the duration of an extraction.

    Usage:
        with COMSession(file_path, password="secret") as session:
            db = session.current_db
            vbe = session.vbe
            project = session.current_project
            # ... pass session to extractors

    The Access application is always quit in __exit__, even on error.
    """

    def __init__(
        self,
        file_path: str | Path,
        password: str | None = None,
        visible: bool = False,
    ) -> None:
        self.file_path = Path(file_path).resolve()
        self.password = password
        self.visible = visible

        self._app: Any = None          # Access.Application COM object
        self._current_db: Any = None   # CurrentDb() DAO Database object
        self._vbe: Any = None          # Application.VBE object
        self._current_project: Any = None  # Application.CurrentProject

        self._msaccess_pid: int | None = None

    # ------------------------------------------------------------------
    # Context manager protocol
    # ------------------------------------------------------------------

    def __enter__(self) -> "COMSession":
        self._open()
        return self

    def __exit__(self, exc_type: Any, exc_val: Any, exc_tb: Any) -> bool:
        self._close()
        return False  # don't suppress exceptions

    # ------------------------------------------------------------------
    # Properties for extractors to use
    # ------------------------------------------------------------------

    @property
    def app(self) -> Any:
        self._assert_open()
        return self._app

    @property
    def current_db(self) -> Any:
        self._assert_open()
        return self._current_db

    @property
    def vbe(self) -> Any:
        self._assert_open()
        return self._vbe

    @property
    def current_project(self) -> Any:
        self._assert_open()
        return self._current_project

    def is_alive(self) -> bool:
        """Ping the COM session with a harmless property read."""
        try:
            _ = self._app.Version
            return True
        except Exception:
            return False

    # ------------------------------------------------------------------
    # Internal open / close
    # ------------------------------------------------------------------

    def _open(self) -> None:
        if not self.file_path.exists():
            raise DatabaseOpenError(f"File not found: {self.file_path}")

        if self.file_path.suffix.lower() not in (".accdb", ".mdb"):
            raise DatabaseOpenError(
                f"File '{self.file_path.name}' does not appear to be an Access database. "
                f"Expected .accdb or .mdb extension."
            )

        try:
            self._app = win32com.client.Dispatch("Access.Application")
        except Exception as e:
            raise AccessNotInstalledError(
                f"Could not create Access.Application COM object: {e}\n"
                "Ensure Microsoft Access or the Access Runtime is installed."
            ) from e

        self._app.Visible = self.visible

        # Record the MSACCESS.EXE process ID so we can forcefully kill it
        # if the COM session becomes unresponsive.
        try:
            handle = self._app.hWndAccessApp()
            if handle:
                _, pid = win32process.GetWindowThreadProcessId(handle)
                self._msaccess_pid = pid
        except Exception:
            pass  # PID tracking is best-effort

        try:
            # OpenCurrentDatabase(Filepath, Exclusive, bstrPassword)
            if self.password:
                self._app.OpenCurrentDatabase(str(self.file_path), False, self.password)
            else:
                self._app.OpenCurrentDatabase(str(self.file_path), False)
        except pywintypes.com_error as e:
            self._force_close()
            hresult = e.hresult if hasattr(e, "hresult") else 0
            if hresult == -2147352567:  # generic dispatch error
                raise DatabaseOpenError(
                    f"Could not open '{self.file_path.name}'. "
                    "If the database is password-protected, provide --password. "
                    f"(COM error: {e})"
                ) from e
            raise DatabaseOpenError(
                f"COM error opening '{self.file_path.name}': {e}"
            ) from e
        except Exception as e:
            self._force_close()
            raise DatabaseOpenError(
                f"Unexpected error opening '{self.file_path.name}': {e}"
            ) from e

        try:
            self._current_db = self._app.CurrentDb()
        except Exception as e:
            self._force_close()
            raise DatabaseOpenError(
                f"Opened '{self.file_path.name}' but could not access CurrentDb(): {e}\n"
                "This may indicate database corruption or workgroup security requirements."
            ) from e

        # VBE is optional — if it fails, VBA extraction will be skipped gracefully
        try:
            self._vbe = self._app.VBE
        except Exception:
            self._vbe = None

        try:
            self._current_project = self._app.CurrentProject
        except Exception:
            self._current_project = None

    def _close(self) -> None:
        try:
            if self._app is not None:
                self._app.Quit()
        except Exception:
            # Even Application.Quit() can fail on corrupt files or already-crashed Access.
            # Fall through to force-kill if we have the PID.
            self._force_close()
        finally:
            self._app = None
            self._current_db = None
            self._vbe = None
            self._current_project = None

    def _force_close(self) -> None:
        """Kill the MSACCESS.EXE process by PID as a last resort."""
        if self._msaccess_pid is not None:
            try:
                subprocess.run(
                    ["taskkill", "/F", "/PID", str(self._msaccess_pid)],
                    capture_output=True,
                    timeout=5,
                )
            except Exception:
                pass

    def _assert_open(self) -> None:
        if self._app is None:
            raise COMSessionError(
                "COMSession is not open. Use 'with COMSession(...) as session:' pattern."
            )

    # ------------------------------------------------------------------
    # Session recovery utilities (used by ExtractionEngine)
    # ------------------------------------------------------------------

    def reset_display_state(self) -> None:
        """
        Call Application.Echo True to reset any display state corruption
        that can occur after DoCmd operations fail.
        """
        try:
            self._app.Echo(True)
        except Exception:
            pass

    def close_all_open_objects(self) -> None:
        """
        Force-close any forms or reports left open in design view.
        Should be called after each form/report extraction cycle.
        """
        # acForm=2, acReport=3
        for obj_type in (2, 3):
            try:
                while self._app.Forms.Count > 0:
                    form_name = self._app.Forms(0).Name
                    self._app.DoCmd.Close(obj_type, form_name)
            except Exception:
                break
