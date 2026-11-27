#!/usr/bin/env python3
"""Isolated LibreOffice headless resave stage for Litchi-changed Office files.

This tool never creates the Litchi mutation and never labels an embedded or
synthetic package as native evidence.  A format-owner test supplies a changed
package, this tool resaves it into a separate directory, and that owner then
reopens the result through the corresponding Litchi API.
"""

from __future__ import annotations

import argparse
from dataclasses import asdict, dataclass
import json
import os
from pathlib import Path
import shutil
import subprocess
import sys
import tempfile


@dataclass(frozen=True)
class Filter:
    extension: str
    filter_name: str
    document_service: str
    registry_file: str | None = None


FILTERS = {
    ".docx": Filter(
        "docx",
        "MS Word 2007 XML",
        "com.sun.star.text.TextDocument",
        "MS_Word_2007_XML",
    ),
    ".xlsx": Filter(
        "xlsx",
        "Calc MS Excel 2007 XML",
        "com.sun.star.sheet.SpreadsheetDocument",
        "calc_MS_Excel_2007_XML",
    ),
    ".pptx": Filter(
        "pptx",
        "Impress MS PowerPoint 2007 XML",
        "com.sun.star.presentation.PresentationDocument",
        "impress_MS_PowerPoint_2007_XML",
    ),
    ".doc": Filter(
        "doc", "MS Word 97", "com.sun.star.text.TextDocument", "MS_Word_97"
    ),
    ".xls": Filter(
        "xls", "MS Excel 97", "com.sun.star.sheet.SpreadsheetDocument", "MS_Excel_97"
    ),
    ".ppt": Filter(
        "ppt",
        "MS PowerPoint 97",
        "com.sun.star.presentation.PresentationDocument",
        "MS_PowerPoint_97",
    ),
    ".rtf": Filter(
        "rtf", "Rich Text Format", "com.sun.star.text.TextDocument", "Rich_Text_Format"
    ),
    ".odt": Filter("odt", "writer8", "com.sun.star.text.TextDocument"),
    ".ods": Filter("ods", "calc8", "com.sun.star.sheet.SpreadsheetDocument"),
    ".odp": Filter(
        "odp", "impress8", "com.sun.star.presentation.PresentationDocument"
    ),
    ".odf": Filter("odf", "math8", "com.sun.star.formula.FormulaProperties"),
    ".odc": Filter("odc", "chart8", "com.sun.star.chart2.ChartDocument"),
    ".odg": Filter("odg", "draw8", "com.sun.star.drawing.DrawingDocument"),
    ".odm": Filter("odm", "writerglobal8", "com.sun.star.text.GlobalDocument"),
    ".oth": Filter(
        "oth", "writerweb8_writer_template", "com.sun.star.text.WebDocument"
    ),
}

UNSUPPORTED = {
    ".xlsb": (
        "LibreOffice's Calc MS Excel 2007 Binary registry filter is IMPORT-only; "
        "same-format XLSB export is unavailable"
    ),
    ".odi": "LibreOffice has no ODI type or import/export filter in its filter registry",
    ".odb": (
        "the StarOffice XML (Base) registry filter is IMPORT-only; a same-package "
        "save needs a live UNO document store, which this CLI harness does not claim"
    ),
}


def find_soffice() -> str | None:
    configured = os.environ.get("LIBREOFFICE_BIN")
    if configured:
        candidate = Path(configured)
        if candidate.is_file() and os.access(candidate, os.X_OK):
            return str(candidate)
        return None
    return shutil.which("libreoffice") or shutil.which("soffice")


def command_for(
    executable: str, source: Path, output_directory: Path, profile: Path
) -> list[str]:
    filter_spec = FILTERS[source.suffix.lower()]
    return [
        executable,
        "--headless",
        "--nologo",
        "--nodefault",
        "--nolockcheck",
        "--nofirststartwizard",
        f"-env:UserInstallation={profile.as_uri()}",
        "--convert-to",
        f"{filter_spec.extension}:{filter_spec.filter_name}",
        "--outdir",
        str(output_directory),
        str(source),
    ]


def resave(source: Path, output_directory: Path, executable: str) -> Path:
    source = source.resolve(strict=True)
    suffix = source.suffix.lower()
    if suffix in UNSUPPORTED:
        raise ValueError(f"{suffix}: {UNSUPPORTED[suffix]}")
    if suffix not in FILTERS:
        raise ValueError(f"unsupported Office extension: {suffix or '<none>'}")
    output_directory = output_directory.resolve(strict=True)
    output = output_directory / source.name
    if output == source:
        raise ValueError("output directory must differ from the input directory")

    with tempfile.TemporaryDirectory(prefix="litchi-lo-profile-") as profile_name:
        profile = Path(profile_name)
        completed = subprocess.run(
            command_for(executable, source, output_directory, profile),
            check=False,
            stdout=subprocess.PIPE,
            stderr=subprocess.PIPE,
            text=True,
        )
    if completed.returncode != 0:
        raise RuntimeError(
            f"LibreOffice exited {completed.returncode}: {completed.stderr.strip()}"
        )
    if not output.is_file() or output.stat().st_size == 0:
        raise RuntimeError(
            "LibreOffice reported success without producing the expected resaved file"
        )
    return output


def probe() -> dict[str, object]:
    return {
        "runtime": find_soffice(),
        "python_uno": _python_uno_available(),
        "filters": {key: asdict(value) for key, value in FILTERS.items()},
        "unsupported": UNSUPPORTED,
    }


def _python_uno_available() -> bool:
    try:
        __import__("uno")
    except ImportError:
        return False
    return True


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--probe", action="store_true", help="print runtime/filter JSON")
    parser.add_argument("source", nargs="?", type=Path)
    parser.add_argument("output_directory", nargs="?", type=Path)
    arguments = parser.parse_args(argv)
    if arguments.probe:
        print(json.dumps(probe(), indent=2, sort_keys=True))
        return 0
    if arguments.source is None or arguments.output_directory is None:
        parser.error("source and output_directory are required unless --probe is used")
    executable = find_soffice()
    if executable is None:
        parser.error(
            "no executable LibreOffice runtime found; set LIBREOFFICE_BIN explicitly"
        )
    try:
        output = resave(arguments.source, arguments.output_directory, executable)
    except (OSError, RuntimeError, ValueError) as error:
        print(f"native ODF resave failed: {error}", file=sys.stderr)
        return 1
    print(output)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
