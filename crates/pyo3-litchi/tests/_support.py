"""Shared helpers for the litchi_py read-API oracle.

The oracle pins the read-only Python API of ``crates/pyo3-litchi`` as it
behaves before M0a step S8a moves the bindings onto the unified
``litchi::File``. Run it with ``crates/pyo3-litchi/tests/run_tests.sh``; see
``docs/unified-ops/s8a-python-oracle.md`` for the test tiers (contract,
snapshot, known drift) and the coverage table.

This module is a helper, not a test module: its name does not match the
``test*.py`` discovery pattern of unittest or pytest.
"""

from __future__ import annotations

import ast
import inspect
import os
import re
import shutil
import subprocess
import sys
from dataclasses import dataclass, field
from pathlib import Path

TESTS_DIR = Path(__file__).resolve().parent
CRATE_DIR = TESTS_DIR.parent
REPO_ROOT = CRATE_DIR.parent.parent
TEST_DATA = Path(os.environ.get("LITCHI_TEST_DATA", str(REPO_ROOT / "test-data")))
STUB_PATH = CRATE_DIR / "python" / "litchi_py" / "__init__.pyi"
README_PATH = CRATE_DIR / "README.md"

try:
    import litchi_py
except ImportError as exc:  # pragma: no cover - environment guard
    raise ImportError(
        "litchi_py is not importable. Build and stage it with "
        "crates/pyo3-litchi/tests/run_tests.sh, or put a directory that holds "
        "litchi_py.abi3.so (a copy of cargo's liblitchi_py.so) on PYTHONPATH."
    ) from exc

L = litchi_py

# ---------------------------------------------------------------------------
# Format vocabularies
# ---------------------------------------------------------------------------

#: The 21 variants of the core ``litchi::FileFormat`` enum
#: (crates/litchi-core/src/detection/types.rs). Record 0003 §5.1 expects the
#: Python enum and the stub to converge on this set in S8a.
CORE_FORMAT_NAMES = frozenset(
    {
        "Doc", "Docx", "Ppt", "Pptx", "Xls", "Xlsx", "Xlsb", "Rtf",
        "Pages", "Keynote", "Numbers",
        "Odt", "Ods", "Odp", "Odg", "Odc", "Odf", "Odi", "Odm", "Oth", "Odb",
    }
)

#: Variants the runtime ``litchi_py.FileFormat`` exposes today: every core
#: variant except ``Odb``, which the bindings map to ``Doc`` (drift D1).
RUNTIME_FORMAT_NAMES_TODAY = CORE_FORMAT_NAMES - {"Odb"}

#: Variants ``__init__.pyi`` declares today: 14 of 21 (drift D2).
STUB_FORMAT_NAMES_TODAY = frozenset(
    {
        "Doc", "Docx", "Ppt", "Pptx", "Xls", "Xlsx", "Xlsb",
        "Odt", "Ods", "Odp", "Pages", "Keynote", "Numbers", "Rtf",
    }
)

#: Public runtime members the stub does not declare, by class (drift D2).
#: ``test_stub_sync`` tolerates exactly these; anything new fails.
KNOWN_RUNTIME_ONLY_MEMBERS = {
    "FileFormat": RUNTIME_FORMAT_NAMES_TODAY - STUB_FORMAT_NAMES_TODAY,
}

#: Base classes the stub declares that the runtime class does not have
#: (drift D3: ``class FileFormat(Enum)``, but the runtime is a pyo3 class).
KNOWN_STUB_BASE_DRIFT = frozenset({("FileFormat", "Enum")})

#: API the README's "API Reference" lists that the module does not have
#: (drift D5): the Workbook/Worksheet section documents an unimplemented API.
KNOWN_README_ONLY_API = frozenset(
    {
        ("Workbook", "worksheets"),
        ("Workbook", "worksheet_by_name"),
        ("Worksheet", "name"),
        ("Worksheet", "row_count"),
        ("Worksheet", "column_count"),
        ("Worksheet", "cell_value"),
        ("Worksheet", "rows"),
    }
)

# ---------------------------------------------------------------------------
# Fixtures and processes
# ---------------------------------------------------------------------------


def fixture(relative: str) -> str:
    """Return the absolute path of a test-data fixture, failing loudly if absent."""
    path = TEST_DATA / relative
    if not path.is_file():
        raise AssertionError(
            f"fixture missing: {path} (set LITCHI_TEST_DATA, or check out test-data/)"
        )
    return str(path)


def read_fixture(relative: str) -> bytes:
    with open(fixture(relative), "rb") as handle:
        return handle.read()


def copy_fixture(relative: str, directory: str, name: str | None = None) -> str:
    """Copy a fixture into ``directory`` (a temporary directory) and return the copy."""
    source = fixture(relative)
    target = os.path.join(directory, name or os.path.basename(source))
    shutil.copyfile(source, target)
    return target


def module_dir() -> str:
    """Directory that holds the imported extension (the staged copy, not its target)."""
    return os.path.dirname(os.path.abspath(L.__file__))


def run_isolated(code: str, timeout: float = 120.0) -> subprocess.CompletedProcess:
    """Run ``code`` in a fresh interpreter that imports the same litchi_py.

    Used for behaviour that panics inside the extension. A dev build turns the
    panic into ``pyo3_runtime.PanicException`` (a ``BaseException``), while a
    release build (``panic = "abort"`` in the workspace profile) terminates the
    interpreter, so such code must never run inside the test process.
    """
    env = dict(os.environ)
    env["PYTHONPATH"] = os.pathsep.join(
        part for part in (module_dir(), env.get("PYTHONPATH", "")) if part
    )
    env["PYTHONDONTWRITEBYTECODE"] = "1"
    env["RUST_BACKTRACE"] = "0"
    return subprocess.run(
        [sys.executable, "-c", code],
        capture_output=True,
        text=True,
        timeout=timeout,
        env=env,
    )


def modify_in_place(path: str) -> None:
    """Change a file without replacing its inode: append one byte, bump mtime.

    The positional file source fingerprints device, inode, length and mtime on
    Unix (crates/litchi-core/src/source/file.rs). Changing the length makes the
    change visible regardless of the filesystem's timestamp granularity.
    """
    before = os.stat(path)
    with open(path, "ab") as handle:
        handle.write(b"\0")
    os.utime(path, ns=(before.st_atime_ns, before.st_mtime_ns + 2_000_000_000))


def replace_atomically(path: str, replacement: str) -> None:
    """Replace ``path`` by renaming a copy of ``replacement`` over it (new inode)."""
    staging = path + ".staging"
    shutil.copyfile(replacement, staging)
    os.replace(staging, path)


# ---------------------------------------------------------------------------
# Read helpers
# ---------------------------------------------------------------------------


def paragraph_texts(document) -> list:
    return [paragraph.text() for paragraph in document.paragraphs()]


def run_facts(paragraph) -> list:
    """(text, bold, italic) per run; ``underline`` is covered by drift tests."""
    return [(run.text(), run.bold(), run.italic()) for run in paragraph.runs()]


def table_grid(table) -> list:
    return [[cell.text() for cell in row.cells()] for row in table.rows()]


def slide_texts(presentation) -> list:
    return [slide.text() for slide in presentation.slides()]


def text_lines(workbook) -> list:
    return workbook.text().split("\n")


# ---------------------------------------------------------------------------
# Stub (.pyi) model
# ---------------------------------------------------------------------------


@dataclass(frozen=True)
class StubMember:
    name: str
    kind: str  # "method", "staticmethod", "property", "attribute", "constructor"
    params: tuple = ()


@dataclass(frozen=True)
class StubClass:
    name: str
    bases: tuple
    members: dict = field(default_factory=dict)


@dataclass(frozen=True)
class Stub:
    classes: dict
    functions: dict
    all: tuple


def _params(node: ast.FunctionDef, drop_first: bool) -> tuple:
    args = node.args
    names = [a.arg for a in (*args.posonlyargs, *args.args, *args.kwonlyargs)]
    if drop_first and names:
        names = names[1:]
    return tuple(names)


def _dotted(node: ast.expr) -> str:
    """Dotted name of a Name/Attribute node (``ast.unparse`` needs Python 3.9)."""
    if isinstance(node, ast.Name):
        return node.id
    if isinstance(node, ast.Attribute):
        return f"{_dotted(node.value)}.{node.attr}"
    return ast.dump(node)


def _decorators(node: ast.FunctionDef) -> set:
    return {_dotted(decorator) for decorator in node.decorator_list}


def parse_stub(path: Path = STUB_PATH) -> Stub:
    """Parse ``__init__.pyi`` into classes, members, functions and ``__all__``."""
    tree = ast.parse(path.read_text(encoding="utf-8"), filename=str(path))
    classes: dict = {}
    functions: dict = {}
    exported: tuple = ()
    for node in tree.body:
        if isinstance(node, ast.ClassDef):
            members: dict = {}
            for item in node.body:
                if isinstance(item, ast.FunctionDef):
                    decorators = _decorators(item)
                    if item.name == "__init__":
                        members["__init__"] = StubMember(
                            "__init__", "constructor", _params(item, True)
                        )
                    elif "staticmethod" in decorators:
                        members[item.name] = StubMember(
                            item.name, "staticmethod", _params(item, False)
                        )
                    elif "property" in decorators:
                        members[item.name] = StubMember(item.name, "property")
                    else:
                        members[item.name] = StubMember(
                            item.name, "method", _params(item, True)
                        )
                elif isinstance(item, ast.AnnAssign) and isinstance(item.target, ast.Name):
                    members[item.target.id] = StubMember(item.target.id, "attribute")
            classes[node.name] = StubClass(
                node.name, tuple(_dotted(base) for base in node.bases), members
            )
        elif isinstance(node, ast.FunctionDef):
            functions[node.name] = _params(node, False)
        elif isinstance(node, ast.Assign):
            if any(isinstance(t, ast.Name) and t.id == "__all__" for t in node.targets):
                exported = tuple(ast.literal_eval(node.value))
    return Stub(classes, functions, exported)


def public_members(cls) -> set:
    return {name for name in dir(cls) if not name.startswith("_")}


def resolve_stub_base(name: str):
    """Runtime class for a base the stub names (only what the stub imports)."""
    import enum

    return {"Enum": enum.Enum, "enum.Enum": enum.Enum, "object": object}[name]


def readme_api_references(path: Path = README_PATH) -> set:
    """``(owner, member)`` pairs listed in the README's "API Reference" section.

    ``owner`` is ``""`` for module-level functions.
    """
    text = path.read_text(encoding="utf-8")
    section = text.split("## API Reference", 1)[1].split("\n## ", 1)[0]
    return {
        (match.group(1) or "", match.group(2))
        for match in re.finditer(r"\*\*`(?:(\w+)\.)?(\w+)\(", section)
    }


def runtime_has(owner: str, member: str) -> bool:
    target = L if not owner else getattr(L, owner, None)
    return target is not None and hasattr(target, member)


def runtime_member_kind(cls, name: str) -> str:
    raw = inspect.getattr_static(cls, name)
    if isinstance(raw, staticmethod):
        return "staticmethod"
    if inspect.isgetsetdescriptor(raw) or isinstance(raw, property):
        return "property"
    if isinstance(raw, cls):
        return "attribute"
    if callable(raw):
        return "method"
    return type(raw).__name__


def runtime_params(callable_object, drop_self: bool) -> tuple:
    names = [p.name for p in inspect.signature(callable_object).parameters.values()]
    if drop_self and names and names[0] == "self":
        names = names[1:]
    return tuple(names)
