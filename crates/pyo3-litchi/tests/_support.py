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
#: The Rust golden files of the unified read API (litchi-ops): node views of
#: real fixtures and capability reports. The Python dicts must equal them.
GOLDEN_DIR = REPO_ROOT / "crates" / "litchi-ops" / "tests" / "golden"

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

#: Variants the runtime ``litchi_py.FileFormat`` exposes: every core variant
#: (S8a added ``Odb``, which the bindings used to map to ``Doc``; drift D1).
RUNTIME_FORMAT_NAMES = CORE_FORMAT_NAMES

#: Public runtime members the stub does not declare, by class. Empty since
#: S8a declared the six missing ODF formats (drift D2); ``test_stub_sync``
#: tolerates exactly these, so anything new fails.
KNOWN_RUNTIME_ONLY_MEMBERS: dict = {}

#: Base classes the stub declares that the runtime class does not have.
#: Empty since S8a dropped the false ``Enum`` base of ``FileFormat`` (drift D3).
KNOWN_STUB_BASE_DRIFT: frozenset = frozenset()

#: API the README's "API Reference" lists that the module does not have.
#: Empty since S8a removed the unimplemented Worksheet API from the README
#: (drift D5).
KNOWN_README_ONLY_API: frozenset = frozenset()

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


def golden(relative: str):
    """Load a Rust golden JSON file under ``GOLDEN_DIR`` (member order kept)."""
    import json

    path = GOLDEN_DIR / relative
    if not path.is_file():
        raise AssertionError(f"golden file missing: {path}")
    with open(path, encoding="utf-8") as handle:
        return json.load(handle)


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
    #: Classes whose name starts with ``_``: typing helpers that exist only in
    #: the stub (the TypedDicts that describe returned dicts), by name.
    typing_only: dict = field(default_factory=dict)


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
    """Parse ``__init__.pyi`` into classes, members, functions and ``__all__``.

    A class whose name starts with ``_`` is a typing helper that exists only
    in the stub; it goes to ``typing_only`` instead of ``classes``. The
    overloads of a method must share its parameter names (the runtime has one
    signature), which this checks.
    """
    tree = ast.parse(path.read_text(encoding="utf-8"), filename=str(path))
    classes: dict = {}
    typing_only: dict = {}
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
                        member = StubMember(item.name, "method", _params(item, True))
                        earlier = members.get(item.name)
                        if earlier is not None and earlier.params != member.params:
                            raise AssertionError(
                                f"the overloads of {node.name}.{item.name} differ in "
                                f"parameter names: {earlier.params} and {member.params}"
                            )
                        members[item.name] = member
                elif isinstance(item, ast.AnnAssign) and isinstance(item.target, ast.Name):
                    members[item.target.id] = StubMember(item.target.id, "attribute")
            target = typing_only if node.name.startswith("_") else classes
            target[node.name] = StubClass(
                node.name, tuple(_dotted(base) for base in node.bases), members
            )
        elif isinstance(node, ast.FunctionDef):
            functions[node.name] = _params(node, False)
        elif isinstance(node, ast.Assign):
            if any(isinstance(t, ast.Name) and t.id == "__all__" for t in node.targets):
                exported = tuple(ast.literal_eval(node.value))
    return Stub(classes, functions, exported, typing_only)


def typed_dict_fields(path: Path = STUB_PATH) -> dict:
    """``{name: {field: required}}`` for each TypedDict the stub declares.

    A field is required unless the class says ``total=False`` or the field's
    annotation is ``NotRequired[...]``.
    """
    tree = ast.parse(path.read_text(encoding="utf-8"), filename=str(path))
    result: dict = {}
    for node in tree.body:
        if not isinstance(node, ast.ClassDef):
            continue
        if not any(_dotted(base).endswith("TypedDict") for base in node.bases):
            continue
        total = all(
            not (keyword.arg == "total" and isinstance(keyword.value, ast.Constant)
                 and keyword.value.value is False)
            for keyword in node.keywords
        )
        fields: dict = {}
        for item in node.body:
            if isinstance(item, ast.AnnAssign) and isinstance(item.target, ast.Name):
                annotation = item.annotation
                optional = (
                    isinstance(annotation, ast.Subscript)
                    and _dotted(annotation.value).endswith("NotRequired")
                )
                fields[item.target.id] = total and not optional
        result[node.name] = fields
    return result


def public_members(cls) -> set:
    """Public names ``cls`` defines itself.

    Members a class inherits are checked on the class that defines them: the
    stub declares ``code``, ``message`` and ``suggestion`` once, on
    ``LitchiError``, and never redeclares what the exception classes inherit
    from the built-in exceptions (``args``, ``errno``, ``with_traceback``, ...).
    For the pyo3 classes, which inherit only from ``object``, this is the same
    set as ``dir()``.
    """
    return {name for name in vars(cls) if not name.startswith("_")}


def resolve_stub_base(name: str):
    """Runtime class for a base the stub names: a litchi_py class or a built-in."""
    import builtins
    import enum

    if name in ("Enum", "enum.Enum"):
        return enum.Enum
    if hasattr(L, name):
        return getattr(L, name)
    return getattr(builtins, name)


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
