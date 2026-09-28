"""0815 source and artifact custody helpers.

This module only records identities and provides paths for the root-owned
drivers.  It does not execute a workload, invoke Cargo, or interpret results.
"""

from pathlib import Path
import hashlib
import json
import subprocess


P = Path(__file__).resolve().parent
ROOT = P.parents[3]
TARGET = Path("/home/zhuhe/code/litchi-target-0815")
ROOT_INPUTS = {
    "Cargo.lock": P / "inputs/root-Cargo.lock",
    "rustfmt.toml": P / "inputs/rustfmt.toml",
}
UNRELATED = {
    "docs/FORMAT_IMPLEMENTATION_REVIEW.md": "bffd00f144c4c1bbb3b0805d21352b40e9ae46f7366c03d6e581ca61b1f27ce5",
    "docs/UNIFIED_OPS_API_DESIGN.md": "f5672c38393a2a6c52f028b2a501ddad93766ef974dbdb46cdc3c45e1db3ef6d",
    "matrix-analysis.json": "e3b0c44298fc1c149afbf4c8996fb92427ae41e4649b934ca495991b7852b855",
}


def sha(path: Path | str) -> str:
    digest = hashlib.sha256()
    with Path(path).open("rb") as stream:
        for chunk in iter(lambda: stream.read(1 << 20), b""):
            digest.update(chunk)
    return digest.hexdigest()


def read(path: Path | str):
    return json.loads(Path(path).read_text())


def write(path: Path | str, value) -> None:
    Path(path).write_text(json.dumps(value, indent=2, sort_keys=True) + "\n")


def artifact(path: Path | str) -> dict[str, object]:
    path = Path(path)
    return {"path": str(path), "bytes": path.stat().st_size, "sha256": sha(path)}


def root_input_hashes() -> dict[str, str]:
    return {name: sha(path) for name, path in ROOT_INPUTS.items()}


def assert_root_inputs() -> dict[str, str]:
    frozen = root_input_hashes()
    for name, copy in ROOT_INPUTS.items():
        source_path = ROOT / name
        assert source_path.is_file(), f"missing root input {source_path}"
        assert sha(source_path) == frozen[name], f"root input changed: {name}"
    return frozen


def tracked_source_names() -> list[str]:
    raw = subprocess.check_output(
        [
            "git",
            "ls-files",
            "-z",
            "--",
            "crates",
            "Cargo.toml",
            "clippy.toml",
            ".cargo/config.toml",
            "rust-toolchain.toml",
        ],
        cwd=ROOT,
    )
    return [name for name in raw.decode().split("\0") if name]


def source() -> dict[str, object]:
    names = tracked_source_names()
    return {
        "revision": subprocess.check_output(
            ["git", "rev-parse", "HEAD"], cwd=ROOT, text=True
        ).strip(),
        "files": {name: sha(ROOT / name) for name in names},
    }


def changed_files(before: dict[str, object], after: dict[str, object]) -> set[str]:
    before_files = before["files"]
    after_files = after["files"]
    assert isinstance(before_files, dict) and isinstance(after_files, dict)
    return {
        name
        for name in before_files.keys() | after_files.keys()
        if before_files.get(name) != after_files.get(name)
    }


def probe_files() -> dict[str, str]:
    return {
        str(path.relative_to(P / "probe-src")): sha(path)
        for path in sorted((P / "probe-src").rglob("*"))
        if path.is_file()
    }


def assert_probe(expected: dict[str, str] | None = None) -> dict[str, str]:
    actual = probe_files()
    if expected is not None:
        assert actual == expected, "probe source changed after freeze"
    assert len(actual) == 6, f"expected six probe files, found {len(actual)}"
    inherited = read(P / "inheritance.json")["probe_reference"]["files"]
    assert actual == inherited, "probe source differs from declared inheritance"
    sealed = P.parent / "change-0813/probe-src"
    sealed_map = {
        str(path.relative_to(sealed)): sha(path)
        for path in sorted(sealed.rglob("*"))
        if path.is_file()
    }
    assert actual == sealed_map, "probe source differs from sealed 0813 copy"
    return actual


def architecture_hashes() -> dict[str, str]:
    declared = read(P / "architecture-inputs.json")
    assert len(declared) == 35, f"expected 35 architecture inputs, found {len(declared)}"
    actual = {name: sha(ROOT / name) for name in declared}
    assert actual == declared, "architecture input changed after packet freeze"
    return actual


def assert_unrelated() -> dict[str, str]:
    actual = {name: sha(ROOT / name) for name in UNRELATED}
    assert actual == UNRELATED, "an unrelated workspace file changed"
    return actual


def assert_codegen_gate(builds: dict[str, dict[str, object]]) -> dict[str, object]:
    gate_path = P / "codegen-gate.json"
    gate = read(gate_path)
    assert gate["schema"] == "litchi.performance.0815.codegen-gate.v1"
    assert gate["advance"] is True
    assert gate["reviewed_after_code_generation"] is True
    for leg in ("before", "after"):
        descriptor = gate[leg]
        path = Path(descriptor["path"])
        assert artifact(path) == descriptor, f"codegen gate {leg} receipt changed"
        receipt = read(path)
        assert receipt["schema"] == "litchi.performance.0815.codegen.v1"
        assert receipt["leg"] == leg
        expected_source = builds[leg]["source"]
        expected_binaries = builds[leg]["binaries"]
        rows = {row["variant"]: row for row in receipt["rows"]}
        assert set(rows) == {"native", "profile"}
        for variant, row in rows.items():
            assert row["source"] == expected_source
            assert row["binary"] == expected_binaries[variant]
            assert artifact(row["binary"]["path"]) == row["binary"]
    return gate
