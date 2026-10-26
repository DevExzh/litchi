#!/usr/bin/env python3
"""Capture all-features rustdoc JSON for every workspace library, then run
``tools/api_debt_inventory.py`` over the whole set.

Usage (after sourcing a worker ``env.sh`` that puts the memory-capped Cargo
shim first on PATH):

    python3 tools/final_api_inventory.py --output <dir outside the checkout>

The capture is sequential (one rustdoc at a time). Each package's JSON is
copied to ``<output>/inventory/<package>.json`` and listed in
``<output>/manifest.tsv`` with its size and SHA-256. A rerun on the same clean
source commit reuses a JSON whose manifest row still matches; any other commit
or a dirty checkout starts fresh. The run ends with ``<output>/inventory.tsv``
(the debt rows) and ``<output>/disposition.tsv`` (a blank review scaffold that
``tools/validate_api_disposition.py`` rejects until every row is decided).
"""
from __future__ import annotations

import argparse
import csv
import hashlib
import json
import os
from dataclasses import dataclass
from pathlib import Path
import shutil
import subprocess
import sys
from typing import Callable, Mapping, Sequence
import tomllib


MANIFEST_FIELDS = ("package", "lib_name", "source_commit", "bytes", "sha256")


class InventoryError(RuntimeError):
    """The capture cannot produce a complete inventory."""


@dataclass(frozen=True)
class Package:
    name: str
    lib_name: str


# runner(command, cwd, env, log_path) -> exit status
CommandRunner = Callable[[Sequence[str], Path, Mapping[str, str], Path], int]


def sha256_file(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as stream:
        for block in iter(lambda: stream.read(1024 * 1024), b""):
            digest.update(block)
    return digest.hexdigest()


def enumerate_packages(root: Path) -> list[Package]:
    """Return every direct ``crates/*`` library package, sorted by manifest path."""
    packages: list[Package] = []
    seen: set[str] = set()
    for manifest in sorted((root / "crates").glob("*/Cargo.toml")):
        try:
            document = tomllib.loads(manifest.read_text(encoding="utf-8"))
        except (OSError, tomllib.TOMLDecodeError) as error:
            raise InventoryError(f"cannot read {manifest}: {error}") from error
        package = document.get("package")
        if not isinstance(package, dict) or not isinstance(package.get("name"), str):
            continue
        name = package["name"]
        lib = document.get("lib") if isinstance(document.get("lib"), dict) else {}
        lib_name = lib.get("name", name.replace("-", "_"))
        if not (manifest.parent / lib.get("path", "src/lib.rs")).is_file():
            continue
        if name in seen:
            raise InventoryError(f"duplicate package name: {name}")
        seen.add(name)
        packages.append(Package(name, lib_name))
    return packages


def rustdoc_command(package: Package, target: Path, cargo: str = "cargo") -> list[str]:
    return [
        cargo, "rustdoc", "--package", package.name, "--lib", "--all-features",
        "--offline", "--locked", "--target-dir", str(target), "--",
        "-Zunstable-options", "--output-format", "json", "--document-hidden-items",
    ]


def validate_json(path: Path, lib_name: str) -> None:
    try:
        document = json.loads(path.read_text(encoding="utf-8"))
    except (OSError, json.JSONDecodeError) as error:
        raise InventoryError(f"invalid rustdoc JSON for {lib_name}: {path}: {error}") from error
    if not isinstance(document, dict) or not isinstance(document.get("index"), dict):
        raise InventoryError(f"rustdoc JSON lacks an index: {path}")
    root_item = document["index"].get(str(document.get("root")))
    if not isinstance(root_item, dict) or root_item.get("name") != lib_name:
        raise InventoryError(f"rustdoc JSON root is not library {lib_name}: {path}")
    if document.get("includes_private") is not False:
        raise InventoryError(f"rustdoc JSON must exclude private items: {path}")


def git_state(root: Path) -> tuple[str, bool]:
    """Return (HEAD commit, checkout is clean)."""
    env = {**os.environ, "GIT_OPTIONAL_LOCKS": "0"}
    commit = subprocess.check_output(["git", "rev-parse", "HEAD"], cwd=root, env=env,
                                     text=True).strip()
    dirty = subprocess.check_output(["git", "status", "--porcelain"], cwd=root, env=env,
                                    text=True).strip()
    return commit, not dirty


def default_runner(command: Sequence[str], cwd: Path, env: Mapping[str, str],
                   log_path: Path) -> int:
    with log_path.open("wb") as log:
        return subprocess.run(command, cwd=cwd, env=dict(env), stdout=log,
                              stderr=subprocess.STDOUT, check=False).returncode


def read_manifest(path: Path) -> dict[str, dict[str, str]]:
    if not path.is_file():
        return {}
    with path.open(newline="", encoding="utf-8") as stream:
        return {row["package"]: row for row in csv.DictReader(stream, delimiter="\t")}


def write_manifest(path: Path, rows: Sequence[Mapping[str, str]]) -> None:
    with path.open("w", newline="", encoding="utf-8") as stream:
        writer = csv.DictWriter(stream, fieldnames=MANIFEST_FIELDS, delimiter="\t",
                                lineterminator="\n")
        writer.writeheader()
        writer.writerows(rows)


def capture(root: Path, output: Path, target: Path, packages: Sequence[Package], *,
            commit: str, reusable: bool, cargo: str = "cargo",
            runner: CommandRunner = default_runner,
            env: Mapping[str, str] | None = None) -> list[Path]:
    """Capture each package's JSON sequentially; return the JSON paths."""
    inventory_dir = output / "inventory"
    log_dir = output / "logs"
    inventory_dir.mkdir(parents=True, exist_ok=True)
    log_dir.mkdir(parents=True, exist_ok=True)
    manifest_path = output / "manifest.tsv"
    previous = read_manifest(manifest_path) if reusable else {}
    environment = {**(env if env is not None else os.environ), "RUSTC_BOOTSTRAP": "1"}
    rows: list[dict[str, str]] = []
    paths: list[Path] = []
    for package in packages:
        json_path = inventory_dir / f"{package.name}.json"
        old = previous.get(package.name)
        if (old is not None and old.get("source_commit") == commit and json_path.is_file()
                and old.get("sha256") == sha256_file(json_path)):
            rows.append(old)
            paths.append(json_path)
            continue
        log_path = log_dir / f"{package.name}.log"
        status = runner(rustdoc_command(package, target, cargo), root, environment, log_path)
        if status != 0:
            write_manifest(manifest_path, rows)
            raise InventoryError(f"rustdoc failed for {package.name} with exit {status}; "
                                 f"see {log_path}")
        generated = target / "doc" / f"{package.lib_name}.json"
        validate_json(generated, package.lib_name)
        shutil.copyfile(generated, json_path)
        rows.append({"package": package.name, "lib_name": package.lib_name,
                     "source_commit": commit, "bytes": str(json_path.stat().st_size),
                     "sha256": sha256_file(json_path)})
        paths.append(json_path)
        write_manifest(manifest_path, rows)
    write_manifest(manifest_path, rows)
    return paths


def run_inventory(root: Path, output: Path, json_paths: Sequence[Path], *,
                  python: str = sys.executable) -> Path:
    inventory_path = output / "inventory.tsv"
    with inventory_path.open("w", encoding="utf-8") as stream:
        result = subprocess.run([python, str(root / "tools" / "api_debt_inventory.py"),
                                 *map(str, json_paths)], cwd=root, stdout=stream,
                                check=False)
    if result.returncode != 0:
        raise InventoryError(f"api_debt_inventory.py failed with exit {result.returncode}")
    write_disposition_scaffold(inventory_path, output / "disposition.tsv")
    return inventory_path


def write_disposition_scaffold(inventory: Path, destination: Path) -> Path:
    """Add an explicit, unresolved decision column to every inventory row."""
    with inventory.open(newline="", encoding="utf-8") as source:
        rows = list(csv.DictReader(source, delimiter="\t"))
    fields = ["crate", "category", "path", "detail", "target", "decision", "evidence", "owner"]
    with destination.open("w", newline="", encoding="utf-8") as stream:
        writer = csv.DictWriter(stream, fieldnames=fields, delimiter="\t", lineterminator="\n")
        writer.writeheader()
        for row in rows:
            if row.get("category") == "mut_draft" and "builder" in row.get("path", "").split("::"):
                decision = "DRAFT_CANDIDATE_REVIEW"
            elif row.get("crate") in {"litchi-cfb", "soapberry-zip"}:
                decision = "SUBSTRATE_SCOPE_REVIEW"
            else:
                decision = "UNASSIGNED_REQUIRED_REVIEW"
            writer.writerow({**row, "decision": decision, "evidence": "", "owner": ""})
    return destination


def main(argv: Sequence[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description=__doc__,
                                     formatter_class=argparse.RawDescriptionHelpFormatter)
    parser.add_argument("--root", type=Path, default=Path(__file__).resolve().parent.parent)
    parser.add_argument("--output", type=Path, required=True,
                        help="output directory outside the checkout")
    parser.add_argument("--target-dir", type=Path,
                        default=Path(os.environ["CARGO_TARGET_DIR"])
                        if os.environ.get("CARGO_TARGET_DIR") else None)
    parser.add_argument("--cargo", default="cargo")
    parser.add_argument("--expect-packages", type=int,
                        help="fail unless exactly this many library packages exist")
    args = parser.parse_args(argv)
    root, output = args.root.resolve(), args.output.resolve()
    try:
        if output.is_relative_to(root):
            raise InventoryError(f"output must be outside the checkout: {output}")
        target = (args.target_dir or root / "target").resolve()
        packages = enumerate_packages(root)
        if args.expect_packages is not None and len(packages) != args.expect_packages:
            raise InventoryError(f"expected {args.expect_packages} library packages, "
                                 f"found {len(packages)}")
        commit, clean = git_state(root)
        paths = capture(root, output, target, packages, commit=commit, reusable=clean,
                        cargo=args.cargo)
        inventory = run_inventory(root, output, paths)
        end_commit, end_clean = git_state(root)
        if (end_commit, end_clean) != (commit, clean):
            raise InventoryError("the checkout changed during the capture; rerun")
    except (InventoryError, OSError, subprocess.CalledProcessError) as error:
        print(f"final API inventory: {error}", file=sys.stderr)
        return 2
    print(f"{len(packages)} packages at {commit}{'' if clean else ' (dirty)'}: {inventory}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
