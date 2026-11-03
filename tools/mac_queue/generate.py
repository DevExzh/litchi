#!/usr/bin/env python3
"""Generate the Litchi-mutated packages of the macOS verification queue.

Record 0430. For every row of `docs/unified-ops/0363-mac-queue.md` this
tool produces the package the native open/save/close/reopen check must
open, groups rows whose packages are byte-identical, and writes a manifest.

    source /home/zhuhe/code/litchi-worktrees/.coord/mac-runbook/env.sh
    python3 tools/mac_queue/generate.py --out <directory>

Steps:

1. `cargo test -p litchi-iwa --all-features --test parity` with
   `LITCHI_MAC_QUEUE_DUMP=<out>/raw/parity`: every harness case writes its
   input, host output and leaf output (`tests/parity/harness.rs`, module
   `mac_queue`). Leaf rows take the leaf output, host rows the host output.
2. `cargo run -p litchi-ops --example mac_queue_file_routes` writes the
   `litchi_ops::File` rows' packages (and the two builder documents) to
   `<out>/raw/example`.
3. Every row is mapped (`rows.map_row`) to its sources; the packages are
   deduplicated by SHA-256 into `<out>/packages/<App>/`, and
   `<out>/manifest.json` and `<out>/queue.tsv` (the Mac runner's input) are
   written. Host-built inputs (inputs that are not a tracked `test-data`
   file) are copied to `<out>/inputs/` and listed in
   `<out>/host-built-inputs.tsv` for the W3 snapshot.

`--reuse-dumps` skips steps 1 and 2. `--check` only proves that every queue
row maps to sources or an explicit grouping (no build, no output).
"""

from __future__ import annotations

import argparse
import hashlib
import json
import os
import shutil
import subprocess
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))

import rows as mac_queue  # noqa: E402  (tools/mac_queue/rows.py)

ROOT = mac_queue.ROOT
DEFAULT_OUT = ROOT.parent / ".coord" / "mac-runbook" / "out"
RUNNER = Path(__file__).resolve().parent / "run_on_mac.sh"
SCHEMA = 1


def sha256_bytes(data: bytes) -> str:
    return hashlib.sha256(data).hexdigest()


def sha256_file(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as handle:
        for chunk in iter(lambda: handle.read(1 << 20), b""):
            digest.update(chunk)
    return digest.hexdigest()


def git_head(root: Path = ROOT) -> str:
    try:
        return subprocess.run(
            ["git", "-C", str(root), "rev-parse", "HEAD"],
            check=True, capture_output=True, text=True,
        ).stdout.strip()
    except (OSError, subprocess.CalledProcessError):
        return "unknown"


def build(out: Path, root: Path = ROOT) -> None:
    """Run the parity harness and the File-route example with dumps on."""
    raw = out / "raw"
    if raw.exists():
        shutil.rmtree(raw)
    env = dict(os.environ, LITCHI_MAC_QUEUE_DUMP=str(raw / "parity"))
    commands = [
        (["cargo", "test", "-p", "litchi-iwa", "--all-features", "--test", "parity"], env),
        (["cargo", "run", "-q", "-p", "litchi-ops", "--features", "pages,keynote,numbers",
          "--example", "mac_queue_file_routes", "--", str(raw / "example")], dict(os.environ)),
    ]
    for command, environment in commands:
        print("+", " ".join(command), file=sys.stderr)
        subprocess.run(command, cwd=root, env=environment, check=True)


class ParityDumps:
    """The harness dumps, indexed by case name, test function and module.

    A run's `test` is libtest's thread name, `module::test_fn`; a case is
    found by its own name or by the test function that ran it.
    """

    def __init__(self, directory: Path) -> None:
        self.by_name: dict[str, list[dict]] = {}
        self.by_module: dict[str, list[dict]] = {}
        # Runs whose case reported failures. In a green parity run these are
        # negative controls (`*_rejects_*` tests expect the case to fail on
        # a corrupted package); their packages never go to the Mac.
        self.failed: list[str] = []
        if not directory.is_dir():
            return
        for sidecar in sorted(directory.glob("*.json")):
            record = json.loads(sidecar.read_text(encoding="utf-8"))
            if not record["passed"]:
                self.failed.append(record["stem"])
                continue
            stem, extension = record["stem"], record["extension"]
            record["paths"] = {
                role: directory / f"{stem}.{role}.{extension}" for role in ("input", "host", "leaf")
            }
            names = [record["case"]]
            module, _, test = record.get("test", "-").rpartition("::")
            if test and test != record["case"]:
                names.append(test)
            for name in names:
                self.by_name.setdefault(name, []).append(record)
            if module:
                self.by_module.setdefault(module, []).append(record)
        for runs in list(self.by_name.values()) + list(self.by_module.values()):
            runs.sort(key=lambda run: (run["stem"].split("~")[0], run["occurrence"]))

    def runs(self, name: str) -> list[dict]:
        return self.by_name.get(name, [])

    def family(self, name: str, source_file: str | None) -> list[dict]:
        """Runs of the module that declares test `name`, of its format.

        Used for tests that compare reads directly instead of through a
        harness case (most `*_read_parity` tests): the packages of the same
        module carry the feature the read covers.
        """
        if not source_file:
            return []
        module = Path(source_file).stem
        runs = self.by_module.get(module, [])
        prefix = name.split("_", 1)[0]
        if prefix in mac_queue.APPS:
            runs = [run for run in runs if run["format"].lower() == prefix]
        return runs


def load_example_dumps(directory: Path) -> dict[str, dict]:
    """Example stem -> its JSON line plus the package path."""
    stems: dict[str, dict] = {}
    listing = directory / "file-routes.jsonl"
    if not listing.is_file():
        return stems
    for line in listing.read_text(encoding="utf-8").splitlines():
        if not line.strip():
            continue
        record = json.loads(line)
        record["path"] = directory / f"{record['stem']}.file.{record['extension']}"
        stems[record["stem"]] = record
    return stems


def tracked_fixture_hashes(root: Path = ROOT) -> dict[str, str]:
    """SHA-256 -> tracked path of every iWork package under test-data."""
    hashes: dict[str, str] = {}
    base = root / "test-data" / "iwork"
    for path in sorted(base.rglob("*")):
        if path.is_file() and path.suffix in (".key", ".pages", ".numbers"):
            hashes.setdefault(sha256_file(path), path.relative_to(root).as_posix())
    return hashes


def format_of_extension(extension: str) -> str:
    return {"key": "keynote", "pages": "pages", "numbers": "numbers"}[extension]


def check_mapping(root: Path = ROOT) -> tuple[list[mac_queue.Row], dict[str, mac_queue.Mapping], list[str]]:
    """Map every row; return the rows, the mappings and the errors."""
    rows = mac_queue.read_queue(root)
    tests = mac_queue.parity_tests(root)
    ledger = mac_queue.ledger_cases(root)
    mappings: dict[str, mac_queue.Mapping] = {}
    errors: list[str] = []
    stems = set(mac_queue.example_stems())
    for row in rows:
        mapping = mac_queue.map_row(row, tests, ledger)
        mappings[row.id] = mapping
        if not mapping.sources and mapping.group is None:
            errors.append(f"line {row.line} ({row.id}): {mapping.reason}: {row.route}")
        for source in mapping.sources:
            known = tests if source.kind == "parity" else stems
            if source.name not in known:
                errors.append(f"line {row.line} ({row.id}): {source.kind} source {source.name} does not exist")
    return rows, mappings, errors


def select_runs(runs: list[dict], limit: int) -> list[dict]:
    """At most `limit` runs of one source (all when `limit` <= 0).

    The first run on each distinct fixture comes first (a corpus case
    covers every fixture once), then runs spread evenly over the rest; the
    result keeps the runs' order. Many cases repeat one route over dozens of
    property values or cells; the native check needs representatives, not
    every variant (`--runs-per-source 0` keeps them all).
    """
    if limit <= 0 or len(runs) <= limit:
        return list(runs)
    chosen: list[int] = []
    seen: set[str] = set()
    for index, run in enumerate(runs):
        if run["fixture"] not in seen:
            seen.add(run["fixture"])
            chosen.append(index)
    chosen = chosen[:limit]
    rest = [index for index in range(len(runs)) if index not in chosen]
    need = limit - len(chosen)
    if need > 0 and rest:
        step = len(rest) / need
        chosen += [rest[min(len(rest) - 1, int(step * k + step / 2))] for k in range(need)]
    return [runs[index] for index in sorted(set(chosen))]


def resolve(rows: list[mac_queue.Row], mappings: dict[str, mac_queue.Mapping], parity: "ParityDumps",
            tests: dict[str, str], limit: int) -> dict[str, list[tuple[mac_queue.Source, list[dict], bool]]]:
    """Row id -> (source, selected runs, family fallback?) per parity source.

    Family fallbacks are resolved last and prefer runs the direct sources of
    the same module already selected, so a read row adds no package of its
    own when a write row of its family has one.
    """
    resolved: dict[str, list[tuple[mac_queue.Source, list[dict], bool]]] = {}
    selected_stems: set[str] = set()
    pending: list[tuple[str, mac_queue.Source]] = []
    for row in rows:
        resolved[row.id] = []
        for source in mappings[row.id].sources:
            if source.kind != "parity":
                continue
            runs = parity.runs(source.name)
            if runs:
                chosen = select_runs(runs, limit)
                selected_stems.update(run["stem"] for run in chosen)
                resolved[row.id].append((source, chosen, False))
            else:
                pending.append((row.id, source))
    for row_id, source in pending:
        family = parity.family(source.name, tests.get(source.name))
        preferred = [run for run in family if run["stem"] in selected_stems]
        chosen = select_runs(preferred or family, limit)
        resolved[row_id].append((source, chosen, bool(chosen)))
    return resolved


def generate(out: Path, root: Path = ROOT, *, runs_per_source: int = 2) -> dict:
    """Build the packages directory and manifest from the dumps in `out/raw`."""
    rows, mappings, errors = check_mapping(root)
    if errors:
        raise SystemExit("unmapped queue rows:\n" + "\n".join(errors))
    parity = ParityDumps(out / "raw" / "parity")
    tests = mac_queue.parity_tests(root)
    example = load_example_dumps(out / "raw" / "example")
    tracked = tracked_fixture_hashes(root)

    for directory in ("packages", "inputs"):
        if (out / directory).exists():
            shutil.rmtree(out / directory)

    packages: dict[str, dict] = {}  # sha -> package
    host_inputs: dict[str, dict] = {}  # sha -> input
    row_entries: list[dict] = []

    def add_package(path: Path, extension: str, source: dict, row: mac_queue.Row) -> str:
        data = path.read_bytes()
        digest = sha256_bytes(data)
        package = packages.get(digest)
        if package is None:
            package = {
                "sha256": digest,
                "format": format_of_extension(extension),
                "app": mac_queue.APPS[format_of_extension(extension)],
                "extension": extension,
                "size": len(data),
                "sources": [],
                "rows": [],
                "_data": data,
            }
            packages[digest] = package
        if source not in package["sources"]:
            package["sources"].append(source)
        if row.id not in package["rows"]:
            package["rows"].append(row.id)
        return digest

    resolved = resolve(rows, mappings, parity, tests, runs_per_source)
    for row in rows:
        mapping = mappings[row.id]
        digests: list[str] = []
        missing: list[str] = []
        family: list[str] = []
        parity_runs = {source: (runs, is_family) for source, runs, is_family in resolved[row.id]}
        for source in mapping.sources:
            if source.kind == "parity":
                runs, is_family = parity_runs[source]
                if is_family:
                    family.append(source.name)
                if not runs:
                    missing.append(source.name)
                    continue
                for run in runs:
                    input_bytes = run["paths"]["input"].read_bytes()
                    input_sha = sha256_bytes(input_bytes)
                    input_tracked = tracked.get(input_sha)
                    if input_tracked is None:
                        host_inputs.setdefault(input_sha, {
                            "sha256": input_sha,
                            "extension": run["extension"],
                            "size": len(input_bytes),
                            "described_as": run["fixture"],
                            "cases": [],
                            "_data": input_bytes,
                        })
                        if run["stem"] not in host_inputs[input_sha]["cases"]:
                            host_inputs[input_sha]["cases"].append(run["stem"])
                    digests.append(add_package(run["paths"][source.role], run["extension"], {
                        "kind": "parity",
                        "case": run["stem"],
                        "role": source.role,
                        "tier": run["tier"],
                        "read_only": run["read_only"],
                        "fixture": run["fixture"],
                        "input_sha256": input_sha,
                        "input_tracked": input_tracked,
                    }, row))
            else:
                record = example.get(source.name)
                if record is None:
                    missing.append(source.name)
                    continue
                fixture = record["fixture"]
                fixture_path = root / fixture
                input_sha = sha256_file(fixture_path) if fixture_path.is_file() else None
                digests.append(add_package(record["path"], record["extension"], {
                    "kind": "example",
                    "case": source.name,
                    "role": "file",
                    "operation": record["operation"],
                    "selector": record["selector"],
                    "property": record["property"],
                    "read_back": record["read_back"],
                    "fixture": fixture,
                    "input_sha256": input_sha,
                    "input_tracked": fixture if input_sha else None,
                }, row))
        unique: list[str] = []
        for digest in digests:
            if digest not in unique:
                unique.append(digest)
        status = "packaged" if unique else ("grouped" if mapping.group else "missing")
        row_entries.append({
            "id": row.id,
            "line": row.line,
            "wave": row.wave,
            "format": row.format,
            "route": row.route,
            "fixture_note": row.cells["fixture"],
            "rule": mapping.rule,
            "status": status,
            "group": mapping.group,
            "group_reason": mapping.reason if mapping.group else None,
            "sources": [f"{s.kind}:{s.name}:{s.role}" for s in mapping.sources],
            "missing_sources": missing,
            "family_fallback": family,
            "packages": unique,
            "checks": list(mac_queue.NATIVE_CHECKS) if unique else [],
            "records": [
                {"record": record, "path": mac_queue.record_path(record, root)}
                for record in mac_queue.cited_records(row.text())
            ],
        })

    missing_rows = [entry for entry in row_entries if entry["status"] == "missing"]
    if missing_rows:
        raise SystemExit(
            "rows without any dumped package (run without --reuse-dumps?):\n"
            + "\n".join(f"line {e['line']} {e['id']}: {e['missing_sources']}" for e in missing_rows)
        )

    # Stable package identifiers: in order of first use by the queue.
    first_use = {}
    for index, entry in enumerate(row_entries):
        for digest in entry["packages"]:
            first_use.setdefault(digest, index)
    ordered = sorted(packages.values(), key=lambda p: (first_use[p["sha256"]], p["sha256"]))
    ids = {}
    for number, package in enumerate(ordered, start=1):
        identifier = f"P{number:04d}"
        ids[package["sha256"]] = identifier
        relative = Path("packages") / package["app"] / f"{identifier}-{package['sha256'][:12]}.{package['extension']}"
        (out / relative).parent.mkdir(parents=True, exist_ok=True)
        (out / relative).write_bytes(package.pop("_data"))
        package["id"] = identifier
        package["path"] = relative.as_posix()
    for entry in row_entries:
        entry["packages"] = [ids[digest] for digest in entry["packages"]]

    inputs = []
    for item in sorted(host_inputs.values(), key=lambda i: i["sha256"]):
        relative = Path("inputs") / f"{item['sha256'][:16]}.{item['extension']}"
        (out / relative).parent.mkdir(parents=True, exist_ok=True)
        (out / relative).write_bytes(item.pop("_data"))
        item["path"] = relative.as_posix()
        inputs.append(item)

    queue_text = (root / mac_queue.MAC_QUEUE).read_bytes()
    manifest = {
        "schema": SCHEMA,
        "record": "0430",
        "source_commit": git_head(root),
        "queue": mac_queue.MAC_QUEUE.as_posix(),
        "queue_sha256": sha256_bytes(queue_text),
        "checks": list(mac_queue.NATIVE_CHECKS),
        "expected_apps": {app: "14.4 or later" for app in mac_queue.APPS.values()},
        "summary": {
            "rows": len(row_entries),
            "rows_packaged": sum(e["status"] == "packaged" for e in row_entries),
            "rows_grouped": sum(e["status"] == "grouped" for e in row_entries),
            "packages": len(ordered),
            "packages_by_app": {
                app: sum(p["app"] == app for p in ordered) for app in sorted(mac_queue.APPS.values())
            },
            "host_built_inputs": len(inputs),
            "negative_control_runs_excluded": len(parity.failed),
            "runs_per_source": runs_per_source,
        },
        "packages": [
            {key: package[key] for key in ("id", "app", "format", "path", "sha256", "size", "rows", "sources")}
            for package in ordered
        ],
        "rows": row_entries,
        "host_built_inputs": inputs,
    }
    (out / "manifest.json").write_text(json.dumps(manifest, indent=1, sort_keys=False) + "\n", encoding="utf-8")
    write_queue_tsv(out / "queue.tsv", ordered)
    write_host_inputs_tsv(out / "host-built-inputs.tsv", inputs)
    if RUNNER.is_file():
        shutil.copy2(RUNNER, out / "run_on_mac.sh")
    return manifest


def write_queue_tsv(path: Path, packages: list[dict]) -> None:
    """The Mac runner's input: one line per distinct package."""
    lines = ["package_id\tapp\tpath\tsha256\trows"]
    for package in packages:
        lines.append("\t".join([package["id"], package["app"], package["path"], package["sha256"], str(len(package["rows"]))]))
    path.write_text("\n".join(lines) + "\n", encoding="utf-8")


def write_host_inputs_tsv(path: Path, inputs: list[dict]) -> None:
    lines = ["sha256\tpath\tsize\tdescribed_as\tcases"]
    for item in inputs:
        lines.append("\t".join([item["sha256"], item["path"], str(item["size"]), item["described_as"], ",".join(item["cases"])]))
    path.write_text("\n".join(lines) + "\n", encoding="utf-8")


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    parser.add_argument("--out", type=Path, default=DEFAULT_OUT, help=f"output directory (default {DEFAULT_OUT})")
    parser.add_argument("--reuse-dumps", action="store_true", help="use <out>/raw as it is; do not run cargo")
    parser.add_argument("--runs-per-source", type=int, default=2,
                        help="representative runs kept per parity source (0 keeps every run; default 2)")
    parser.add_argument("--check", action="store_true", help="only check that every queue row is mapped")
    args = parser.parse_args(argv)

    if args.check:
        rows, mappings, errors = check_mapping()
        for error in errors:
            print(error, file=sys.stderr)
        grouped = sum(1 for m in mappings.values() if m.group)
        print(f"mac queue: {len(rows)} rows, {len(rows) - grouped - len(errors)} mapped to packages, "
              f"{grouped} explicitly grouped, {len(errors)} errors")
        return 1 if errors else 0

    out = args.out.resolve()
    try:
        out.relative_to(ROOT)
    except ValueError:
        pass
    else:
        raise SystemExit(f"{out} is inside the repository; generated packages must stay outside git")
    out.mkdir(parents=True, exist_ok=True)
    if not args.reuse_dumps:
        build(out)
    manifest = generate(out, runs_per_source=args.runs_per_source)
    summary = manifest["summary"]
    print(f"mac queue: {summary['rows']} rows -> {summary['packages']} distinct packages "
          f"{summary['packages_by_app']}; {summary['rows_grouped']} rows grouped without a package; "
          f"{summary['host_built_inputs']} host-built inputs; manifest {out / 'manifest.json'}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
