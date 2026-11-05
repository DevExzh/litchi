#!/usr/bin/env python3
"""Control kit for the macOS failures of queue run 1 (record 0431).

Three modes, all without third-party modules:

    control_kit.py audit FILE...          sparse arrays and unlisted references
    control_kit.py build --queue-out OUT --fixed DUMP --native TEST_DATA --out KIT
    control_kit.py add KIT LIST           add another record's packages to a kit
    control_kit.py analyze KIT RESULTS    expected against actual Mac outcomes
    control_kit.py next-run --queue-out OUT --unproven LIST --out KIT
                                          a kit of the queue packages no run has seen
    control_kit.py variant-no-uuid-map IN OUT
                                          IN with the engine's uuid_reference_map emptied

`audit` reports the two structural rules the failures isolate: an entry of a
chart `TSP.SparseReferenceArray` beyond its `count`, and a reference to a
non-root object of another component that the owner's component does not list
in `external_references`. `build` writes a directory that
`tools/mac_queue/run_on_mac.sh` runs (same `queue.tsv` layout as the 0430
output). It reads OUT, which it never writes. `analyze` never touches the
queue: a queue row passes only on a queue package's Mac result.

`add` takes a TSV with the header `id app file expect source`. `file` is relative
to the list; `expect` is opens, refused, crash or unknown. It refuses an id or a
file hash already in the kit, and an existing kit keeps its order: added packages
run last, so a kit's own controls still run first. Re-tar the kit afterwards.
"""

from __future__ import annotations

import argparse
import csv
import hashlib
import io
import json
import shutil
import struct
import sys
import zipfile
from pathlib import Path

CHART_DRAWABLE = 5021
CALCULATION_ENGINE = 4000
UUID_REFERENCE_MAP = 7
CHART_EXTENSION = 10000
SERIES_NON_STYLES = 19
PACKAGE_METADATA = 11006
ENGINE = "Index/CalculationEngine.iwa"
METADATA = "Index/Metadata.iwa"


# --- wire ---------------------------------------------------------------


def read_varint(data: bytes, at: int) -> tuple[int, int]:
    value = shift = 0
    while True:
        byte = data[at]
        at += 1
        value |= (byte & 0x7F) << shift
        shift += 7
        if not byte & 0x80:
            return value, at


def varint(value: int) -> bytes:
    output = bytearray()
    while True:
        byte = value & 0x7F
        value >>= 7
        if value == 0:
            output.append(byte)
            return bytes(output)
        output.append(byte | 0x80)


def fields(data: bytes) -> list[tuple[int, int, object]]:
    """(number, wire type, value) of a protobuf message."""
    output: list[tuple[int, int, object]] = []
    at = 0
    while at < len(data):
        key, at = read_varint(data, at)
        number, wire = key >> 3, key & 7
        if wire == 0:
            value, at = read_varint(data, at)
        elif wire == 1:
            value, at = data[at : at + 8], at + 8
        elif wire == 2:
            length, at = read_varint(data, at)
            value, at = data[at : at + length], at + length
        elif wire == 5:
            value, at = data[at : at + 4], at + 4
        else:
            raise ValueError(f"wire type {wire}")
        output.append((number, wire, value))
    return output


def encode(items: list[tuple[int, int, object]]) -> bytes:
    output = bytearray()
    for number, wire, value in items:
        output += varint(number << 3 | wire)
        if wire == 0:
            output += varint(value)  # type: ignore[arg-type]
        elif wire == 2:
            output += varint(len(value))  # type: ignore[arg-type]
            output += value  # type: ignore[operator]
        else:
            output += value  # type: ignore[operator]
    return bytes(output)


def packed(items: list[tuple[int, int, object]], number: int) -> list[int]:
    values: list[int] = []
    for field, wire, value in items:
        if field != number:
            continue
        if wire == 0:
            values.append(value)  # type: ignore[arg-type]
        else:
            at = 0
            while at < len(value):  # type: ignore[arg-type]
                item, at = read_varint(value, at)  # type: ignore[arg-type]
                values.append(item)
    return values


# --- IWA ----------------------------------------------------------------


def unsnappy(data: bytes) -> bytes:
    size, at = read_varint(data, 0)
    output = bytearray()
    while at < len(data):
        tag = data[at]
        at += 1
        kind = tag & 3
        if kind == 0:
            length = tag >> 2
            if length >= 60:
                extra = length - 59
                length = int.from_bytes(data[at : at + extra], "little")
                at += extra
            length += 1
            output += data[at : at + length]
            at += length
            continue
        if kind == 1:
            length = ((tag >> 2) & 7) + 4
            offset = ((tag >> 5) << 8) | data[at]
            at += 1
        elif kind == 2:
            length = (tag >> 2) + 1
            offset = int.from_bytes(data[at : at + 2], "little")
            at += 2
        else:
            length = (tag >> 2) + 1
            offset = int.from_bytes(data[at : at + 4], "little")
            at += 4
        for _ in range(length):
            output.append(output[-offset])
    if len(output) != size:
        raise ValueError("snappy length mismatch")
    return bytes(output)


def iwa_decode(data: bytes) -> bytes:
    output = bytearray()
    at = 0
    while at < len(data):
        if data[at] != 0:
            raise ValueError("not an IWA chunk")
        length = int.from_bytes(data[at + 1 : at + 4], "little")
        output += unsnappy(data[at + 4 : at + 4 + length])
        at += 4 + length
    return bytes(output)


def iwa_encode(data: bytes) -> bytes:
    """Valid IWA: 64 KiB chunks, each a single Snappy literal."""
    output = bytearray()
    for at in range(0, len(data), 65536):
        chunk = data[at : at + 65536]
        body = bytearray(varint(len(chunk)))
        size = len(chunk) - 1
        if size < 60:
            body.append(size << 2)
        elif size < 256:
            body += bytes([60 << 2, size])
        else:
            body += bytes([61 << 2]) + struct.pack("<H", size)
        body += chunk
        output += b"\0" + len(body).to_bytes(3, "little") + body
    return bytes(output)


class Archive:
    """The objects of one `.iwa`: header fields and message payloads, in order."""

    def __init__(self, data: bytes) -> None:
        self.objects: list[dict] = []
        at = 0
        while at < len(data):
            length, at = read_varint(data, at)
            header = fields(data[at : at + length])
            at += length
            identifier = next(v for n, _, v in header if n == 1)
            messages = []
            for number, _, value in header:
                if number != 2:
                    continue
                info = fields(value)  # type: ignore[arg-type]
                size = next(v for n, _, v in info if n == 3)
                messages.append({"info": info, "data": data[at : at + size]})
                at += size  # type: ignore[operator]
            self.objects.append({"id": identifier, "header": header, "messages": messages})

    def to_bytes(self) -> bytes:
        output = bytearray()
        for item in self.objects:
            messages = iter(item["messages"])
            header = []
            payload = bytearray()
            for number, wire, value in item["header"]:
                if number == 2:
                    message = next(messages)
                    info = [
                        (n, w, len(message["data"]) if n == 3 else v)
                        for n, w, v in message["info"]
                    ]
                    value = encode(info)
                    payload += message["data"]
                header.append((number, wire, value))
            encoded = encode(header)
            output += varint(len(encoded)) + encoded + payload
        return bytes(output)


class Package:
    def __init__(self, data: bytes) -> None:
        self.order: list[str] = []
        self.members: dict[str, bytes] = {}
        with zipfile.ZipFile(io.BytesIO(data)) as source:
            for name in source.namelist():
                self.order.append(name)
                self.members[name] = source.read(name)
        self.archives = {
            name: Archive(iwa_decode(body))
            for name, body in self.members.items()
            if name.endswith(".iwa")
        }

    def objects(self):
        for name, archive in self.archives.items():
            for item in archive.objects:
                yield name, item

    def to_bytes(self, changed: set[str]) -> bytes:
        output = io.BytesIO()
        with zipfile.ZipFile(output, "w", zipfile.ZIP_STORED) as target:
            for name in self.order:
                body = self.members[name]
                if name in changed:
                    body = iwa_encode(self.archives[name].to_bytes())
                info = zipfile.ZipInfo(name, date_time=(2026, 11, 4, 0, 0, 0))
                info.compress_type = zipfile.ZIP_STORED
                target.writestr(info, body)
        return output.getvalue()

    def metadata(self) -> dict:
        for _, item in self.objects():
            for message in item["messages"]:
                if next(v for n, _, v in message["info"] if n == 1) == PACKAGE_METADATA:
                    return message
        raise ValueError("no package metadata")


# --- audit --------------------------------------------------------------


def sparse_arrays(package: Package) -> list[tuple[int, int, list[int]]]:
    """(chart object, count, entry indexes) of every series_non_styles array."""
    output = []
    for _, item in package.objects():
        for message in item["messages"]:
            if next(v for n, _, v in message["info"] if n == 1) != CHART_DRAWABLE:
                continue
            for number, _, chart in fields(message["data"]):
                if number != CHART_EXTENSION:
                    continue
                for inner, _, value in fields(chart):  # type: ignore[arg-type]
                    if inner != SERIES_NON_STYLES:
                        continue
                    sparse = fields(value)  # type: ignore[arg-type]
                    count = next((v for n, _, v in sparse if n == 1), 0)
                    entries = [
                        next(v for n, _, v in fields(entry) if n == 1)  # type: ignore[arg-type]
                        for n, _, entry in sparse
                        if n == 2
                    ]
                    output.append((item["id"], count, entries))
    return output


def components(package: Package) -> dict[str, dict]:
    table: dict[str, dict] = {}
    for number, _, value in fields(package.metadata()["data"]):
        if number != 3:
            continue
        info = fields(value)  # type: ignore[arg-type]
        component = {
            "id": next(v for n, _, v in info if n == 1),
            "listed": set(),
        }
        for field, _, external in info:
            if field == 6:
                external = dict((n, v) for n, _, v in fields(external))  # type: ignore[arg-type]
                if 2 in external:
                    component["listed"].add((external[1], external[2]))
        locator = next((v for n, _, v in info if n == 3), None)
        preferred = next(v for n, _, v in info if n == 2)
        for stem in filter(
            None,
            (
                locator.decode() if locator else None,  # type: ignore[union-attr]
                preferred.decode(),  # type: ignore[union-attr]
                f"{preferred.decode()}-{component['id']}",  # type: ignore[union-attr]
            ),
        ):
            table.setdefault(f"Index/{stem}.iwa", component)
    return table


def unlisted_references(package: Package) -> list[tuple[int, int, int, int]]:
    """(owner, owner type, target, target type) the owner's component misses."""
    table = components(package)
    roots = {component["id"] for component in table.values()}
    home = {}
    kinds = {}
    for name, item in package.objects():
        home[item["id"]] = name
        kinds[item["id"]] = next(v for n, _, v in item["messages"][0]["info"] if n == 1)
    output = []
    for name, item in package.objects():
        component = table.get(name)
        if component is None:
            continue
        for message in item["messages"]:
            for target in packed(message["info"], 5):
                if target not in home or target in roots:
                    continue
                other = table.get(home[target])
                if other is None or other is component:
                    continue
                if (other["id"], target) not in component["listed"]:
                    output.append((item["id"], kinds[item["id"]], target, kinds[target]))
    return output


# --- variants -----------------------------------------------------------


def variant_sparse_count(source: bytes, count: int) -> bytes:
    """Only the `count` of each non-empty series_non_styles array changes."""
    package = Package(source)
    changed = set()
    for name, item in package.objects():
        for message in item["messages"]:
            if next(v for n, _, v in message["info"] if n == 1) != CHART_DRAWABLE:
                continue
            outer = fields(message["data"])
            rebuilt_outer = []
            touched = False
            for number, wire, chart in outer:
                if number == CHART_EXTENSION:
                    inner = []
                    for field, inner_wire, value in fields(chart):  # type: ignore[arg-type]
                        if field == SERIES_NON_STYLES:
                            sparse = fields(value)  # type: ignore[arg-type]
                            if len(sparse) > 1:
                                sparse = [(1, 0, count), *[s for s in sparse if s[0] != 1]]
                                value = encode(sparse)
                                touched = True
                        inner.append((field, inner_wire, value))
                    chart = encode(inner)
                rebuilt_outer.append((number, wire, chart))
            if touched:
                message["data"] = encode(rebuilt_outer)
                changed.add(name)
    if not changed:
        raise ValueError("no populated sparse array")
    return package.to_bytes(changed)


def variant_list_references(source: bytes) -> bytes:
    """Only the missing `external_references` entries are appended."""
    package = Package(source)
    missing = unlisted_references(package)
    if not missing:
        raise ValueError("nothing to list")
    table = components(package)
    home = {item["id"]: name for name, item in package.objects()}
    additions: dict[int, list[tuple[int, int]]] = {}
    for owner, _, target, _ in missing:
        pair = (table[home[target]]["id"], target)
        bucket = additions.setdefault(table[home[owner]]["id"], [])
        if pair not in bucket:
            bucket.append(pair)
    message = package.metadata()
    rebuilt = []
    for number, wire, value in fields(message["data"]):
        if number == 3:
            info = fields(value)  # type: ignore[arg-type]
            identifier = next(v for n, _, v in info if n == 1)
            extra = additions.get(identifier, [])  # type: ignore[arg-type]
            if extra:
                last = max(i for i, (n, _, _) in enumerate(info) if n == 6)
                entries = [
                    (6, 2, encode([(1, 0, component), (2, 0, target)]))
                    for component, target in extra
                ]
                info = info[: last + 1] + entries + info[last + 1 :]
                value = encode(info)
        rebuilt.append((number, wire, value))
    message["data"] = encode(rebuilt)
    return package.to_bytes({METADATA})


def variant_drop_objects(source: bytes, drop: set[int]) -> bytes:
    """The synthetic pivot objects in `drop` leave the calculation engine."""
    package = Package(source)
    archive = package.archives[ENGINE]
    before = len(archive.objects)
    archive.objects = [item for item in archive.objects if item["id"] not in drop]
    if len(archive.objects) != before - len(drop):
        raise ValueError("an object to drop is absent")
    return package.to_bytes({ENGINE})


# --- the kit ------------------------------------------------------------

EXTENSION = {"Keynote": "key", "Numbers": "numbers", "Pages": "pages"}

# failing package, app, parity stem, passing sibling
GROUP_A = [
    ("P0422", "Keynote", "keynote_slide_chart_pie_wedge_explosions_parity~3", "P0420"),
    ("P0427", "Keynote", "keynote_chart_pie_wedge_explosions_partial_release_parity", "P0420"),
    ("P0971", "Keynote", "keynote_slide_chart_pie_leader_lines_parity~3", None),
    ("P0425", "Numbers", "numbers_sheet_chart_pie_wedge_explosions_parity~3", "P0421"),
    ("P0623", "Numbers", "numbers_sheet_chart_pie_leader_lines_parity~3", None),
    ("P0428", "Pages", "pages_body_chart_pie_wedge_explosions_parity~3", None),
    ("P0928", "Pages", "pages_body_chart_pie_leader_lines_parity~3", None),
]
GROUP_B = [
    ("P1126", "numbers_sheet_chart_caption_replace_parity", "P1122"),
    ("P1128", "numbers_sheet_chart_caption_remove_parity", "P1124"),
    ("P1141", "numbers_sheet_chart_caption_set_parity", "P1130"),
    ("P1148", "numbers_sheet_chart_caption_set_parity~2", "P1133"),
]
GROUP_C = [
    ("P1142", "numbers_pivot_formula_author_referenced_tree_parity"),
    ("P1143", "numbers_pivot_formula_author_without_cache_parity"),
    ("P1144", "numbers_pivot_formula_author_parity"),
]
# Objects of the synthetic pivot fixture (crates/litchi-numbers/tests/pivot_synthetic_group_bys.rs).
CYCLE = {9_100_020, 9_100_021, 9_100_022, 9_100_023}
OTHERS = {9_100_002, 9_100_003, 9_100_030, 9_100_031, 9_100_040}
INLINE = {9_100_010}


class Kit:
    def __init__(self, queue_out: Path, fixed: Path, native: Path, out: Path) -> None:
        self.queue_out = queue_out
        self.raw = queue_out / "raw" / "parity"
        self.fixed = fixed
        self.native = native
        self.out = out
        manifest = json.loads((queue_out / "manifest.json").read_text())
        self.packages = {package["id"]: package for package in manifest["packages"]}
        self.rows: list[dict] = []
        self.seen: dict[str, str] = {}

    def add(self, group, failing, kind, app, data, expect, source) -> None:
        identifier = f"{group}-{failing}-{kind}" if failing else f"{group}-{kind}"
        digest = hashlib.sha256(data).hexdigest()
        if digest in self.seen:
            self.rows.append(
                {
                    "id": identifier,
                    "group": group,
                    "failing": failing,
                    "kind": kind,
                    "app": app,
                    "same_as": self.seen[digest],
                    "sha256": digest,
                    "expect": expect,
                    "source": source,
                }
            )
            return
        self.seen[digest] = identifier
        path = f"packages/{app}/{identifier}.{EXTENSION[app]}"
        target = self.out / path
        target.parent.mkdir(parents=True, exist_ok=True)
        target.write_bytes(data)
        self.rows.append(
            {
                "id": identifier,
                "group": group,
                "failing": failing,
                "kind": kind,
                "app": app,
                "path": path,
                "sha256": digest,
                "size": len(data),
                "expect": expect,
                "source": source,
            }
        )

    def queue_package(self, identifier: str) -> bytes:
        return (self.queue_out / self.packages[identifier]["path"]).read_bytes()

    def dump(self, root: Path, stem: str, role: str, app: str) -> bytes:
        return (root / f"{stem}.{role}.{EXTENSION[app]}").read_bytes()

    def family(self, group, failing, app, stem, *, input_expect="opens", fixed=True) -> None:
        self.add(
            group, failing, "1-input", app, self.dump(self.raw, stem, "input", app),
            input_expect, f"raw/parity/{stem}.input",
        )
        self.add(
            group, failing, "2-host-current", app, self.dump(self.raw, stem, "host", app),
            "crash" if group == "C" else "refused", f"raw/parity/{stem}.host",
        )
        self.add(
            group, failing, "3-leaf-current", app, self.queue_package(failing),
            "crash" if group == "C" else "refused", f"queue package {failing}",
        )
        if fixed:
            self.add(
                group, failing, "4-leaf-fixed", app, self.dump(self.fixed, stem, "leaf", app),
                "opens", f"fixed parity dump {stem}.leaf",
            )
            self.add(
                group, failing, "6-host-fixed", app, self.dump(self.fixed, stem, "host", app),
                "opens", f"fixed parity dump {stem}.host",
            )

    def sibling(self, group, failing, app, identifier) -> None:
        self.add(
            group, failing, f"5-sibling-{identifier}", app, self.queue_package(identifier),
            "opens", f"queue package {identifier} (passed in run 1)",
        )

    def stem_package(self, stem: str) -> str:
        for identifier, package in self.packages.items():
            if any(source.get("case") == stem for source in package["sources"]):
                return identifier
        raise KeyError(stem)

    def build(self) -> None:
        for failing, app, stem, sibling in GROUP_A:
            self.family("A", failing, app, stem)
            sibling = sibling or self.stem_package(stem.split("~")[0])
            self.sibling("A", failing, app, sibling)
            self.add(
                "A", failing, "7-min-only-count-changed", app,
                variant_sparse_count(self.queue_package(failing), 3),
                "opens", f"queue package {failing} with only count set to 3",
            )
        for failing, stem, sibling in GROUP_B:
            # All but the first new caption start from a host-captioned chart,
            # which already carries the unlisted reference.
            expect = "opens" if failing == "P1141" else "refused"
            self.family("B", failing, "Numbers", stem, input_expect=expect)
            self.sibling("B", failing, "Numbers", sibling)
            self.add(
                "B", failing, "7-min-only-reference-listed", "Numbers",
                variant_list_references(self.queue_package(failing)),
                "opens", f"queue package {failing} with only the missing external reference",
            )
            self.add(
                "B", failing, "8-input-fixed", "Numbers",
                self.dump(self.fixed, stem, "input", "Numbers"),
                "opens", f"fixed parity dump {stem}.input",
            )
        synthetic = self.queue_package("P1096")
        self.add("C", "P1096", "1-input-host-leaf", "Numbers", synthetic, "crash",
                 "queue package P1096: read-only, the bytes of all three")
        for failing, stem in GROUP_C:
            self.family("C", failing, "Numbers", stem, input_expect="crash", fixed=False)
        self.sibling("C", "P1096", "Numbers", "P1146")
        basic = (self.native / "iwork/numbers/basic.numbers").read_bytes()
        self.add("C", "", "5-native-basic", "Numbers", basic, "opens",
                 "test-data/iwork/numbers/basic.numbers")
        for kind, drop in (
            ("7-no-cycle", CYCLE),
            ("7-only-inline-tree", CYCLE | OTHERS),
            ("7-without-inline-tree", INLINE),
            ("7-none-of-the-objects", CYCLE | OTHERS | INLINE),
        ):
            # Without any synthetic object this is basic.numbers again, written
            # by this tool: the control for the tool's own package writer.
            expect = "opens" if kind == "7-none-of-the-objects" else "unknown"
            self.add("C", "", kind, "Numbers", variant_drop_objects(synthetic, drop),
                     expect, f"P1096 without objects {sorted(drop)}")
        stem = "numbers_pivot_formula_author_inline_only_parity"
        for role, kind in (("input", "8-inline-only-input"), ("leaf", "8-inline-only-leaf"),
                           ("host", "8-inline-only-host")):
            self.add("C", "", kind, "Numbers", self.dump(self.fixed, stem, role, "Numbers"),
                     "unknown", f"fixed parity dump {stem}.{role}")
        # Run-1 passes whose bytes the count fix changes.
        for identifier, package in sorted(self.packages.items()):
            if package["app"] not in EXTENSION:
                continue
            data = self.queue_package(identifier)
            arrays = sparse_arrays(Package(data))
            if not any(e and max(e) >= count for _, count, e in arrays):
                continue
            if identifier in {row["failing"] for row in self.rows}:
                continue
            source = next(s for s in package["sources"] if s.get("case"))
            stem = source["case"]
            self.add("D", identifier, "revalidate-fixed", package["app"],
                     self.dump(self.fixed, stem, source["role"], package["app"]),
                     "opens", f"fixed parity dump {stem}.{source['role']}")
        self.write()

    def write(self) -> None:
        # Controls first: a refused package can leave a dialog that blocks the next.
        rank = {"5": 0, "7": 1, "8": 2, "1": 3, "4": 4, "6": 5, "2": 6, "3": 7, "r": 1}
        runnable = [row for row in self.rows if "path" in row]
        runnable.sort(key=lambda row: (row["group"] == "C", rank.get(row["kind"][0], 9), row["id"]))
        with (self.out / "queue.tsv").open("w", encoding="utf-8") as output:
            output.write("package_id\tapp\tpath\tsha256\trows\n")
            for row in runnable:
                output.write(f"{row['id']}\t{row['app']}\t{row['path']}\t{row['sha256']}\t0\n")
        (self.out / "manifest.json").write_text(
            json.dumps({"schema": 1, "record": "0431", "packages": self.rows}, indent=1) + "\n"
        )
        shutil.copy2(Path(__file__).parent / "run_on_mac.sh", self.out / "run_on_mac.sh")


def add_packages(kit: Path, listing: Path) -> int:
    """Append the packages of `listing` to the kit; nothing is written on a refusal."""
    manifest_path = kit / "manifest.json"
    manifest = json.loads(manifest_path.read_text())
    rows = manifest["packages"]
    ids = {row["id"] for row in rows}
    hashes = {row["sha256"] for row in rows}
    additions = []
    with listing.open(encoding="utf-8", newline="") as source:
        for line in csv.DictReader(source, delimiter="\t", quoting=csv.QUOTE_NONE):
            if line["app"] not in EXTENSION:
                raise ValueError(f"{line['id']}: unknown app {line['app']}")
            if line["expect"] not in {"opens", "refused", "crash", "unknown"}:
                raise ValueError(f"{line['id']}: unknown expectation {line['expect']}")
            data = (listing.parent / line["file"]).read_bytes()
            digest = hashlib.sha256(data).hexdigest()
            if line["id"] in ids or digest in hashes:
                raise ValueError(f"{line['id']}: id or bytes already in the kit")
            Package(data)
            ids.add(line["id"])
            hashes.add(digest)
            additions.append((line, data, digest))
    with (kit / "queue.tsv").open("a", encoding="utf-8") as queue:
        for line, data, digest in additions:
            path = f"packages/{line['app']}/{line['id']}.{EXTENSION[line['app']]}"
            target = kit / path
            target.parent.mkdir(parents=True, exist_ok=True)
            target.write_bytes(data)
            queue.write(f"{line['id']}\t{line['app']}\t{path}\t{digest}\t0\n")
            rows.append(
                {
                    "id": line["id"],
                    "group": "added",
                    "failing": "",
                    "kind": "added",
                    "app": line["app"],
                    "path": path,
                    "sha256": digest,
                    "size": len(data),
                    "expect": line["expect"],
                    "source": line["source"],
                }
            )
    manifest_path.write_text(json.dumps(manifest, indent=1) + "\n")
    print(f"added {len(additions)} packages; {len(rows)} in the kit")
    return 0


def next_run(queue_out: Path, unproven: Path, out: Path) -> int:
    """A kit of the queue packages `unproven` lists (`apply_results --unproven`),
    in the layout of `out/`. `apply_results.py` folds its results by hash, against
    the queue manifest, not this kit's."""
    manifest = json.loads((queue_out / "manifest.json").read_text())
    by_id = {package["id"]: package for package in manifest["packages"]}
    wanted = unproven.read_text(encoding="utf-8").split()
    rows = []
    queue = ["package_id\tapp\tpath\tsha256\trows"]
    for identifier in wanted:
        package = by_id[identifier]
        data = (queue_out / package["path"]).read_bytes()
        digest = hashlib.sha256(data).hexdigest()
        if digest != package["sha256"]:
            raise ValueError(f"{identifier}: the file does not match the manifest")
        target = out / package["path"]
        target.parent.mkdir(parents=True, exist_ok=True)
        target.write_bytes(data)
        queue.append("\t".join([identifier, package["app"], package["path"], digest, str(len(package["rows"]))]))
        rows.append(
            {
                "id": identifier,
                "group": "queue",
                "kind": "queue",
                "app": package["app"],
                "path": package["path"],
                "sha256": digest,
                "size": len(data),
                "expect": "opens",
                "source": ", ".join(source["case"] for source in package["sources"]),
            }
        )
    (out / "queue.tsv").write_text("\n".join(queue) + "\n", encoding="utf-8")
    (out / "manifest.json").write_text(
        json.dumps({"queue_manifest_sha256": hashlib.sha256((queue_out / "manifest.json").read_bytes()).hexdigest(),
                    "packages": rows}, indent=1) + "\n"
    )
    shutil.copy2(Path(__file__).resolve().parent / "run_on_mac.sh", out / "run_on_mac.sh")
    print(f"{len(rows)} queue packages in {out}")
    return 0


def variant_no_uuid_map(source: bytes) -> bytes:
    """Only the calculation engine's `uuid_reference_map` (field 7) is emptied."""
    package = Package(source)
    for _, item in package.objects():
        for message in item["messages"]:
            if next(v for n, _, v in message["info"] if n == 1) != CALCULATION_ENGINE:
                continue
            message["data"] = encode(
                [(n, w, b"" if n == UUID_REFERENCE_MAP else v) for n, w, v in fields(message["data"])]
            )
            return package.to_bytes({ENGINE})
    raise ValueError("no calculation engine")


def analyze(kit: Path, results: Path) -> int:
    rows = json.loads((kit / "manifest.json").read_text())["packages"]
    last = {}
    with results.open(encoding="utf-8", newline="") as source:
        for line in csv.DictReader(source, delimiter="\t", quoting=csv.QUOTE_NONE):
            last[line["package_id"]] = line
    surprises = 0
    for row in rows:
        result = last.get(row.get("same_as") or row["id"])
        actual = result["result"] if result else "absent"
        stage = f" {result['stage']}" if result and actual == "fail" else ""
        expect = row["expect"]
        surprising = (expect == "opens" and actual != "pass") or (
            expect in {"refused", "crash"} and actual == "pass"
        )
        surprises += surprising
        marker = "UNEXPECTED" if surprising else ""
        print(f"{row['id']}\t{expect}\t{actual}{stage}\t{marker}")
    return 1 if surprises else 0


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    commands = parser.add_subparsers(dest="command", required=True)
    audit = commands.add_parser("audit")
    audit.add_argument("files", nargs="+", type=Path)
    build = commands.add_parser("build")
    build.add_argument("--queue-out", type=Path, required=True)
    build.add_argument("--fixed", type=Path, required=True)
    build.add_argument("--native", type=Path, required=True)
    build.add_argument("--out", type=Path, required=True)
    extra = commands.add_parser("add")
    extra.add_argument("kit", type=Path)
    extra.add_argument("listing", type=Path)
    check = commands.add_parser("analyze")
    check.add_argument("kit", type=Path)
    check.add_argument("results", type=Path)
    following = commands.add_parser("next-run")
    following.add_argument("--queue-out", type=Path, required=True)
    following.add_argument("--unproven", type=Path, required=True)
    following.add_argument("--out", type=Path, required=True)
    no_map = commands.add_parser("variant-no-uuid-map")
    no_map.add_argument("input", type=Path)
    no_map.add_argument("output", type=Path)
    arguments = parser.parse_args()
    if arguments.command == "audit":
        for path in arguments.files:
            package = Package(path.read_bytes())
            print(path.name)
            print("  sparse (chart, count, indexes):", sparse_arrays(package))
            print("  unlisted (owner, type, target, type):", unlisted_references(package))
        return 0
    if arguments.command == "add":
        return add_packages(arguments.kit, arguments.listing)
    if arguments.command == "analyze":
        return analyze(arguments.kit, arguments.results)
    if arguments.command == "next-run":
        if arguments.out.exists():
            print(f"{arguments.out} exists; choose a new directory", file=sys.stderr)
            return 2
        return next_run(arguments.queue_out, arguments.unproven, arguments.out)
    if arguments.command == "variant-no-uuid-map":
        arguments.output.write_bytes(variant_no_uuid_map(arguments.input.read_bytes()))
        return 0
    if arguments.out.exists():
        print(f"{arguments.out} exists; choose a new directory", file=sys.stderr)
        return 2
    arguments.out.mkdir(parents=True)
    Kit(arguments.queue_out, arguments.fixed, arguments.native, arguments.out).build()
    return 0


if __name__ == "__main__":
    sys.exit(main())
