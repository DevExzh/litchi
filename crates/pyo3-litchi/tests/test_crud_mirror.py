"""Python mirror of the record 0289 CRUD checklist (design §14 layer 3).

``crates/litchi-ops/tests/crud_matrix.rs`` exercises every claimed
(format, element, verb) cell of the 21 claimed formats through the Rust
unified API and writes what passed to
``docs/unified-ops/results/0289/crud-checklist.tsv`` (columns ``format``,
``element``, ``verb``, ``claim``, ``follow_up``, ``observed``). This module
reads that TSV and replays every successful cell (``observed`` starting
``ok `` with claim ``supported`` or ``format-verb``) through ``litchi_py``,
so a cell that works in Rust but not through the bindings fails here.

Evidence formats (see ``run_format`` in crud_matrix.rs):

* get/find/view: ``ok N instances, first `<fixture>` <path>``; the first
  instance is replayed (``view`` is ``get`` at depth 0 with its text and
  ``format``, as the Rust cell reads ``view.text()`` and ``view.format()``).
* query: ``ok matches in <fixture>``; ``query(element)`` must match.
* set: ``ok `<fixture>` <path> {json}[ (value unchanged: no-op only)]``;
  edit (``ok tx.set``) and batch (``ok set item``) replay the same set
  through ``File.edit`` and an atomic ``File.batch``.
* add: ``ok <new path> on `<fixture>` under <parent> {json}[ at start]``;
  the returned path must equal the recorded one. ``<PNG>`` stands for
  :data:`PNG`.
* remove: ``ok `<fixture>` <path>``; move/swap: ``ok `<fixture>` <a> <b>``
  (move a after b; swap a and b).
* create/save: per format, ``create`` → ``to_bytes`` → ``open_bytes`` and
  ``save_as`` → ``open``; a format without a creator saves its first pinned
  fixture, as the Rust test does.

Every write result is published with ``to_bytes`` and reopened with
``open_bytes``. Fixture labels are test-data paths, ``create()``, ``seed``
(:data:`SEEDS`) and chains ``<base> + add <element>``, rebuilt from that
element's add cell exactly as ``run_format`` builds them (the add is made a
second time and kept when it succeeds and republishes).

Cells the TSV marks ``FAILED`` are known failures (``KNOWN_FAILURES`` in
crud_matrix.rs): they are replayed on the element's first instance (the
get cell's evidence) and must fail with a ``LitchiError``.

A second test checks that the Python capability report claims exactly the
cells the TSV calls ``supported``.

A third test replays the per-format ADR 0008 rows of record 0341
(``crud-adr0008.tsv``: atomic failure, patch inverse, composition). The
atomic-failure row runs the same two-item batch on a file opened from a
copy of the fixture and requires the refusal to carry the recorded code
with ``item == 1``, the File, its file on disk and two refused save
destinations unchanged. The composition row makes both edits on two
threads from one File and requires both orders to hold both values and to
give the same bytes (or, where the Rust row says the bytes differ, the same
text and node views), and each edit's ``Commit.patch`` applied to the other
edit's result to be refused with the recorded codes. The patch-inverse row
(record 0347) requires ``Commit.patch`` to replay onto its source File
(the commit's bytes), its ``inverse()`` to restore the exact source bytes,
the patch on its own result to be refused with the recorded code, and a
same-format byte-identical, separately opened copy to accept the patch
through the bounded content proof.

Rerun: ``crates/pyo3-litchi/tests/run_tests.sh --no-build -k crud_mirror``;
set ``LITCHI_MIRROR_VERBOSE=1`` to print the replayed cells per verb.
"""

from __future__ import annotations

import collections
import json
import os
import re
import sys
import tempfile
import unittest

from _support import L, REPO_ROOT, TEST_DATA

TSV = REPO_ROOT / "docs" / "unified-ops" / "results" / "0289" / "crud-checklist.tsv"
ADR_TSV = REPO_ROOT / "docs" / "unified-ops" / "results" / "0289" / "crud-adr0008.tsv"

#: ``PNG`` of crud_matrix.rs; the TSV abbreviates it to ``<PNG>``.
PNG = (
    "data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAQAAAACCAYAAAB/qH1jAAAAEklEQVR4nGP4z8DwHxkzoAsAAA8hD/"
    "EEN8afAAAAAElFTkSuQmCC"
)

#: ``SEEDS`` of crud_matrix.rs (keep in step): documents built from
#: ``create()`` with the unified verbs, labelled ``seed``. Each step is
#: ``("add", parent, type, props)`` or ``("set", path, props)``.
SEEDS = {
    "docx": [
        ("add", "/body", "paragraph", {"text": "One"}),
        ("add", "/body", "paragraph", {"text": "Two"}),
    ],
    "pptx": [
        ("add", "/", "slide", {"layout": "Blank"}),
        ("add", "/", "slide", {"layout": "Blank"}),
        ("add", "/slide[1]", "shape", {"text": "Box", "width": "1in", "height": "1in"}),
    ],
    "rtf": [
        ("add", "/body", "paragraph", {"text": "One"}),
        ("add", "/body", "paragraph", {"text": "Two"}),
    ],
    "ods": [("set", "/Sheet1/A1:B2", {"merge": "true"})],
    # Keep in step with crud_matrix.rs: Pages needs a named, non-empty second
    # section so section-removal and paragraph rows exercise the same seed.
    "pages": [
        ("add", "/body", "paragraph", {"text": "Intro"}),
        ("add", "/", "section", {"name": "Appendix"}),
        ("set", "/section[2]", {"text": "Tail"}),
    ],
    # Record 0309: character properties are written on runs of text boxes.
    "keynote": [("add", "/slide[1]", "shape", {"text": "Box\nSecond"})],
}

#: ``Spec.extension`` of crud_matrix.rs where it differs from the format name.
EXTENSIONS = {"keynote": "key"}

#: Verbs the capability cross-check compares, mapped to the per-element
#: operation ``element_claim`` (crud_matrix.rs) reads; ``None`` means the
#: verb applies to every element without a gap. ``move`` and ``swap`` are
#: format verbs the report does not declare per element, so they cannot be
#: mapped and are not compared.
CLAIM_OPERATION = {
    "get": "get", "find": "get", "view": "get", "query": "query",
    "set": "set", "edit": "set", "batch": "set",
    "add": "add", "remove": "remove", "create": None, "save": None,
}

GET_RE = re.compile(r"^ok (\d+) instances, first `([^`]+)` (.+)$")
QUERY_RE = re.compile(r"^ok matches in (.+)$")
SET_RE = re.compile(r"^ok `([^`]+)` (.+?) (\{.*\})( \(value unchanged: no-op only\))?$")
ADD_RE = re.compile(r"^ok (\S+) on `([^`]+)` under (.+?) (\{.*\})( at start)?$")
PAIR_RE = re.compile(r"^ok `([^`]+)` (\S+) (\S+)$")
ONE_RE = re.compile(r"^ok `([^`]+)` (.+)$")


def load_rows():
    with open(TSV, encoding="utf-8") as handle:
        lines = handle.read().splitlines()
    header = lines[0].split("\t")
    return [dict(zip(header, line.split("\t"))) for line in lines[1:] if line]


def fmt_name(file) -> str:
    return str(file.format).lower() if file.format is not None else "none"


def describe(error: BaseException) -> str:
    code = getattr(error, "code", None)
    head = f"{type(error).__name__}"
    if code:
        head += f"[{code}]"
    return f"{head}: {str(error)[:200]}"


def guarded(fn):
    """``(value, None)`` or ``(None, error)``; panics (BaseException) count too."""
    try:
        return fn(), None
    except (KeyboardInterrupt, SystemExit):
        raise
    except BaseException as error:  # noqa: BLE001 - PanicException is a BaseException
        return None, error


def republish(file):
    return L.open_bytes(file.to_bytes())


def props_of(text: str) -> dict:
    props = json.loads(text)
    return {key: (PNG if value == "<PNG>" else value) for key, value in props.items()}


class Checklist:
    """The TSV, indexed, and the fixtures its evidence names (cached)."""

    def __init__(self):
        self.rows = load_rows()
        self.cells = {(r["format"], r["element"], r["verb"]): r for r in self.rows}
        self.formats = list(dict.fromkeys(r["format"] for r in self.rows))
        self._fixtures: dict = {}

    def fixture(self, fmt: str, label: str):
        key = (fmt, label)
        if key not in self._fixtures:
            self._fixtures[key] = self._build(fmt, label)
        return self._fixtures[key]

    def _build(self, fmt: str, label: str):
        if label == "create()":
            return L.create(fmt)
        if label == "seed":
            file = self.fixture(fmt, "create()")
            for step in SEEDS[fmt]:
                if step[0] == "add":
                    file, _ = file.add(step[1], step[2], step[3])
                else:
                    file = file.set(step[1], step[2])
            return file
        if " + add " in label:
            base, element = label.rsplit(" + add ", 1)
            add = self.add_evidence(fmt, element)
            if add is None or add["base"] != base:
                raise AssertionError(f"chain {label!r}: no add cell for {element} on {base!r}")
            added = self.replay_add(fmt, element, add)
            # crud_matrix.rs: a second, identical add, kept when it succeeds
            # and republishes.
            twice, error = guarded(lambda: self._add(added, element, add))
            if error is None and guarded(lambda: republish(twice[0]))[1] is None:
                return twice[0]
            return added
        path = TEST_DATA / label
        with open(path, "rb") as handle:
            file = L.open_bytes(handle.read())
        if fmt_name(file) != fmt:
            raise AssertionError(f"{label} opened as {file.format}")
        return file

    def add_evidence(self, fmt: str, element: str):
        row = self.cells.get((fmt, element, "add"))
        if row is None:
            return None
        match = ADD_RE.match(row["observed"])
        if match is None:
            return None
        return {
            "path": match.group(1),
            "base": match.group(2),
            "parent": match.group(3),
            "props": props_of(match.group(4)),
            "start": bool(match.group(5)),
        }

    @staticmethod
    def _add(file, element, add):
        if add["start"]:
            return file.add(add["parent"], element, add["props"], index=0)
        return file.add(add["parent"], element, add["props"])

    def replay_add(self, fmt: str, element: str, add):
        """The add cell's result, checked against its recorded path."""
        base = self.fixture(fmt, add["base"])
        file, path = self._add(base, element, add)
        if path != add["path"]:
            raise AssertionError(f"add returned {path}, Rust recorded {add['path']}")
        republish(file)
        return file


_CHECKLIST = None


def checklist() -> Checklist:
    global _CHECKLIST
    if _CHECKLIST is None:
        _CHECKLIST = Checklist()
    return _CHECKLIST


class CrudMirrorTest(unittest.TestCase):
    """Every ``ok`` cell of the TSV replays through litchi_py."""

    def test_replay_ok_cells(self):
        cl = checklist()
        failures: list = []
        counts: collections.Counter = collections.Counter()
        formats_seen: set = set()

        def fail(row, reason):
            failures.append(f"{row['format']} {row['element']} {row['verb']}: {reason}")

        def run(row, fn):
            _, error = guarded(fn)
            if error is not None:
                fail(row, describe(error))
            else:
                counts[row["verb"]] += 1
                formats_seen.add(row["format"])

        for row in cl.rows:
            obs, verb, fmt, element = row["observed"], row["verb"], row["format"], row["element"]
            if row["claim"] not in ("supported", "format-verb") or not obs.startswith("ok "):
                continue
            if verb in ("create", "save"):
                continue  # per format, below
            if verb in ("get", "find", "view"):
                match = GET_RE.match(obs)
                if not match:
                    fail(row, f"unparsed evidence {obs!r}")
                    continue
                label, path = match.group(2), match.group(3)

                def read(label=label, path=path, verb=verb, fmt=fmt):
                    file = cl.fixture(fmt, label)
                    if verb == "get":
                        file.get(path)
                    elif verb == "find":
                        if file.find(path) is None:
                            raise AssertionError(f"find({path}) returned None")
                    else:
                        view = file.get(path, 0)
                        if "format" not in view:
                            raise AssertionError(f"view of {path} has no format")

                run(row, read)
            elif verb == "query":
                match = QUERY_RE.match(obs)
                label = match.group(1)

                def query(label=label, fmt=fmt, element=element):
                    if not cl.fixture(fmt, label).query(element, limit=3):
                        raise AssertionError(f"query({element}) matched nothing in {label}")

                run(row, query)
            elif verb in ("set", "edit", "batch"):
                source = row if verb == "set" else cl.cells[(fmt, element, "set")]
                match = SET_RE.match(source["observed"])
                if not match:
                    fail(row, f"unparsed set evidence {source['observed']!r}")
                    continue
                label, path, props = match.group(1), match.group(2), props_of(match.group(3))

                def write(label=label, path=path, props=props, verb=verb, fmt=fmt):
                    file = cl.fixture(fmt, label)
                    if verb == "set":
                        result = file.set(path, props)
                    elif verb == "edit":
                        result = file.edit(lambda tx: tx.set(path, props)).file
                    else:
                        item = {"command": "set", "path": path, "props": props}
                        result = file.batch([item], mode="atomic").file
                    republish(result)

                run(row, write)
            elif verb == "add":
                add = cl.add_evidence(fmt, element)
                if add is None:
                    fail(row, f"unparsed evidence {obs!r}")
                    continue
                run(row, lambda fmt=fmt, element=element, add=add: cl.replay_add(fmt, element, add))
            elif verb == "remove":
                match = ONE_RE.match(obs)
                label, path = match.group(1), match.group(2)
                run(row, lambda label=label, path=path, fmt=fmt: republish(
                    cl.fixture(fmt, label).remove(path)))
            elif verb in ("move", "swap"):
                match = PAIR_RE.match(obs)
                if not match:
                    fail(row, f"unparsed evidence {obs!r}")
                    continue
                label, a, b = match.groups()

                def pair(label=label, a=a, b=b, verb=verb, fmt=fmt):
                    file = cl.fixture(fmt, label)
                    if verb == "move":
                        result = file.move(a, after=b)
                    else:
                        result = file.swap(a, b)
                    republish(result)

                run(row, pair)
            else:
                fail(row, f"unknown verb with evidence {obs!r}")

        # Known failures (KNOWN_FAILURES of crud_matrix.rs): the evidence
        # has no path, so replay on the element's first instance (the get
        # cell's evidence) and require a typed refusal.
        known = 0
        for row in cl.rows:
            if not row["observed"].startswith("FAILED"):
                continue
            fmt, element, verb = row["format"], row["element"], row["verb"]
            get = GET_RE.match(cl.cells[(fmt, element, "get")]["observed"])
            if verb != "remove" or get is None:
                continue  # not replayable from the evidence
            label, path = get.group(2), get.group(3)
            known += 1
            _, error = guarded(lambda: republish(cl.fixture(fmt, label).remove(path)))
            if error is None:
                fail(row, f"known failure now succeeds on `{label}` {path}")
            elif not isinstance(error, L.LitchiError):
                fail(row, f"known failure raised untyped {describe(error)}")

        # create/save, per format (the create/save cells of every element
        # row carry the same format result).
        for fmt in cl.formats:
            rows = [r for r in cl.rows if r["format"] == fmt]
            create_ok = any(r["verb"] == "create" and r["observed"] == "ok format" for r in rows)
            save_ok = any(r["verb"] == "save" and r["observed"] == "ok format" for r in rows)
            row = {"format": fmt, "element": "*", "verb": "create"}
            created, error = guarded(lambda: L.create(fmt))
            if create_ok:
                if error is not None:
                    fail(row, describe(error))
                else:
                    reopened, error = guarded(lambda: republish(created))
                    if error is not None:
                        fail(row, describe(error))
                    elif fmt_name(reopened) != fmt:
                        fail(row, f"reopened as {reopened.format}")
                    else:
                        counts["create"] += 1
                        formats_seen.add(fmt)
            elif error is None:
                fail(row, "create succeeds although the TSV records a refusal")
            elif not isinstance(error, L.LitchiError):
                fail(row, f"untyped {describe(error)}")
            if not save_ok:
                continue
            row = {"format": fmt, "element": "*", "verb": "save"}
            base = created
            if base is None:
                # crud_matrix.rs saves the first pinned fixture instead.
                label = next(
                    (GET_RE.match(r["observed"]).group(2) for r in rows
                     if r["verb"] == "get" and GET_RE.match(r["observed"])
                     and "/" in GET_RE.match(r["observed"]).group(2)),
                    None,
                )
                base, error = guarded(lambda: cl.fixture(fmt, label))
                if error is not None:
                    fail(row, f"no base to save: {describe(error)}")
                    continue

            def save(base=base, fmt=fmt):
                with tempfile.TemporaryDirectory() as directory:
                    path = os.path.join(directory, f"saved.{EXTENSIONS.get(fmt, fmt)}")
                    base.save_as(path)
                    reopened = L.open(path)
                    if fmt_name(reopened) != fmt:
                        raise AssertionError(f"reopened as {reopened.format}")
                    reopened.close()

            _, error = guarded(save)
            if error is not None:
                fail(row, describe(error))
            else:
                counts["save"] += 1
                formats_seen.add(fmt)

        summary = ", ".join(f"{verb} {counts[verb]}" for verb in sorted(counts))
        if os.environ.get("LITCHI_MIRROR_VERBOSE"):
            print(f"\ncrud mirror: replayed {summary}; known failures {known}; "
                  f"formats {len(formats_seen)}", file=sys.stderr)
        if failures:
            self.fail(
                f"{len(failures)} replay mismatches (replayed: {summary}; known failures "
                f"checked: {known}):\n" + "\n".join(failures)
            )
        # Every format has at least one replayed cell.
        self.assertEqual(sorted(formats_seen), sorted(cl.formats))

    def test_capabilities_agree_with_claims(self):
        """The TSV's ``supported`` claims equal the Python capability report.

        Mirrors ``element_claim``: a verb the format refuses, or an element
        with a ``gap``, is never supported; otherwise the element's
        operation decides (``CLAIM_OPERATION``). move/swap are not compared.
        """
        cl = checklist()
        failures: list = []
        for fmt in cl.formats:
            try:
                file = L.create(fmt)
            except L.LitchiError:
                label = next(
                    GET_RE.match(r["observed"]).group(2) for r in cl.rows
                    if r["format"] == fmt and r["verb"] == "get" and GET_RE.match(r["observed"])
                )
                file = cl.fixture(fmt, label)
            caps = file.capabilities()
            elements = {element["element"]: element for element in caps["elements"]}
            tsv_elements = list(dict.fromkeys(
                r["element"] for r in cl.rows if r["format"] == fmt))
            if list(elements) != tsv_elements:
                failures.append(
                    f"{fmt} * elements: Python {list(elements)} != TSV {tsv_elements}")
            for (cell_fmt, name, verb), row in cl.cells.items():
                if cell_fmt != fmt or verb not in CLAIM_OPERATION or name not in elements:
                    continue
                element = elements[name]
                operation = CLAIM_OPERATION[verb]
                supported = caps["verbs"].get(verb, {}).get("supported", False) and \
                    "gap" not in element and \
                    (operation is None or bool(element["operations"].get(operation)))
                claimed = row["claim"] == "supported"
                if supported != claimed:
                    failures.append(
                        f"{fmt} {name} {verb}: TSV claim {row['claim']}, "
                        f"Python report says {'supported' if supported else 'not supported'}")
        if failures:
            self.fail(f"{len(failures)} claim mismatches:\n" + "\n".join(failures))


def load_adr_rows():
    with open(ADR_TSV, encoding="utf-8") as handle:
        lines = handle.read().splitlines()
    header = lines[0].split("\t")
    return [dict(zip(header, line.split("\t"))) for line in lines[1:] if line]


class Adr0008MirrorTest(unittest.TestCase):
    """The ADR 0008 rows of ``crud-adr0008.tsv`` replay through litchi_py."""

    def test_replay_adr_rows(self):
        cl = checklist()
        rows = load_adr_rows()
        failures: list = []
        counts: collections.Counter = collections.Counter()
        self.assertEqual(
            sorted({row["format"] for row in rows}), sorted(cl.formats),
            "every checklist format has its ADR 0008 rows")
        for row in rows:
            fmt, kind = row["format"], row["row"]
            if row["status"] != "ok":
                # n/a rows carry their typed reason in the Rust report.
                if row["status"] != "n/a":
                    failures.append(f"{fmt} {kind}: status {row['status']}")
                continue
            replay = {
                "atomic failure": self._atomic,
                "patch inverse": self._inverse,
                "composition": self._compose,
            }[kind]
            _, error = guarded(lambda replay=replay, row=row: replay(cl, row))
            if error is not None:
                failures.append(f"{fmt} {kind}: {describe(error)}")
            else:
                counts[kind] += 1
        if os.environ.get("LITCHI_MIRROR_VERBOSE"):
            print(f"\nADR 0008 mirror: {dict(counts)}", file=sys.stderr)
        if failures:
            self.fail(f"{len(failures)} ADR 0008 mismatches:\n" + "\n".join(failures))

    @staticmethod
    def _atomic(cl, row):
        fmt = row["format"]
        extension = EXTENSIONS.get(fmt, fmt)
        file = cl.fixture(fmt, row["fixture"])
        props = props_of(row["props"])
        before = file.to_bytes()
        with tempfile.TemporaryDirectory() as directory:
            source = os.path.join(directory, f"source.{extension}")
            with open(source, "wb") as handle:
                handle.write(before)
            origin = L.open(source)
            try:
                items = [
                    {"command": "set", "path": row["path"], "props": props},
                    {"command": "set", "path": row["path2"], "props": props},
                ]
                try:
                    origin.batch(items, mode="atomic")
                except L.LitchiError as error:
                    if error.code != row["code"] or error.item != 1:
                        raise AssertionError(
                            f"refused {error.code} at item {error.item}, Rust recorded "
                            f"{row['code']} at item 1") from error
                else:
                    raise AssertionError("the batch with a failing item 1 succeeded")
                if origin.to_bytes() != before:
                    raise AssertionError("the File changed after the refused batch")
                with open(source, "rb") as handle:
                    if handle.read() != before:
                        raise AssertionError("the File's source file changed")
                other = "xlsx" if extension == "docx" else "docx"
                sentinel = b"litchi sentinel"
                wrong = os.path.join(directory, f"destination.{other}")
                blocker = os.path.join(directory, "blocker")
                for path in (wrong, blocker):
                    with open(path, "wb") as handle:
                        handle.write(sentinel)
                for path in (wrong, os.path.join(blocker, f"saved.{extension}")):
                    try:
                        origin.save_as(path)
                    except L.LitchiError:
                        pass
                    else:
                        raise AssertionError(f"save_as {path} succeeded")
                for path in (wrong, blocker):
                    with open(path, "rb") as handle:
                        if handle.read() != sentinel:
                            raise AssertionError(f"{path} changed")
                expected = sorted(["blocker", f"destination.{other}", f"source.{extension}"])
                if sorted(os.listdir(directory)) != expected:
                    raise AssertionError(f"refused saves left {sorted(os.listdir(directory))}")
            finally:
                origin.close()

    @staticmethod
    def _inverse(cl, row):
        file = cl.fixture(row["format"], row["fixture"])
        before = file.to_bytes()
        props = props_of(row["props"])
        commit = file.edit(lambda tx: tx.set(row["path"], props))
        after = commit.file.to_bytes()
        if not commit.changed or after == before:
            raise AssertionError("the edit changed nothing")
        if file.to_bytes() != before:
            raise AssertionError("the edit changed its source File")
        republish(commit.file)
        patch = commit.patch
        if patch is None or patch.is_noop():
            raise AssertionError("the changing edit has no patch")
        if patch.apply(file).to_bytes() != after:
            raise AssertionError("the patch replayed onto its source differs from the commit")
        inverse = patch.inverse()
        if inverse.is_noop():
            raise AssertionError("the inverse of a change is a no-op")
        if inverse.apply(commit.file).to_bytes() != before:
            raise AssertionError("the inverse does not restore the exact source bytes")
        if file.to_bytes() != before or commit.file.to_bytes() != after:
            raise AssertionError("applying a patch changed its input")
        try:
            patch.apply(commit.file)
        except L.LitchiError as error:
            if error.code != row["code"]:
                raise AssertionError(
                    f"the patch on its own result refused {error.code}, Rust recorded "
                    f"{row['code']}") from error
        else:
            raise AssertionError("the patch applied to its own result")
        foreign = L.open_bytes(before)
        try:
            applied = patch.apply(foreign)
            if applied.to_bytes() != after:
                raise AssertionError("a byte-identical copy took the patch to other bytes")
        except L.LitchiError as error:
            raise AssertionError(
                f"a byte-identical reopened copy refused {error.code}; Rust: {row['note']}"
            ) from error
        finally:
            foreign.close()
        if "a byte-identical reopened copy accepts (content-checked)" not in row["note"]:
            raise AssertionError(f"Rust did not record content-checked replay: {row['note']}")

    @staticmethod
    def _compose(cl, row):
        import concurrent.futures

        file = cl.fixture(row["format"], row["fixture"])
        a, b = (row["path"], props_of(row["props"])), (row["path2"], props_of(row["props2"]))
        with concurrent.futures.ThreadPoolExecutor(max_workers=2) as pool:
            commits = list(pool.map(
                lambda edit: file.edit(lambda tx: tx.set(edit[0], edit[1])), [a, b]))
        first, second = commits[0].file, commits[1].file
        ab, ba = first.set(*b), second.set(*a)
        # Neither edit can be silently dropped: each patch refuses the
        # other's result, with the codes the Rust row records.
        codes = []
        for commit, other in ((commits[0], second), (commits[1], first)):
            try:
                commit.patch.apply(other)
            except L.LitchiError as error:
                codes.append(error.code)
            else:
                raise AssertionError("a patch applied to the other edit's result")
        if "/".join(codes) != row["code"]:
            raise AssertionError(f"stale patches refused {'/'.join(codes)}, Rust {row['code']}")

        def value(target, edit):
            view = target.get(edit[0], 0)
            key = next(iter(edit[1]))
            return view.get("format", {}).get(key), view.get("text")

        want = (value(first, a), value(second, b))
        for order, composed in (("a then b", ab), ("b then a", ba)):
            reopened = republish(composed)
            got = (value(reopened, a), value(reopened, b))
            if got != want:
                raise AssertionError(f"{order} dropped an edit: {got} != {want}")
        if first.set(*b).to_bytes() != ab.to_bytes():
            raise AssertionError("a then b is not repeatable")
        if "the same bytes" in row["note"]:
            if ab.to_bytes() != ba.to_bytes():
                raise AssertionError("the two orders give different bytes")
        else:
            left, right = republish(ab), republish(ba)
            if left.text() != right.text() or left.get("/", 64) != right.get("/", 64):
                raise AssertionError("the two orders give different documents")


if __name__ == "__main__":
    unittest.main()
