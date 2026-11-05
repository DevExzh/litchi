"""Python mirror of the record 0289 corpus fidelity report (design §14 layer 3).

``crates/litchi-ops/tests/corpus_fidelity.rs`` opens every claimed-format
file under ``test-data/``, runs an empty edit, saves it and requires the
saved bytes to equal the input; it writes the outcome of each file to
``docs/unified-ops/results/0289/fidelity.tsv`` (column ``noop``: ``exact``,
``refused``, ``signed-exact``, ``locked``, ...). This module replays a
deterministic sample of that report through ``litchi_py``:

* ``exact``: the :data:`PER_FORMAT` smallest files of each format (ties by
  path): ``open`` → ``File.edit`` with no operation → ``save_as`` must
  write the input bytes.
* ``signed-exact``: every such file: ``open`` → ``save_as`` must write the
  input bytes (a signed file refuses edits).
* ``refused``: every typed empty-edit refusal: the empty edit must raise a
  ``LitchiError`` with the recorded code, and ``open`` → ``save_as``, when
  it succeeds, must be exact (the Rust test fails on ``save-only DIFFERS``).
* ``locked``: every encrypted file must open locked.

The saved file keeps the input's extension, as ``save_bytes`` in the Rust
test does. Mismatches are collected and reported in one assertion message.

Rerun: ``crates/pyo3-litchi/tests/run_tests.sh --no-build -k fidelity_mirror``.
"""

from __future__ import annotations

import os
import tempfile
import unittest

from _support import L, REPO_ROOT, TEST_DATA

TSV = REPO_ROOT / "docs" / "unified-ops" / "results" / "0289" / "fidelity.tsv"

#: Exact no-op files replayed per format.
PER_FORMAT = 4


def load_rows():
    with open(TSV, encoding="utf-8") as handle:
        lines = handle.read().splitlines()
    header = lines[0].split("\t")
    rows = [dict(zip(header, line.split("\t"))) for line in lines[1:] if line]
    # Package directories (record 0306) are saved in place by the Rust test;
    # this mirror replays single files only.
    return [row for row in rows if not (TEST_DATA / row["file"]).is_dir()]


def sample(rows):
    """The exact files to replay: the smallest per format, ties by path."""
    by_format: dict = {}
    for row in rows:
        if row["noop"] == "exact":
            size = os.path.getsize(TEST_DATA / row["file"])
            by_format.setdefault(row["format"], []).append((size, row["file"], row))
    picked = []
    for fmt in sorted(by_format):
        picked.extend(row for _, _, row in sorted(by_format[fmt])[:PER_FORMAT])
    return picked


def describe(error: BaseException) -> str:
    code = getattr(error, "code", None)
    return f"{type(error).__name__}{f'[{code}]' if code else ''}: {str(error)[:200]}"


def saved_bytes(file, directory: str, relative: str) -> bytes:
    extension = relative.rsplit(".", 1)[-1].lower()
    path = os.path.join(directory, f"saved.{extension}")
    if os.path.exists(path):
        os.remove(path)
    file.save_as(path)
    with open(path, "rb") as handle:
        return handle.read()


def first_difference(a: bytes, b: bytes) -> str:
    offset = next((i for i, (x, y) in enumerate(zip(a, b)) if x != y), min(len(a), len(b)))
    return f"{len(a)} -> {len(b)} bytes, first difference at {offset}"


class FidelityMirrorTest(unittest.TestCase):
    def test_noop_saves_are_byte_identical(self):
        rows = load_rows()
        failures: list = []
        replayed = {"exact": 0, "signed-exact": 0, "refused": 0, "locked": 0}
        targets = sample(rows) + [r for r in rows if r["noop"] in ("signed-exact", "refused", "locked")]
        with tempfile.TemporaryDirectory() as directory:
            for row in targets:
                relative, noop = row["file"], row["noop"]
                label = f"{row['format']} {relative} {noop}"
                path = str(TEST_DATA / relative)
                with open(path, "rb") as handle:
                    data = handle.read()
                try:
                    file = L.open(path)
                    if noop == "locked":
                        if not file.locked:
                            failures.append(f"{label}: opened unlocked")
                        replayed[noop] += 1
                        continue
                    if noop == "signed-exact":
                        out = saved_bytes(file, directory, relative)
                    elif noop == "refused":
                        try:
                            file.edit(lambda tx: None)
                        except L.LitchiError as error:
                            code = row["detail"].split(": ", 2)[1] if ": " in row["detail"] else ""
                            if code and error.code != code:
                                failures.append(f"{label}: refused with {error.code}, Rust {code}")
                        else:
                            failures.append(f"{label}: the empty edit succeeds; Rust refused it")
                        # The detail is cut at 140 characters, so the
                        # save-only result may be missing; the Rust test
                        # fails on "save-only DIFFERS", so a save that
                        # succeeds must be exact, and a typed save refusal
                        # is accepted unless the report says it was exact.
                        try:
                            out = saved_bytes(file, directory, relative)
                        except L.LitchiError:
                            if "save-only exact" in row["detail"]:
                                raise
                            replayed[noop] += 1
                            continue
                    else:
                        commit = file.edit(lambda tx: None)
                        out = saved_bytes(commit.file, directory, relative)
                    file.close()
                except (KeyboardInterrupt, SystemExit):
                    raise
                except BaseException as error:  # noqa: BLE001 - PanicException
                    failures.append(f"{label}: {describe(error)}")
                    continue
                if out != data:
                    failures.append(f"{label}: saved bytes differ ({first_difference(data, out)})")
                replayed[noop] += 1
        if os.environ.get("LITCHI_MIRROR_VERBOSE"):
            import sys

            print(f"\nfidelity mirror: replayed {replayed}", file=sys.stderr)
        if failures:
            self.fail(
                f"{len(failures)} fidelity mismatches (replayed {replayed}):\n"
                + "\n".join(failures)
            )
        self.assertGreater(replayed["exact"], 0)


if __name__ == "__main__":
    unittest.main()
