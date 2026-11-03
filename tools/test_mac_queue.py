"""Unit tests of the macOS queue generator, runner and result folding (record 0430)."""

from __future__ import annotations

import hashlib
import json
import os
import shutil
import subprocess
import sys
import tempfile
import unittest
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent / "mac_queue"))

import apply_results  # noqa: E402
import generate  # noqa: E402
import rows as mac_queue  # noqa: E402

RUNNER = Path(__file__).resolve().parent / "mac_queue" / "run_on_mac.sh"

HEADER = (
    "| Wave | Format | Route | Reason queued | Fixture | Result | Artifact SHA-256 |\n"
    "|---|---|---|---|---|---|---|\n"
)


def queue_line(wave: str, fmt: str, route: str, reason: str, fixture: str, result: str = "pending", sha: str = "-") -> str:
    return f"| {wave} | {fmt} | {route} | {reason} | {fixture} | {result} | {sha} |\n"


def sha(data: bytes) -> str:
    return hashlib.sha256(data).hexdigest()


TESTS = {
    "keynote_chart_title_set_parity": "x.rs",
    "numbers_chart_title_set_parity": "x.rs",
    "pages_chart_title_set_parity": "x.rs",
    "pages_chart_title_remove_parity": "x.rs",
    "keynote_chart_series_trendlines_allocate_parity": "x.rs",
    "numbers_chart_series_trendlines_patch_parity": "x.rs",
    "pages_chart_series_trendlines_remove_parity": "x.rs",
    "pages_unrelated_reset_parity": "x.rs",
    "pages_text_box_remove_parity": "x.rs",
}


class ParseTests(unittest.TestCase):
    def test_rows_have_content_ids_and_duplicates_are_numbered(self) -> None:
        text = "# title\n\n" + HEADER + queue_line("W1", "Pages", "`a::b`", "r", "f") + queue_line("W1", "Pages", "`a::b`", "r", "f")
        rows = mac_queue.parse_queue(text)
        self.assertEqual(len(rows), 2)
        self.assertEqual(rows[1].id, rows[0].id + "-2")
        self.assertEqual(rows[0].line, 5)
        self.assertEqual(rows[0].route, "a::b")

    def test_id_ignores_result_columns(self) -> None:
        pending = mac_queue.parse_queue(HEADER + queue_line("W1", "Pages", "r", "x", "f"))
        done = mac_queue.parse_queue(HEADER + queue_line("W1", "Pages", "r", "x", "f", "pass (0430)", "ab" * 32))
        self.assertEqual(pending[0].id, done[0].id)

    def test_unexpected_header_is_refused(self) -> None:
        with self.assertRaises(ValueError):
            mac_queue.parse_queue("| A | B |\n|---|---|\n| 1 | 2 |\n")

    def test_braces_expand(self) -> None:
        self.assertEqual(mac_queue.expand_braces("a_{x,y}_b_{1,2}"), ["a_x_b_1", "a_x_b_2", "a_y_b_1", "a_y_b_2"])

    def test_citations(self) -> None:
        cited = mac_queue.cited_tests("tier A (pages_chart_title_{set,remove}_parity)", TESTS)
        self.assertEqual(cited, ["pages_chart_title_set_parity", "pages_chart_title_remove_parity"])
        star = mac_queue.cited_tests("(*_chart_title_set_parity, *_reset_parity)", TESTS)
        self.assertEqual(star, ["keynote_chart_title_set_parity", "numbers_chart_title_set_parity", "pages_chart_title_set_parity"])
        siblings = mac_queue.cited_tests("(keynote_chart_series_trendlines_allocate_parity and siblings)", TESTS)
        self.assertEqual(set(siblings), {
            "keynote_chart_series_trendlines_allocate_parity",
            "numbers_chart_series_trendlines_patch_parity",
            "pages_chart_series_trendlines_remove_parity",
        })

    def test_records(self) -> None:
        self.assertEqual(mac_queue.cited_records("no evidence (0375b); record 0372, records 0371 and 0368a"),
                         ["0375b", "0372", "0371", "0368a"])


class MappingTests(unittest.TestCase):
    def row(self, wave: str, fmt: str, route: str, reason: str = "r", fixture: str = "f") -> mac_queue.Row:
        return mac_queue.parse_queue(HEADER + queue_line(wave, fmt, route, reason, fixture))[0]

    def test_file_route_maps_to_example_stems(self) -> None:
        mapping = mac_queue.map_row(self.row("D7", "Pages", "`litchi_ops::File::set` (`chart.font.size=null`)",
                                             "covered by pages_chart_title_set_parity"), TESTS)
        self.assertEqual(mapping.sources, (mac_queue.Source("example", "pages-chart-font-size-reset", "file"),))

    def test_native_authoring_is_an_explicit_group(self) -> None:
        mapping = mac_queue.map_row(self.row("W1-F", "Numbers", "native-created volatile highlight (input)"), TESTS)
        self.assertEqual(mapping.sources, ())
        self.assertEqual(mapping.group, "native-authoring")

    def test_host_rows_take_the_host_output(self) -> None:
        mapping = mac_queue.map_row(self.row("W4-P", "Pages", "`litchi_iwa::pages::Draft::remove_text_box` (host fix)"), TESTS)
        self.assertEqual(mapping.sources, (mac_queue.Source("parity", "pages_text_box_remove_parity", "host"),))

    def test_ledger_fallback(self) -> None:
        mapping = mac_queue.map_row(self.row("W1", "Pages", "`litchi_pages::Package::x`", "as above"), TESTS,
                                    {"litchi_pages::Package::x": ["pages_chart_title_set_parity", "gone_parity"]})
        self.assertEqual(mapping.rule, "ledger")
        self.assertEqual(mapping.sources, (mac_queue.Source("parity", "pages_chart_title_set_parity", "leaf"),))

    def test_unmapped_row_is_reported(self) -> None:
        mapping = mac_queue.map_row(self.row("W9", "Pages", "`litchi_pages::Package::y`"), TESTS)
        self.assertEqual(mapping.rule, "unmapped")


class RealQueueTests(unittest.TestCase):
    """The tracked queue: every row maps to packages or an explicit group."""

    def test_every_queue_row_maps(self) -> None:
        rows, mappings, errors = generate.check_mapping()
        self.assertEqual(errors, [])
        self.assertGreater(len(rows), 700)
        for row in rows:
            mapping = mappings[row.id]
            self.assertTrue(mapping.sources or mapping.group, f"line {row.line}: {row.route}")
            for source in mapping.sources:
                if source.kind == "example":
                    self.assertIn(source.name, mac_queue.example_stems(), f"line {row.line}")

    def test_explicit_cases_exist(self) -> None:
        tests = mac_queue.parity_tests()
        self.assertEqual([case for case in mac_queue.explicit_cases() if case not in tests], [])

    def test_example_writes_every_stem(self) -> None:
        source = (mac_queue.ROOT / "crates/litchi-ops/examples/mac_queue_file_routes.rs").read_text(encoding="utf-8")
        for stem in set(mac_queue.FILE_ROUTE_STEMS.values()) | {"cell-value-set"}:
            self.assertIn(f'"{stem}"', source, stem)
        self.assertIn('"pages-builder-identity-{suffix}"', source)
        self.assertEqual(mac_queue.BUILDER_STEMS, ("pages-builder-identity-a", "pages-builder-identity-b"))


class FakeRepository:
    """A minimal repository root plus dump directory for `generate`."""

    def __init__(self, base: Path) -> None:
        self.root = base / "repo"
        self.out = base / "out"
        (self.root / "docs/unified-ops").mkdir(parents=True)
        (self.root / "docs/unified-ops/0375b-chart-text.md").write_text("record\n")
        parity = self.root / "crates/litchi-iwa/tests/parity"
        parity.mkdir(parents=True)
        parity.joinpath("x.rs").write_text("".join(f"#[test]\nfn {name}() {{}}\n" for name in TESTS))
        (self.root / "test-data/iwork/pages").mkdir(parents=True)
        self.fixture = b"tracked pages fixture"
        (self.root / "test-data/iwork/pages/basic.pages").write_bytes(self.fixture)
        queue = "# 0363\n\n" + HEADER
        queue += queue_line("W1-E", "Pages", "`litchi_pages::Package::edit_title`", "set (0375b)", "f; pages_chart_title_set_parity")
        queue += queue_line("W1-E", "Pages", "`litchi_pages::Package::edit_title`", "remove", "f; pages_chart_title_remove_parity")
        queue += queue_line("W1-E", "Pages", "`litchi_pages::Package::title`", "read", "pages_chart_title_{set,remove}_parity")
        queue += queue_line("W4-P", "Pages", "`litchi_iwa::pages::Draft::remove_text_box` (host fix)", "r", "basic.pages")
        queue += queue_line("W1-F", "Numbers", "native-created volatile highlight (input)", "r", "author it")
        (self.root / mac_queue.MAC_QUEUE).write_text(queue)
        self.dump = self.out / "raw/parity"
        self.dump.mkdir(parents=True)
        (self.out / "raw/example").mkdir(parents=True)
        (self.out / "raw/example/file-routes.jsonl").write_text("")

    def case(self, name: str, *, input_bytes: bytes, host: bytes, leaf: bytes, passed: bool = True,
             test: str | None = None) -> None:
        for role, data in (("input", input_bytes), ("host", host), ("leaf", leaf)):
            (self.dump / f"{name}.{role}.pages").write_bytes(data)
        (self.dump / f"{name}.json").write_text(json.dumps({
            "case": name, "test": f"x::{test or name}", "stem": name, "occurrence": 1, "format": "Pages", "tier": "A",
            "fixture": "iwork/pages/basic.pages", "read_only": False, "passed": passed, "extension": "pages",
        }))


class GenerateTests(unittest.TestCase):
    def setUp(self) -> None:
        self.base = Path(tempfile.mkdtemp())
        self.repo = FakeRepository(self.base)

    def tearDown(self) -> None:
        shutil.rmtree(self.base)

    def test_manifest_groups_identical_packages(self) -> None:
        host_input = b"host-built input"
        self.repo.case("pages_chart_title_set_parity", input_bytes=host_input, host=b"set", leaf=b"set")
        self.repo.case("pages_chart_title_remove_parity", input_bytes=host_input, host=b"removed", leaf=b"removed")
        self.repo.case("pages_text_box_remove_parity", input_bytes=self.repo.fixture, host=b"host fix", leaf=b"set")
        manifest = generate.generate(self.repo.out, self.repo.root)

        self.assertEqual(manifest["summary"]["rows"], 5)
        self.assertEqual(manifest["summary"]["packages"], 3)
        self.assertEqual(manifest["summary"]["rows_grouped"], 1)
        by_sha = {package["sha256"]: package for package in manifest["packages"]}
        set_package = by_sha[sha(b"set")]
        self.assertEqual(set_package["id"], "P0001")
        # The set row and the read row share the "set" package.
        self.assertEqual(len(set_package["rows"]), 2)
        rows = {entry["line"]: entry for entry in manifest["rows"]}
        self.assertEqual(rows[5]["packages"], ["P0001"])
        self.assertEqual(rows[7]["packages"], ["P0001", "P0002"])
        self.assertEqual(rows[8]["packages"], [by_sha[sha(b"host fix")]["id"]])
        self.assertEqual(rows[9]["status"], "grouped")
        self.assertEqual(rows[5]["records"], [{"record": "0375b", "path": "docs/unified-ops/0375b-chart-text.md"}])
        self.assertEqual(rows[5]["checks"], list(mac_queue.NATIVE_CHECKS))
        source = set_package["sources"][0]
        self.assertEqual(source["input_sha256"], sha(host_input))
        self.assertIsNone(source["input_tracked"])
        self.assertEqual(by_sha[sha(b"host fix")]["sources"][0]["input_tracked"], "test-data/iwork/pages/basic.pages")
        # Host-built inputs are kept for W3, tracked fixtures are not.
        self.assertEqual([item["sha256"] for item in manifest["host_built_inputs"]], [sha(host_input)])
        for package in manifest["packages"]:
            self.assertEqual(sha((self.repo.out / package["path"]).read_bytes()), package["sha256"])
        tsv = (self.repo.out / "queue.tsv").read_text().splitlines()
        self.assertEqual(tsv[0], "package_id\tapp\tpath\tsha256\trows")
        self.assertEqual(len(tsv), 4)
        self.assertTrue(tsv[1].startswith("P0001\tPages\tpackages/Pages/P0001-"))
        self.assertEqual(json.loads((self.repo.out / "manifest.json").read_text())["summary"], manifest["summary"])

    def test_failed_runs_are_negative_controls(self) -> None:
        for name in ("pages_chart_title_set_parity", "pages_chart_title_remove_parity", "pages_text_box_remove_parity"):
            self.repo.case(name, input_bytes=self.repo.fixture, host=b"h" + name.encode(), leaf=b"l" + name.encode())
        self.repo.case("pages_chart_title_rejects_corruption", input_bytes=self.repo.fixture, host=b"bad",
                       leaf=b"bad", passed=False, test="pages_chart_title_set_parity")
        manifest = generate.generate(self.repo.out, self.repo.root)
        self.assertEqual(manifest["summary"]["packages"], 3)
        self.assertEqual(manifest["summary"]["negative_control_runs_excluded"], 1)
        self.assertNotIn(sha(b"bad"), {package["sha256"] for package in manifest["packages"]})

    def test_cases_are_found_by_test_function_and_family(self) -> None:
        # The set test runs a case under another name; the remove test runs
        # no harness case, so its row falls back to module x's Pages runs.
        self.repo.case("pages_chart_text_set_case", input_bytes=self.repo.fixture, host=b"s", leaf=b"s",
                       test="pages_chart_title_set_parity")
        self.repo.case("pages_text_box_remove_parity", input_bytes=self.repo.fixture, host=b"h", leaf=b"l")
        manifest = generate.generate(self.repo.out, self.repo.root)
        rows = {entry["line"]: entry for entry in manifest["rows"]}
        by_sha = {package["sha256"]: package["id"] for package in manifest["packages"]}
        self.assertEqual(rows[5]["packages"], [by_sha[sha(b"s")]])
        self.assertEqual(rows[5]["family_fallback"], [])
        self.assertEqual(rows[6]["family_fallback"], ["pages_chart_title_remove_parity"])
        self.assertEqual(set(rows[6]["packages"]), {by_sha[sha(b"s")], by_sha[sha(b"l")]})

    def test_run_selection_covers_fixtures_first(self) -> None:
        runs = [{"fixture": f, "stem": f"c~{i}"} for i, f in enumerate("AAAABBC")]
        self.assertEqual([r["stem"] for r in generate.select_runs(runs, 2)], ["c~0", "c~4"])
        self.assertEqual([r["stem"] for r in generate.select_runs(runs, 3)], ["c~0", "c~4", "c~6"])
        four = [r["stem"] for r in generate.select_runs(runs, 4)]
        self.assertEqual(len(four), 4)
        self.assertTrue({"c~0", "c~4", "c~6"} <= set(four))
        self.assertEqual(generate.select_runs(runs, 0), runs)
        self.assertEqual(generate.select_runs(runs[:1], 2), runs[:1])

    def test_row_without_any_dump_is_refused(self) -> None:
        # No dump at all: the family fallback has nothing either.
        with self.assertRaises(SystemExit):
            generate.generate(self.repo.out, self.repo.root)


def result_line(package: str, digest: str, result: str = "pass", stage: str = "-", resaved: str = "c" * 64,
                date: str = "2026-11-10T09:00:00Z") -> str:
    return "\t".join([package, digest, "Pages", "14.4", "macOS 15.1", result, stage, resaved, "file", "d" * 64, date, "-"])


class ApplyTests(unittest.TestCase):
    def setUp(self) -> None:
        self.queue = "# 0363\n\n" + HEADER
        self.queue += queue_line("W1", "Pages", "`a`", "one", "f")
        self.queue += queue_line("W1", "Pages", "`b`", "two", "f")
        self.queue += queue_line("W1", "Pages", "`c`", "three", "f", "pass (earlier)", "e" * 64)
        self.queue += queue_line("W1", "Numbers", "native-created", "four", "f")
        rows = mac_queue.parse_queue(self.queue)
        self.ids = [row.id for row in rows]
        self.manifest = {
            "packages": [
                {"id": "P0001", "sha256": "1" * 64},
                {"id": "P0002", "sha256": "2" * 64},
            ],
            "rows": [
                {"id": self.ids[0], "packages": ["P0001"]},
                {"id": self.ids[1], "packages": ["P0001", "P0002"]},
                {"id": self.ids[2], "packages": ["P0002"]},
                {"id": self.ids[3], "packages": []},
            ],
        }
        self.base = Path(tempfile.mkdtemp())

    def tearDown(self) -> None:
        shutil.rmtree(self.base)

    def results(self, *lines: str) -> dict[str, dict[str, str]]:
        path = self.base / "results.tsv"
        path.write_text("\t".join(apply_results.RESULT_COLUMNS) + "\n" + "\n".join(lines) + "\n")
        return apply_results.read_results(path)

    def test_pass_fills_both_columns(self) -> None:
        results = self.results(result_line("P0001", "1" * 64, resaved="a" * 64))
        text, changes = apply_results.apply(self.queue, self.manifest, results)
        lines = text.splitlines()
        self.assertEqual(lines[4], f"| W1 | Pages | `a` | one | f | pass (0430: Pages 14.4, 2026-11-10) | {'a' * 64} |")
        # Row two waits for P0002; row four has no package.
        self.assertEqual(lines[5], queue_line("W1", "Pages", "`b`", "two", "f").rstrip("\n"))
        self.assertEqual(len(changes), 1)
        # The row id survives the update.
        self.assertEqual([row.id for row in mac_queue.parse_queue(text)], self.ids)

    def test_any_failure_fails_the_row_and_done_rows_are_kept(self) -> None:
        results = self.results(
            result_line("P0001", "1" * 64, resaved="a" * 64),
            result_line("P0002", "2" * 64, result="fail", stage="reopen", resaved="b" * 64),
        )
        text, changes = apply_results.apply(self.queue, self.manifest, results)
        lines = text.splitlines()
        self.assertIn("| FAIL (0430: P0002 reopen) | P0001 " + "a" * 64 + ", P0002 " + "b" * 64 + " |", lines[5])
        self.assertIn("pass (earlier)", lines[6])
        self.assertEqual(len(changes), 2)
        forced, _ = apply_results.apply(self.queue, self.manifest, results, force=True)
        self.assertIn("FAIL (0430: P0002 reopen)", forced.splitlines()[6])

    def test_last_result_wins_and_dry_runs_are_ignored(self) -> None:
        dry = "\t".join(["P0001", "1" * 64, "Pages", "14.4", "macOS", "dry-run", "-", "-", "-", "-", "t", "plan"])
        results = self.results(
            result_line("P0001", "1" * 64, result="fail", stage="open", resaved="-"),
            dry,
            result_line("P0001", "1" * 64, resaved="a" * 64),
        )
        self.assertEqual(results["P0001"]["result"], "pass")

    def test_hash_mismatch_and_stale_rows_are_errors(self) -> None:
        results = self.results(result_line("P0001", "9" * 64))
        with self.assertRaises(ValueError):
            apply_results.apply(self.queue, self.manifest, results)
        manifest = dict(self.manifest, rows=self.manifest["rows"] + [{"id": "q-000000000000", "packages": []}])
        with self.assertRaises(ValueError):
            apply_results.apply(self.queue, manifest, self.results())

    def test_command_line_dry_run_does_not_write(self) -> None:
        queue = self.base / "queue.md"
        queue.write_text(self.queue)
        (self.base / "manifest.json").write_text(json.dumps(self.manifest))
        (self.base / "results.tsv").write_text(
            "\t".join(apply_results.RESULT_COLUMNS) + "\n" + result_line("P0001", "1" * 64) + "\n")
        args = ["--manifest", str(self.base / "manifest.json"), "--results", str(self.base / "results.tsv"),
                "--queue", str(queue)]
        self.assertEqual(apply_results.main(args + ["--dry-run"]), 0)
        self.assertEqual(queue.read_text(), self.queue)
        self.assertEqual(apply_results.main(args), 0)
        self.assertIn("pass (0430", queue.read_text())


@unittest.skipUnless(shutil.which("bash") and shutil.which("shasum"), "needs bash and shasum")
class RunnerTests(unittest.TestCase):
    def test_syntax(self) -> None:
        subprocess.run(["bash", "-n", str(RUNNER)], check=True)

    def test_dry_run_verifies_hashes_without_an_application(self) -> None:
        base = Path(tempfile.mkdtemp())
        try:
            (base / "packages/Pages").mkdir(parents=True)
            good = base / "packages/Pages/P0001-x.pages"
            good.write_bytes(b"good")
            (base / "packages/Pages/P0002-y.pages").write_bytes(b"changed")
            (base / "queue.tsv").write_text(
                "package_id\tapp\tpath\tsha256\trows\n"
                f"P0001\tPages\tpackages/Pages/P0001-x.pages\t{sha(b'good')}\t2\n"
                f"P0002\tPages\tpackages/Pages/P0002-y.pages\t{sha(b'original')}\t1\n")
            run = subprocess.run(["bash", str(RUNNER), "--dry-run", str(base)], capture_output=True, text=True,
                                 env=dict(os.environ, HOME=str(base)))
            self.assertEqual(run.returncode, 1, run.stdout + run.stderr)
            results = apply_results.read_results(base / "results.tsv")
            self.assertEqual(list(results), ["P0002"])
            self.assertEqual((results["P0002"]["result"], results["P0002"]["stage"]), ("fail", "verify"))
            lines = (base / "results.tsv").read_text().splitlines()
            self.assertEqual(lines[0].split("\t"), list(apply_results.RESULT_COLUMNS))
            self.assertEqual(lines[1].split("\t")[5], "dry-run")
            self.assertFalse((base / "mac-run/P0001").exists())
        finally:
            shutil.rmtree(base)


if __name__ == "__main__":
    unittest.main()
