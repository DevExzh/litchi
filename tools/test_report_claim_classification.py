"""Fail-closed tests for the historical REPORT claim classification sidecar."""

from __future__ import annotations

import copy
import io
import tempfile
import unittest
from contextlib import redirect_stderr
from pathlib import Path

from tools import check_report_claim_classification as checker


REPO_ROOT = Path(__file__).resolve().parents[1]
REGISTRY_PATH = REPO_ROOT / "docs" / "performance" / "report-claim-classification-v1.json"


def load_seed() -> dict:
    return checker.load_json(REGISTRY_PATH)


def write_report_fixture(root: Path, report_text: str) -> None:
    report = root / checker.REPORT_RELATIVE_PATH
    report.parent.mkdir(parents=True, exist_ok=True)
    report.write_text(report_text, encoding="utf-8")


def write_claim_fixture(root: Path, claim_text: str | None = None) -> None:
    claim_registry = root / checker.CLAIM_REGISTRY_RELATIVE_PATH
    claim_registry.parent.mkdir(parents=True, exist_ok=True)
    claim_registry.write_text(
        claim_text
        if claim_text is not None
        else (REPO_ROOT / checker.CLAIM_REGISTRY_RELATIVE_PATH).read_text(encoding="utf-8"),
        encoding="utf-8",
    )


class ReportClassificationTests(unittest.TestCase):
    def test_seed_registry_is_valid_and_exhaustive(self) -> None:
        summary = checker.validate_registry(load_seed(), repo_root=REPO_ROOT)
        self.assertEqual(summary["table_count"], 2)
        self.assertEqual(summary["row_count"], 167)
        self.assertEqual(
            summary["state_counts"],
            {"strict_claim": 0, "historical": 145, "descriptive": 14, "withheld": 8},
        )

    def test_cli_smoke(self) -> None:
        status, message = checker.lint_registry(REGISTRY_PATH, repo_root=REPO_ROOT)
        self.assertEqual(status, 0, message)
        self.assertIn("167 REPORT rows", message)

    def test_table_identity_and_order_are_bound(self) -> None:
        registry = load_seed()
        self.assertEqual(
            [table["id"] for table in registry["report"]["tables"]],
            ["historical-stable-tranche", "historical-accepted-results"],
        )
        self.assertEqual(registry["report"]["tables"][0]["row_count"], 88)
        self.assertEqual(registry["report"]["tables"][1]["row_count"], 79)
        self.assertEqual(registry["strict_claim_registry"]["links"], [])

    def test_row_digest_tampering_is_rejected(self) -> None:
        registry = load_seed()
        registry["report"]["tables"][0]["rows"][0]["row_sha256"] = "0" * 64
        with self.assertRaises(checker.ClassificationPolicyError):
            checker.validate_registry(registry, repo_root=REPO_ROOT)

    def test_report_row_tampering_is_rejected(self) -> None:
        registry = load_seed()
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            report = root / checker.REPORT_RELATIVE_PATH
            report.parent.mkdir(parents=True)
            original = (REPO_ROOT / checker.REPORT_RELATIVE_PATH).read_text(encoding="utf-8")
            report.write_text(original.replace("XLSX row-start index", "XLSX row-start tampered", 1), encoding="utf-8")
            (root / checker.CLAIM_REGISTRY_RELATIVE_PATH).parent.mkdir(parents=True, exist_ok=True)
            (root / checker.CLAIM_REGISTRY_RELATIVE_PATH).write_text(
                (REPO_ROOT / checker.CLAIM_REGISTRY_RELATIVE_PATH).read_text(encoding="utf-8"),
                encoding="utf-8",
            )
            with self.assertRaises(checker.ClassificationPolicyError):
                checker.validate_registry(registry, repo_root=root)

    def test_duplicate_label_is_rejected(self) -> None:
        registry = load_seed()
        rows = registry["report"]["tables"][0]["rows"]
        rows[1]["label"] = rows[0]["label"]
        with self.assertRaises(checker.ClassificationPolicyError):
            checker.validate_registry(registry, repo_root=REPO_ROOT)

    def test_duplicate_digest_is_rejected(self) -> None:
        registry = load_seed()
        rows = registry["report"]["tables"][0]["rows"]
        rows[1]["row_sha256"] = rows[0]["row_sha256"]
        with self.assertRaises(checker.ClassificationPolicyError):
            checker.validate_registry(registry, repo_root=REPO_ROOT)

    def test_ordinal_reordering_is_rejected(self) -> None:
        registry = load_seed()
        rows = registry["report"]["tables"][0]["rows"]
        rows[0], rows[1] = rows[1], rows[0]
        with self.assertRaises(checker.ClassificationPolicyError):
            checker.validate_registry(registry, repo_root=REPO_ROOT)

    def test_unknown_state_is_rejected(self) -> None:
        registry = load_seed()
        registry["report"]["tables"][0]["rows"][0]["state"] = "accepted"
        with self.assertRaises(checker.ClassificationPolicyError):
            checker.validate_registry(registry, repo_root=REPO_ROOT)

    def test_header_identity_is_rejected_when_changed(self) -> None:
        registry = load_seed()
        registry["report"]["tables"][0]["header"][0] = "Current evidence"
        with self.assertRaises(checker.ClassificationInputError):
            checker.validate_registry(registry, repo_root=REPO_ROOT)

    def test_duplicate_json_keys_are_rejected(self) -> None:
        with tempfile.TemporaryDirectory() as directory:
            path = Path(directory) / "duplicate.json"
            path.write_text('{"schema_version": 1, "schema_version": 1}\n', encoding="utf-8")
            with self.assertRaises(checker.ClassificationInputError):
                checker.load_json(path)

    def test_nonfinite_json_is_rejected(self) -> None:
        with tempfile.TemporaryDirectory() as directory:
            path = Path(directory) / "nan.json"
            path.write_text('{"value": NaN}\n', encoding="utf-8")
            with self.assertRaises(checker.ClassificationInputError):
                checker.load_json(path)

    def test_registry_path_traversal_is_rejected(self) -> None:
        registry = load_seed()
        registry["report"]["path"] = "../REPORT.md"
        with self.assertRaises(checker.ClassificationInputError):
            checker.validate_registry(registry, repo_root=REPO_ROOT)

    def test_report_path_raw_normalization_spellings_are_rejected(self) -> None:
        for path in (
            "docs//performance/REPORT.md",
            "docs/./performance/REPORT.md",
            "docs/performance/../performance/REPORT.md",
        ):
            with self.subTest(path=path):
                registry = load_seed()
                registry["report"]["path"] = path
                with self.assertRaises(checker.ClassificationInputError):
                    checker.validate_registry(registry, repo_root=REPO_ROOT)

    def test_strict_claim_registry_path_must_be_exact(self) -> None:
        registry = load_seed()
        registry["strict_claim_registry"]["path"] = "docs/performance/claim-registry-copy.json"
        with self.assertRaises(checker.ClassificationInputError):
            checker.validate_registry(registry, repo_root=REPO_ROOT)

    def test_strict_claim_registry_path_raw_normalization_spellings_are_rejected(self) -> None:
        for path in (
            "docs//performance/claim-registry-v1.json",
            "docs/./performance/claim-registry-v1.json",
            "docs/performance/../performance/claim-registry-v1.json",
        ):
            with self.subTest(path=path):
                registry = load_seed()
                registry["strict_claim_registry"]["path"] = path
                with self.assertRaises(checker.ClassificationInputError):
                    checker.validate_registry(registry, repo_root=REPO_ROOT)

    def test_report_symlink_rebinding_is_rejected_inside_or_outside_repository(self) -> None:
        registry = load_seed()
        original = (REPO_ROOT / checker.REPORT_RELATIVE_PATH).read_text(encoding="utf-8")
        for target_kind in ("in-repository", "out-of-repository"):
            with self.subTest(target=target_kind):
                with tempfile.TemporaryDirectory() as directory:
                    root = Path(directory)
                    report = root / checker.REPORT_RELATIVE_PATH
                    report.parent.mkdir(parents=True, exist_ok=True)
                    if target_kind == "in-repository":
                        target = report.parent / "report-real.md"
                        target.write_text(original, encoding="utf-8")
                    else:
                        target = REPO_ROOT / checker.REPORT_RELATIVE_PATH
                    report.symlink_to(target)
                    with self.assertRaises(checker.ClassificationInputError):
                        checker.validate_registry(registry, repo_root=root)

    def test_claim_registry_symlink_rebinding_is_rejected_inside_or_outside_repository(self) -> None:
        registry = load_seed()
        report_text = (REPO_ROOT / checker.REPORT_RELATIVE_PATH).read_text(encoding="utf-8")
        claim_text = (REPO_ROOT / checker.CLAIM_REGISTRY_RELATIVE_PATH).read_text(encoding="utf-8")
        for target_kind in ("in-repository", "out-of-repository"):
            with self.subTest(target=target_kind):
                with tempfile.TemporaryDirectory() as directory:
                    root = Path(directory)
                    write_report_fixture(root, report_text)
                    claim = root / checker.CLAIM_REGISTRY_RELATIVE_PATH
                    if target_kind == "in-repository":
                        target = claim.parent / "claim-real.json"
                        target.write_text(claim_text, encoding="utf-8")
                    else:
                        target = REPO_ROOT / checker.CLAIM_REGISTRY_RELATIVE_PATH
                    claim.symlink_to(target)
                    with self.assertRaises(checker.ClassificationInputError):
                        checker.validate_registry(registry, repo_root=root)

    def test_cli_rejects_sidecar_symlink_rebinding_inside_or_outside_repository(self) -> None:
        sidecar_text = REGISTRY_PATH.read_text(encoding="utf-8")
        for target_kind in ("in-repository", "out-of-repository"):
            with self.subTest(target=target_kind):
                with tempfile.TemporaryDirectory() as directory:
                    root = Path(directory)
                    sidecar = root / "docs" / "performance" / "report-claim-classification-v1.json"
                    sidecar.parent.mkdir(parents=True, exist_ok=True)
                    if target_kind == "in-repository":
                        target = sidecar.parent / "sidecar-real.json"
                        target.write_text(sidecar_text, encoding="utf-8")
                    else:
                        target = REGISTRY_PATH
                    sidecar.symlink_to(target)
                    output = io.StringIO()
                    with redirect_stderr(output):
                        status = checker.main(
                            ["--registry", str(sidecar), "--repo-root", str(root)]
                        )
                    self.assertNotEqual(status, 0)
                    self.assertIn("symbolic link", output.getvalue())

    def test_table_after_next_heading_is_rejected(self) -> None:
        registry = load_seed()
        for boundary in (
            "## inserted boundary\n",
            "   ## inserted boundary\n",
            "Setext inserted boundary\n---\n",
        ):
            with self.subTest(boundary=boundary):
                with tempfile.TemporaryDirectory() as directory:
                    root = Path(directory)
                    report = root / checker.REPORT_RELATIVE_PATH
                    report.parent.mkdir(parents=True)
                    original = (REPO_ROOT / checker.REPORT_RELATIVE_PATH).read_text(encoding="utf-8")
                    marker = "\n| Change | Historical/descriptive evidence (not a current claim) | Historical scope / limitation |"
                    moved = original.replace(marker, "\n" + boundary + marker, 1)
                    report.write_text(moved, encoding="utf-8")
                    claim_registry = root / checker.CLAIM_REGISTRY_RELATIVE_PATH
                    claim_registry.parent.mkdir(parents=True, exist_ok=True)
                    claim_registry.write_text(
                        (REPO_ROOT / checker.CLAIM_REGISTRY_RELATIVE_PATH).read_text(encoding="utf-8"),
                        encoding="utf-8",
                    )
                    with self.assertRaises(checker.ClassificationInputError):
                        checker.validate_registry(registry, repo_root=root)

    def test_first_table_requires_exact_preamble_and_header_line(self) -> None:
        registry = load_seed()
        original = (REPO_ROOT / checker.REPORT_RELATIVE_PATH).read_text(encoding="utf-8")
        marker = "\n| Change | Historical/descriptive evidence (not a current claim) | Historical scope / limitation |"
        for insertion, name in (
            ("\n<h2>Inserted heading</h2>\n", "html heading"),
            ("\n<div>Inserted prose</div>\n", "html div"),
            ("\nInserted prose\n", "prose"),
        ):
            with self.subTest(insertion=name):
                with tempfile.TemporaryDirectory() as directory:
                    root = Path(directory)
                    write_report_fixture(root, original.replace(marker, insertion + marker, 1))
                    write_claim_fixture(root)
                    with self.assertRaises(checker.ClassificationInputError):
                        checker.validate_registry(registry, repo_root=root)

    def test_wrapping_heading_and_table_in_common_html_blocks_is_rejected(self) -> None:
        registry = load_seed()
        original = (REPO_ROOT / checker.REPORT_RELATIVE_PATH).read_text(encoding="utf-8")
        heading = "### Historical stable-tranche table (descriptive; not current claims)"
        section_start = original.index(heading)
        section_end = original.index("\n\nRaw evidence:", section_start)
        section = original[section_start:section_end]
        for opening, closing, name in (
            ("<div>", "</div>", "div"),
            ("<details>", "</details>", "details"),
            ("<table>", "</table>", "table"),
            ("<div>\n<details>", "</details>\n</div>", "nested div/details"),
        ):
            with self.subTest(block=name):
                with tempfile.TemporaryDirectory() as directory:
                    root = Path(directory)
                    wrapped = (
                        original[:section_start]
                        + opening
                        + "\n"
                        + section
                        + "\n"
                        + closing
                        + original[section_end:]
                    )
                    write_report_fixture(root, wrapped)
                    write_claim_fixture(root)
                    with self.assertRaises(checker.ClassificationInputError):
                        checker.validate_registry(registry, repo_root=root)

    def test_html_comment_and_raw_tag_ordering_is_fail_closed_for_both_tables(self) -> None:
        registry = load_seed()
        original = (REPO_ROOT / checker.REPORT_RELATIVE_PATH).read_text(encoding="utf-8")
        sections = (
            (
                "### Historical stable-tranche table (descriptive; not current claims)",
                "| Change | Historical/descriptive evidence (not a current claim) | Historical scope / limitation |",
            ),
            (
                "## Historical accepted results (descriptive; not current claims)",
                "| Workload group | Historical before | Historical after | Historical result (not a current claim) | Historical memory result |",
            ),
        )
        variants = (
            ("<div><!-- comment -->", "</div>", "same-line comment after opening tag"),
            ("<div><!--\ncomment\n-->", "</div>", "multiline comment after opening tag"),
            ("<div><!--\n", "</div>\n-->", "closing tag inside comment"),
            ("<!-- comment --><div>", "</div>", "comment then tag on same line"),
            (
                "<div>\n<!-- <details> -->\n<details>",
                "</details>\n</div>",
                "nested tag and comment",
            ),
        )
        original_lines = original.splitlines(keepends=True)
        offsets = []
        for heading, header in sections:
            start_line = next(index for index, line in enumerate(original_lines) if line.rstrip("\n") == heading)
            header_line = next(
                index
                for index, line in enumerate(original_lines[start_line:], start=start_line)
                if line.rstrip("\n") == header
            )
            end_line = header_line
            while end_line < len(original_lines) and original_lines[end_line].startswith("|"):
                end_line += 1
            start_offset = sum(len(line) for line in original_lines[:start_line])
            end_offset = sum(len(line) for line in original_lines[:end_line])
            offsets.append((start_offset, end_offset))
        for section_index, (start_offset, end_offset) in enumerate(offsets):
            for opening, closing, name in variants:
                with self.subTest(section=section_index + 1, variant=name):
                    with tempfile.TemporaryDirectory() as directory:
                        root = Path(directory)
                        wrapped = (
                            original[:start_offset]
                            + opening
                            + "\n"
                            + original[start_offset:end_offset]
                            + "\n"
                            + closing
                            + original[end_offset:]
                        )
                        write_report_fixture(root, wrapped)
                        write_claim_fixture(root)
                        with self.assertRaises(checker.ClassificationInputError):
                            checker.validate_registry(registry, repo_root=root)

    def test_fenced_table_is_ignored(self) -> None:
        registry = load_seed()
        original = (REPO_ROOT / checker.REPORT_RELATIVE_PATH).read_text(encoding="utf-8")
        start_marker = "| Change | Historical/descriptive evidence (not a current claim) | Historical scope / limitation |"
        start = original.index(start_marker)
        end = original.index("\n\nRaw evidence:", start)
        table = original[start:end]
        for opening, closing in (("```markdown", "```"), ("~~~markdown", "~~~")):
            with self.subTest(fence=opening[0]):
                with tempfile.TemporaryDirectory() as directory:
                    root = Path(directory)
                    report = root / checker.REPORT_RELATIVE_PATH
                    report.parent.mkdir(parents=True)
                    fenced = original[:start] + opening + "\n" + table + "\n" + closing + original[end:]
                    report.write_text(fenced, encoding="utf-8")
                    claim_registry = root / checker.CLAIM_REGISTRY_RELATIVE_PATH
                    claim_registry.parent.mkdir(parents=True, exist_ok=True)
                    claim_registry.write_text(
                        (REPO_ROOT / checker.CLAIM_REGISTRY_RELATIVE_PATH).read_text(encoding="utf-8"),
                        encoding="utf-8",
                    )
                    with self.assertRaises(checker.ClassificationInputError):
                        checker.validate_registry(registry, repo_root=root)

    def test_html_and_raw_html_tables_are_ignored(self) -> None:
        registry = load_seed()
        original = (REPO_ROOT / checker.REPORT_RELATIVE_PATH).read_text(encoding="utf-8")
        start_marker = "| Change | Historical/descriptive evidence (not a current claim) | Historical scope / limitation |"
        start = original.index(start_marker)
        end = original.index("\n\nRaw evidence:", start)
        table = original[start:end]
        wrappers = (
            ("<!--", "-->", "comment"),
            ("<pre>", "</pre>", "pre"),
            ("<script>", "</script>", "script"),
            ("<style>", "</style>", "style"),
        )
        for opening, closing, name in wrappers:
            with self.subTest(block=name):
                with tempfile.TemporaryDirectory() as directory:
                    root = Path(directory)
                    report = root / checker.REPORT_RELATIVE_PATH
                    report.parent.mkdir(parents=True)
                    wrapped = original[:start] + opening + "\n" + table + "\n" + closing + original[end:]
                    report.write_text(wrapped, encoding="utf-8")
                    claim_registry = root / checker.CLAIM_REGISTRY_RELATIVE_PATH
                    claim_registry.parent.mkdir(parents=True, exist_ok=True)
                    claim_registry.write_text(
                        (REPO_ROOT / checker.CLAIM_REGISTRY_RELATIVE_PATH).read_text(encoding="utf-8"),
                        encoding="utf-8",
                    )
                    with self.assertRaises(checker.ClassificationInputError):
                        checker.validate_registry(registry, repo_root=root)

    def test_fake_minimal_claim_registry_is_rejected(self) -> None:
        registry = load_seed()
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            report = root / checker.REPORT_RELATIVE_PATH
            report.parent.mkdir(parents=True)
            report.write_text(
                (REPO_ROOT / checker.REPORT_RELATIVE_PATH).read_text(encoding="utf-8"),
                encoding="utf-8",
            )
            claim_registry = root / checker.CLAIM_REGISTRY_RELATIVE_PATH
            claim_registry.parent.mkdir(parents=True, exist_ok=True)
            claim_registry.write_text(
                "{\n"
                '  "schema_version": 1,\n'
                '  "registry_kind": "litchi-performance-claim-registry",\n'
                '  "claims": [{"id": "claim-0251-xlsx-xml-borrowed", "status": "landed", "code_state": "landed"}]\n'
                "}\n",
                encoding="utf-8",
            )
            with self.assertRaises(checker.ClassificationInputError):
                checker.validate_registry(registry, repo_root=root)

    def test_claim_id_on_non_strict_row_is_rejected(self) -> None:
        registry = load_seed()
        registry["report"]["tables"][0]["rows"][0]["claim_id"] = "claim-0251-xlsx-xml-borrowed"
        with self.assertRaises(checker.ClassificationPolicyError):
            checker.validate_registry(registry, repo_root=REPO_ROOT)

    def test_strict_row_requires_bijective_linkage(self) -> None:
        registry = load_seed()
        registry["report"]["tables"][0]["rows"][0]["state"] = "strict_claim"
        with self.assertRaises(checker.ClassificationPolicyError):
            checker.validate_registry(registry, repo_root=REPO_ROOT)

    def test_strict_row_with_nonlanded_claim_is_rejected(self) -> None:
        registry = load_seed()
        registry["report"]["tables"][0]["rows"][0]["state"] = "strict_claim"
        registry["report"]["tables"][0]["rows"][0]["claim_id"] = "claim-0248-cfb-streaming"
        registry["strict_claim_registry"]["links"] = [
            {
                "claim_id": "claim-0248-cfb-streaming",
                "table_id": "historical-stable-tranche",
                "ordinal": 1,
                "row_sha256": registry["report"]["tables"][0]["rows"][0]["row_sha256"],
            }
        ]
        with self.assertRaises(checker.ClassificationPolicyError):
            checker.validate_registry(registry, repo_root=REPO_ROOT)

    def test_claim_linkage_accepts_explicit_landed_link_shape(self) -> None:
        # This is a synthetic future link only; the checked-in registry has no
        # strict rows because the audit found no REPORT-row bijection.
        registry = load_seed()
        registry["report"]["tables"][0]["rows"][0]["state"] = "strict_claim"
        registry["report"]["tables"][0]["rows"][0]["claim_id"] = "claim-0251-xlsx-xml-borrowed"
        registry["strict_claim_registry"]["links"] = [
            {
                "claim_id": "claim-0251-xlsx-xml-borrowed",
                "table_id": "historical-stable-tranche",
                "ordinal": 1,
                "row_sha256": registry["report"]["tables"][0]["rows"][0]["row_sha256"],
            }
        ]
        summary = checker.validate_registry(registry, repo_root=REPO_ROOT)
        self.assertEqual(summary["state_counts"]["strict_claim"], 1)

    def test_canonical_digest_is_stable(self) -> None:
        registry = load_seed()
        self.assertEqual(
            checker.sha256_bytes(checker.canonical_bytes(registry)),
            checker.sha256_bytes(checker.canonical_bytes(copy.deepcopy(registry))),
        )


if __name__ == "__main__":
    unittest.main()
