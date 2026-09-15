#!/usr/bin/env python3
"""Drive the change 0626 smoke-baseline pipeline over two real harness reports.

The GitHub Actions run cannot be executed in this repository, so this script
substitutes for the parts of the workflow that are not network access: it plays
the `full` job's descriptor step and the `smoke` job's select, compare and
report steps over two reports captured by the real release allocator harness,
once per scenario, and writes every artifact the workflow would upload.

Usage:

    python3 docs/performance/results/change-0626/verify-pipeline.py \\
      --reference <full-run allocator report> \\
      --current <smoke allocator report> \\
      --work <scratch directory> \\
      --out docs/performance/results/change-0626/pipeline

Scenario directories written under `--out`:

    fallback/      no reference artifact was fetched
    fetched/       a compatible reference is fetched and compared
    regression/    the same, with the current report's allocation counters
                   raised 10%, so the comparator reports a regression
    drift/         the same, with the reference runner's CPU model changed, so
                   the comparator fails closed on build identity
    incompatible/  the reference was captured over a different corpus, so the
                   case/corpus key manifest digest rejects it before comparison

Each directory holds the fetch status, the reference descriptor, the selection,
the comparator's machine and human output, the classification and the rendered
job summary, plus `transcript.txt` with each command's exit status. The
scenario's own copies of the two 117 KB harness reports and of the baseline the
selector chose go to `--work`, because they are byte-for-byte reproducible from
the two retained source reports.
"""

from __future__ import annotations

import argparse
import copy
import io
import json
import sys
from contextlib import redirect_stdout
from pathlib import Path

ROOT = Path(__file__).resolve().parents[4]
sys.path.insert(0, str(ROOT))

from tools import perf_compare  # noqa: E402
from tools import perf_smoke_baseline  # noqa: E402

SMOKE_POLICY = ROOT / "docs" / "performance" / "perf-smoke-baseline-policy-v1.json"
ALLOCATOR_POLICY = (
    ROOT / "docs" / "performance" / "perf-regression-policy-allocator-v1.json"
)
RUNNER_LABELS = "ubuntu-latest"
REFERENCE_RUN_ID = "11111111"
CURRENT_RUN_ID = "22222222"


def write_json(path: Path, document) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(
        json.dumps(document, indent=2, sort_keys=True) + "\n", encoding="utf-8"
    )


def run(transcript: list[str], label: str, argv: list[str]) -> int:
    buffer = io.StringIO()
    with redirect_stdout(buffer):
        status = perf_smoke_baseline.main(argv)
    transcript.append(f"$ perf_smoke_baseline.py {' '.join(argv)}")
    transcript.append(buffer.getvalue().rstrip("\n"))
    transcript.append(f"exit status {status}")
    transcript.append("")
    return status


def compare(transcript: list[str], argv: list[str]) -> int:
    buffer = io.StringIO()
    with redirect_stdout(buffer):
        status = perf_compare.main(argv)
    transcript.append(f"$ perf_compare.py {' '.join(argv)}")
    transcript.append(buffer.getvalue().rstrip("\n"))
    transcript.append(f"exit status {status}")
    transcript.append("")
    return status


def scenario(
    name: str,
    out: Path,
    work: Path,
    *,
    reference: dict | None,
    current: dict,
    fetched: bool,
    reason: str,
) -> dict:
    directory = out / name
    directory.mkdir(parents=True, exist_ok=True)
    workspace = work / name
    workspace.mkdir(parents=True, exist_ok=True)
    transcript: list[str] = [f"# scenario: {name}", ""]
    reference_dir = workspace / "reference"
    reference_dir.mkdir(exist_ok=True)
    current_path = workspace / "current.json"
    write_json(current_path, current)

    if reference is not None:
        report_path = workspace / "full-run-report.json"
        write_json(report_path, reference)
        descriptor_status = run(
            transcript,
            "descriptor",
            [
                "descriptor",
                "--policy",
                str(SMOKE_POLICY),
                "--comparator-policy",
                str(ALLOCATOR_POLICY),
                "--report",
                str(report_path),
                "--runner-labels",
                RUNNER_LABELS,
                "--run-id",
                REFERENCE_RUN_ID,
                "--event",
                "schedule",
                "--out",
                str(reference_dir / "allocator-baseline-descriptor.json"),
            ],
        )
        assert descriptor_status == 0, name
        write_json(reference_dir / "allocator-baseline.json", reference)

    status_path = directory / "fetch-status.json"
    status_argv = ["fetch-status", "--out", str(status_path), "--reason", reason]
    if fetched:
        status_argv.extend(["--fetched", "--run-id", REFERENCE_RUN_ID])
    run(transcript, "fetch-status", status_argv)

    baseline_path = workspace / "baseline.json"
    selection_path = directory / "selection.json"
    run(
        transcript,
        "select",
        [
            "select",
            "--policy",
            str(SMOKE_POLICY),
            "--comparator-policy",
            str(ALLOCATOR_POLICY),
            "--current",
            str(current_path),
            "--reference-dir",
            str(reference_dir),
            "--fetch-status",
            str(status_path),
            "--runner-labels",
            RUNNER_LABELS,
            "--current-run-id",
            CURRENT_RUN_ID,
            "--baseline-out",
            str(baseline_path),
            "--selection-out",
            str(selection_path),
        ],
    )

    comparison_path = directory / "comparison.json"
    comparator_status = compare(
        transcript,
        [
            "--policy",
            str(ALLOCATOR_POLICY),
            "--baseline",
            str(baseline_path),
            "--current",
            str(current_path),
            "--json-out",
            str(comparison_path),
            "--summary-out",
            str(directory / "comparison.txt"),
        ],
    )

    classification_path = directory / "classification.json"
    report_status = run(
        transcript,
        "report",
        [
            "report",
            "--policy",
            str(SMOKE_POLICY),
            "--selection",
            str(selection_path),
            "--comparison",
            str(comparison_path),
            "--comparator-exit-status",
            str(comparator_status),
            "--classification-out",
            str(classification_path),
            "--summary-out",
            str(directory / "outcome.md"),
            "--step-summary",
            str(directory / "step-summary.md"),
        ],
    )

    # The fetched report is a byte copy of the retained source report, so only
    # the descriptor — what the selector actually binds against — is retained.
    descriptor_path = reference_dir / "allocator-baseline-descriptor.json"
    if descriptor_path.exists():
        retained = directory / "reference"
        retained.mkdir(exist_ok=True)
        (retained / descriptor_path.name).write_text(
            descriptor_path.read_text(encoding="utf-8"), encoding="utf-8"
        )
    (directory / "transcript.txt").write_text(
        "\n".join(transcript) + "\n", encoding="utf-8"
    )
    selection = json.loads(selection_path.read_text(encoding="utf-8"))
    classification = json.loads(classification_path.read_text(encoding="utf-8"))
    return {
        "scenario": name,
        "mode": selection["mode"],
        "fallback_reasons": selection["fallback_reasons"],
        "comparator_status": classification["comparator_status"],
        "comparator_exit_status": comparator_status,
        "outcome": classification["outcome"],
        "blocking": classification["blocking"],
        "job_exit_status": report_status,
    }


def scale_allocation(report: dict, factor: float) -> dict:
    scaled = copy.deepcopy(report)
    for result in scaled["results"]:
        allocation = result["operation_metrics"]["allocation"]
        for field, vector in allocation.items():
            if not isinstance(vector, dict) or "values" not in vector:
                continue
            vector["values"] = [round(value * factor) for value in vector["values"]]
    for evidence in scaled.get("filesystem_evidence", []):
        for sample in evidence["samples"]:
            metrics = sample["allocation_metrics"]
            for field, value in metrics.items():
                if isinstance(value, int) and not isinstance(value, bool):
                    metrics[field] = round(value * factor)
    return scaled


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--reference", type=Path, required=True)
    parser.add_argument("--current", type=Path, required=True)
    parser.add_argument("--out", type=Path, required=True)
    parser.add_argument("--work", type=Path, required=True)
    args = parser.parse_args()

    reference = json.loads(args.reference.read_text(encoding="utf-8"))
    current = json.loads(args.current.read_text(encoding="utf-8"))

    drifted = copy.deepcopy(reference)
    drifted["environment"]["cpu_model"] = "Intel(R) Xeon(R) Platinum 8370C CPU"

    other_corpus = copy.deepcopy(reference)
    for result in other_corpus["results"]:
        result["corpus"]["archive_sha256"] = "0" * 64

    summary = [
        scenario(
            "fallback",
            args.out,
            args.work,
            reference=None,
            current=current,
            fetched=False,
            reason="no successful full run has published a reference artifact",
        ),
        scenario(
            "fetched",
            args.out,
            args.work,
            reference=reference,
            current=current,
            fetched=True,
            reason=f"downloaded the baseline artifact of run {REFERENCE_RUN_ID}",
        ),
        scenario(
            "regression",
            args.out,
            args.work,
            reference=reference,
            current=scale_allocation(current, 1.10),
            fetched=True,
            reason=f"downloaded the baseline artifact of run {REFERENCE_RUN_ID}",
        ),
        scenario(
            "drift",
            args.out,
            args.work,
            reference=drifted,
            current=current,
            fetched=True,
            reason=f"downloaded the baseline artifact of run {REFERENCE_RUN_ID}",
        ),
        scenario(
            "incompatible",
            args.out,
            args.work,
            reference=other_corpus,
            current=current,
            fetched=True,
            reason=f"downloaded the baseline artifact of run {REFERENCE_RUN_ID}",
        ),
    ]
    write_json(args.out / "summary.json", summary)
    for item in summary:
        print(
            f"{item['scenario']:<13} mode={item['mode']:<17} "
            f"comparator={item['comparator_status']:<10} "
            f"outcome={item['outcome']:<32} job_exit={item['job_exit_status']}"
        )
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
