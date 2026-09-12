"""Generate the auditable DOCX publication allocation probe.

The generated Rust remains a close copy of the current public benchmark.  A
source hash and exact anchor counts make accidental drift or a broad rewrite
fail before any generated file is replaced.
"""

from __future__ import annotations

import difflib
import hashlib
import json
from pathlib import Path


HERE = Path(__file__).resolve().parent
REPO = HERE.parents[4]
SOURCE = REPO / "crates/litchi-docx/examples/managed_paragraph_batch_perf.rs"
EXPECTED_SOURCE_SHA256 = (
    "89ffdd564f6a76814fcc8b769f324557bd15a5201cb366ce9e75ade0a1c61181"
)
GENERATED = HERE / "src/main.rs"
DIFF = HERE / "source-diff.patch"
BINDING = HERE / "source-binding.json"

INSTRUMENTATION = """extern crate self as litchi_perf_baseline;

#[allow(
    dead_code,
    reason = "The shared observer includes harness-only identity helpers."
)]
#[path = "../../../../../../tools/perf-baseline/src/allocation_metrics.rs"]
pub mod allocation_metrics;
#[cfg(feature = "allocator-metrics")]
#[path = "../../../../../../tools/perf-baseline/src/bin/support/counting_allocator.rs"]
mod counting_allocator;

"""

ALLOCATION_RECORDS = """#[derive(Serialize)]
struct AllocationSampleRecord<'a> {
    tag: &'static str,
    scope: &'static str,
    case: String,
    repeat: usize,
    ordinal: usize,
    warmup: bool,
    #[serde(rename = "allocationSample")]
    allocation_sample: &'a allocation_metrics::Sample,
}

fn allocation_case(config: &Config) -> String {
    format!(
        "p{}-k{}-{}-{}",
        config.paragraphs,
        config.replacements,
        config.source.name(),
        config.mode.name()
    )
}

fn emit_allocation_sample(config: &Config, sample: &Sample) -> AnyResult<()> {
    let record = AllocationSampleRecord {
        tag: "allocationSample",
        scope: "publish_document_commit_to_stream_method_only_before_returned_snapshot_drop",
        case: allocation_case(config),
        repeat: sample.repeat,
        ordinal: sample.ordinal,
        warmup: sample.warmup,
        allocation_sample: &sample.allocation_sample,
    };
    let stdout = io::stdout();
    let mut output = stdout.lock();
    serde_json::to_writer(&mut output, &record)?;
    output.write_all(b"\\n")?;
    output.flush()?;
    Ok(())
}

"""

RUN_AND_EMIT = """fn run_sample_and_emit(
    config: &Config,
    fixture: &Fixture,
    preflight: &Preflight,
    fixture_path: &Path,
    repeat: usize,
    ordinal: usize,
    warmup: bool,
) -> AnyResult<Sample> {
    let sample = run_sample(
        config,
        fixture,
        preflight,
        fixture_path,
        repeat,
        ordinal,
        warmup,
    )?;
    emit_allocation_sample(config, &sample)?;
    Ok(sample)
}

"""


def sha256(data: bytes) -> str:
    return hashlib.sha256(data).hexdigest()


def require_count(source: str, anchor: str, label: str, expected: int = 1) -> None:
    count = source.count(anchor)
    if count != expected:
        raise SystemExit(
            f"{label} anchor count is {count}, expected exactly {expected}"
        )


def write_text(path: Path, text: str) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(text, encoding="utf-8", newline="")


def transform(source: str) -> str:
    anchors = {
        "imports": "use std::error::Error;\n",
        "serde import": "use sha2::{Digest, Sha256};\n",
        "sample field": "    publish_ns: u64,\n",
        "main": "fn main() -> AnyResult<()> {\n",
        "run sample calls": "            rows.push(run_sample(\n",
        "run sample function": "fn run_sample(\n",
        "publication": (
            "    let phase_started = Instant::now();\n"
            "    let published = package.publish_document_commit_to_stream(&mut output, &commit)?;\n"
            "    drop(published);\n"
        ),
        "sample construction": "        publish_ns,\n",
        "elapsed": "    let elapsed_ns = elapsed_ns(operation_started)?;\n",
    }
    for label, anchor in anchors.items():
        require_count(source, anchor, label, 2 if label == "run sample calls" else 1)

    transformed = source.replace(anchors["imports"], INSTRUMENTATION + anchors["imports"], 1)
    transformed = transformed.replace(
        anchors["serde import"],
        anchors["serde import"] + "use serde::Serialize;\n",
        1,
    )
    transformed = transformed.replace(
        anchors["sample field"],
        anchors["sample field"] + "    allocation_sample: allocation_metrics::Sample,\n",
        1,
    )
    transformed = transformed.replace(
        anchors["main"],
        anchors["main"]
        + "    #[cfg(feature = \"allocator-metrics\")]\n"
        + "    allocation_metrics::enable();\n",
        1,
    )
    transformed = transformed.replace(
        anchors["run sample calls"],
        "            rows.push(run_sample_and_emit(\n",
    )
    transformed = transformed.replace(
        anchors["publication"],
        "    let phase_started = Instant::now();\n"
        "    let allocation_region = allocation_metrics::begin();\n"
        "    let published = package.publish_document_commit_to_stream(&mut output, &commit)?;\n"
        "    let allocation_sample = allocation_region\n"
        "        .finish()\n"
        "        .unwrap_or_else(allocation_metrics::unavailable_sample);\n"
        "    drop(published);\n",
        1,
    )
    transformed = transformed.replace(
        anchors["sample construction"],
        anchors["sample construction"] + "        allocation_sample,\n",
        1,
    )
    transformed = transformed.replace(
        anchors["run sample function"],
        ALLOCATION_RECORDS + RUN_AND_EMIT + anchors["run sample function"],
        1,
    )
    return transformed


def main() -> None:
    source_bytes = SOURCE.read_bytes()
    source_sha256 = sha256(source_bytes)
    if source_sha256 != EXPECTED_SOURCE_SHA256:
        raise SystemExit(
            f"source SHA-256 changed: {source_sha256}, expected {EXPECTED_SOURCE_SHA256}"
        )
    source = source_bytes.decode("utf-8")
    generated = transform(source)

    write_text(GENERATED, generated)
    diff = "".join(
        difflib.unified_diff(
            source.splitlines(keepends=True),
            generated.splitlines(keepends=True),
            fromfile="crates/litchi-docx/examples/managed_paragraph_batch_perf.rs",
            tofile="docs/performance/results/change-0518/allocator-probe/src/main.rs",
        )
    )
    write_text(DIFF, diff)
    binding = {
        "source": "crates/litchi-docx/examples/managed_paragraph_batch_perf.rs",
        "source_sha256": source_sha256,
        "source_lines": len(source.splitlines()),
        "generated": "docs/performance/results/change-0518/allocator-probe/src/main.rs",
        "generated_sha256": sha256(generated.encode("utf-8")),
        "generated_lines": len(generated.splitlines()),
        "diff": "docs/performance/results/change-0518/allocator-probe/source-diff.patch",
        "canonical_modules": {
            "allocation_metrics": "tools/perf-baseline/src/allocation_metrics.rs",
            "counting_allocator": "tools/perf-baseline/src/bin/support/counting_allocator.rs",
        },
        "anchors": {
            "publication_region_start": "allocation_metrics::begin() immediately before publish_document_commit_to_stream",
            "publication_region_finish": "allocation_region.finish() immediately after publish returns and before drop(published)",
            "json_emission": "run_sample_and_emit after run_sample lifecycle and semantic oracles",
            "csv": "source CSV writer copied unchanged",
        },
    }
    write_text(BINDING, json.dumps(binding, indent=2) + "\n")
    print(json.dumps(binding, indent=2))


if __name__ == "__main__":
    main()
