# DOCX source-backed SVG lifecycle performance evidence

This directory contains the bounded profile harness and measurement contract
for the source-backed DOCX SVG lifecycle introduced by
`892441d95db29da4351390716ef5c65b4c7c97de`. The harness is outside the
production workspace dependency graph and calls only the committed public
`litchi_docx::source_backed` APIs.

The first command to run is the correctness smoke. It builds in an isolated
committed checkout and runs one sample through the selected matrix. It does
not produce a performance claim:

```bash
COMMITTED_ROOT=/tmp/litchi-docx-svg-lifecycle-892441d95 \
  bash docs/report/spec-gap-validation-evidence/docx-svg-lifecycle-performance/run_smoke.sh
```

The committed checkout must be clean and resolve to the baseline commit. The
runner copies the retained source snapshot, this evidence harness, and the two
required native fixtures into that temporary checkout; the production checkout is never used for the
profile build. The isolated Cargo target is disposable and is removed when
the runner created it.

The full profile remains intentionally gated until the scenario contract and
smoke receipts have been reviewed:

```bash
PROFILE_FROZEN=1 \
COMMITTED_ROOT=/tmp/litchi-docx-svg-lifecycle-892441d95 \
  bash docs/report/spec-gap-validation-evidence/docx-svg-lifecycle-performance/run_profile.sh
```

The runner retains raw JSON receipts, `/usr/bin/time -v` RSS receipts, source
manifests, build provenance, and a report recomputed from the raw samples.
Smoke and full outputs use separate `smoke-*` and `full-*` report, provenance,
and receipt names; rerunning one mode removes only its known outputs and keeps
the other mode and unrelated review evidence intact.
`baseline-source/` retains the exact committed source inputs used to build the
profile, so manifest paths remain readable after disposable staging is
removed. `requirements.md` defines the metric boundaries and expected
64-owner typed refusals. No report row is a before/after claim.
