# 0362ul57b — DrawingML DOCX/PPTX consumer adapter

Date: 2026-10-25. This source-only follow-up composes the DOCX/PPTX
validation-module exports with the existing DrawingML historical consumer
adapter. It does not run Cargo, a native probe, or a timing workload. The
earlier `0362ul57` and `0362ul57a` failure/repair records remain unchanged;
the coordinator must create a fresh preparation receipt after this source
fence is composed.

## Candidate and historical source identities

The checked-in `tools/perf-baseline/src/lib.rs` is the candidate source for the
DOCX/PPTX composition:

```text
candidate SHA-256: 4acd072fd68c094e24524862194da9676260e1fb25f6ea3dbb539197d2838b5b
historical SHA-256: 280d61ae1f714af78be6cecf21a4432a9ea2a27478e89c237b5d358bc02db02a
```

The candidate includes the three actual current API exports:

```text
litchi_docx::validate_read_at
  -> litchi_docx::validation::validate_read_at
litchi_pptx::validate_source
  -> litchi_pptx::validation::validate_source
litchi_pptx::PptxValidationLimits::default()
  -> litchi_pptx::validation::Limits::default()
```

The historical copy is produced only by the role-aware preparer. It applies
the six reviewed substitutions below, in addition to the existing RTF/XLS
compatibility mappings:

```text
litchi_rtf::streaming::Limits::new
  -> litchi_rtf::write::StreamingRtfLimits::new
litchi_rtf::streaming::Writer::new
  -> litchi_rtf::write::StreamingRtfWriter::new
litchi_xls::validation::Limits::default()
  -> litchi_xls::validation::XlsValidationLimits::default()
litchi_docx::validation::validate_read_at
  -> litchi_docx::validate_read_at
litchi_pptx::validation::validate_source
  -> litchi_pptx::validate_source
litchi_pptx::validation::Limits::default()
  -> litchi_pptx::PptxValidationLimits::default()
```

The six substitutions change the historical adapter copy by seven bytes in
total. Applying them to the candidate must reproduce the historical hash
`280d61…` exactly; the candidate is never rewritten. All other adapter paths,
including the DrawingML probe, allocator, fixtures, and timed operation code,
remain byte-identical.

## Source-bound checks and handoff

The preparer and focused unit test bind the checked-in candidate hash, require
one occurrence of each candidate spelling and no pre-adapted spelling, then
prove the exact six-substitution historical bytes and both role hashes. The
focused suite is expected to contain 14 tests:

```sh
PYTHONDONTWRITEBYTECODE=1 python3 -m unittest tools.test_drawingml_raw_qualification
```

No native or performance result is claimed here. After the 14 tests pass, the
coordinator must run the existing preparer with a fresh external output path,
validate its role-specific receipt, and only then build or measure the raw
lanes. No receipt produced for the pre-DOCX/PPTX source may be reused.
