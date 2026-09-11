# Theme-family host hardening

This evidence records the Theme-family source through commit `4814a84d7`.
Gates were run on base `f4f29c9306d182072aabe8eca08145470fdcfadd` with the
prefix correction applied; recorded source hashes identify the tested bytes.
An unrelated OPC import-order change was present during validation and is
recorded in the input manifest, but is not included in this batch.
It follows the original XLSB integration and MCE receipts; those earlier
measurements remain evidence for their recorded source revisions only.

The complete Theme scanner now validates XML names, references, characters,
declarations, and reserved namespace bindings throughout the source. Direct
Theme, extension-list, and recognized-extension containers enforce element-only
content. Direct extension lists reject foreign element children; MCE envelopes
remain subject to conservative mutation refusal. Unknown extension payloads
remain opaque, including legal comments and CDATA containing `<?`.

Standalone family fragments may start with a UTF-8 BOM. Embedding strips that
leading BOM; embedded XML declarations and processing instructions remain
refused. Inherited namespace bindings make a projected Family source usable as
standalone XML. The complete Snapshot bytes and family range retain the original
fragment spelling. Selective edits preserve the current opaque payload, and
XLSB patches include Family in semantic state and restore exact source on inverse.

Removal closes whitespace-only selected extension/list containers while retaining
unrelated metadata, comments, and payload. Hidden MCE ownership blocks selective
mutation. Both normative and observed native URI profiles are recognized using
XML token normalization while existing URI spelling remains in source bytes.

Namespace scope totals are maintained incrementally. Only admitted candidates
capture inherited scope, and duplicate supported families are refused before a
second capture. The documented active scope limit is 256 user declarations,
including shadowed entries. Resolver lookup remains bounded rather than claimed
linear; quick-xml may transiently inspect one further event's declarations.
Limited edits preflight both complete output and the standalone family
intermediate before constructing replacement buffers. Scalar sizing and source
validation may still allocate bounded scratch space before limit rejection.

These choices follow [ADR 0005](../../../adr/0005-io-memory-and-performance.md)
and [ADR 0006](../../../adr/0006-validation-security-and-compatibility.md).
The scanner is not a complete Theme schema validator, an MCE evaluator, or
evidence of native Office application acceptance. Other host integrations and
durable family patches remain outside this batch.

The fresh compile-first gates passed 1,156 tests including doctests, with 12
ignored tests, followed by strict Clippy, rustdoc, formatting, and diff checks.
No `RUSTFLAGS` warning suppression was used. `gates/receipt.json` records the
commands and raw-log hashes; `gates/source-hashes.json` records their inputs.

Reproduce from the repository root:

```sh
python3 docs/report/spec-gap-validation-evidence/theme-family-hardening/run_checks.py
python3 docs/report/spec-gap-validation-evidence/theme-family-hardening/verify.py
```

The downstream standalone probe covers 17 cases, including native fixture
identity, legal opaque token preservation, explicit namespace errors, recognized
container grammar, MCE refusal, BOM stripping, and caller output limits. Four
normal outputs (native, scalar-edited, removed, and normative-added) pass offline
XSD validation. Its opaque diagnostic output is intentionally excluded from
schema-validity claims. The retained probe uses otherwise valid attributes and
checks the exact prefix-limit error at 4,097 bytes, with a successful 4,096-byte
control.

Independent review approved the bounded correction, recorded in `review.json`.
The [current profile](../theme-family-hardening-performance/report.md) retains
540 samples across nine lanes and three fresh processes. Native shared read
has a 0.137 ms median and replacement 0.415 ms on this host. Allocation/status
checks pass for both namespace-heavy unknown-owner workloads, duplicate refusal,
and one-byte caller-limit refusal. These are absolute measurements, with no
before/after speedup claim. Its verifier rehashes 4,256 build inputs and
recomputes report percentiles from raw samples.
