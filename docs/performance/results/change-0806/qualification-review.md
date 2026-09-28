# 0806 before qualification review

Result: **PASS for the before-only qualification oracle.** I independently
replayed the fresh qualification reports, the sealed 0792 identities, the
valid-4attr preservation fields, and the current probe-quality/amendment
records. This review is complete before candidate application and before any
after source, build, or capture work. It makes no latency, RSS, allocation
improvement, profile, or adoption claim.

## Packet identity and execution separation

The fresh qualification packet contains exactly 18 report/receipt rows: one
`capture`, `commit`, and `lifecycle` row for each of `tiny`, `medium`,
`large`, `vendor`, `unicode-vendor`, and `valid-4attr`. Every row has
`leg=before`, `block=0`, exit code zero, one requested sample, and an empty
stderr log. All rows use the same before-allocation executable
(`14,358,120` bytes,
`4f4b7e1cabccd683cc1c884972c65ba40f1b0aa5a6c1f6afe74a9a6486cb5954`). The
qualification source manifest is `1,182,921` bytes with SHA-256
`fbbead35ddcc01aae68676d9f47b8f7bf1d94108ffa56825979e8340fe25fc0d`, matching
the current `build-before` source manifest. The receipt and completion
identities are retained in `qualification/receipts.json` and
`qualification/complete.json`.

The three before quality receipts all exit zero: `cargo fmt --check`, the
offline locked release all-features test run, and warning-denied all-targets
Clippy. The test log records **36 passed, 0 failed**. The amendment audit is
also passed and reports `counter_logic_and_layout_preserved=true` and
`fixture_logic_and_tests_preserved=true`. Its recorded changes are limited to
test-only allocator registration isolation, the non-allocator fallback
feature gate, seven narrow retained-support `dead_code` allowances, test
module relocation, the explicit `valid-4attr` Serde name, the six-shape
round-trip test, and comments/documentation whitespace.

The report source identity and quality records use the current probe source:

| Probe input | SHA-256 |
| --- | --- |
| `probe-src/src/main.rs` | `3eea5b00b2f62ee26ec2a24b8c3604054b3de60df7b035d0012f28a167c1433f` |
| `probe-src/src/allocation_metrics.rs` | `b34e7a16a0cd29eff40e1bd3be038fbe20f720560243f1978439f93db93b2d67` |
| `probe-src/src/counting_allocator.rs` | `896edb6f6837235b8a17bdf0007d7ec756a63389a3750ec46f2c2681a28fce5e` |
| `probe-src/Cargo.lock` | `3828dd46cddb7b5f3b0e838051b1c8d70140c1deb0ead723cd7ce991e0c5fb07` |

The qualification lane is allocation-only. Every sample reports
`status=measured` with scope `operation_global_system_allocator`; the source
retains the 0794 timing strings and `capture_region_0793` profile owner. The
separate before build records native and profile binaries, but neither binary
is substituted into this allocation qualification and no profile result is
used as a timing or resource result.

## Matrix and semantic checks

All 18 reports have schema
`litchi.pptx.public-workflow-probe-0806.v1`, tool
`public-pptx-probe-0806`, the expected dimensions, exact mode timing scope,
`warmup=0`, and one sample at index zero. For every sample, the source digest
matches the report source, semantic verification is true, reopening succeeds,
expected and actual semantic text are equal, and readback byte count and
SHA-256 equal the output object. Capture rows have `marker_matches=null`;
commit and lifecycle rows have `marker_matches=true`.

The source/output identities per shape are:

| Shape | Source bytes / SHA-256 | Capture output bytes / SHA-256 | Commit and lifecycle output bytes / SHA-256 |
| --- | --- | --- | --- |
| `tiny` | 31,530 / `26b1487517882224f820f9f2384c6428dcc40fa2d12c306d74df2809f4003f49` | 31,530 / `26b1487517882224f820f9f2384c6428dcc40fa2d12c306d74df2809f4003f49` | 31,549 / `10d6819120dc881e89ae4d191a85b550d6bb11efc20418565fac605b79adc07a` |
| `medium` | 40,788 / `50ad2f81099ee29d4768d7080b7fc51ea2b5ca2aadd531031efcab65e8d5409e` | 40,788 / `50ad2f81099ee29d4768d7080b7fc51ea2b5ca2aadd531031efcab65e8d5409e` | 40,806 / `eee54f80d033e9e423625157dff22943f146e5989f814e61adc9fbbbda2ddcf4` |
| `large` | 215,220 / `9c46542b763fc4bef63dfe4336cadd2bfba2b7e7b3f18a376c3924eb5643b3e9` | 215,220 / `9c46542b763fc4bef63dfe4336cadd2bfba2b7e7b3f18a376c3924eb5643b3e9` | 215,240 / `af76af57c33a47c69cf921d41fd1b961a566557c372ca0e546e2e4085157ee15` |
| `vendor` | 42,433 / `e8b519bbbcf4140fe3d57f0cedbeb0f266fa0bbc3d1509d4b392151f323f2be8` | 42,433 / `e8b519bbbcf4140fe3d57f0cedbeb0f266fa0bbc3d1509d4b392151f323f2be8` | 42,450 / `4ceca02399d8e993511be03046ed0970a23ad0d942ba95925c769eed24fdb4b4` |
| `unicode-vendor` | 42,370 / `1b18f9837140109e64aba0f28ab23e69e86e52df71360d9773d5b8fd4aae426a` | 42,370 / `1b18f9837140109e64aba0f28ab23e69e86e52df71360d9773d5b8fd4aae426a` | 42,386 / `e6b516c103604d8f1de75861eb1bd36bbf2c39f055d2a11c75acf3f1b1231ac8` |
| `valid-4attr` | 42,304 / `a15d498e5201902f749b9be08e3d698bf1c1ebc8a73aba467558105c382519f3` | 42,304 / `a15d498e5201902f749b9be08e3d698bf1c1ebc8a73aba467558105c382519f3` | 42,323 / `7af3fecea9679c1f7e713a7b0c5aa4b0e261e5a30b2f1ff6947c637fdc419133` |

The valid row's fixture metadata and every sample's preservation oracle agree
on 12 slide parts, 96 replaced text tags, four root declarations per slide,
and four namespaced attributes per text tag. The exact four qualified
name/value pairs are:

| Prefix/name | Namespace URI | Value |
| --- | --- | --- |
| `lx1:probeOne` | `urn:litchi:perf:0806:extension:one` | `litchi-perf-0806-valid-4attr-one` |
| `lx2:probeTwo` | `urn:litchi:perf:0806:extension:two` | `litchi-perf-0806-valid-4attr-two` |
| `lx3:probeThree` | `urn:litchi:perf:0806:extension:three` | `litchi-perf-0806-valid-4attr-three` |
| `lx4:probeFour` | `urn:litchi:perf:0806:extension:four` | `litchi-perf-0806-valid-4attr-four` |

All three valid rows report `extension_preservation_check=true`,
`extension_text_tags=96`, `extension_attributes_per_text_tag=4`,
`extension_attribute_occurrences=384`, `extension_value_occurrences=384`,
and `extension_namespace_declarations_per_slide=4`. Their URI, qualified
name, and value lists are exactly the lists above. The fixture and source
oracle enforce well-formed XML names, double-quoted values, distinct keys,
root-only declarations, and no local declarations; the report separates the
four declarations from the four text-tag attributes rather than counting
namespace declarations as attributes.

The vendor and unicode-vendor controls each report the unchanged unknown
namespace oracle with 96 occurrences. Ordinary shapes have only null extension
oracle fields. This gives independent controls for the new valid extension path and
the existing unknown-namespace preservation path.

## Sealed 0792 parity

For each of the 15 original shape/mode rows, I compared the fresh report with
the corresponding sealed `change-0792/qualification` report. The complete
`source` object and the sample `output` object match byte identity in all
15/15 rows. Fixture metadata and every pre-existing verification field also
match in all 15/15 rows, including marker behavior, expected and actual semantic
text, readback bytes and hash, and the vendor unknown-namespace check and
occurrence count. The new extension verification fields are present but null
on these original rows. Schema/tool labels and run-specific measurement fields
are not fixture-parity claims. The original fixture bytes and preservation
oracles therefore have no observed drift. Root independently checked this
field-by-field boundary while reviewing the handoff.

## Disposition

The fresh before qualification is sufficient to freeze the strict
four-attribute oracle and proceed to the root-owned after source/build gates.
The candidate is still preflight-only. After source identity, after quality,
after captures, resource/profile custody, cross-format veto, cleanup, and the
final seal remain required before any adoption decision.
