# 0806 public PPTX probe source review

The 0806 probe is frozen under `probe-src/` as a source-only extension of the
qualified 0794 public PPTX harness. It keeps the three public operation modes,
the allocator target, the capture-profile feature, the `capture_region_0793`
owner symbol, and the 0794 timing scopes. No generated `Cargo.toml` is stored;
the root build driver materializes it from `Cargo.toml.template`.

The original fifteen rows remain unchanged: `tiny`, `medium`, `large`,
`vendor`, and `unicode-vendor`, each in `capture`, `commit`, and `lifecycle`.
The only new shape is `valid-4attr`, with the same 12-slide by 8-text-box
dimensions as `vendor` and `unicode-vendor`. The resulting matrix has eighteen
rows. Its report identities are `litchi.pptx.public-workflow-probe-0806.v1`
and `public-pptx-probe-0806`; allocator-enabled and ordinary fallback binary
identities use the 0806 suffix as well.

The ordinary fixture still uses the unchanged source-text generator and the
unchanged 0780 edit marker, `litchi-perf-0780-static-mce-capabilities`. The
vendor and Unicode-vendor generators retain their six prefixes, six local
names, URI construction, root declaration insertion, text-tag insertion, and
legacy attribute values (`litchi-perf-0785-{note}`). The added `value` field in
the internal entry record reconstructs those exact legacy attribute bytes; it
does not change them. The existing unknown-namespace readback oracle remains
the same path for those two shapes.

`valid-4attr` generates exactly these four distinct, double-quoted,
namespaced attributes on every `a:t` start tag:

| Prefix and URI | Local name | Value |
| --- | --- | --- |
| `lx1`, `urn:litchi:perf:0806:extension:one` | `probeOne` | `litchi-perf-0806-valid-4attr-one` |
| `lx2`, `urn:litchi:perf:0806:extension:two` | `probeTwo` | `litchi-perf-0806-valid-4attr-two` |
| `lx3`, `urn:litchi:perf:0806:extension:three` | `probeThree` | `litchi-perf-0806-valid-4attr-three` |
| `lx4`, `urn:litchi:perf:0806:extension:four` | `probeFour` | `litchi-perf-0806-valid-4attr-four` |

The generator validates distinct prefixes, URIs, names, and values, XML name
characters, UTF-8 strings, and double-quoted assignments before rewriting a
slide. It has no malformed-input case. The new readback oracle independently
checks semantic text, exactly four parsed name/value pairs on every `a:t`, all
four declarations on the `<p:sld>` root start tag, one occurrence of each
declaration per slide, exactly four `xmlns:lx…` declarations globally per
slide, and one occurrence of each complete name/value attribute per text tag.
The oracle records total text tags, attributes per text tag, total attribute
and value occurrences, namespace URIs, names, and values in every sample's
verification object. A local declaration or an extra/duplicate text-tag
attribute fails the operation.

The source lineage is frozen by these SHA-256 identities:

| File | SHA-256 |
| --- | --- |
| `probe-src/Cargo.toml.template` | `7fdb50bb6f5103ddca97ec2fee66561526ca08fd757ff248fdcd0543fc46a11f` |
| `probe-src/Cargo.lock` | `3828dd46cddb7b5f3b0e838051b1c8d70140c1deb0ead723cd7ce991e0c5fb07` |
| `probe-src/src/main.rs` | `3eea5b00b2f62ee26ec2a24b8c3604054b3de60df7b035d0012f28a167c1433f` |
| `probe-src/src/allocation_metrics.rs` | `b34e7a16a0cd29eff40e1bd3be038fbe20f720560243f1978439f93db93b2d67` |
| `probe-src/src/counting_allocator.rs` | `896edb6f6837235b8a17bdf0007d7ec756a63389a3750ec46f2c2681a28fce5e` |

The template and lock are byte-identical to 0794. The allocator support files
retain their 0794 runtime implementation and differ in identity strings plus a
test-only cfg on the global allocator. Unit tests call the wrapper directly,
while synthetic counter tests run with the process allocator unwrapped so
test-harness allocations cannot enter their regions; release binaries still
install the wrapper. The unavailable-sample helper is compiled only for the
non-allocator feature, matching its uninstrumented-binary purpose and keeping
the all-features test target warning-free. The copied counter APIs use
per-item, non-test-only `allow(dead_code)` annotations with comments explaining
their retained allocator-test/profile purpose; the counter implementation and
layout remain unchanged. The main module's tests are at file end and its
module documentation has the required list paragraph break. `Shape::Valid4Attr`
has an explicit `valid-4attr` Serde rename, and the end-of-file tests round-trip
all six CLI shape identities through parse, JSON serialization, and JSON
deserialization. `main.rs` is the only harness logic change: the new shape,
its generator, its explicit preservation oracle, the report fields, and the
0806 report identities.

Validation completed for this source-only handoff is `rustfmt --edition 2024`
and its `--check` pass over all three Rust files. Cargo materialization,
compilation, production-source application, native/allocator/profile capture,
staging, and commit remain root-owned.

The deserialization and equality derives needed only by the shape round-trip
test are gated with `cfg_attr(test, ...)`; ordinary binaries retain their
original derive set.
