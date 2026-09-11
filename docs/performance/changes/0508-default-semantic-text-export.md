# 0508: default semantic text-export coverage

Four existing semantic text-export selectors now run in the checked default
baseline: plain RTF, ODT, ODS and ODP. The default grows from 37 cases/201 rows/
31 corpora to **41 cases/213 rows/43 corpora**, retaining every old result key
and every old corpus metadata entry exactly. No timed runner, corpus builder,
format implementation, public API or dependency changes. Two existing OPC
oracle conditions use the equivalent `is_some()` spelling to satisfy Clippy.

The CRUD index still has 15 representative categories and 33 mapped selectors.
It now binds 15 measured selectors to **60 case/corpus rows**, up from 11/48;
18 selectors remain correctness-only, down from 22. The conversion category's
XML-compaction references were incorrect for plain text output and now point
to explicit scope and measured-claim requirements. This promotes a timing
contract, not full format, XML conversion or native-producer certification.

## Evidence and timing scope

The owned temporary target uses repository-pinned Rust 1.95.0, release mode,
no debug information, no incremental build and two build jobs. The source
manifest binds 7,195 Rust/TOML/lock files; build and measurement receipts verify
no source change. The release binary SHA-256 is
`ab874929448dfa3a4c17fd9072e702081b4e78b84cae39f3be120518a0c7f6f2`.
This is a shared AMD EPYC 9R45 host without an isolated-core claim. Both full
measurement children ran serially after compilation and before tests.

One zero-warmup/one-sample preflight supplies identity only. Two full runs each
use three warmups and fifteen measured samples per row: **6,390 full-run
samples**, including **360 new export samples**, plus 213 preflight samples.
All raw rows, sorted samples and sample-order permutations are retained.
The full runs took 48.16 and 47.68 seconds respectively, including setup and
post-clock oracles. These times are descriptive, not a comparison against
historical builds.

RTF opens before each sample and copies output into a bounded pre-reserved
retained vector; exact bytes and byte/object counts are checked after timing.
ODF opens and validates before its sample loop and writes into a bounded
hashing discard sink; digest updates are timed, final digest/progress checks
are outside timing. ODF reports zero retained output bytes. These different
sink scopes cannot support a pure cross-format encoder ranking. Neither
measures open-to-export lifecycle or whole-process memory efficiency.

The large-shape medians are descriptive baselines:

| Existing selector | R1 p50 (ms) | R2 p50 (ms) | Semantic output bytes |
|---|---:|---:|---:|
| `rtf_semantic_text_to_sink` | 0.953184 | 0.939894 | 499,999 |
| `odt_semantic_text_to_sink` | 1.885247 | 1.947448 | 499,999 |
| `ods_semantic_text_to_sink` | 1.847017 | 1.830637 | 1,802,239 |
| `odp_semantic_text_to_sink` | 0.240561 | 0.233021 | 8,998 |

ODT-large issues 19,999 sink writes and ODS-large 65,535, with largest writes
of 49 and 54 bytes respectively. These are observed output boundaries; a
future write-coalescing or traversal change needs its own profile and matched
experiment. This batch makes no speedup, allocation reduction or regression
comparison claim. Tail samples remain available in the raw reports; fifteen
samples make the nearest-rank p95 and p99 equal to the maximum.

## Identity, provenance and validation

Rust and Python catalog derivations agree exactly. New metadata records the
four deterministic generator contracts. An explicit generator discriminator
keeps existing ODP append metadata unchanged. The checked-in RTF watermark
fixture remains conservatively unmapped; the default uses only plain RTF.
Unknown security/producer properties remain unknown.

- Result-key digest: `63cdfaaf094b744baf9a6bb770c676c7752671a3d08a711fa0e2c98538efc8cd`.
- Catalog digest: `f03c9f56846f3c7a22d012189e40126dd65be50efbfd24275c3f9ab40854e41a`.
- Content-set digest: `8e629fef89f7eccf023ebc6a6e8b8aebdbcf91207f6d6dfedd2e2fe59c3452a5`.

The report/catalog validator and CRUD timing validator accept both full runs.
The replay verifier checks all 639 raw row statistics, the old identities,
source and binary bindings, unchanged sink controls and the checked policy.
Missing-export-row and short-export-sample negative probes are both rejected.
The initial Python run exposed two tests selecting the newly promoted category
as a generated-per-run fixture, plus a stale pre-existing allocator guard that
still expected the wrapper inline after its support-module extraction. The
fixtures now select an actual remaining generated-per-run row, and the guard
checks the explicit support module while still excluding it from the normal
entry and safe shared library. The initial failed log is retained. Strict Clippy also found two pre-existing
`!is_none()` expressions in OPC oracles; their equivalent `is_some()` spelling
is included in the final source. The entire build, capture and Rust gate set
is repeated after that cleanup; `initial-capture/` retains superseded evidence.
A repeat-promotion guard failure is also retained; promotion now derives the
index and policy from the frozen original inputs so rerunning it is repeatable.

The final gate set passes: 483 Rust harness tests (one opt-in real-producer
security test remains ignored), 199 Python tests, full harness formatting,
warnings-denied Clippy and rustdoc, doctests, allocator-feature all-target
checking, boundaries, strict claim-registry validation and report claim
classification. Production format implementations are unchanged; this is not
a new full-workspace or native-producer certification.

Final cleanup removed 3,320 owned temporary files and 1,348,038,656
unique-inode allocated bytes after checking for live process references.
Read-only report verification remains available without the removed binaries.

See [the retained evidence and replay instructions](../results/change-0508/README.md),
[the source review](../results/change-0508/source-review.md), and
[the ADR dispositions](../results/change-0508/adr-review.md). The historical
ODG priority question is separately assessed in
[the comparison review](../results/change-0508/odg-priority-review.md).
The full non-iWork goal remains active: provider/native-producer coverage,
remaining correctness-only scenarios, and the 0499/0500 follow-ups remain.
