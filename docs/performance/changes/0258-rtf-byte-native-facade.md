# Change 0258: byte-native unified RTF ingress

## Status

Landed in `a968449f6`, with the opt-in harness introduced in `8a6f0bc0e`
and hardened in `107060085`. The production change is accepted for
correctness and lossless byte transport. No latency or speedup claim is
accepted.

## Production scope

The ordinary `litchi::Document::from_bytes` RTF path now passes the owned
input directly to `litchi_rtf::RtfDocument::parse_bytes`. It no longer
requires `String::from_utf8` before native parsing. This preserves literal
CP-1252 input and the native RTF compressed transports instead of rejecting
them at the facade boundary.

The byte, reader, and smart detectors recognize native `LZFu` and stored
`MELA` framing when the `rtf` feature is enabled. The reader probe was widened
from eight to twelve bytes and restores the caller's cursor. ZIP and OLE2
signatures retain precedence, including rtf-only builds, so a container header
with `LZFu` bytes at offset eight does not become an RTF false positive.

The native parser remains the preservation authority. Focused tests cover
plain exact-source bytes, literal CP-1252 text, LZFu and MELA semantic parity
and native exact-source round trip, truncated compressed-frame refusal, and
ZIP/OLE2 precedence. The facade still maps native RTF parse failures into its
established `ParseError` category; this change does not broaden edit or
publication capabilities.

## Harness scope

Two opt-in selectors were added without changing the 36-case default matrix:

- `rtf_file_open` times `litchi::Document::from_bytes(owned_bytes)`.
- `rtf_file_open_lifecycle` additionally times facade full-text projection.

Input cloning, source hashes, exact native source round trip, and semantic
parity are outside the timed interval. The native seam and facade both use the
same `litchi_rtf` parser/model, so this is adapter-parity evidence rather than
an independent parser oracle. CP-1252, LZFu, and MELA runs are correctness
evidence only because the historical control cannot open those inputs through
the facade. The controlled comparison therefore uses generated plain RTF.

## Controlled evidence

Both captures reuse these clean, distinct release binaries:

- control revision `9426734501c29d4b41b86ac66683efa1ca3e290d`,
  49,568,584 bytes, SHA-256
  `bdba203782d4b7ca4d0e58830e0cb41a60b8393a761366e962e3cc840236153b`;
- candidate revision `107060085595da22c01572ecbc6e04c172ca5cd1`,
  49,564,504 bytes, SHA-256
  `aa7e29e407ec66925b8a2ac7be6620862bdb4687a198498b2fd8443d45922cee`.

Each A1/B1/B2/A2 capture is pinned to CPU 2 with one worker, warm filesystem
state, 20 warmups, and 500 retained samples for each selector over tiny,
medium, and large plain generated RTF. Configuration, corpus/source identity,
binary identity, clean-worktree identity, and recomputation checks pass.

The first capture recorded the host-default `rustc 1.97.1` in runtime metadata.
It accepted 0 of 24 statistic cells. The capture is retained because its
executable hashes and workload identities are exact, but its runtime toolchain
label does not establish compiler provenance; build logs are not part of the
compact package.

The immediate rerun resolved runtime metadata from the pinned workspace and
records `rustc 1.95.0 (59807616e 2026-04-14)` in all four legs. Its strict
summary accepts five in-run cells:

| Selector / shape | Accepted in pinned rerun | A1 -> B1 reduction | A2 -> B2 reduction |
|---|---|---:|---:|
| `rtf_file_open` / large | p50 | 3.812737% | 3.444067% |
| `rtf_file_open` / large | mean | 3.639829% | 4.368652% |
| `rtf_file_open` / large | p95 | 2.194055% | 12.010203% |
| `rtf_file_open` / large | p99 | 1.392353% | 12.278874% |
| `rtf_file_open_lifecycle` / medium | p50 | 0.804682% | 1.280085% |

The other 19 pinned-run cells are rejected, including 14 adverse-both cells.
More importantly, none of these five cells was accepted in the immediately
preceding matched capture. The executable bytes, revisions, corpus/source
identities, selectors, CPU affinity, worker count, warmups, and retained sample
counts match between captures. Because the accepted set does not reproduce,
all latency results are withheld and no claim-registry entry is added.

## Retained artifacts

Both compact packages are retained under
[`results/rtf-unified-ingress-0258-evidence/`](../results/rtf-unified-ingress-0258-evidence/):

- `initial/summary.json` SHA-256
  `cafc7c0bfede05818fe73319d2e4815e64baceffa98121b793f2bcc42872350c`;
  manifest SHA-256
  `1e236c4eee4e9c88ff2620b43741092203ac5e65a2bb45dd3b6c4469a7743f5b`;
- `pinned/summary.json` SHA-256
  `f53feda283865975999c6437fcaef13ae8ae1d208cf1e2cb39ce22d8e964c64a`;
  manifest SHA-256
  `db44cac304762cef91b55064c3c5a5d27fc5d25dde75ac5f48076001fda8841a`.

The packages contain the strict summary, four deterministic zstd-compressed
raw reports, and their manifest. They do not contain the binaries; binary
sizes, modes, paths, revisions, and SHA-256 identities are embedded in each
summary.

## Verification

- `cargo test -p litchi --no-default-features --features rtf --lib`: 36/36
  passed.
- Focused compressed RTF plus DOCX tests and rtf-only, docx-only, and combined
  feature checks passed.
- The focused selector, 380-case registry, complete name/parse registry,
  help output, CP-1252/LZFu smokes, unsupported-shape refusal, and harness
  formatting checks passed.
- The strict summary and package tools reproduced both packages byte-for-byte;
  their 46 focused unit tests passed.
- Rustfmt, RTF-only Clippy, the crate-boundary gate, and performance/policy
  Python gates passed. Combined RTF+DOCX Clippy remains blocked by six
  pre-existing `litchi-docx` lint findings. A broad harness suite was stopped
  at an unrelated existing ODP assertion while concurrent ODF files were
  dirty; no RTF failure was observed.

## Claim boundary

This change establishes byte-native facade ingress and narrow adapter/source
preservation correctness. It makes no latency, throughput, allocation, RSS,
copy-byte, physical-I/O, decompression, cold-cache, compressed-input,
real-producer, rich-RTF, edit/save, or end-to-end format claim. A future
performance claim requires a reproducible controlled capture and independent
resource evidence where applicable.
