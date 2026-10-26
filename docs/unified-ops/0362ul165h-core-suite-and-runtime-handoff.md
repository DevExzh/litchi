# 0362ul165h — core suite and runtime handoff

Root verifies 35 completed checks from the same live C2a full65 gate. The
core suite completes all six chunks, with 332 passing tests reported in the
logs, including 253 library tests and 59 doctests. Every chunk log matches
its producer's size and SHA-256; command order, successful exits and resource
checks pass. The reviewed root verifier validates the new top-level row
alongside the previous 34. C/P remain clean and frozen at
`33a0a63d3553712c1276da592f7846a0c7dcabd6`. Remaining gate checks, census,
corpus and native qualification still require their own evidence.

## ODG v4 reaches the next consumer

Independent review verifies both repaired v3 producer seams in v4. ODG's
public preparation, bounded synthetic build and sealing succeed; ODP's
producer now records its required 4 MiB output cap, and sealing rejects
changed artifacts. Forty-eight source tests and Python AST checks pass.
No Cargo, native build, archive or timing run is claimed.

The ODG manifest nevertheless fails the actual runtime validator. Its fixture
is an external packet copy with the correct bytes, but the runtime requires
the fixed reviewed fixture under the harness checkout. The handoff test stops
at `builder.validate_build_receipt`, so it misses that later consumer.
Independent review reproduces the runtime refusal and confirms that the
reviewed harness fixture passes with the same 738-byte identity/provenance.

V4 is returned for a v5 correction that preserves the fixed fixture binding
and exercises runtime admission after sealing. The future integration recipe
must bind paths to the actual reviewed harness checkout: its root changes
when the tooling moves from the leaf into protocol P. An external copy cannot
silently replace a fixed reviewed input. All prior packets remain immutable.
The existing ODF leaf is retained for this repair and subsequent compilation
and integration; it is not merged on the basis of synthetic tests alone.

## Evidence

Paths are relative to `/home/zhuhe/code/litchi-worktrees/.coord/`.

| Artifact | Bytes | SHA-256 |
| --- | ---: | --- |
| `final-gate-C2a-root-verifier-20261026-v2/partial-3-core-complete/verification.json` | 26,604 | `4e8e1e6ac8a26e3221c00cb9213cabdb7eafab8cf1786a18f7cab346607a496d` |
| `final-gate-C2a-root-verifier-20261026-v2/core-6-chunks-verified.json` | 5,295 | `7fb9bda9cb19cd929c1b27bfdf238336a8369ae28a7e80f40f344a92215c188b` |
| `final-C2a-three-lanes-source-only-20261026-v4-independent-review-20261026/review-receipt.json` | 6,546 | `826bcf03419c47e4faba180bf2f6b1aeaf2a8cded5a70c031260b8d5ee6f9ceb` |

The restored flat-parts and W6 lanes remain in repair. The preserved e4a
baseline at `.coord/odp-0362yd-baseline-clean` contains a 139,425-byte root
lock, SHA-256
`04665b965a919412e29c8772956ddf6267264fa40e3e0ef7c9967fd2eb584171`;
the W6 worker is checking whether its provenance and dependency graph can
support that lane. Finding the file does not admit it or authorize modifying
the existing ODP baseline. All 28 final qualification rows remain open.
Only memo 22 is deferred. Nothing is pushed.
