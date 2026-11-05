# 0361h — Required character-style policies

- Contract: [0351 §2.5](0351-api-contract.md#25-configured-limits-items-18-19).
- Base: `61515148e1`.
- Status: accepted after the [0361ha](0361ha-character-style-output-repair.md)
  buffer repair, correctness checks and matched performance review.

The seven raw character-style entry points in `litchi-iwa-text-wire` now
require `WireLimits`: `char_style_runs`, `styles_in_span`, `restyle_span`,
`character_style_variation`, `character_style_properties`,
`character_style_parent` and `stylesheet_with_style`. Variation construction
now returns a result. No implicit-policy wrapper, alias, renamed duplicate
or deprecation remains. Seven compile-fail examples protect the required
argument, alongside the existing glossary suffix bans.

Pages, Numbers and Keynote migrate every caller with this change. Pages and
Numbers derive input/output bounds from the retained archive message policy,
capped by the wire layer's hard ceilings; metadata-header field limits are
not repurposed as body-field limits. Keynote passes its existing workspace
wire profile. Each owner preserves typed resource and allocation failures.
Pages and Numbers expose the common resource kind inside their own family
`LimitKind`; Keynote maps it into its existing structural resource kinds.
Malformed payloads and invalid spans retain their existing classifications.

Each raw operation accumulates input bytes, decoded fields and scan work
across the known nested messages it visits. A nested view receives the
remaining field/work allowance before allocating its span list. Root depth
is zero. Unknown length-delimited fields remain opaque. Style deduplication
also charges its comparison work. Restyling checks work before invoking the
caller callback. The private budget cannot be reused after a failed call;
successful counters stay within the finite hard ceilings of `WireLimits`.

Encoded payloads use fallible reservation after output-byte admission. Font
decoding bounds the owned name before reserving it. Generated variations
admit field, work and depth counts. Object references use a fixed stack
buffer. Restyling reuses the parsed storage root; stylesheet registration
avoids the previous duplicate field-reference vector and writes the extended
children entry directly into its output. Unknown source bytes retain their
original framing when their field is not replaced.

The current correctness checks pass: 57 raw unit tests, eight compile-fail
doctests, 1,015 all-feature Pages/Numbers/Keynote library tests, and all-target,
all-feature Clippy for those owners and the raw crate. Five pre-existing
owner tests remain ignored. Boundary tests cover aggregate bytes, fields,
work and nesting, exact output admission, callback suppression, opaque byte
preservation, malformed-input classification and typed owner propagation.
The first boundary build found two now-fixed redundant test qualifications;
the first Clippy pass found the now-fixed placement of a test module.

The performance probe compares all seven operations, with small and larger
run tables and stylesheets. Its fixture/timed bodies are shared between
endpoints; only the explicit-policy call adapter is conditional. Separate
allocation binaries keep instrumentation out of native timing. Both library
builds use the same lockfile and release profile, and their source manifests
identify every file. Exact fixture and output comparisons precede native
qualification. The runner requires five ABBA blocks, common call counts,
500 ms calibration, at least 150 ms per measured leg, two warmups, and host
checks before, during and after each child. Raw receipts remain external.

The accepted v3 source matches all seven measured Rust files. All 60 native
fixture/output files match the corresponding base bytes. Every allocation
route has equal or fewer calls, requested bytes and peak live bytes than the
base, with zero retained bytes after dropping the result. The v2 memory
regression and its v3 repair remain recorded in
[0361ha](0361ha-character-style-output-repair.md).

All qualified ABBA blocks improve native latency. Ratios below are candidate
time divided by base time, using the median of each route's complete blocks:

| Operation | Two-run fixture | 64-run fixture |
| --- | ---: | ---: |
| Read runs | 0.9317 | 0.9606 |
| Find span styles | 0.8953 | 0.9609 |
| Restyle span | 0.6374 | 0.8562 |
| Register stylesheet style | 0.7903 | 0.9618 |
| Construct variation | — | 0.3521 |
| Read properties | — | 0.9318 |
| Read parent | — | 0.8530 |

The first attempt stopped during a warmup when the guard detected an
unrelated `test-adblock-ubo-header.py` process. Its five already-completed,
clean two-run read blocks are included alongside all 55 blocks from the
complete second attempt. The interrupted warmup is unqualified. The raw
first manifest retains its last `running` status; the traceback and blocked
host snapshots document its termination. Both attempts used identical
protected binaries and the unchanged runner. The complete attempt has 2,263
clean host snapshots; the shortest qualified measured leg across both
attempts is 550.4 ms. Allocation counts come from separate instrumented
processes; these measurements do not make process-RSS claims.

This record does not close other raw API audits, retained footnote policy
propagation, pending wave integrations or the final workspace gate.

The [evidence summary](results/0361h/evidence.json) retains source-manifest
file names, sizes and SHA-256 values, both build commands, identical lock and
Cargo profile identities, protected binary hashes, validation logs and all
allocator attempts. The full source manifests remain outside the checkout.
Regenerate the summary with:

```sh
python3 docs/unified-ops/results/0361h/collect.py \
  --coord (deleted-coord-artifact) --out docs/unified-ops/results/0361h \
  --native (deleted-coord-artifact)
```

The source snapshots contain `git archive 61515148e1 Cargo.toml
rust-toolchain.toml crates`, the protected common `Cargo.lock`, and the seven
changed Rust files listed in the source delta. Regenerate each full file
manifest by hashing every snapshot file in sorted relative-path order with
SHA-256 and recording its byte length. `build.py` records the exact pinned
Cargo/Rust compiler commands for both probe variants; `abba.py` records every
calibration, warmup and measured invocation.
