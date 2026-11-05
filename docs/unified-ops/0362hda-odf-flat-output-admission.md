# 0362hda — ODF flat output admission follow-up

- Date: 2026-10-22.
- Parent: 0362hd (`90d88df9d8`, based on `8d6abdab69`).
- Status: validated source follow-up; coordinator integration and the final
  workspace inventory remain required.

The first flat raw-limits slice had three correctness gaps. `BoundedBytes`
reserved the difference between its target and capacity even though
`try_reserve_exact` measures additional bytes from the current length. A
large write after a spare-capacity prefix could therefore make the final
`Vec` grow infallibly beyond the sink bound. The reservation now uses the
current length, with a regression that writes 196 bytes after a four-byte
prefix in a 64-byte-capacity, 256-byte-bounded sink and proves the capacity
stays bounded.

The XML scanner now checks the attempted child depth for `Event::Empty` as it
does for `Event::Start`, including the exact next-depth refusal. `Parts`
retains the maximum depth admitted by its split and rechecks the proof in
`to_bytes` and `merge`, so a package split under a loose profile cannot be
published through an unchanged route under a tighter depth profile. The
output paths also recheck the retained flat source/content size admission.

`Parts` now retains a SHA-256 content proof of its source rather than another
full source buffer. `merge` verifies the digest after the cheap length check
and returns the typed `Error::PatchConflict` for a foreign source, including
a same-length source whose bytes differ. This closes the old range-reuse
behavior while retaining the bounded source memory model. The digest is
computed once at split and once per merge; no source copy is retained.

The synthetic ZIP projection has an explicit `Limits::projection_options`
carrier: it uses the hard synthetic source scratch ceiling while retaining the
caller's semantic limits. ODT, ODP, ODS and ODG projection consumers now pass
that carrier when reopening generated ZIP bytes. This preserves exact flat
input ceilings without incorrectly applying them to a larger internal ZIP.

Focused and affected validation completed with the capped environment from
`.coord/w8-odf-flat-limits/env.sh`, `CARGO_BUILD_JOBS=1`, and offline mode:

* `cargo test -p litchi-odf-common --all-features --offline flat::tests` —
  11 passed.
* `cargo test -p litchi-odt -p litchi-ods -p litchi-odp -p litchi-odg
  --all-features --offline --lib` — 807 + 694 + 221 + 22 passed.
* `cargo test -p litchi-ops --all-features --offline --lib` — 965 passed.
  An earlier run at the same command stopped at 964 passed and one failure:
  `tests::document::odt::flat_odt_source_limits_remain_typed`; that exposed
  the projection reopen using the original flat source ceiling and was fixed
  by the consumer migration above.
* `cargo clippy -p litchi-odf-common -p litchi-odt -p litchi-ods -p
  litchi-odp -p litchi-odg -p litchi-ops --all-targets --all-features
  --offline -- -D warnings` — passed.
* `cargo fmt --all -- --check` and `git diff --check` — passed.
* The exact glossary rustdoc command that previously exposed the iWork-only
  detector warning was rerun and passed:
  `RUSTC_BOOTSTRAP=1 cargo rustdoc --package litchi --lib --target-dir
  /home/zhuhe/code/litchi-worktrees/.coord/w8-odf-flat-limits/iwork-rustdoc
  --no-default-features --features iwork --offline -- -Zunstable-options
  --output-format json --document-hidden-items`.

The matched comparison is prepared against source `8d6abdab69`: split the
same tracked flat fixtures, publish a package, merge an exact no-op and a
single-child replacement, and compare exact bytes, refusal kinds, and
fallible-reservation paths. The source digest and retained-depth checks are
covered by the focused tests; any native timing or allocator qualification is
intentionally deferred until the coordinator grants a quiet measurement
window. Target size was 4.6 GiB at handoff; no benchmark binaries or large
evidence files were added.

## Full ODS gate and initial matched allocation receipt

The first matched allocation capture below used an embedded fixture and one
fresh process per route. It remains useful as historical evidence, but the
portable-source, four-leg receipt in the final section supersedes it for
qualification. In particular, the embedded probe's source hash and binary
hashes must not be used as the reproducibility inputs for the final result.

The complete bounded ODS gate was rerun after this follow-up, rather than only
the library target:

```text
source /home/zhuhe/code/litchi-worktrees/.coord/w8-odf-flat-limits/env.sh
export CARGO_NET_OFFLINE=true
python3 /home/zhuhe/code/litchi-worktrees/.coord/scripts/chunked-cargo-tests.py -p litchi-ods --all-features
```

The gate ran 18 bounded commands and 120 result blocks. It reported 1,768
passed, 0 failed and 13 filtered compile/doctest cases. The full output is
kept outside the checkout at
`/home/zhuhe/code/litchi-worktrees/.coord/w8-odf-flat-limits/ods-gate-full.log`
(183,230 bytes, SHA-256
`32921e8cfdab15c4eeaccb52f3dcdef47b7ea4e49661a0aa73a65ad1ba718062`). The
largest reported target was 5,892,730,880 bytes; the final target directory
was 5.5 GiB.

The initial matched non-timing probe uses the same source file and fixture for both
endpoints. The base endpoint is `8d6abdab69`; the final endpoint is
`384205d218`. Both probe lockfiles are 38,627 bytes with SHA-256
`5b181fd886d5315fd56250db98226fa44aa2fc5d5bd0d5f3eb895732e65dfeab`, the
probe source is 17,532 bytes with SHA-256
`a6516e98df099c12fc7150b3ef5709b68a5f06e8ba50e07e87fc5859f194a256`, and
the tracked fixture is 31,441 bytes with SHA-256
`9b65ff2a65ceb92409a78c831ae74bbff82c51abe0203684467903091956ef53`.
The exact build commands were:

```text
source /home/zhuhe/code/litchi-worktrees/.coord/w8-odf-flat-limits/env.sh
export CARGO_TARGET_DIR=/home/zhuhe/code/litchi-worktrees/.coord/w8-odf-flat-limits/probe-target-base
cargo build --release --offline                 # probe dependency paths at 8d6abdab69
export CARGO_TARGET_DIR=/home/zhuhe/code/litchi-worktrees/.coord/w8-odf-flat-limits/probe-target-final
cargo build --release --offline --features final # probe dependency paths at 384205d218
```

The protected base binary is 3,591,112 bytes, SHA-256
`65c7e435017f72620229f2a26e17ed70d5cc40d4c86536ee0c6ab23b374947a5`; the
final binary is 2,822,328 bytes, SHA-256
`2f2dd856c687ab4d38b9a28d7e49d85251f1e065fb4b21abc324e0bfebc794d2`.
The source and binary artifacts remain outside the repository.

Each route ran in a fresh process with a counting allocator. `requested` is
the sum of allocation and reallocation requests, `retained` is the live-byte
delta while the result owner remained alive, and `peak` is the interval peak
live-byte delta. Output hashes and sizes matched for every successful route.

| route | base result | final result | base calls/requested/retained/peak | final calls/requested/retained/peak |
| --- | --- | --- | --- | --- |
| split | success | success | 56 / 120,147 / 65,388 / 67,895 | 34 / 49,113 / 42,640 / 45,147 |
| to_package | success | success | 143 / 498,582 / 13,664 / 443,184 | 143 / 498,582 / 13,664 / 443,184 |
| merge_noop, independent equal source | exact source | exact source | 1 / 31,441 / 31,441 / 31,441 | 1 / 31,441 / 31,441 / 31,441 |
| merge_edit | exact changed bytes | exact changed bytes | 17 / 34,388 / 31,441 / 33,250 | 17 / 34,388 / 31,441 / 33,250 |
| merge_package | exact source | exact source | 791 / 1,282,589 / 31,441 / 136,403 | 791 / 1,282,589 / 31,441 / 136,403 |
| owner_roundtrip | exact source, 2 sheets | exact source, 2 sheets | 1,599 / 1,358,254 / 1,080,056 / 1,091,322 | 1,599 / 1,358,254 / 1,080,056 / 1,091,322 |

The final `split` reduced the observed allocation calls to 0.607× base,
requested bytes to 0.409×, retained bytes to 0.652× and peak bytes to 0.665×.
All other successful routes had identical allocation counters and exact
output hashes. This is allocation evidence only; no native latency claim is
made.

The proof and limit routes also ran in both endpoints. A same-length foreign
source was accepted by the base range-only check and returned the foreign
bytes; the final endpoint returned the typed `PatchConflict` after its digest
check. A loose split followed by no-change merge of an independently owned
source with an empty child at depth 3 succeeded on base, while the final
endpoint returned `Limit(Semantic(Depth), observed=3, maximum=2)`. Direct
split under that same tight profile returned the same typed depth refusal on
the final endpoint. These routes confirm the empty-element attempted-depth
check, retained-depth recheck, and independent-source content proof without
retaining a second flat source.

The initial JSON receipts are outside the checkout at
`/home/zhuhe/code/litchi-worktrees/.coord/w8-odf-flat-limits/matched-results-20261022/`;
the base and final JSONL files are 4,359 and 3,604 bytes. Native latency and
allocator qualification beyond this fresh-process count/retention probe
remain coordinator work.

## Portable four-leg qualification receipt

The committed probe source is
[`flat-qualification-probe.rs`](results/0362hda/flat-qualification-probe.rs)
(17,783 bytes, SHA-256
`d101b76120a33438bf99c113bfce3892607d1e28706166c610fb4ee673a9a5f8`). It
accepts the tracked fixture path as a runtime argument, so the source no
longer embeds a checkout-relative file. The compact receipt is
[`flat-qualification-receipt.json`](results/0362hda/flat-qualification-receipt.json)
(17,617 bytes, SHA-256
`e76aabd2de362870c5637879565815effb2a299620e401edba5bb13af9dde674`).

Both endpoints were rebuilt from the same probe source and lockfile. The base
binary is 3,563,672 bytes (SHA-256
`7c6e576e709b9bdc3e1aaacb1a22b016c0eeeae921039b5decf1b16217973fef`), and the
final binary is 2,794,872 bytes (SHA-256
`694c9eb532dc6516cdd48d569b9e0f8460161defe344f465a17de7253725684d`). The
lockfile is 38,627 bytes with SHA-256
`5b181fd886d5315fd56250db98226fa44aa2fc5d5bd0d5f3eb895732e65dfeab`; the
tracked fixture is 31,441 bytes with SHA-256
`9b65ff2a65ceb92409a78c831ae74bbff82c51abe0203684467903091956ef53`.

Each of the nine routes ran in four fresh processes in the order
`base → final → final → base`. Every endpoint's two legs had identical route
status, output metadata and allocation counters, including the refusal routes.
The final endpoint rejected same-length foreign content with
`PatchConflict` and rejected the retained-depth output with typed
`Limit(Semantic(Depth), observed=3, maximum=2)`; the base endpoint accepted
both routes under its historical range-only behavior. Successful output bytes
and hashes matched within each endpoint. The split route's final allocation
ratios against base were 0.607 calls, 0.409 requested bytes, 0.652 retained
bytes and 0.665 peak bytes. The other successful routes had identical
allocation counters. This is fresh-process allocation evidence only; no
latency claim is made.

## Coordinator receipt verification

The [coordinator receipt](results/0362hda/coordinator-verification.json)
verifies all 36 raw allocation legs against the committed summary, including
paired endpoint equality. It records each raw file's size and SHA-256. Both
protected binary hashes, portable probe copies and equal lockfiles match the
worker receipt. The 606 tracked source files in each endpoint's workspace,
core, ODF-common and ODS source set match their recorded Git blob IDs at
`8d6abdab69` and `384205d218`. This is verification of the retained artifacts
and current build inputs, not a reconstructed claim about an unrecorded
compiler command. The existing record preserves the exact intended build
commands and earlier projection-policy failure.

Successful-route outputs and allocation counters support the stated memory
results. Native latency qualification and integration remain open; this
verification does not accept a latency regression or waive the final gate.
