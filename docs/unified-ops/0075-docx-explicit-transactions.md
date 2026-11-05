# 0075 — Bounded DOCX closure transactions and process-local replay

This checkpoint advances M2 W6 from the retained native commits in 0074.
The preceding turn made verified implementation progress. Full command-batch
execution, coalescing, performance acceptance and the broader goal remain open.

## Public contract

`File::edit` and `File::edit_with_limits` expose
`transaction::{Edit, Commit, Patch, Diagnostics, Error, LimitKind}`. The facade
re-exports the module. `Commit<T>` carries public `file`, `patch`, `diagnostics`
and closure `value` fields, following record 0015's normalized result shape.
`Edit` follows the adopted glossary instead of the draft name `Transaction`.

The proxy exposes the five integrated DOCX main-document mutations and staged
`get`/`find`. Each operation resolves against the preceding native result;
`add` returns its current canonical positional path. Subsequent structural
changes can change what that path addresses. Reads take `&mut self` as settled
in record 0015, allowing future run barriers without changing the public shape.

A failed mutation poisons the whole closure, including when its error is caught.
Later calls report that first failure and its zero-based mutation index. Read
errors can be handled by the closure without poisoning. Returning `Err` or
unwinding discards the candidate. A private invocation identity also rejects
swapping proxies between nested closures, even over the same original file;
limits, failure state and history cannot transfer into another invocation.

`SemanticLimits` defaults to 10,000 mutation calls and a 256 MiB summed charge
for each step's before/after main XML. The checked consuming builder rejects
zero and values above `isize::MAX`. Shared XML is charged again per step. This
is a conservative logical retention bound, not a measurement or bound on total
heap use. Native owners retain package/source data under their existing limits.
The byte charge is checked after the isolated native candidate commit and before
retaining it; native admission/execution limits govern transient staging.

`Source::Transaction` retains typed invalid-limit, limit, allocation and
patch-conflict errors. They map to the existing closed unified codes
`invalid_argument`, `limit_exceeded`, `internal_error` and `invalid_input`.
No new global code or authored diagnostic content is introduced.

## Replay and security

The opaque patch retains ordered native commits plus their immutable before and
after handles. Forward and inverse application reuse already-validated endpoint
handles without serializing or reopening ZIPs. A patch is process-local, not a
durable interchange format. Its bound input must be the original handle or a
cheap clone; a separate open conflicts even when all bytes match. Identity uses
the outer owner including signature disposition and origin sidecars.

Every retained source guard is checked before and after selecting the replay
endpoint. This includes an original source-backed signed owner whose authorized
edit promoted to an eager unsigned owner. Checking only the final owner would
lose that guard. The closure also checks source freshness after errors and after
building its complete result. Signed wrappers, their signed inner handles,
locked inputs and read-only admission refuse before editing.

The explicitly unsigned handle is the logical baseline. The patch never
reverses signature removal or recreates a signed state. For the current sealed
main-document-only handler family, final change is determined by comparing the
first source XML and last target XML, rather than counting operations. A
cancelled or exact-no-op transaction canonically returns the original logical
handle, preserving its exact archive publication. Its patch replays those
canonical top-level endpoints; it does not execute the retained intermediate
steps. Future graph-editing families must replace this main-only change proof.

Diagnostics report final logical `changed`, successful mutation-call count,
leaf-operation count (saturating on overflow), and the retained XML charge.
Operation counts include no-ops. Debug output omits document and closure values.
Advisory diagnostic entries remain future work alongside the wider executor.

## Discoverability and remaining work

`capabilities::Verb::Edit` now reports this closure operation separately from
`Batch`, which remains unsupported. DOCX reports the same signature and
retained-source admission policies as its five individual write verbs. Other
formats remain explicit follow-ups. Decrypted DOCX owners still refuse native
editing to avoid losing their encryption profile. This existing leaf limit is
now listed for all five write verbs and the closure in encryption-enabled
capability reports, with a follow-up. A regression checks refusal before the
closure and continued refusal of plaintext byte publication. Capability goldens and the generated
OfficeCLI delta are refreshed; this adds no OfficeCLI batch-execution claim.

This implementation closes a native commit at every operation. It establishes
correctness and bounded retention for the integrated subset. Homogeneous family
coalescing/fusion, command admission and best-effort isolation, larger graphs,
advisory diagnostics, and paired 1/10/100/1k/10k performance evidence remain
required. M2 and the full goal are not complete.

## Verification

- All-feature ops suite: 1,021 tests/doctests passed. After adding the final
  encrypted-owner refusal and discovery regression, the final targeted run
  passed all 16 transaction tests and eight capability tests.
- DOCX-only ops suite: 646 tests/doctests passed; minimal-feature suite: 417.
- The facade transaction test passed, including module exports and inverse replay.
- Final all-feature all-target Clippy and rustdoc passed with warnings denied.
- Scoped rustfmt, `git diff --check`, crate-boundary validation and the generated
  OfficeCLI delta check passed. The boundary checker was rerun offline after
  stopping an initial metadata check that was waiting on network access.
- Independent read-only review examined replay identity, guard ordering, logical
  unsigned baselines and no-op collapse. The final implementation and tests cover
  those conditions; no review worktree or temporary artifact was created.

The transaction regressions compare all five verbs to a single-operation byte
oracle, exercise eager and source owners, enforce foreign-owner rejection and
nested-proxy identity, prove rollback after errors/panics and exact archive
preservation after cancellation, cover bound equality/refusal, and retain
origin and source fences across signature-removal promotion.

The rewritten local base remains an ancestor of HEAD. No rebase onto that
revision is necessary. The required staged size and binary-directory checks
printed nothing. An audit of 331 changed/untracked paths found no prohibited
large evidence, binary directories or compiled ELF artifacts.

`git add -A` failed because Git cannot create
`/home/zhuhe/code/litchi/.git/worktrees/unified-ops-api/index.lock` on the read-only
metadata filesystem. No commit was created; changes remain in this checkout.
The full goal remains active.

Reproduction commands:

```sh
CARGO_NET_OFFLINE=true cargo test -p litchi-ops --no-default-features --features docx
CARGO_NET_OFFLINE=true cargo test -p litchi-ops --no-default-features
CARGO_NET_OFFLINE=true cargo test -p litchi-ops --all-features
CARGO_NET_OFFLINE=true cargo clippy -p litchi-ops --all-features --all-targets -- -D warnings
CARGO_NET_OFFLINE=true RUSTDOCFLAGS='-D warnings' cargo doc -p litchi-ops --all-features --no-deps
python3 tools/check_crate_boundaries.py
python3 tools/officecli_capability_delta.py --officecli /home/zhuhe/code/litchi/3rdparty/OfficeCLI --check
```

Cleanup removed 25,012 build files (34.5 GiB) and 13 scratch/cache
paths, including generated untracked lockfiles. No temporary worktrees or
external fixture links were created by this checkpoint.
