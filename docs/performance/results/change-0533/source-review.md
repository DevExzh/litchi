# 0533 CFB claim-sector success-path source review

`scope: read-only OLE2/CFB candidate review before implementation`

`baseline: 0dd079b95982aa1fbef5eddd1ec3743261dcaf20`

`performance_claim: none`

This review covers the frozen [0533 plan](plan.json), the current baseline
[`claim_sector`](../../../../crates/litchi-cfb/src/file.rs#L1025), and the
accepted CFB constraints in ADR 0001, 0002, 0003, 0005, 0006, 0008, 0010,
0011, 0024, and 0026. The baseline `file.rs` SHA-256 is
`62f7b357d35b5bb2921336bb6a983a8d7e9ec58c30ae694a3be4c89e9c251788`, matching
the retained 0532 source. No Rust source, test, build, or benchmark was run or
changed for this review.

## Candidate boundary

The proposed production delta has one purpose: keep the successful sector-role
claim small enough for the compiler to inline at ordinary call sites while
placing the three formatted failure paths in private cold functions:

```text
claim_sector_index_error(sector, role)
claim_sector_bounds_error(sector, role)
claim_sector_conflict_error(sector, existing, role)
```

Each helper is private and carries `#[cold]` and `#[inline(never)]`. The
`claim_sector` method receives only an ordinary `#[inline]` hint. It must not
become `#[inline(always)]`, public, generic, unsafe, or dependent on a new
crate. The proposal is a code-layout experiment; the attributes do not by
themselves prove that the release callers inline the checked success path.

The three helpers must format the existing messages byte-for-byte:

```text
{role} sector {sector} does not fit usize
{role} sector {sector} is outside the file
Sector {sector} is claimed by both {existing} and {role}
```

`PhysicalSectorRole::label` remains the sole label mapping. The helper taking
the prior role must receive a copy of the role before assignment, so the
conflict diagnostic cannot observe a role that was already overwritten.

## Required checked behavior

The method's logical and observable order is part of the CFB validation
contract:

1. Convert the hostile `u32` sector to `usize`, returning the index error on
   conversion failure.
2. Perform checked `sector_roles.get_mut`, returning the bounds error when the
   physical role map has no slot.
3. Reject any role other than `Unclaimed`, returning the conflict error with
   the old and requested labels.
4. Assign the requested role and return `Ok(())`.

The candidate must leave all four operations in this order. In particular, it
must not use unchecked indexing, pre-check a different length, move the role
assignment before the conflict branch, or replace the `Result` with an
infallible helper. A failed conversion, missing slot, or conflict must leave
the role map unchanged. `claim_chain` currently retains earlier assignments
when a later sector fails; the candidate must not add rollback or otherwise
change that private call-sequence behavior.

Calling a cold helper from the existing `map_err` and `ok_or_else` closures is
compatible with this boundary only if it leaves the borrow and evaluation
order equivalent. The conversion closure must run before `get_mut`; the
bounds closure must run only for `None`; and the conflict helper must run
before any write. Helpers should take copied `u32`/enum values rather than
borrow `self` or the mutable slot. This keeps the hot method free of formatter
state and avoids extending the mutable borrow across the role assignment.

The conversion failure is not dynamically reachable on the current 64-bit
benchmark host because every `u32` fits in `usize`. It remains a required
portable branch and must stay visible in source and in any supported narrower
target build. The out-of-file and conflict branches are reachable on the
current host and require exact-message tests.

## Callers and ownership proof

The direct callers and their role arguments are unchanged:

| caller | role(s) claimed | required boundary |
| --- | --- | --- |
| `load_fat` | `Fat`, `Difat` | Header/DIFAT duplicate and physical bounds checks remain before publication of the FAT index. |
| `load_minifat` | `MiniFat` through `claim_chain` | The complete checked chain is collected before physical roles are changed. |
| `load_directory` | `Directory` through `claim_chain` | Directory chain validation and fallible data handling retain their existing order. |
| `validate_stream_allocations` | `MiniStream`, `RegularStream` through `claim_chain` | Root and regular chains remain collect-then-claim; mini-sector ownership stays a separate checked namespace. |
| `validate_physical_sector_layout` | none | The full FAT/role reconciliation pass remains present and unchanged. |

The chain collectors still validate table bounds, conversion, cycles, terminal
markers, fallible reservations, and reset-on-error behavior before the claim
step. Their checks do not prove that a collected sector fits the physical role
map, so removing `get_mut` would weaken a real defense. The final physical
pass still performs `fat.get(sector)` before its unclaimed-marker test; this
candidate must not fuse, reorder, or remove that pass.

`validate_stream_allocations` must retain its collect-then-claim order. A
success-path inline hint cannot be used as a reason to fuse collection with
ownership publication, bypass the checked scratch reservation, or alter the
partial failure state of `claim_chain`. No stream, directory, FAT, MiniFAT,
source, or public API ownership changes are part of this experiment.

## ADR and resource compatibility

| ADR | obligation for this candidate |
| --- | --- |
| 0001, 0002, 0024, 0026 | Keep CFB validation in `litchi-cfb`, preserve crate and directory ownership, and expose no new API or dependency. |
| 0003, 0006 | Preserve deterministic typed errors, validation precedence, role mutation boundaries, and non-mutating failure behavior. |
| 0005 | Move only formatting code out of the successful path; measure representative latency, instructions, allocations, and RSS rather than inferring a gain from attributes or static instruction counts. |
| 0008 | Run the full applicable quality gates and retain source, binary, corpus, and assembly custody for any result. |
| 0010, 0011 | Keep physical CFB ownership and preservation boundaries intact; this is not an archive or source-backed ownership change. |

The standalone performance harness is its own workspace and has no explicit
release LTO profile. The root workspace's LTO setting cannot explain or
guarantee this candidate's generated code. Assembly must be captured from the
actual baseline, candidate, and final binaries. It should check the six
`claim_sector` variants, the helper symbols, stack reservation, role stores,
and the relevant `load_fat`/`claim_chain` callers. A smaller helper alone is
not evidence of inlining, and a larger caller or duplicated cold block is a
candidate regression.

The frozen plan's four primary workflows retain the 0524 admission rule:
each primary p50 must improve by at least 3% in both paired repeats, and
constructor instruction count must decrease in both repeats. CFB tiny,
many-small, and few-large rows remain guards; all rows require individual
review for adverse latency, mean/tail, RSS, and allocation behavior. A same-
build variation or matched regression over 5% is a review trigger. The
operation-local allocator lane must confirm that a successful claim-layout
change does not add allocation calls, bytes, or incremental peak; it is not a
substitute for the native gate. If the primary gate fails, the production
change should be rejected even if Callgrind or assembly appears favorable.

No cold-cache, physical-I/O, producer, concurrency, scaling, generic OLE2,
or OOXML speedup follows from this source review. OLE2/OOXML remains active;
ODF remains deferred and iWork is outside scope.

## Source disposition

The candidate is source-level safe to stage for matched measurement, provided
the implementation follows the exact helper boundary and the checked body
above. This review does not approve production retention. The final decision
requires differential behavior tests, all applicable quality checks, actual
candidate/final binary inspection, and the frozen ABBA performance and
allocation gates. A failed candidate may retain only tests that exercise the
unchanged private contract without depending on the new helper symbols.

## Applied source review

The applied production patch matches the proposed helper boundary exactly: three private cold, non-inlined functions and the ordinary inline hint, with conversion, checked lookup, conflict refusal, assignment and exact strings retained. The remainder of the production file is unchanged. The independent tests occupy only the existing private test module. `candidate/source.patch` records the complete runtime-plus-test diff; `candidate.patch` records the production-only delta.
