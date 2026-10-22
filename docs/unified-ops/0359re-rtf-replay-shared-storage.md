# 0359re — RTF replay shared change storage

- Date: 2026-10-22.
- Status: source repair recorded; final build, heap qualification and native
  timing remain open.
- Source base: `b0eb83e41e`.
- Final source in this commit has not been rebuilt because the coordinator's
  host-wide Cargo hold is active.

The replay facade must return an owned `Commit` containing its own reversible
`Patch`. The old implementation copied `self.changes` for every replay. This
change keeps ordinary commits in their original `Box<[Change]>` allocation and
promotes a source patch to one cached `Arc<[Change]>` only when replay first
needs an owned result. Every later replay commit clones that Arc pointer. The
returned commit therefore does not allocate a second shared change buffer.

## Ownership boundary

`Package::apply(&self, &Patch)` borrows its input while returning a `Commit`
that may outlive it. A `Box<[Change]>` owned by the input cannot be moved into
the returned patch through that borrow. Consuming `Patch` would change the
public API; moving the Box through interior synchronization or an intrusive
unsafe reference count would add a new ownership mechanism. The implementation
uses a private `OnceLock` instead: the source patch retains its ordinary Box,
and the first replay creates the cache that both owners can safely retain.

`Patch::Clone` resets the cache when its storage is still owned. If the patch
has already entered replay, its `Shared` storage clones the Arc pointer. The
patch before/after packages, diagnostics, exact-source checks, inverse and
durable conversion keep their existing behavior.

The final source uses `Arc::from(&changes[..])`, which removes the intermediate
Vec/Box conversion from the promotion path. Rust's slice conversion still
allocates an Arc buffer and clones each `Change`; this source edit has not yet
been measured. The measured prototype used the equivalent
`to_vec().into_boxed_slice()` bridge and is reported below, without treating it
as the final candidate result.

## Measured prototype

The prototype binary was built before the Cargo hold from the same probe and
has SHA-256
`74cfeddc1fef9669bca6a2faea89a4a3a96f2fbfc84e7545b3a2d98a81f4ed91`. It was
compared with the original 0359b binary
`4693bd12232d92a61ef2f437de2d623e345a106288d5d0a5c5638cf27677420` using the
fixed external executable path, CPU 7, `MemoryMax=20G`, and A-B-B-A heaptrack
captures. Compiler processes were active, so these are allocation-attribution
receipts and not native qualification.

| Route set | Allocations | Temporary | Peak heap | Leaked |
| --- | ---: | ---: | ---: | ---: |
| original, runs 1 and 2 | 177,971 | 1,258 | 3.57 MiB | 544 B |
| storage prototype, runs 1 and 2 | 177,975 | 1,258 | 3.57 MiB | 544 B |

The four additional allocation calls are all on the first
`ChangeStorage::to_shared` promotion edge. The replay route made 3,400 calls;
the first promotion contributed 1,191 inclusive Ir and the remaining 3,399
replays reused the cached Arc. The heap diff has no peak or leak change.

Callgrind used the same protected old and prototype binaries. Whole-program
Ir was 626,297,809 and 626,297,658 for the old runs, versus 624,865,766 and
624,866,831 for the storage runs. Inclusive `Package::apply` fell from
1,679,600 Ir (3,400 calls) to 511,281 Ir; `Edit::commit_staged` changed from
347,723,593 to 347,842,333 Ir. These deterministic instruction results are
diagnostic and do not replace native timing.

The external receipts are under
`/home/zhuhe/code/litchi-worktrees/.coord/rtf-refusals/`:

- `heap-repair-storage-20261022/`, whose four compressed captures are 18,737
  to 18,989 bytes;
- `callgrind-repair-storage-20261022/`, whose four callgrind files are below
  1 MiB.

Their SHA-256 values are retained in the coordinator scratch area and in the
commit record below. The final source still requires a release rebuild and a
new matched original-to-candidate heap run. The five-block clean-window native
runner also remains required; no timing or performance acceptance is claimed.
