# 0359re — RTF replay shared change storage

- Date: 2026-10-22.
- Status: final source measured for focused tests, allocation and callgrind
  attribution; clean-window native timing remains open.
- Source commit: `f75941fc7d`.
- Final release binary SHA-256:
  `6d47549df0be90cee8f5830685fc18d108824b1c68d0db8d84f800697b01bb52`.

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
allocates an Arc buffer and clones each `Change`. The final measurement below
shows that this removes one of the prototype's four promotion allocations.

## Final-source measurement

The final binary was built with one Cargo job, no incremental compilation,
release debug information disabled, a 20 GiB systemd memory cap, and a
`TMPDIR` under `/home/zhuhe`. The command was:

```text
CARGO_TARGET_DIR=(deleted-coord-artifact) CARGO_BUILD_JOBS=1 CARGO_INCREMENTAL=0 CARGO_PROFILE_RELEASE_DEBUG=0 LITCHI_MEM_MAX=20G TMPDIR=/home/zhuhe/.cache/rtf-refusal-perf cargo build --release -p litchi-rtf --example refusal_contract_bench
```

It was compared with the original 0359b binary
`4693bd12232d92a61ef2f437de2d623e345a106288d5d0a5c5638cf27677420a` using the
fixed external executable path, CPU 7, `MemoryMax=20G`, and A-B-B-A heaptrack
captures. Compiler processes were active, so these are allocation-attribution
receipts and not native qualification.

| Route set | Allocations | Temporary | Peak heap | Leaked |
| --- | ---: | ---: | ---: | ---: |
| original, runs 1 and 2 | 177,971 | 1,258 | 3.57 MiB | 544 B |
| final source, runs 1 and 2 | 177,974 | 1,258 | 3.57 MiB | 544 B |

The final source has three additional allocation calls, all on the first
`ChangeStorage::to_shared` promotion edge. The replay route made 3,400 calls;
callgrind reports one `to_shared` call at 1,008 inclusive Ir, with the
remaining 3,399 replays reusing the cached Arc. The heap diff has no peak or
leak change. The previous prototype used the Vec/Box bridge and measured four
additional calls and 1,191 inclusive Ir for that promotion.

The final callgrind A-B-B-A runs used the same fixed path and CPU. Whole-program
Ir was 627,542,591 and 627,541,901 for the original runs, versus 624,872,284
and 624,871,461 for the final runs. The final median is 2,670,374 Ir lower
than the original median. Against the prior current API binary, final
`Package::apply` is 511,098 Ir for 3,400 calls versus 1,679,600; the final
`to_shared` body is 1,008 Ir for one call versus the prototype's 1,191.
`Edit::commit_staged` is 347,843,299 Ir versus 347,723,593, and
`Commit::new` is 24,847 versus 24,709. The final `Commit` drop is 12,968,314
versus 13,382,208, while Arc drop is 25,400,592 versus 25,398,727. These
deterministic instruction results are diagnostic and do not replace native
timing.

The final binary exits successfully on all eight probe routes. The focused
exact-source tests passed: destination splice (9), paragraph lifecycle (16),
native run format (5), and canonical primary replay (4), including repeated
replay after dropping the source patch owner. No error, output or inverse
assertion changed.

The external receipts are under
`(deleted-coord-artifact)`:

- `heap-repair-final-20261022/`, whose four compressed captures are 18,654 to
  19,293 bytes;
- `callgrind-repair-final-20261022/`, whose four callgrind files are below
  137 KiB;
- `heap-repair-storage-20261022/` and `callgrind-repair-storage-20261022/`,
  which retain the earlier four-allocation prototype receipts.

Their SHA-256 values are retained in the coordinator scratch area and in the
commit record below. Every final heap capture was rejected for native use
because compiler processes were observed. The five-block clean-window native
runner remains required; no native timing or performance acceptance is claimed.
