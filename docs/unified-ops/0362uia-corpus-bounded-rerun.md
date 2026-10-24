# 0362uia — bounded corpus and focused rerun

This addendum corrects the validation provenance for the corpus-fidelity
refresh in `0362ui`. The old unblessed diagnostic was run from the removed
`uo2/final-corpus-fidelity` leaf without an explicit target path: Cargo printed
relative `target/debug/deps/...` paths, and cleanup later found that leaf's
local `target/` at 14,428 files and more than 20.3 billion apparent bytes.
That pass remains diagnostic only. Its evidence and the quarantined old target
were preserved; neither was overwritten or deleted.

The new validation leaf is
`/home/zhuhe/code/litchi-worktrees/uo2-corpus-bound-rerun`, branch
`uo2/corpus-bound-rerun`, at source commit
`2510bf5b40d2db17e774cf68734f7e08d91700ec`. The pinned `Cargo.lock` was
copied from the known f6fe receipt source
`/home/zhuhe/code/litchi-worktrees/.coord/odp-final-f6fe-source-20261024/Cargo.lock`
and has SHA-256
`52ed0cacac15b5e440349bfc0f4aca2665bcf50d59d4321624e596bb410783c6`.
Record ID `0362uia` was unused before this addendum.

## Explicit bounded envelope

Every command below supplied its own environment; no shell environment from a
previous command was relied upon. The first `PATH` entry is the memory-capped
Cargo shim `/home/zhuhe/code/litchi-worktrees/.coord/bin/cargo`, which delegates
to `/home/zhuhe/.cargo/bin/cargo` under a systemd scope. `LITCHI_CAPPED` was
explicitly unset so the shim was engaged. Both `CARGO_TARGET_DIR` and Cargo's
`--target-dir` point to the same external target:

```text
PATH=/home/zhuhe/code/litchi-worktrees/.coord/bin:$PATH
LITCHI_MEM_MAX=20G
CARGO_BUILD_JOBS=1
RUST_TEST_THREADS=1
CARGO_INCREMENTAL=0
CARGO_PROFILE_DEV_DEBUG=0
CARGO_PROFILE_TEST_DEBUG=0
CARGO_PROFILE_RELEASE_DEBUG=0
CARGO_NET_OFFLINE=true
CARGO_TARGET_DIR=/home/zhuhe/code/litchi-worktrees/.coord/final-corpus-fidelity-target
TMPDIR=/home/zhuhe/code/litchi-worktrees/.coord/corpus-bound-rerun-tmp-20261024-2510bf5b40
```

The target measured 3,483,757,354 bytes and 3,217 files before the first
command. After both commands it measured 3,483,689,639 bytes and 3,217 files,
well below the 20,000,000,000-byte job limit. The new leaf had no local
`target/` directory. The temporary directory was empty after the runs.

## Unblessed report equality

The exact command was:

```text
env -u LITCHI_CAPPED \
  PATH=/home/zhuhe/code/litchi-worktrees/.coord/bin:$PATH \
  LITCHI_MEM_MAX=20G CARGO_BUILD_JOBS=1 RUST_TEST_THREADS=1 \
  CARGO_INCREMENTAL=0 CARGO_PROFILE_DEV_DEBUG=0 CARGO_PROFILE_TEST_DEBUG=0 \
  CARGO_PROFILE_RELEASE_DEBUG=0 CARGO_NET_OFFLINE=true \
  CARGO_TARGET_DIR=/home/zhuhe/code/litchi-worktrees/.coord/final-corpus-fidelity-target \
  TMPDIR=/home/zhuhe/code/litchi-worktrees/.coord/corpus-bound-rerun-tmp-20261024-2510bf5b40 \
  cargo test --offline --locked \
  --target-dir /home/zhuhe/code/litchi-worktrees/.coord/final-corpus-fidelity-target \
  -p litchi-ops --all-features --test corpus_fidelity --test crud_matrix -- --nocapture
```

`LITCHI_BLESS` was absent. The child exited 0: corpus fidelity opened 1,020
files in 82.30 seconds, and CRUD reported zero unreached claimed cells in
28.30 seconds. The target changed from 3,483,757,354 to 3,483,689,458 bytes;
the file count stayed 3,217. The receipt, including the complete changed-file
manifest and resolved test-artifact paths, is
`/home/zhuhe/code/litchi-worktrees/.coord/corpus-bound-rerun-receipt-20261024-2510bf5b40/01-corpus-crud-unblessed.log`
(68,009 bytes, SHA-256
`64a8cec6b1d318cf378d2c8006afdfa8ac65d0ae97ef44e88dc0a180292c6a42`). The
resolved test binaries were:

```text
/home/zhuhe/code/litchi-worktrees/.coord/final-corpus-fidelity-target/debug/deps/corpus_fidelity-d11d840ec08ebf56
/home/zhuhe/code/litchi-worktrees/.coord/final-corpus-fidelity-target/debug/deps/crud_matrix-c1ad7948b32cfc9c
```

## Focused two-fixture consumers

The exact command was:

```text
env -u LITCHI_CAPPED \
  PATH=/home/zhuhe/code/litchi-worktrees/.coord/bin:$PATH \
  LITCHI_MEM_MAX=20G CARGO_BUILD_JOBS=1 RUST_TEST_THREADS=1 \
  CARGO_INCREMENTAL=0 CARGO_PROFILE_DEV_DEBUG=0 CARGO_PROFILE_TEST_DEBUG=0 \
  CARGO_PROFILE_RELEASE_DEBUG=0 CARGO_NET_OFFLINE=true \
  CARGO_TARGET_DIR=/home/zhuhe/code/litchi-worktrees/.coord/final-corpus-fidelity-target \
  TMPDIR=/home/zhuhe/code/litchi-worktrees/.coord/corpus-bound-rerun-tmp-20261024-2510bf5b40 \
  cargo test --offline --locked \
  --target-dir /home/zhuhe/code/litchi-worktrees/.coord/final-corpus-fidelity-target \
  -p litchi-ops --all-features --test odp_unified_write --test xlsx_cells_runs -- --nocapture
```

`LITCHI_BLESS` was absent. The child exited 0: all 7 ODP unified-write tests
and all 6 XLSX rich-cell tests passed. The target changed from 3,483,689,458
to 3,483,689,639 bytes; the file count stayed 3,217. The receipt, including
the changed-file manifest and resolved test-artifact paths, is
`/home/zhuhe/code/litchi-worktrees/.coord/corpus-bound-rerun-receipt-20261024-2510bf5b40/02-focused-odp-xlsx.log`
(4,905 bytes, SHA-256
`8aa1822ca35bda369d1931b88756fc1be3a0c863e30719029e30de554b04121d`). The
resolved test binaries were:

```text
/home/zhuhe/code/litchi-worktrees/.coord/final-corpus-fidelity-target/debug/deps/odp_unified_write-bed179e84f0f5589
/home/zhuhe/code/litchi-worktrees/.coord/final-corpus-fidelity-target/debug/deps/xlsx_cells_runs-05dfbb3732ae831b
```

No golden, source, or `NORMALIZING` data changed during either unblessed run.
The earlier `0362ui` text now labels its relative-target pass diagnostic only
and points to this corrected bounded provenance. No `PROGRESS.md` change is
included.
