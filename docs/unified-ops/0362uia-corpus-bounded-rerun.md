# 0362uia — bounded corpus and focused rerun

This addendum supersedes the earlier diagnostic receipts referenced by
`0362ui`. The removed `uo2/final-corpus-fidelity` leaf had a local `target/`
with 14,428 files and more than 20.3 billion apparent bytes; its unblessed
Cargo output printed relative `target/debug/deps/...` paths. That pass remains
diagnostic only. Its evidence and quarantined target were preserved.

The validation leaf is
`/home/zhuhe/code/litchi-worktrees/uo2-corpus-bound-rerun`, branch
`uo2/corpus-bound-rerun`, source commit
`2510bf5b40d2db17e774cf68734f7e08d91700ec`. The pinned `Cargo.lock` was copied
from `/home/zhuhe/code/litchi-worktrees/.coord/odp-final-f6fe-source-20261024/Cargo.lock`
and has SHA-256
`52ed0cacac15b5e440349bfc0f4aca2665bcf50d59d4321624e596bb410783c6`.
Record ID `0362uia` was unused before this addendum.

Every final child supplied its own environment; no previous shell environment
was relied upon. The memory-capped Cargo shim is
`/home/zhuhe/code/litchi-worktrees/.coord/bin/cargo`, and it delegates to
`/home/zhuhe/.cargo/bin/cargo` under a systemd scope. The final children used
`LITCHI_BLESS=<unset via -u>`, `LITCHI_CAPPED=<unset via -u>`,
`LITCHI_MEM_MAX=20G`, `CARGO_BUILD_JOBS=1`, `RUST_TEST_THREADS=1`,
`CARGO_INCREMENTAL=0`, `CARGO_PROFILE_DEV_DEBUG=0`,
`CARGO_PROFILE_TEST_DEBUG=0`, `CARGO_PROFILE_RELEASE_DEBUG=0`, and
`CARGO_NET_OFFLINE=true`.

Both `CARGO_TARGET_DIR` and `--target-dir` were the same absolute external
path `/home/zhuhe/code/litchi-worktrees/.coord/final-corpus-fidelity-target`.
The target measured 3,483,689,639 bytes and 3,217 files before and after both
final cached reruns, below the 20,000,000,000-byte limit. The leaf had no local
`target/` directory. Tracked status before and after each child was clean
(`## uo2/corpus-bound-rerun`). The final receipt headers recorded source
`c9daaef6289a0fcdf455f4dde3efeb15ab7d558e`; this is the record-only commit
over the code source `2510bf5b40d2db17e774cf68734f7e08d91700ec`.

The earlier bounded-looking receipts remain immutable and are superseded by
the final pair below:

- `01-corpus-crud-unblessed.log`: SHA-256
  `64a8cec6b1d318cf378d2c8006afdfa8ac65d0ae97ef44e88dc0a180292c6a42`.
- `02-focused-odp-xlsx.log`: SHA-256
  `8aa1822ca35bda369d1931b88756fc1be3a0c863e30719029e30de554b04121d`.

## Final corpus and CRUD child

The final child used this exact captured argv (the expanded `PATH` is retained
verbatim from the receipt):

    /usr/bin/env -u LITCHI_BLESS -u LITCHI_CAPPED PATH=/home/zhuhe/code/litchi-worktrees/.coord/bin:/home/zhuhe/.local/bin:/home/zhuhe/.codex/packages/standalone/releases/0.164.0-x86_64-unknown-linux-musl/codex-path:/home/zhuhe/.codex/tmp/arg0-linux/codex-arg0FCLdkN:/home/zhuhe/.codex/tmp/arg0-linux/codex-arg00dw7C8:/home/zhuhe/.local/bin:/home/zhuhe/.kimi-code/bin:/home/zhuhe/.bun/bin:/home/zhuhe/.local/bin:/run/user/1001/fnm_multishells/146209_1792754085282/bin:/home/zhuhe/.local/bin:/home/zhuhe/.moon/bin:/home/zhuhe/.kimi-code/bin:/home/zhuhe/.bun/bin:/home/zhuhe/.local/bin:/run/user/1001/fnm_multishells/52327_1792753151407/bin:/home/zhuhe/.cargo/bin:/usr/local/sbin:/usr/local/bin:/usr/sbin:/usr/bin:/sbin:/bin:/usr/games:/usr/local/games:/snap/bin:/home/zhuhe/.fzf/bin:/home/zhuhe/go/bin:/home/zhuhe/go/bin LITCHI_MEM_MAX=20G CARGO_BUILD_JOBS=1 RUST_TEST_THREADS=1 CARGO_INCREMENTAL=0 CARGO_PROFILE_DEV_DEBUG=0 CARGO_PROFILE_TEST_DEBUG=0 CARGO_PROFILE_RELEASE_DEBUG=0 CARGO_NET_OFFLINE=true CARGO_TARGET_DIR=/home/zhuhe/code/litchi-worktrees/.coord/final-corpus-fidelity-target TMPDIR=/home/zhuhe/code/litchi-worktrees/.coord/corpus-bound-rerun-tmp-20261024-c9daaef628-final cargo test --offline --locked --target-dir /home/zhuhe/code/litchi-worktrees/.coord/final-corpus-fidelity-target -p litchi-ops --all-features --test corpus_fidelity --test crud_matrix -- --nocapture

The receipt is
`/home/zhuhe/code/litchi-worktrees/.coord/corpus-bound-rerun-receipt-20261024-c9daaef628-final/01-final-corpus-crud-unblessed.log`
(4,686 bytes, SHA-256
`14f9f76ca1bfd05b31162690806ad7b3a958989be36548bd2101f7783713a6e3`). It
ran from `07:08:41Z` to `07:11:13Z`, exited 0, opened 1,020 corpus files in
101.78 seconds, and reported zero unreached CRUD cells in 42.85 seconds.
Resolved test binaries were
`/home/zhuhe/code/litchi-worktrees/.coord/final-corpus-fidelity-target/debug/deps/corpus_fidelity-d11d840ec08ebf56`
and
`/home/zhuhe/code/litchi-worktrees/.coord/final-corpus-fidelity-target/debug/deps/crud_matrix-c1ad7948b32cfc9c`.
The temporary directory ended with exactly one zero-byte marker,
`01-final-corpus-crud.start`; it was not an empty directory.

## Final focused ODP/XLSX child

The final child used this exact captured argv:

    /usr/bin/env -u LITCHI_BLESS -u LITCHI_CAPPED PATH=/home/zhuhe/code/litchi-worktrees/.coord/bin:/home/zhuhe/.local/bin:/home/zhuhe/.codex/packages/standalone/releases/0.164.0-x86_64-unknown-linux-musl/codex-path:/home/zhuhe/.codex/tmp/arg0-linux/codex-arg0FCLdkN:/home/zhuhe/.codex/tmp/arg0-linux/codex-arg00dw7C8:/home/zhuhe/.local/bin:/home/zhuhe/.kimi-code/bin:/home/zhuhe/.bun/bin:/home/zhuhe/.local/bin:/run/user/1001/fnm_multishells/146209_1792754085282/bin:/home/zhuhe/.local/bin:/home/zhuhe/.moon/bin:/home/zhuhe/.kimi-code/bin:/home/zhuhe/.bun/bin:/home/zhuhe/.local/bin:/run/user/1001/fnm_multishells/52327_1792753151407/bin:/home/zhuhe/.cargo/bin:/usr/local/sbin:/usr/local/bin:/usr/sbin:/usr/bin:/sbin:/bin:/usr/games:/usr/local/games:/snap/bin:/home/zhuhe/.fzf/bin:/home/zhuhe/go/bin:/home/zhuhe/go/bin LITCHI_MEM_MAX=20G CARGO_BUILD_JOBS=1 RUST_TEST_THREADS=1 CARGO_INCREMENTAL=0 CARGO_PROFILE_DEV_DEBUG=0 CARGO_PROFILE_TEST_DEBUG=0 CARGO_PROFILE_RELEASE_DEBUG=0 CARGO_NET_OFFLINE=true CARGO_TARGET_DIR=/home/zhuhe/code/litchi-worktrees/.coord/final-corpus-fidelity-target TMPDIR=/home/zhuhe/code/litchi-worktrees/.coord/corpus-bound-rerun-tmp-20261024-c9daaef628-final cargo test --offline --locked --target-dir /home/zhuhe/code/litchi-worktrees/.coord/final-corpus-fidelity-target -p litchi-ops --all-features --test odp_unified_write --test xlsx_cells_runs -- --nocapture

The receipt is
`/home/zhuhe/code/litchi-worktrees/.coord/corpus-bound-rerun-receipt-20261024-c9daaef628-final/02-final-focused-odp-xlsx.log`
(5,241 bytes, SHA-256
`570318e13a55b8e77d2effdec8e56a6e8698bebb5b0fa45eb8d0ec0a614243fa`). It
ran from `07:11:21Z` to `07:11:35Z`, exited 0, and passed all 7 ODP
unified-write tests and all 6 XLSX rich-cell tests. Resolved test binaries
were
`/home/zhuhe/code/litchi-worktrees/.coord/final-corpus-fidelity-target/debug/deps/odp_unified_write-bed179e84f0f5589`
and
`/home/zhuhe/code/litchi-worktrees/.coord/final-corpus-fidelity-target/debug/deps/xlsx_cells_runs-05dfbb3732ae831b`.

The same temporary directory ended with exactly two zero-byte markers:
`01-final-corpus-crud.start` and `02-final-focused-odp-xlsx.start`; it had no
other files. The final receipt pair is authoritative for bounded provenance;
all earlier receipts remain preserved but superseded. No golden, source,
`NORMALIZING`, or `PROGRESS.md` data changed during these unblessed runs.
