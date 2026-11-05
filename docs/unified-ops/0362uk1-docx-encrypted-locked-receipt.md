# 0362uk1 — locked DOCX encrypted-admission receipt

This receipt supplements `0362uk` for source
`6b05b8063cd8f2e248caee7ec0858896de7f3da3` (`fix(ops): refuse encrypted
DOCX edits before closure`).  The earlier receipt incorrectly said that the
checks omitted `--locked` because `Cargo.lock` is ignored.  The lock is an
external, pinned build input even though it is not tracked in this checkout;
the reruns below use `--offline --locked`.

The existing ignored `Cargo.lock` was 139,592 bytes with SHA-256
`6d7895e6056c931fc19780ff4d0a031230e2ac79b9b30a9060bc8b6fad6747da` before
the reruns.  It was copied to
`/home/zhuhe/code/litchi-worktrees/.coord/docx-encrypted-locked-receipts/Cargo.lock.before-6b05`.
After both commands, `cmp` was identical, the SHA-256 was unchanged, and the
lock mtime remained `2026-10-24 06:05:52.555870342 +0000`; neither command
mutated the lock.

Both commands used the existing external target
`/home/zhuhe/code/litchi-worktrees/.coord/docx-encrypted-target`, one Cargo
build job, one test thread, `LITCHI_MEM_MAX=20G`, `CARGO_INCREMENTAL=0`, and
`CARGO_NET_OFFLINE=true`.  The target was 1,058,588,195 bytes after the
reruns.  The source worktree was clean at the tested commit.

## Checks

The exact measured invocations were:

```text
source /home/zhuhe/code/litchi-worktrees/.coord/final-acceptance-2026-10-24/env.sh
export CARGO_TARGET_DIR=/home/zhuhe/code/litchi-worktrees/.coord/docx-encrypted-target
export TMPDIR=/home/zhuhe/code/litchi-worktrees/.coord/docx-encrypted-tmp
/usr/bin/time -f 'time_status=%x elapsed_seconds=%e maxrss_kb=%M' -o <receipt>.time cargo test --offline --locked -p litchi-ops --features 'docx encryption' --test docx_transactions -- --nocapture
/usr/bin/time -f 'time_status=%x elapsed_seconds=%e maxrss_kb=%M' -o <receipt>.time cargo test --offline --locked -p litchi-ops --features docx --test docx_transactions
```

The encrypted suite passed 22 tests, exit 0, elapsed 2.53 seconds, and max
RSS 73,160 KB.  The no-encryption suite passed 20 tests, exit 0, elapsed
2.49 seconds, and max RSS 73,064 KB.  These are correctness receipts only;
no performance qualification is claimed.

The command inputs were identified as follows:

| input | bytes | SHA-256 |
| --- | ---: | --- |
| `/home/zhuhe/code/litchi-worktrees/.coord/final-acceptance-2026-10-24/env.sh` | 503 | `c69f8b9e9c5f8bed6303fe75a3d263bd6f33e78dde619b116d2af9141db17a46` |
| `/home/zhuhe/code/litchi-worktrees/.coord/bin/cargo` | 520 | `1b22a69fd625d57565041bf6c634d8ebd1a82be8b3d0006390472c86324230d8` |
| `.cargo/config.toml` | 470 | `50b044c83828f2c448ee59c362afaa79de2dd47fbd390f0cd2888391f1d1119d` |
| Rust 1.95.0 `rustc` commit `59807616e` | — | — |
| Cargo 1.95.0 commit `f2d3ce0bd` | — | — |

## External receipts

These files remain outside Git:

| receipt | bytes | SHA-256 |
| --- | ---: | --- |
| `/home/zhuhe/code/litchi-worktrees/.coord/docx-encrypted-locked-receipts/identity-before.log` | 2592 | `882fba76bc1b12dc9cb03f9e180f69a338a50643223119912d7e69de40755e58` |
| `/home/zhuhe/code/litchi-worktrees/.coord/docx-encrypted-locked-receipts/docx-transactions-encryption-locked.log` | 2797 | `b8614476fb797ddbb72f0b396ff60fec064ef12a68461d44e20ee53d8b08f091` |
| `/home/zhuhe/code/litchi-worktrees/.coord/docx-encrypted-locked-receipts/docx-transactions-encryption-locked.time` | 51 | `9905745c09071a8d30d70138febf465227f1f6ce9d4b1426566e9a89cbff5cab` |
| `/home/zhuhe/code/litchi-worktrees/.coord/docx-encrypted-locked-receipts/docx-transactions-no-encryption-locked.log` | 2207 | `978de9d0708254f6d7ad054430d5710793c27f8e61430270f788aae0aee2bd4c` |
| `/home/zhuhe/code/litchi-worktrees/.coord/docx-encrypted-locked-receipts/docx-transactions-no-encryption-locked.time` | 51 | `ecc0c473bcc9ca7c5baccc5b89d4269cc2d0271aeab8a72bb7462024cf43f958` |
| `/home/zhuhe/code/litchi-worktrees/.coord/docx-encrypted-locked-receipts/Cargo.lock.before-6b05` | 139592 | `6d7895e6056c931fc19780ff4d0a031230e2ac79b9b30a9060bc8b6fad6747da` |

The initial unlocked diagnostic logs from `0362uk` were retained unchanged;
the locked receipts above use distinct names.
