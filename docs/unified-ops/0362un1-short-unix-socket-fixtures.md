# 0362un1 — short Unix socket fixture roots

The final facade gate's `special_nodes_are_refused_without_blocking` test
could panic before opening the input when the gate's `TMPDIR` made the socket
path exceed Unix `SUN_LEN`.  The two iWork archive tests that create sockets
had the same latent dependency through their custom temporary roots.

Only those three fixture parents now use
`tempfile::tempdir_in("/tmp")`, with a comment documenting the Unix socket
path bound.  `TempDir` retains unique names and RAII cleanup; no process-wide
environment, current directory, or production code changed.  All other test
directories continue using their existing helpers.

## Checks

The tested source was `dd78c02a23a202b6fa25562872a58113c4b20cb8`
(`Merge reviewed DOCX encrypted edit admission repair`).  The ignored
`Cargo.lock` was copied from the pinned external receipt
`/home/zhuhe/code/litchi-worktrees/.coord/docx-encrypted-locked-receipts/Cargo.lock.before-6b05`
(139,592 bytes,
`6d7895e6056c931fc19780ff4d0a031230e2ac79b9b30a9060bc8b6fad6747da`) and
all commands used `--offline --locked`, one Cargo build job, one test thread,
and `LITCHI_MEM_MAX=20G`.  The deliberately long `TMPDIR` was 167 characters;
the target was the existing
`/home/zhuhe/code/litchi-worktrees/.coord/docx-encrypted-target` and measured
2,202,031,473 bytes after the checks.

* `cargo test --offline --locked -p litchi --features iwork --test iwork_file special_nodes_are_refused_without_blocking -- --nocapture`: 1 passed, exit 0, max RSS 3,099,696 KB.
* `cargo test --offline --locked -p litchi-iwa-archive --lib pages_metadata_profile_rejects_selected_symlink_and_directory -- --nocapture`: 1 passed, exit 0, max RSS 3,106,804 KB.
* `cargo test --offline --locked -p litchi-iwa-archive --lib rejects_special_nodes_in_loose_index -- --nocapture`: 1 passed, exit 0, max RSS 73,196 KB.
* `rustfmt --edition 2024 --check crates/litchi/tests/iwork_file.rs crates/litchi-iwa-archive/src/directory.rs`: passed.
* `git diff --check`: passed.

This is a fixture portability correction only; no performance qualification is
claimed.  The earlier DOCX locked-input logs remain outside Git and were not
renamed or overwritten.

## External receipts

| receipt | bytes | SHA-256 |
| --- | ---: | --- |
| `/home/zhuhe/code/litchi-worktrees/.coord/short-unix-socket-fixtures-receipts/iwork-file-special-nodes-long-tmpdir.log` | 5736 | `c42583b811bdb2fc3fbcb0c51cabacc94bb866d7fba1ac592205b97e6c663996` |
| `/home/zhuhe/code/litchi-worktrees/.coord/short-unix-socket-fixtures-receipts/iwork-file-special-nodes-long-tmpdir.time` | 55 | `0dc6343d5e13e5f07a6b05938ae6e24fe99297bbaa112cae564a9ca62c72cf27` |
| `/home/zhuhe/code/litchi-worktrees/.coord/short-unix-socket-fixtures-receipts/iwa-pages-metadata-socket-long-tmpdir.log` | 1896 | `5dfdbc17635c7b427b4f9d50ce684b2d158a61f5d94515d2cdee7f37ddc960d7` |
| `/home/zhuhe/code/litchi-worktrees/.coord/short-unix-socket-fixtures-receipts/iwa-pages-metadata-socket-long-tmpdir.time` | 54 | `3850f523394fed80072e6e28005ab7bdee38d63b70bfb1a6de5100fd83cea440` |
| `/home/zhuhe/code/litchi-worktrees/.coord/short-unix-socket-fixtures-receipts/iwa-loose-special-node-long-tmpdir.log` | 943 | `827cdeb7d9a45d21c5f4c561dea69036a2c223e082f8068b44cd4bb24012b4bb` |
| `/home/zhuhe/code/litchi-worktrees/.coord/short-unix-socket-fixtures-receipts/iwa-loose-special-node-long-tmpdir.time` | 51 | `820f682618d62100f252ffd168b9b907135dc9806d923d57bf66432ff84bafce` |
