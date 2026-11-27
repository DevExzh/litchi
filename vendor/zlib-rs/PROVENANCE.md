# vendored zlib-rs 0.6.8

This directory is a pinned source snapshot of the crates.io `zlib-rs` 0.6.8
package, copied from the approved local registry source at revision
`0254e07204884bfc902d6ebb68c8b703d0ea49cf` (checksum
`b268e58e7c693d7c271f93ffc4ba3b380412554231c85bf61ca7af91042a4112`). The
upstream repository is
`https://github.com/trifectatechfoundation/zlib-rs`; the package is licensed
under the Zlib license, retained in `LICENSE`.

The only source adaptation adds the public `inflate_workspace_size` query. It
returns the exact compiled backing-store request already used by zlib-rs's
inflate allocator. The workspace `[patch.crates-io]` entry selects this path
for the repository's `flate2_zlib-rs` build; it intentionally does not rewrite
the ignored `Cargo.lock`. The integration owner must record that source
identity change when regenerating the lock.

The imported upstream bytes were not whitespace-normalized. `git diff --check`
therefore reports the original trailing U+0020 on `LICENSE:1`, and the mixed
binary/text PDF fixture at
`src/deflate/test-data/paper-100k.pdf` on lines 291--313, 356--357, 389--390,
404, 417--422, 451, and 501. Those diagnostics are pre-existing registry
content; changing them would alter the upstream fixture or license bytes.
