# 0362ul86b — scoped ZIP alias formatter repair

- Date: 2026-10-25.
- Source: `a84f5761388ab5f891d4592e7feea8c30da076ab` plus the scoped
  formatter repair in this commit.
- Status: source-ready; the coordinator must review the fresh validation
  manifest before starting its eight rows.

The first guarded validation of the alias-closure source stopped at its
workspace-wide formatting row. Its preserved log reported formatting drift in
both vendor and workspace paths, including ZIP, ODF, OPC and ops consumers;
that row is intentionally retained as a failed result and is not relabeled as
a package-only failure. Active lanes make a workspace-wide formatter repair
unsafe here, so this record limits the repair to the nineteen Rust paths
changed by `0362ul86` and `0362ul86a`.

Each of those paths was passed explicitly to the pinned Rust formatter with
edition 2024 and `skip_children=true`, preventing module traversal into
unrelated files. Eleven paths required whitespace/layout changes; the other
eight were already normalized. No vendor file or Rust path outside the
0362ul86/86a change set was edited. The full workspace formatting pass,
including current vendor and other-lane paths, remains a coordinator-owned
requirement before the full gate.

## Source checks

The scoped formatter command was:

```text
rustfmt --edition 2024 --config skip_children=true <the nineteen changed Rust paths>
```

The same explicit path list passes the corresponding `rustfmt --check`
invocation and `git diff --check`. No Cargo, rustdoc, native, corpus or
performance command was run for this repair. The next validation capture is
prepared with a scoped formatter check as row zero and the seven previously
approved focused ZIP/ODF/OPC/ops/glossary rows; it must not start until the
coordinator reviews its new source, lock and runner manifest.
