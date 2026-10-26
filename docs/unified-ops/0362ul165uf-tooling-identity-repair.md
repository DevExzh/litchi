# 0362ul165uf — tooling import and reviewed-adapter identity repair

The current remaining-check diagnostic exposed two source-level tooling
defects in its combined Python unit row.

`tools.ppt_lazy_native_qualification` preferred top-level imports whenever an
earlier test put `tools/` on `sys.path`. Its builder then used the package
qualified `RunFailure` while the PPT module caught a distinct top-level class,
so an intentional dirty-worktree refusal escaped `assertRaises`. Package
qualified imports are now preferred, with the top-level form retained only for
direct execution from the tools directory.

The DrawingML reviewed-adapter table still pinned the pre-DOCX/ODF
`tools/perf-baseline/Cargo.toml` identity. The checked-in shared consumer now
has SHA-256
`f18a17d5ca2bb0848e73c53ace39550ba38dbd7a75fab76c5a0b0af4f435984a`, so the
candidate identity is updated while detached historical RTF manifest pins
remain unchanged.

Source-only validation on the leaf passes `git diff --check`, the direct and
combined PPT lazy unit tests, the DrawingML reviewed-hash test, the legacy and
DrawingML protocol test modules, and the repository size/binary guards. No
Cargo, rustc, native probe, or timing command was run.

The diagnostic row's separate `rustc` failure is an orchestration environment
issue: its systemd launch has no `/home/zhuhe/.cargo/bin` on `PATH`. The gate's
Cargo wrapper is found through its own explicit path, while the glossary
fixture invokes `rustc` directly. A future gate launch must bind the reviewed
toolchain bin directory into `PATH` (or provide an absolute reviewed `RUSTC`)
before rerunning the row; this source leaf does not weaken that requirement.
