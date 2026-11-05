# 0362ul94v — OPC publication compiler repair

The final OOXML four-row packet stopped at the first Clippy invocation on
source `b73ee18a69be4122b9097588f146e2c3b3b8b713`. The captured log was
`ooxml-combined-0362ul94-f5-final-b73ee18a/01-clippy-leaves-attempt-actual-f5-01.log`,
24,829 bytes, SHA-256
`4f71abc505886f58e798b469473cbb6046ed605c92f28fdaf394b2f44a29b867`.

The compiler stopped with fifteen `E0433` diagnostics in
`crates/litchi-opc/src/package.rs`: the new authored-publication code used
`BlobPart` and `SourceLimitKind` without importing them. The repair adds the
canonical module imports alongside the existing `Part` and `SourceLimits`
imports. No behavior, feature gate, limit policy, allocation path, or public
API changed.

This source-only repair was prepared in a fresh sparse worktree from the
exact failed source commit. Rustfmt/static checks and repository commit guards
were run; no Cargo, compiler, build, rustdoc, metadata, native, or performance
command was run.
