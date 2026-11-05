# 0362ul69d — Formula raw-policy compile-fix return

Date: 2026-10-25. Parent: 0362ul69c. Status: source repair complete; the
coordinator still owns the post-fix Formula and consumer validation.

## First actual validation blocker

After the sparse checkout was hydrated with the complete tracked `crates`
tree, the approved f5 runner reached Cargo on source `328fe19328`. The first
attempt before hydration had stopped at manifest loading because a workspace
member was absent; that is a hydration failure, not a source result. The first
hydrated Formula all-target attempt then exposed a real compile failure:
the new preflight matches grouped `MathNode::Row(Vec<_>)` with boxed wrapper
variants. Rust correctly rejected the mixed bindings with 29 `E0308` errors,
and the MTEF node writer also retained one now-unused `LimitKind` import.

The external guarded receipt is:

    (deleted-coord-artifact)/00-formula-all-targets-tests-attempt-formula-all-targets-v2.log
    23637 bytes, SHA-256 46729f580b1a8651b5cc148369cd62c41e8e9cc9dae4d733bbf53c39affe6fb6

The runner was the approved f5 `run_safe_gate.py`, SHA-256
`1904177040fc7c523582433028d9bba56efc30fe0d9d96beb554a2fab343691e`, with
offline/locked Cargo, one build and test thread, debug and incremental output
disabled, and the 19/20 GiB target guards. The command exited 101 before any
test executed; the target was 71,292,179 bytes and no resource or provenance
guard failed.

## Source repair

The MTEF depth walk and output preflight now give `Row` its native `Vec` arm
and handle the boxed `Phantom`, `Degree`, `Base`, `Argument`, `Numerator`,
`Denominator`, `Integrand`, `LowerLimit`, `UpperLimit`, and `Limit` arms in a
separate pattern. The OMML output preflight uses the same typed split. This is
only a representation-level compile repair: it preserves the exact traversal,
depth accounting, work charging, string admission, and emitted bytes intended
by 0362ul69b/c. The unused MTEF import is removed. No public API, limit value,
wire spelling, or semantic AST behavior changes.

Source checks after the repair were `rustfmt --edition 2024` on the three
affected writer files and `git diff --check`. The guarded Formula suite,
affected DOC/DOCX/ops/facade tests, and strict Clippy must be rerun on this
clean repair commit before any validation claim is made.

The exact 4100 root lock and original tool locks remain hydrated and byte
checked; no lock generation is permitted. No native, rustdoc, broad ops, or
performance command is part of this return.
