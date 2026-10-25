# 0362ul69h — Formula strict-Clippy test repair

Date: 2026-10-25. Parent: 0362ul69g. Status: source repair complete; the
coordinator still owns the next guarded validation attempt.

The second scoped strict-Clippy lane passed the Formula library and stopped
in its test target on two `clippy::cloned_ref_to_slice_refs` diagnostics. The
new MTEF boundary tests cloned a single `MathNode` only to construct a
one-element slice. Both call sites now use `std::slice::from_ref`, retaining
the original node for the following boundary assertion while avoiding the
unnecessary allocation and clone.

The external f5 receipt is:

    /home/zhuhe/code/litchi-worktrees/.coord/formula-raw-policy-328fe19328/state/08-formula-affected-clippy-rerun-attempt-affected-clippy-rerun.log
    1179 bytes, SHA-256 4309bc37abb6d0c31372619439c2a58945f8c3ad5414100c6c64dcc332a6d6d7

The command exited 101 at the first test target, with no resource or
provenance guard failure. Exact 4100 and original tool locks remain byte
checked. No lock generation, native, rustdoc, broad ops, or performance
command is part of this repair.
