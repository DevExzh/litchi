# 0362ul69g — Formula strict-Clippy repair

Date: 2026-10-25. Parent: 0362ul69f. Status: source repair complete; the
coordinator still owns the next guarded validation attempt.

The first scoped strict-Clippy lane stopped in `litchi-formula` before any
consumer crate ran. `mtef/writer/mod.rs` charged a Unicode symbol with an
`if ch.is_ascii()` expression whose two branches were identical. The
workspace's normal lint configuration permits this existing warning, but the
required `-D warnings` consumer lane correctly rejects it as
`clippy::if_same_then_else`. The repair removes the redundant conditional and
retains the same `CHAR` charge and emitted bytes.

The external f5 receipt is:

    (deleted-coord-artifact)/07-formula-affected-clippy-attempt-affected-clippy.log
    829 bytes, SHA-256 75c6fd4c291eb7f8a9d1048cffe09f6bf7cb5f06e066fd734d1d5c9a78695f87

The command exited 101 at the first package, with no resource or provenance
guard failure. Exact 4100 and original tool locks remain byte-checked. No lock
generation, native, rustdoc, broad ops, or performance command is part of
this repair.
