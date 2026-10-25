# 0362ul157a — PPT owned record view correction

- Date: 2026-10-25.
- Status: source-ready; Cargo/rustdoc validation remains with the coordinator.
- Base: `bbb629dbcfd1fa8217147c9c6eae275ac9ced04f` (`0362ul157`).

The canonical `litchi_ppt::RecordParser::find_records_owned()` method now has
no argument. The former deprecated method accepted a `RecordType` value but
never filtered by it; retaining that unused parameter would advertise a
capability the method did not provide. The replacement therefore states its
actual all-record behavior while preserving the owned-clone result and the
borrowed `find_records_ref()` alternative.

The focused test builds a small `Document` → `Slide` → `TextCharsAtom` tree
with a sibling `TextHeaderAtom`, checks the exact flattened order and payload
bytes, then drops the parser before inspecting the returned owned records.
This proves nested traversal, fidelity and ownership without a fixture or
large allocation.

Source checks performed:

- `rustfmt --edition 2024 --check` on the changed PPT parser;
- `git diff --check`;
- focused search confirming the canonical method has no argument and no old
  public spelling remains;
- the staged AGENTS.md size and binary checks.

Cargo tests, rustdoc, clippy, native fixtures and performance measurements are
intentionally outstanding for the coordinator's independent build slot. No
binary, fixture, or large performance evidence is committed.
