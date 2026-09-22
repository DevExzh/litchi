# 0735 candidate implementation

The candidate changes only the private persisted-record staging owner in
`litchi-ppt`. `replace_persisted_record` and `insert_persisted_record` retain
their existing identity, removal, range, and record-framing checks in the same
order. Once those checks return successfully, the old implementation cloned
the complete `Editor`, inserted the already-owned `Vec<u8>` into the clone's
`staged_storage`, set `changed`, and assigned the clone back. The candidate
performs that same map insertion and flag update directly on the editor.

The direct transition has no typed operation after validation: `BTreeMap::insert`
takes ownership of the caller-provided record and `changed = true` is a scalar
store. Allocation aborts remain process failures under Rust's ordinary
allocation model; they are not converted into a recoverable error or used to
justify a partial mutation contract. Every refusal therefore returns before
either field is touched.

The focused tests cover unknown, removed, unavailable, and malformed record
refusals, including short, truncated, trailing, and over-limit records, with
an exhaustive editor-state comparison. Successful replacement, insertion, and
sequential staging use correctly framed records, compare the direct transition
with a test-only implementation of the former clone transition, and assert
unrelated staged records remain unchanged. A fixture-backed replacement also
compares the finished package bytes produced by both transitions.

No Cargo command, formatter, native Office run, or profiler was run by the
candidate owner. The coordinator must copy this source into the live crate,
run the required gates, and decide retention from the frozen evidence.
