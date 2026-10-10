# 0069 — Property input and registry foundation

## Scope

This checkpoint advances M2 W1/W2 from record 0015. It adds property input
values, strict registry admission and structured error attribution. The nine
DOCX readable property tables now derive from registry entries; their existing
order, names, scalar types and read-only access flags remain the authority for
node views and capability output. No File write verb is advertised by this
checkpoint: real setters, staging and commit publication are still required.

`litchi_ops` and the `litchi` facade expose `Key`, `PropValue`, `Props`,
`IntoProps` and `props!`. The macro returns a stack array and supports mixed
scalar values. Array/vector and borrowed input routes preserve assignment
order. String assignments split at the first `=`, retaining the rest of the
value; malformed string lists fail before an iterator is returned. Ordered
lists retain duplicate keys rather than discarding them. Explicit collection
into `Props` owns a vector; borrowed inputs do not require that collection.

Wide integers retain their exact integer representation until coercion. The
registry checks typed values in the destination property's context, so a bare
integer font size means points, just as its textual spelling does. Conversion
uses a fixed stack scratch buffer and a callback that cannot retain references
to that buffer. Non-finite numbers are refused even for text destinations, and
scratch overflow is a typed refusal. Wide 128-bit integers serialize as decimal
strings to avoid JSON consumers rounding them through floating point;
non-finite values cannot silently serialize as null.

## Registry and errors

`registry::Table` retains declaration order plus compile-time sorted canonical
and alias lookup indices. Lookup folds ASCII without allocating. Construction
rejects duplicate folded names, ambiguous/shadowing aliases, invalid targets,
and constraints with incompatible types or reversed bounds. The generated read
projection derives its set/add flags from the same entry that owns the typed
setter enum. Existing DOCX entries are read-only until handlers land.

Write admission distinguishes unknown properties, read-only properties,
operation-specific access and declared leaf gaps. It validates coercion,
clearability and domain constraints before invoking a setter callback. Null
clearing is explicit. Error messages omit rejected values; suggestions use the
known vocabulary. Registry aliases and typed setter variants are tested on a
small executable table; no fake setter is registered for an actual format.

`Error` gains item/property attribution without losing its classification or
shared typed source. Clone preserves that context. Debug output excludes
messages, suggestions, authored property keys and values; explicit accessors
remain available to callers.

## Verification

- DOCX-enabled ops tests, integration goldens, allocation checks and doctests:
  529 passed. No capability or node golden was re-blessed.
- Minimal ops tests/doctests: 359 passed.
- Minimal and DOCX all-target Clippy passed with warnings denied.
- Minimal ops rustdoc passed with warnings denied; conditional opener links were
  adjusted so documentation also builds without format features.
- The DOCX-only `litchi` facade compiled with the new re-exports, and the
  workspace library compilation passed.
- The counting-allocator regression observed zero allocations for borrowed
  mixed-value macro construction/iteration and borrowed `key=value` iteration.

This allocation evidence covers input construction, not an end-to-end write.
The existing parser's allocation fallback for unusually long grouped numeric
text remains; no universal zero-allocation coercion claim is made. Setter
liveness, batch performance, mutation atomicity and all-format parity require
subsequent implementation and measurements.

## Remaining work

Connect concrete typed DOCX setters to these declarations and the existing leaf
Edit/Commit owners; implement File write verbs, sequential atomic batches,
publication, the other format handlers and Python bindings. This foundation
does not close M2 or the complete goal.

## Commit and cleanup

The rewritten base remains an ancestor of HEAD. Both required staged evidence
checks are silent. `git add -A` still fails because the linked worktree's
`index.lock` cannot be created on the read-only Git metadata filesystem; no
commit was created. The changed/untracked artifact audit found no prohibited evidence or compiled
executables. Removed 10.0 GiB of generated Cargo output, generated untracked
lockfiles and checkpoint scratch logs. No temporary worktree was created.
