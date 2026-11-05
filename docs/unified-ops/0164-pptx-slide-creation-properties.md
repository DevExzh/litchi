# PPTX slide creation properties

Unified slide insertion now accepts optional `name` and `hidden` properties
alongside the required exact, nonempty `layout`. The static property registry
supplies the same coercion and clear rules to add and set. Names preserve XML
character semantics, including line endings; null clears the authored name.
Hidden requires boolean coercion and cannot be cleared. Omitted values retain
native creation defaults. Read views may omit the default visible value.

For example:

```rust
let added = file.add(
    "/",
    NewNode::new("slide")
        .prop("layout", "Blank")
        .prop("name", "Quarterly roadmap")
        .prop("hidden", true),
)?;
```

All properties are admitted before native staging. The adapter rejects repeated
case-insensitive keys, invalid XML characters and invalid values, and accounts
for aggregate layout/name/boolean staging bytes. It inserts the slide and applies
both metadata values in one native edit before a single commit. The native owner
retains authoring, relationship, source-preservation and patch-proof ownership.

Native-owner acquisition follows facade property and placement admission.
The eager slide route acquires the owner directly rather than constructing
and discarding an admission-only edit before the real edit. The actual edit
still performs strict graph and encryption admission. The shape insertion route
is unchanged.

## Verification

The final combined DOCX/XLSX/PPTX/encryption suite passed 1,125 tests and
doctests, including the new atomic/best-effort batch cases. Ten native
slide-property tests passed, including serialized topology replay and inverse
after reopening. Facade all-target Clippy, the affected native test target, and
all harness targets passed with warnings denied. Warning-denied facade rustdoc,
minimal-feature compilation, scoped formatting and diff checks passed. The
crate-boundary checker passed for 66 packages and 268 internal dependency
declarations. The regenerated OfficeCLI delta passes its deterministic check.

Review found no blocking API or replay issue. Its admission-order and batch
coverage follow-ups were incorporated before the final facade run.

## Paired measurement

[Results](0164-pptx-slide-creation-properties-results.json) retain every sample,
corpus hash, source digest and binary provenance. The native leg inserts at
position one, assigns both metadata properties, and commits once. The unified
leg adds an equivalent prepared node. Each source deck has one existing slide
with the indicated number of shapes; the new slide is empty.

Three ABBA rounds provide six samples per leg, pinned to CPU 0 after validation
jobs finished. The failure matrix is empty. All source cases have exact-byte
equivalence. All eager cases have identical decoded ZIP members, while archive
bytes differ; exact eager archive equivalence is not claimed.

| Shapes in existing slide | Eager unified/native | Source unified/native |
|---:|---:|---:|
| 1 | 0.996 | 1.001 |
| 10 | 0.996 | 1.004 |
| 100 | 0.995 | 1.009 |
| 1,000 | 0.999 | 1.002 |
| 10,000 | 1.000 | 0.997 |

One-shape native/unified medians were 2.307/2.299 milliseconds for eager owners
and 1.800/1.803 milliseconds for source owners. Ratios below one are measurement
variation. The benchmark reuses warmed owners and excludes opening,
serialization, allocation/RSS measurement, and real-world corpus acceptance.
It measures one insertion, not a 10,000-operation batch.

## Cleanup

Temporary build targets, raw logs and measurement scratch were removed after
retaining every sample and the binary name, size, SHA-256 and regeneration
command. No worktrees or borrowed fixtures were created. The inherited workspace
lockfile was preserved and the rewritten base remains an ancestor.

This milestone does not close broader all-format CRUD, Python, fuzz or global
performance acceptance.
