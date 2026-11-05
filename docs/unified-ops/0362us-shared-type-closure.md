# 0362us — Close shared-type duplicates in the final sweep

Record 0351 §2.2 requires core Commit aliases for exact triads and shared
Diagnostics for changed-only summaries. The source audit found remaining
copies in PPT/PPTX, XLS/XLSB, Pages and ODF. The reviewed migrations and their
consumer changes are recorded in [0362usa](0362usa-ppt-shared-types.md),
[0362usb](0362usb-xls-shared-types.md),
[0362usc](0362usc-pages-shared-types.md), and
[0362usd](0362usd-odf-shared-types.md). These canonical aliases are the shape
prescribed by the contract; they do not retain obsolete API names.

PPTX opened commits retain their additional snapshot state. OOXML properties
and other genuinely richer diagnostics retain their custom types. Pages
metadata replay now returns its Commit triad, with changed/no-op/replay and
conflict consumers migrated. ODI flat semantic planning and source provenance
remain available on its Patch after removal of derived Commit helpers.

The glossary gate now pins 30 nominated shared public type paths to the
nominal core Commit or Diagnostics identity. It resolves aliases/re-exports
by rustdoc identity and fails closed on missing, unrelated, ambiguous or
cyclic targets. Richer custom types are not nominated. The Pages replay rule
and ODI removed-helper bans remain present after conflict resolution. The
composed glossary and inventory tool suites pass 54 tests on `ad52437f5e`;
actual final-source rustdoc captures and the complete gate remain required.

## Preserve the redaction boundary

Adopting core Commit can make a previously non-Debug format commit printable.
For example, flat ODI Image and Patch already derive Debug over retained
state. Core Commit's derived Debug would expose those payloads through the
new alias. Its Debug implementation now renders only the family's existing
content-free diagnostics, never the package or patch. A regression uses
payloads whose Debug panics to prove they are not formatted. This is a narrow
obligation of the changed API; the comprehensive item-22 audit remains deferred.
The commit data, accessors, equality, cloning and publication behavior do not
change, and normal editing performs no additional work.

## Remaining qualification

ODM now stores its changed flag during commit instead of computing exact-byte
equality lazily in diagnostics(). Record 0362usd identifies that moved
comparison for native qualification. No performance exception is accepted.
The owner-replay sweep also returned ODG Package::apply_joined plan publication
for migration into an Edit staging verb (0362ut); that independent correction,
all 68 final inventory captures/dispositions, the complete gate, corpus and
native qualification remain required. This record does not claim final
workspace acceptance.

## Composed consumer verification

The six core edit tests pass on `5dae0abf78`, including the non-formatting
payload regression. The complete workspace Clippy run then finds three ODF
benchmark result-discard sites and three DOCX error-expectation assertions.
Commit-only benchmark measurements now explicitly discard their black-boxed
results at the same statement boundary. The DOCX assertions use `expect_err`
now that their successful Commit payload has content-free Debug. Assertion
messages and failure conditions are unchanged.
[The v24 receipt](results/0362up/core-clippy-shared-v24.json) retains both full
commands, tested source and external log hashes. Final combined verification
and native qualification remain required.

The complete composed workspace rerun on
`8225b82dde7bac272e9e77789a78751ae07aaa2a` passes:

```text
cargo clippy --offline --locked --workspace --all-targets --all-features --keep-going --message-format=json-render-diagnostics -- -D warnings
```

The run completed in 199.602 seconds with the one-job 20G memory cap and an
external target of 4,397,247,763 bytes. The 2,403,707-byte raw log remains
external at `(deleted-coord-artifact)/check-0.log`,
SHA-256 `bbd34401c4ee66836da9fe8c66066c3f4308192f64ac81bdc08bbec5830a40c5`.
The reduced receipt is `results/0362up/clippy-composed-v25.json`; source and
lock identity were rechecked after completion. The glossary and inventory
tool suites also pass 60 tests with the reviewed replay proposal helper.
This closes the composed consumer lint returns. ODG joined composition is
still returned for ordered-state corrections; the complete final gate,
inventory and native qualification remain required.
