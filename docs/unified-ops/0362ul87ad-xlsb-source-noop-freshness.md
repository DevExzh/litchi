# 0362ul87ad — XLSB source no-op freshness fence

Status: source-only follow-up to 0362ul87ac.  No Cargo, rustdoc, native
qualification, performance measurement or final gate was run for this repair.

## Source authentication

The XLSB source-backed metadata edit now checks the retained source version
before publishing an exact no-op.  A no-op edit therefore cannot publish an
unchanged owner after the underlying source has advanced, while the ordinary
path still uses the source's bounded version check before property planning.

No-op replay on an owner with a retained metadata revision now independently
checks the underlying source version before composing the in-memory revision
patch.  Composing an empty patch alone only authenticates the property
revision and cannot detect an external source change.  Replay against an
unchanged owner remains accepted; a stale owner returns the typed source
change error.

The XLSB source-backed integration suite adds regressions for both stale
no-op commit and stale revised-owner no-op replay.

## Canonical registration

The API glossary and both glossary/debt inventory registrations now name the
canonical `metadata::Edit::{set_custom_props, remove_custom_props}` pair.
`set_custom_props` remains semantic assignment, including its exact no-op for
an equal empty collection; `remove_custom_props` is the explicit graph-removal
intent.

## Source-only checks

`git diff --check` and scoped `rustfmt --edition 2024 --check` pass for the
changed Rust sources.  The coordinator should run the XLSB metadata source
tests and the full API/glossary gate after integration.
