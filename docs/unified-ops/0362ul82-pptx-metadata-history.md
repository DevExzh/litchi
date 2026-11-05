# 0362ul82 — PPTX metadata history policies

- Date: 2026-10-25.
- Status: source-ready for independent review; native Cargo validation remains
  with the coordinator's isolated build slot.
- Base: `fda2b5d223452548d92201ec4945af9e4d5835a5`
  (`uo2/w4-xlsx` integration base).

This slice closes the explicit raw-policy contract for the five remaining
PPTX metadata-history families: Changes Information, Revision Information,
slide synchronization, WebVTT caption tracks, and inert slide-show events.
Each public raw parser/serializer and package loader takes the family's
typed `Limits`; there is no default-only twin, deprecated alias, or
`*_with_limits` entry point. Each edit has an `EditOptions::limits` profile;
the profile is selected from `OpenOptions`, retained by the package and
projection paths, copied into source snapshots and patches, and used for
changed-payload replay admission on the recipient package.

The raw consumers migrated in the same change include the package/presentation
facades, slide projections, modern-comment command reader, package-signature
tests, and the metadata unit/integration tests. Changes, Revision and slide
sync preserve their MCE and format validation behavior, including the fixed
slide-synchronization four-digit year and calendar checks. Track discovery
retains opaque slide XML and validates every changed track island under the
recipient profile; WebVTT targets are never fetched. Event discovery remains
inert, and opaque extension payloads are not interpreted. Exact source bytes,
relationship topology, no-op behavior, and the existing staged commit/patch
triad remain intact. The Changes, Revision, slide-sync and WebVTT writers
count their complete output with checked arithmetic before one exact,
fallible reservation; event rewrites likewise calculate the final replacement
size before copying source bytes once. Emission length is asserted before the
independent output is parsed under the caller's selected profile.

The replay hook validates only changed payloads. A patch whose recipient
content matches the recorded base but comes from an independent lineage is
accepted under the recipient's selected family profile; an exact no-op keeps
the recipient unchanged. Physical source limits still bound every XML
payload before the family semantic check.

Source checks performed here:

- `rustfmt --edition 2024 --check` over all changed Rust files;
- `git diff --check`;
- focused `rg` sweeps for every migrated raw call site and for removed
  default/alias spellings;
- JSON syntax validation of the glossary policy.

Cargo tests, rustdoc, clippy, native fixture processing and ABBA measurements
remain intentionally outstanding for the coordinator's independent build
slot. No binary, fixture, or large performance evidence is committed.
