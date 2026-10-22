# 0362hc — ODF semantic options through prepared and flat owners

- Date: 2026-10-22.
- Parent: W8 Part A candidate `4d464d60b7`.
- Status: source candidate; final integrated qualification remains required.

`File` applied selected ODF source limits while preparing an archive, but
several adapters then constructed the format owner with default semantic
limits. Flat adapters also discarded the selected ODF options. Consequently,
the same content ceiling could reject a path open and admit a byte open, or
be ignored for an ODF object family.

Prepared-package constructors now accept `OpenOptions` in ODT, ODS, ODP, OTH,
ODG, ODC, ODI, ODB and ODF formula, following the existing ODM constructor.
The default constructors delegate to them. The prepared source keeps the
source policy already applied to it; the selected semantic policy governs
the format projection. Every File adapter migrates in the same change. The
constructors transfer the prepared archive, reuse its index and source
allocation, and add no owner fields or additional archive index.

The common File ODF options are available when only an object format is
enabled. Packaged object/text-family path preparation also uses the selected
source policy and preserves the typed owner failure if preparation refuses.
Flat ODT, ODS and ODP use their existing option-bearing native constructors;
flat ODP publication retains the flat owner's options. Flat ODG applies the
physical input ceiling to the original XML and semantic limits to its
packaged projection. The generated internal ZIP is not charged as a second
physical input.

Flat ODT now distinguishes semantic `ContentBytes` from physical `InputBytes`
when their minimum bounds the reader. Byte and reader failures retain their
actual observed and maximum counts. This is an error-payload correction;
both kinds continue to classify as `limit_exceeded`.

Boundary tests cover all ten packaged ODF formats from bytes and paths,
exact source publication at the content ceiling, one-under typed failures,
all four supported flat families, native flat ODT byte/reader boundaries,
and prepared ODT index identity. The
[validation receipt](results/0362hc/validation.json) records exact source and
log identities, successful checks, and earlier unsuccessful attempts.

The affected library suites pass 2,752 tests, including all 965 File tests.
The full ODS suite passes 1,768 tests with none ignored, and all 299 affected
doctests pass. All-target/all-feature Clippy passes for the ten affected
crates; ODT-, ODS-, ODP-, ODG-, ODC- and formula-only File Clippy, the
no-format build, facade/CLI/Python all-target/all-feature compilation,
formatting, ODT/ODP/File glossary profiles, crate boundaries and tracked
corpus disposition checks also pass. The ODS and documentation totals overlap
the broader library checks and are reported separately.

This record closes the prepared/flat ingress propagation defect. It does
not establish native performance acceptance or complete the wider raw API
audit. In particular, the existing public `odf_common::flat::Parts`
split/merge helpers still need the required raw-policy review. The final
File workload, W4/W4b integration, corpus dispositions, whole-workspace
inventory and gate remain required.
