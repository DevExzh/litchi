# 0362ul59 — bounded font dependency assembly

This preparation combines the source-reviewed Allsorts/zlib-rs adaptation
`21fa93a194d84c2444717215782e07c147c8d3e8` with the first-party font consumer
`c8beeed05d2ece1abf81f61478a9f72e37e8386e`. It is an isolated assembly on
`uo2/shared-font-raw`, not final integration or validation acceptance.

The Allsorts fork remains version 0.17.0. It admits cumulative requested
working allocation bytes and independently admits exact output bytes. Its
zlib-rs 0.6.8 bridge returns the same compiled allocation size used by the
inflate allocator; this includes its alignment and window padding. Independent
source review confirms the workspace selects that backend for both Allsorts
and flate2. The first-party scope begins before glyph mapping, parsing and
provider construction and continues through subsetting. The output ceiling
is the smaller of the per-program ceiling and remaining retained-program
budget. Error classification and complete consumer behavior still require
review and actual tests.

Upstream provenance and license files are retained. The imported zlib-rs
license has one trailing space, and its PDF fixture is a binary document with
text-like regions. Exact-path Git attributes preserve those upstream bytes:
only the license's end-of-line whitespace check is disabled, and only that PDF
is declared binary. Source-code whitespace checks remain unchanged. This is
not a source, API, fidelity or performance exception.

No prior registry lock or performance receipt is relabeled. Root will retain
all old lockfiles, prepare new locks deliberately, and verify that existing
registry name/version/checksum identities are preserved except for the two
explicit path-source substitutions. Standalone consumer manifests and locks
must select the same reviewed sources. Metadata preparation is separate from
bounded compile/test validation; no native or performance result is claimed.
