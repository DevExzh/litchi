# 0362ul47 — imgconv raw record limits

This leaf migrates the variable-size PICT, EMF, and WMF raw parser boundaries
to the existing `litchi_imgconv::Limits` profile.  The profile is now a
required argument at every public variable parser/record entry owned here;
fixed-width headers and scalar zerocopy helpers remain unchanged.

## Canonical API

* `PictParser::new(data, limits)` replaces the public `new_bounded` twin.
* `EmfParser::{new,new_compatible,from_owned,from_owned_compatible,new_lazy}`
  require `&Limits` and retain the selected profile for zero-copy iteration.
* `EmfRecord::{parse}` and `EmfRecordRef::parse_ref` require `&Limits`;
  `EmfRecordRef::to_owned` also checks the selected allocation ceiling.
* `WmfParser::{new,from_owned}` and `WmfRecord::parse` require `&Limits`.
* Allocating EMF point/count array helpers require `&Limits`; fixed-width
  scalar/header readers retain their existing signatures.  Borrowed DIB
  extraction also requires `&Limits` so a public variable fragment boundary
  cannot bypass input/selected-range admission.

The raw paths reject input/record bytes before copying or reserving and cap
record/path counts before allocation.  Owner converters pass their retained
`Options::limits`; convenience conversion functions select their documented
owner default only at the high-level facade boundary.

## Consumer and test migration

PICT, EMF SVG, WMF SVG, and the shared metafile codec were migrated in the
same source change.  Compatibility raster converters retain and forward the
parser-selected policy into their SVG playback path, so output/object/state
ceilings cannot silently fall back to defaults.  Focused parser and converter
tests cover exact input-size acceptance, record-count rejection, record-byte
rejection, allocating path-array boundaries, and compatibility output-limit
refusal.  No `3rdparty/` fixture is referenced.  Native/Cargo validation is
deferred to the coordinator's capped validation slot; this record makes no
performance or full-workspace gate claim.
