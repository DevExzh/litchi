# 0730 DOC handoff pilot probe

This directory is a source-compatible measurement harness copied from the
sealed 0728 `build-5-probe`. The crate name is the only Rust namespace change:
the timing binary is `ole_format_save_probe` and the allocator binary is
`ole_format_save_probe_alloc`. Both binaries use the ordinary default feature
set and compile against the workspace crates selected by their relative path.
The probe intentionally contains the inherited container and PPT controls for
source compatibility, but this 0730 matrix measures only the public DOC
`--operation format` route for `docfloat` and `docnohf`.

The bound edit is the exact 0728 replacement text:

```
litchi copy-through baseline replacement text
```

It is 45 UTF-16 code units. The expected bytes are produced once by the same
public edit before any measured sample. That expected output is the pinned
identity used by each sample's oracle; the candidate binary is required to
produce the same source, expected-output, replacement, semantic-witness and
inventory identities as the baseline run. The root analyzer binds those
identities to the current fixture and source manifests.

The timed `format` interval starts immediately before
`Snapshot::open(source.to_vec(), ...)` and ends after `edit.commit()` and the
returned snapshot bytes have been copied into the output `Vec`. It therefore
covers the public open, edit construction, paragraph replacement, commit and
output-copy lifecycle. The expected-output construction, source/expected
inventory collection, and all output validation happen outside that interval.
Warmups execute the same lifecycle and are discarded; only requested samples
are serialized.

The allocator binary installs the inherited counting system allocator. Its
`allocation_format` region opens immediately before the same public lifecycle
and closes after the output `Vec` is returned. The output is kept alive outside
the region for validation, so `retained_bytes` describes ownership at the
region boundary rather than a peak or RSS measurement. Allocation counters are
process-local and report total allocated/deallocated bytes, calls, relative
peak live bytes, and boundary-retained bytes. The probe does not claim phase
allocation costs for the DOC route.

Each measured output is checked after its interval. The inherited strong
oracles validate CFB completeness and structure, exact stream paths and bytes,
unchanged streams, allowed changed streams, semantic directory metadata, root
and storage CLSIDs, normalized raw directory bytes, a length-changing stream
proof, and the direct DOC paragraph/collection witness. The probe also runs
the inherited negative controls (missing stream, untouched same-length
mutation, directory CLSID/state/timestamp mutations, source swap, and a wrong
DOC target). A failure in any expected oracle or control fails the process.

The two inputs are the same fixtures and paths used by 0728:

* `test-data/ole/doc/FloatingPictures.doc` (`docfloat`)
* `test-data/ole/doc/NoHeadFoot.doc` (`docnohf`)

The root runner owns Cargo builds, native process execution, sample ordering,
candidate integration, and packet sealing. This directory contains no
candidate-only API calls, so the same probe source and command shape can build
the pre-change baseline and the post-change candidate.
