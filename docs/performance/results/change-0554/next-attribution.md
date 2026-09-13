# 0554 next OLE2/XLS attribution: physical-role reconciliation

Status: attribution-only. This note selects a follow-up experiment; it does
not report a native speedup, an adopted change, or a missing-function result.

The fresh 0554 baseline owner profile places the largest independent cost after
the directory-name work in `OleFile::validate_physical_sector_layout`. The
selected profile is the XLS-owned source-open case from repeat 1, with five
positive timed owner dumps, no warmup, and Callgrind `Ir` only. The frozen plan
is `docs/performance/results/change-0554/plan.json`, SHA-256
`d901c87af4e8d2585af997dbc850832db21db0ce0e359d3dcf7d90c7b415e98c`. Its
selected owner is
`litchi_xls::workbook::source::SourceBackedWorkbook::from_read_at_with_limits`
and its case is `xls_owned_source_open_one_cell`.

## Owner boundary and call path

Every numbered dump has one positive incoming edge into the selected owner:

```text
litchi_xls::workbook::source::SourceBackedWorkbook::from_read_at
  -> SourceBackedWorkbook::from_read_at_with_limits       [selected owner]
     -> litchi_cfb::shared::SharedOleFile::open_with_limits
        -> SharedOleFile::open_source_with_limits
           -> litchi_cfb::file::OleFile<R>::open_with_limits
              -> load_fat
              -> load_directory
              -> validate_stream_allocations
                 -> SectorChainScratch::collect_exact
              -> validate_physical_sector_layout
     -> SourceBackedWorkbook::from_shared_ole_file_with_limits
```

The owner edge is the scope authority: its call count is one and its incoming
inclusive `Ir` equals the dump summary in all five positive dumps. The
runner-to-wrapper collection-off ancestry is context only; any zero call count
in that ancestry is not treated as an operation count. The CFB functions above
are descendants inside the selected XLS owner, so their costs must not be
presented as standalone XLS latency.

The corresponding baseline source paths at revision
`53330ff6745dd7416f3bfb48ef891a020ee9d151` are:

| Path | Boundary |
| --- | --- |
| `crates/litchi-xls/src/workbook/source.rs:733-742` | selected XLS owner and shared CFB open |
| `crates/litchi-cfb/src/shared.rs:472-505` | shared-reader wrapper and `OleFile` construction |
| `crates/litchi-cfb/src/file.rs:749-750` | validation order inside `OleFile::open_with_limits` |
| `crates/litchi-cfb/src/file.rs:1079-1164` | stream-chain validation and role claims |
| `crates/litchi-cfb/src/file.rs:1178-1192` | final physical role/FAT reconciliation pass |
| `crates/litchi-cfb/src/file.rs:1459-1498` | public directory entry parse, including name decode |
| `crates/litchi-cfb/src/file.rs:2911-2937` | duplicate public UTF-16LE decoder |

The baseline source manifest binds the relevant source files as follows. The
live checkout may temporarily contain the 0554 candidate while this evidence
is being collected; these are the hashes of the source used by the baseline
profile, not hashes read from that mutable candidate checkout.

| Source file | Baseline SHA-256 |
| --- | --- |
| `crates/litchi-cfb/src/file.rs` | `72fb5295dfbba5f3ed2192a9c24ecbf7d8aa2759f418b77563a4e2ff1042d03b` |
| `crates/litchi-cfb/src/shared.rs` | `2aef59e5e7d20a8984b0c2f07d561865a87a74d927cb088d4a08ff85532c0624` |
| `crates/litchi-cfb/src/directory_name.rs` | `3aa3565c9fd9e697115fa2ba76bb7f293e9bb5468492335f7dff33793dd604e5` |
| `crates/litchi-xls/src/workbook/source.rs` | `3e7dbde9f197771ee7766a597f3f27f6a0143b79f1cf2575dd52a9f8493f6d92` |

## Fresh profile evidence

The 0554 profile driver validated these dumps with the immutable raw parser
from change 0536. The 0554 driver SHA-256 is
`af789e447d0093d1e2ad6f2f01a44109ea6701735517c5d1cf2ed9754705f315`; the
retained 0536 parser SHA-256 is
`5ce0d3a0c9f9246f207b9be063791cf6ccf014f6582da47562de42c7e0d364a`.
The baseline receipt SHA-256 is
`66991507f940ac9e0cbf90247292a64dcad473ac2e161b89fbf7f900b08d05d1`, and
the baseline source-manifest SHA-256 is
`c6ec30f2e7c8116c6e9bd77d1622295d8245a36261daf9fabae657be4fe9a8e9`.
The retained baseline executable is 60,099,224 bytes with SHA-256
`075a61ea31d9441e12dcdfd03fb0d4e0b1f7d264f460eac465d65864cc4ae459`; the
profile catalog binds the XLS/CFB archive to SHA-256
`6a57231ba681bc7bdd38d447ebd5348ef3b1fefedeefb1e61c97f22faa074e53`.

| Positive dump | Summary/owner `Ir` | Raw SHA-256 |
| --- | ---: | --- |
| `baseline/profile-r1-xls-owned.callgrind.1` | 2,263,310 | `e1062ccec1059976592fd9dd39d457f690db3f960d31820ea45bee7fce82b527` |
| `baseline/profile-r1-xls-owned.callgrind.2` | 2,263,116 | `af2418bc07bf8e8ed29bc2f6754657b8434b564d2345d2adc0b1efc5cc525bb1` |
| `baseline/profile-r1-xls-owned.callgrind.3` | 2,265,417 | `2b5d91bcbc4df58180b7cab76257ec5b2b2556dbfc3d9bf4539dadda63eb8b82` |
| `baseline/profile-r1-xls-owned.callgrind.4` | 2,263,697 | `63235eb34d53cf7aa5a31a300d98fe9eb4c8e667cd5faad2a4b73ca5c42fdbf8` |
| `baseline/profile-r1-xls-owned.callgrind.5` | 2,263,711 | `00ca0ae63cf6ee5927a5600b4ba10421e51c2bd392a47445c1dbc700d5734a8e` |

The selected-owner summary median is 2,263,697 `Ir`. For each function below,
`self` is exclusive function `Ir`; `direct` is the sum of its direct child
edge inclusive `Ir`; and `inclusive` is `self + direct`. Inclusive rows
overlap when one function is a descendant of another and must not be added.
The values are the median of the five positive dumps:

| Rank | Function | Self `Ir` | Direct-child `Ir` | Inclusive `Ir` | Self / owner |
| ---: | --- | ---: | ---: | ---: | ---: |
| 1 (excluded) | `SectorChainScratch::collect_exact` | 1,120,228 | 43,250 | 1,163,478 | 49.4867% |
| 2 | `OleFile<R>::validate_physical_sector_layout` | 398,342 | 0 | 398,342 | 17.5970% |
| 3 | `OleFile<R>::validate_stream_allocations` | 362,897 | 1,168,790 | 1,531,687 | 16.0312% |
| 4 | `OleFile<R>::load_fat` | 50,101 | 71,539 | 121,640 | 2.2132% |
| 5 | `OleFile<R>::load_directory` | 8,046 | 51,689 | 59,735 | 0.3554% |

The first row is retained only as a boundary and exclusion. The 0548
checkpoint and 0549 test-and-mark variants of this collector family were
already rejected on fresh native gates; this note does not select, revive, or
modify that code. The stream-validation inclusive row contains the same
collector descendants: its 1,168,790 direct-child `Ir` is mostly the
`collect_exact` edge, so the 16.0312% self row is the independent local loop
cost and the 67.6631% inclusive row is not a second removable budget.

The next independent row is exact across all five dumps:

```text
OleFile<R>::open_with_limits
  -> validate_physical_sector_layout   calls=1, inclusive Ir=398,342
                                             self Ir=398,342
```

It is a leaf in this profile. Its source loop at `file.rs:1179-1189` walks the
physical `sector_roles` prefix, obtains the matching FAT entry, and rejects an
unclaimed sector whose marker is not `FREESECT`. The XLS input catalog reports
16,995,840 archive bytes; for the 512-byte CFB layout this is 33,194 physical
sectors, making the observed pass about 12.0004 `Ir` per physical sector for
this heavy input. That scaling observation is specific to this corpus and is
not a latency estimate.

For comparison, the name work addressed by 0554 is much smaller in this owner
profile: `parse_directory_entry` has 1,368 self and 6,312 direct-child `Ir`,
`decode_utf16le` has 4,404 self and 1,356 direct-child `Ir`, and
`directory_name_data` has 14,532 self and 11,325 direct-child `Ir`. The
decoder's direct incoming edge is 12 calls per positive dump in the baseline,
but its outgoing `try_reserve` edge is not a decoder call count. These figures
explain why the physical pass is the next attribution target without claiming
that its whole inclusive fraction is removable.

## Recommended next experiment

Measure a source-bound, proof-first **physical marker accounting** candidate.
The candidate would ask whether the final full physical-prefix scan can be
replaced by exact state carried through the existing FAT materialization and
role claims. While decoding the FAT, it could retain the physical-prefix
non-`FREESECT` markers; as later directory, MiniFAT, root, and stream claims
assign roles, it could retire the corresponding pending markers. The final
validation point would then check the retained offending marker state rather
than rereading every physical role/FAT pair.

This is a feasibility hypothesis, not a design approval. The first review
must prove all of the following before any timed candidate build:

* FAT length and missing-entry errors retain their current final precedence;
  markers beyond the physical file keep their tolerated padding behavior.
* FAT and DIFAT claims made before `self.fat` is installed are accounted for
  without an unrecorded second full scan; later `claim_sector` calls remain
  checked and duplicate claims still fail at the same boundary.
* The exact offending sector and marker remain available for the existing
  corruption text and ordering, including all-free, unclaimed non-free,
  claimed non-free, short-FAT, and malformed-chain fixtures.
* `validate_stream_allocations`, `SectorChainScratch`, its reservation labels,
  zero-fill/reset behavior, and the root/classic-Mac and public directory
  views remain unchanged. No public API, unsafe code, or new unbounded state
  is justified.

The 0534 paired-prefix rewrite already made the physical loop shorter and was
rejected because every primary XLS p50 row regressed. This experiment must not
reuse that patch or repeat its loop rewrite; it should be evaluated as a
different accounting/removal hypothesis and closed if the bookkeeping costs
or error-state requirements erase the opportunity. If a source review passes,
the candidate still needs fresh CFB shape controls, actual XLS owner profiles,
native latency, allocation/RSS, malformed-input, correctness, and final
quality gates under the 0554 admission contract. No Callgrind `Ir` result in
this note is a production or adoption claim.
