# Retained scanner instruction decomposition

Root recomputed all four accepted 0525 candidate commit profiles. The analyzer
checks the exact 550-file prior seal, source hashes, raw incoming commit edge,
raw program total, direct owner chain, every direct child edge of the five
selected owners, and inclusive = self + direct equations against both retained
annotation forms. Parser helpers and raw inputs are hash-bound in
`profile-analysis.json`. This is new attribution of retained measurements,
not a new performance capture.

Across four profiles, commit totals 782,799,138 Ir and scanner
totals 416,939,582 Ir (53.2627% of commit).
Scanner self work totals 43,317,612 Ir. Its largest direct children are:

| Scanner direct child | Aggregate Ir | Share of commit |
| --- | ---: | ---: |
| `litchi_xlsx::raw::worksheet::edit::codec::snapshot::scan::Scanner::start_cell` | 126,302,261 | 16.1347% |
| `quick_xml::reader::Reader<R>::read_event_impl` | 80,289,669 | 10.2567% |
| `quick_xml::name::NamespaceResolver::resolve_event` | 31,526,144 | 4.0274% |
| `quick_xml::reader::ns_reader::NsReader<R>::process_event` | 31,132,640 | 3.9771% |
| `quick_xml::name::QName::local_name` | 28,409,696 | 3.6292% |
| `litchi_xlsx::raw::worksheet::edit::codec::snapshot::scan::Scanner::scan_guard` | 23,704,144 | 3.0281% |
| `__rustc::__rust_realloc` | 15,594,283 | 1.9921% |
| `litchi_xlsx::raw::namespace::is_spreadsheetml_name` | 13,879,184 | 1.7730% |
| `alloc::raw_vec::RawVec<T,A>::grow_one` | 8,179,777 | 1.0449% |
| `quick_xml::events::BytesRef::decode` | 6,131,372 | 0.7833% |

`start_cell` contains `cell_address` and `cell_tag`; those nested costs must
not be added again to the scanner total. The whole rewrite branch totals
444,163,798 Ir, including scanner work. Layout destruction is a
sibling of the scanner under that branch and totals 11,180,418 Ir.

Scanner direct `__rust_realloc` and `RawVec::grow_one` total 23,774,060 Ir.
Those rows and layout destruction identify allocation/storage work worth a
bounded pilot, but include work beyond the cell primary-span representation.
They are not a removable-cost estimate. The arena does not target most of the
scanner's parser, attribute, namespace or guard costs; native usefulness is
therefore uncertain. The all-event namespace resolver edge includes required
start/empty lookups, so it cannot quantify the removable end-event subset.

No function-call metadata is reported as an allocation count. Collection-off
setup/readback can contribute that metadata. No instruction ratio is a native
latency, memory, I/O, producer, scaling or end-to-end speedup claim. The draft
has not been built or timed. Failed native admission must reject it even if
allocation or instruction counters improve.

`analyze_test.py` verifies a valid control and rejects in-memory changes to raw
edge cost and owner self cost. `verify.py` replays this decomposition into a
temporary output and requires byte equality, alongside source/ADR, patch
reconstruction and cleanup checks. Temporary replay directories are removed
by their context managers; no retained evidence is overwritten by verification.
