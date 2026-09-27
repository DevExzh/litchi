# 0784 Callgrind profile analysis

This is an offline replay of the six retained Callgrind captures. Ir is
guest instruction attribution for the exact `capture_region_0784` wrapper;
it is not native latency or a production performance claim.

## Scoped totals

| Pass | Shape | Summary / owner inclusive Ir | Wrapper self Ir | Immediate child Ir | Dominant child | Dominant child Ir |
| ---: | --- | ---: | ---: | ---: | --- | ---: |
| 0 | tiny | 7,915,062 | 2 | 7,915,060 | `litchi_pptx::package::model::Package::opened_presentation_with_limits` | 7,915,060 |
| 0 | medium | 14,291,409 | 2 | 14,291,407 | `litchi_pptx::package::model::Package::opened_presentation_with_limits` | 14,291,407 |
| 0 | large | 594,662,717 | 2 | 594,662,715 | `litchi_pptx::package::model::Package::opened_presentation_with_limits` | 594,662,715 |
| 1 | large | 594,727,095 | 2 | 594,727,093 | `litchi_pptx::package::model::Package::opened_presentation_with_limits` | 594,727,093 |
| 1 | medium | 14,285,406 | 2 | 14,285,404 | `litchi_pptx::package::model::Package::opened_presentation_with_limits` | 14,285,404 |
| 1 | tiny | 7,913,175 | 2 | 7,913,173 | `litchi_pptx::package::model::Package::opened_presentation_with_limits` | 7,913,173 |

The partition checked for every row is `wrapper self Ir + immediate-child
inclusive Ir = owner inclusive Ir`. Descendant inclusive rows overlap and
are excluded from that sum.

## Dominant inclusive paths

- **0-tiny**: `pptx_capture_probe::capture_region_0784` (2 self) -> `litchi_pptx::package::model::Package::opened_presentation_with_limits` (83 self) -> `litchi_pptx::opened::model::capture_internal` (2,056 self) -> `litchi_pptx::opened::model::package_fingerprint_with_memo` (7,140 self) -> `litchi_pptx::opened::model::feed` (12,150 self) -> `sha2::sha256::compress256` (5,291,677 self)
- **0-medium**: `pptx_capture_probe::capture_region_0784` (2 self) -> `litchi_pptx::package::model::Package::opened_presentation_with_limits` (83 self) -> `litchi_pptx::opened::model::capture_internal` (6,511 self) -> `litchi_pptx::opened::model::package_fingerprint_with_memo` (9,799 self) -> `litchi_pptx::opened::model::feed` (17,250 self) -> `sha2::sha256::compress256` (7,479,568 self)
- **0-large**: `pptx_capture_probe::capture_region_0784` (2 self) -> `litchi_pptx::package::model::Package::opened_presentation_with_limits` (83 self) -> `litchi_pptx::opened::model::capture_internal` (50,359 self) -> `litchi_pptx::parts::slide::SlidePart::finish_from_processed` (48,200 self) -> `litchi_pptx::notes::codec::scan_processed_xml` (22,308,021 self) -> `litchi_pptx::notes::codec::inspect_element` (34,965,460 self) -> `<litchi_opc::xml_attributes::CheckedAttributes as core::iter::traits::iterator::Iterator>::next` (13,856,850 self) -> `quick_xml::events::attributes::IterState::next` (38,981,486 self)
- **1-large**: `pptx_capture_probe::capture_region_0784` (2 self) -> `litchi_pptx::package::model::Package::opened_presentation_with_limits` (83 self) -> `litchi_pptx::opened::model::capture_internal` (50,143 self) -> `litchi_pptx::parts::slide::SlidePart::finish_from_processed` (48,200 self) -> `litchi_pptx::notes::codec::scan_processed_xml` (22,308,021 self) -> `litchi_pptx::notes::codec::inspect_element` (34,965,460 self) -> `<litchi_opc::xml_attributes::CheckedAttributes as core::iter::traits::iterator::Iterator>::next` (13,856,850 self) -> `quick_xml::events::attributes::IterState::next` (38,981,486 self)
- **1-medium**: `pptx_capture_probe::capture_region_0784` (2 self) -> `litchi_pptx::package::model::Package::opened_presentation_with_limits` (83 self) -> `litchi_pptx::opened::model::capture_internal` (6,511 self) -> `litchi_pptx::opened::model::package_fingerprint_with_memo` (9,799 self) -> `litchi_pptx::opened::model::feed` (17,250 self) -> `sha2::sha256::compress256` (7,479,568 self)
- **1-tiny**: `pptx_capture_probe::capture_region_0784` (2 self) -> `litchi_pptx::package::model::Package::opened_presentation_with_limits` (83 self) -> `litchi_pptx::opened::model::capture_internal` (2,056 self) -> `litchi_pptx::opened::model::package_fingerprint_with_memo` (7,140 self) -> `litchi_pptx::opened::model::feed` (12,150 self) -> `sha2::sha256::compress256` (5,291,677 self)

## Top self Ir rows

### 0-tiny

| Function | Self Ir | Immediate child Ir |
| --- | ---: | ---: |
| `sha2::sha256::compress256` | 5,291,677 | 0 |
| `quick_xml::events::attributes::IterState::next` | 322,298 | 121,768 |
| `core::str::converts::from_utf8` | 286,357 | 0 |
| `litchi_pptx::notes::codec::inspect_element` | 177,938 | 981,991 |
| `quick_xml::reader::Reader<R>::read_event_impl` | 163,476 | 357,315 |
| `litchi_pptx::notes::codec::scan_processed_xml` | 88,941 | 1,948,798 |
| `quick_xml::name::NamespaceResolver::resolve_prefix` | 85,469 | 14,789 |
| `memchr::arch::x86_64::memchr::memchr3_raw::find_avx2` | 80,280 | 82,480 |
| `<litchi_opc::xml_attributes::CheckedAttributes as core::iter::traits::iterator::Iterator>::next` | 79,194 | 264,260 |
| `quick_xml::reader::state::ReaderState::emit_start` | 78,699 | 20,019 |

### 0-medium

| Function | Self Ir | Immediate child Ir |
| --- | ---: | ---: |
| `sha2::sha256::compress256` | 7,479,568 | 0 |
| `core::str::converts::from_utf8` | 775,887 | 0 |
| `quick_xml::events::attributes::IterState::next` | 739,646 | 262,714 |
| `litchi_pptx::notes::codec::inspect_element` | 492,360 | 2,515,038 |
| `quick_xml::reader::Reader<R>::read_event_impl` | 486,039 | 1,031,018 |
| `litchi_pptx::notes::codec::scan_processed_xml` | 288,933 | 5,206,573 |
| `quick_xml::name::NamespaceResolver::resolve_prefix` | 259,589 | 50,819 |
| `quick_xml::reader::state::ReaderState::emit_start` | 239,736 | 52,461 |
| `memchr::arch::x86_64::memchr::memchr3_raw::find_avx2` | 210,960 | 216,698 |
| `<litchi_opc::xml_attributes::CheckedAttributes as core::iter::traits::iterator::Iterator>::next` | 204,690 | 587,852 |

### 0-large

| Function | Self Ir | Immediate child Ir |
| --- | ---: | ---: |
| `sha2::sha256::compress256` | 207,605,811 | 0 |
| `core::str::converts::from_utf8` | 54,540,311 | 0 |
| `quick_xml::events::attributes::IterState::next` | 38,981,486 | 14,103,277 |
| `litchi_pptx::notes::codec::inspect_element` | 34,965,460 | 169,039,178 |
| `quick_xml::reader::Reader<R>::read_event_impl` | 34,766,287 | 70,188,418 |
| `litchi_pptx::notes::codec::scan_processed_xml` | 22,308,021 | 354,821,610 |
| `quick_xml::name::NamespaceResolver::resolve_prefix` | 19,119,253 | 4,036,179 |
| `quick_xml::reader::state::ReaderState::emit_start` | 17,427,424 | 1,631,179 |
| `<litchi_opc::xml_attributes::CheckedAttributes as core::iter::traits::iterator::Iterator>::next` | 13,856,850 | 34,438,655 |
| `memchr::arch::x86_64::memchr::memchr3_raw::find_avx2` | 13,750,320 | 13,790,912 |

### 1-large

| Function | Self Ir | Immediate child Ir |
| --- | ---: | ---: |
| `sha2::sha256::compress256` | 207,605,811 | 0 |
| `core::str::converts::from_utf8` | 54,540,311 | 0 |
| `quick_xml::events::attributes::IterState::next` | 38,981,486 | 14,102,909 |
| `litchi_pptx::notes::codec::inspect_element` | 34,965,460 | 169,038,675 |
| `quick_xml::reader::Reader<R>::read_event_impl` | 34,766,287 | 70,175,994 |
| `litchi_pptx::notes::codec::scan_processed_xml` | 22,308,021 | 354,839,270 |
| `quick_xml::name::NamespaceResolver::resolve_prefix` | 19,119,253 | 4,036,179 |
| `quick_xml::reader::state::ReaderState::emit_start` | 17,427,424 | 1,630,166 |
| `<litchi_opc::xml_attributes::CheckedAttributes as core::iter::traits::iterator::Iterator>::next` | 13,856,850 | 34,438,287 |
| `memchr::arch::x86_64::memchr::memchr3_raw::find_avx2` | 13,750,320 | 13,790,902 |

### 1-medium

| Function | Self Ir | Immediate child Ir |
| --- | ---: | ---: |
| `sha2::sha256::compress256` | 7,479,568 | 0 |
| `core::str::converts::from_utf8` | 775,887 | 0 |
| `quick_xml::events::attributes::IterState::next` | 739,646 | 262,373 |
| `litchi_pptx::notes::codec::inspect_element` | 492,360 | 2,514,658 |
| `quick_xml::reader::Reader<R>::read_event_impl` | 486,039 | 1,030,765 |
| `litchi_pptx::notes::codec::scan_processed_xml` | 288,933 | 5,205,498 |
| `quick_xml::name::NamespaceResolver::resolve_prefix` | 259,589 | 50,819 |
| `quick_xml::reader::state::ReaderState::emit_start` | 239,736 | 52,585 |
| `memchr::arch::x86_64::memchr::memchr3_raw::find_avx2` | 210,960 | 216,703 |
| `<litchi_opc::xml_attributes::CheckedAttributes as core::iter::traits::iterator::Iterator>::next` | 204,690 | 587,511 |

### 1-tiny

| Function | Self Ir | Immediate child Ir |
| --- | ---: | ---: |
| `sha2::sha256::compress256` | 5,291,677 | 0 |
| `quick_xml::events::attributes::IterState::next` | 322,298 | 121,899 |
| `core::str::converts::from_utf8` | 286,357 | 0 |
| `litchi_pptx::notes::codec::inspect_element` | 177,938 | 980,628 |
| `quick_xml::reader::Reader<R>::read_event_impl` | 163,476 | 357,654 |
| `litchi_pptx::notes::codec::scan_processed_xml` | 88,941 | 1,947,348 |
| `quick_xml::name::NamespaceResolver::resolve_prefix` | 85,469 | 14,789 |
| `memchr::arch::x86_64::memchr::memchr3_raw::find_avx2` | 80,280 | 82,455 |
| `<litchi_opc::xml_attributes::CheckedAttributes as core::iter::traits::iterator::Iterator>::next` | 79,194 | 264,391 |
| `quick_xml::reader::state::ReaderState::emit_start` | 78,699 | 20,299 |

Raw `.callgrind.1` hashes, zero-summary termination hashes, receipt
artifacts, native report parity, and 0780 capture fixture identities are
bound in `profile-analysis.json`.
