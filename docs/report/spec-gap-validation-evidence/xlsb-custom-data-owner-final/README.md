# XLSB Custom Data owner validation

The current candidate is the nine-file capture in `source-v6.json`, based on
`f2daff703`. It adds the ordinary `Package` and `Workbook` Custom Data lifecycle
and a source-preserving BIFF12 ExtConn14 binding lens. The shared XML model and
codec were committed separately in `c00798d63`. The approved owner design and
pinned record grammar are retained in [the design](../xlsb-custom-data-design.md)
and [grammar evidence](../xlsb-extconn14-grammar/README.md).

## Public lifecycle

`litchi_xlsb::custom_data` exports semantic storage values, selectors, limits,
snapshots, transactions, commits and patches. `Package` and `Workbook` expose
`custom_data`, `edit_custom_data`, their limit/context variants,
`apply_custom_data` and `apply_custom_data_patch`. Transactions support insert,
upsert, scalar properties/payload replacement, rename and explicit
referenced-storage removal policies. Payloads remain inert bytes; public CRUD
calls require no OPC identifiers or BIFF record IDs.

Rename and retarget update every recognized many-to-one binding. Unsupported
reference-bearing wrappers block identity-changing operations. Patches check
the source dependency closure, apply scoped changes to a detached candidate,
and perform semantic readback before publication. Exact inverse restoration
retains Properties, relationships and content-type source bytes. Incoming graph
proof scanning is shared; source memory reservations remain lifetime-bound.
UID-indexed readback handles insertions and renames across catalog sort order.
Collection cardinality and replacement byte limits are checked before allocation.

## Validation

Root validation used an isolated checkout and target with only the captured
files and retained lockfile overlaid. The final run passed 654 library tests
and 18 public lifecycle tests, with one library test ignored. All-target Clippy
and rustdoc passed with warnings denied. All nine hashes matched the built
checkout and shared source after the gates. This test selection does not claim
that every XLSB integration test ran.

```sh
cargo test --locked -p litchi-xlsb --lib --test custom_data_lifecycle --offline
cargo clippy --locked -p litchi-xlsb --all-targets --offline -- -D warnings
RUSTDOCFLAGS="-D warnings" cargo doc --locked -p litchi-xlsb --no-deps --offline
```

The lifecycle suite covers ordinary CRUD, many-to-one bindings, opaque/signed
refusal, stale sources, exact inverses, unrelated member preservation or
fail-closed behavior, source CRLF/indentation/PI/comment/CDATA preservation,
UID-order-changing rename and insertion, storage cardinality, exact/one-under
XML/BIFF/output limits, cancellation and temporary-limit refusal during inverse
source-template restoration. The six binding unit tests are in the library
selection; separate binding review also probed collision work/cancellation,
allocation charging, tighter rewrite limits and opaque wrapper ranges.

All fixtures are synthetic. No native producer, application open/save, formula
execution, external data retrieval or runtime speedup claim is made.

The final capture adds fallible reservation across the audited temporary collections, including the complete before/after relationship-ID union. Independent audit cleared the functional and broad collection findings; root reviewed the last checked union-reservation correction and approves this synthetic lifecycle scope (`v6-review.json`). The binding audit and subsequent lint-only corrections are recorded in `bindings-review.json`. `v4-review.json` and `v5-review.json` record historical rejected captures. Earlier captures and `v4-review.json`
record rejected revisions; their green tests do not approve the current source.
The separate user-owned `tests/custom_data.rs` was not modified or included in
this test selection. Its fixture uses compact `XmlPart` construction with
formatting whitespace, omits `END_EXT_CONN14`, and uses an unsupported direct
extension namespace. The new valid formatted physical-source regression passes.
