# 0362ul45 — DOC and OLE raw codec limits

Status: implementation and bounded validation are complete on final source tip
`983ab2cfb2d95b71562babdc468bcae46873da89`; the first repair was prepared from
predecessor source tip `75841b8fbdcd3e20f79239a3ce53713865244933`.  The API
work was prepared from integration source
`30aaafbf98e097aeec8cf6fe3df6caaae62a37d6`.  This packet covers the DOC
binary-part and legacy OLE Custom XML rows in the raw/record limit audit
`raw-record-limit-audit-3431ba566c.md`, under 0351 §2.5.3/4.  It does not
claim native qualification or ABBA measurement.

## API boundary

The public variable-size DOC fragment entry points now require a borrowed
`litchi_doc::SourceLimits`; no default twin or `_with_limits` public alias was
added:

- annotation bookmarks: the FIB/table and byte parsers, including the `Tags`
  projections;
- command bars: the FIB/table and byte parsers;
- OLE controls: the `RgxOcxInfo` FIB/table and byte parsers;
- route slip: the `Metadata` FIB/table and byte parsers; and
- XML schemas: `Collection::parse` and `parse_custom_xml_transform`.

The standalone variable helpers use the same required profile:
`route_slip::Recipient::parse_bytes` and command-bar `XString` parsing and
serialization (`parse_prefix`, `parse`, `parse_terminated_prefix`,
`to_bytes`, and `to_terminated_bytes`).

The corresponding variable-size writers now take the same required profile:
annotation bookmarks, command bars (including individual controls and the
package append seam), OLE-control arrays, and route slip (including
recipients).  Command-bar count parsers admit their minimum wire footprint
before reserving count-sized vectors.

Route-slip recipients and annotation bookmarks also preflight their fixed
minimum per-entry footprint before reserving count-sized vectors.

Each byte boundary charges `SourceLimitKind::EntryBytes` before a count or
length field can drive allocation.  The document package retains its ingress
`SourceLimits` in `DocumentState`; deferred annotation parsing and the route,
command-bar, and annotation package editors pass that retained profile through
their whole lifecycle, including reparse validation.  Existing package
defaults remain only at package/editor ingress, where they are immediately
retained.  The fixed 4/6-byte ObjectPool metadata parser and lexical/fixed
record helpers were not widened or churned.

The public OLE Custom XML codec now requires `&custom_xml::Limits` on
`write`, `parse_properties`, and `write_properties`.  The previous public
defaulting implementations are gone.  The model constructor uses its own
existing default model policy, while package inspection and transaction paths
continue to use their retained `Limits`.  The crypto test consumers were
migrated to pass an explicit profile.

## Correctness checks

The existing DOC malformed, truncation, round-trip, and atomic editor tests
were migrated to explicit source profiles.  New focused checks prove that
annotation bookmarks, command bars, route slip, OLE controls, XML schema
references, and the custom XML transform refuse a profile whose entry limit is
one byte below the source.  OLE Custom XML tests additionally exercise public
property-size refusal and public writer refusal for XML depth and aggregate
store bytes; the pre-existing item-count, item-size, schema-reference, and
malformed-input checks remain in place.

The held review repair adds exact/one-under output checks for every DOC raw
writer, bounded directory visitors for Custom XML inspection, retained-profile
Item XML edits, early attribute-string admission, and a schema-reference
count-change regression that retains opaque attributes, comments, whitespace,
and namespace topology.  The follow-up source repair additionally computes
decoded/normalized attribute lengths before owning XML values, resolves
property splice spans by namespace (including default-namespace elements and
same-local-name extensions), preserves UTF-16 Properties source encoding, and
adds exact DOC minimum-footprint and public-helper limit coverage.  The
opaque-topology follow-up now tracks the parser's opaque-subtree state and
whether the first typed schemaRefs container has been seen, preserving
same-namespace extension wrappers before and after the typed container.  These
follow-up tests were executed in the final bounded cohort below.

Static checks completed in this leaf (without Cargo or native tooling):

```text
rustfmt --edition 2024 --check <all changed Rust files>  PASS
git diff --check                                      PASS
targeted repository consumer search                   PASS
```

The glossary registers all newly required public signatures: 34 DOC raw
entry points in each of the `doc-minimal` and `doc-full` profiles, plus the
three OLE Custom XML codec entry points in `ole-common`.  JSON parsing and the
external-type-aware `checked_policy` from foundation `939d4016092409575bb280cf3d3d1b8623cf0994`
both pass.  This policy validation did not rebuild Rust or change the tested
source tree.

## Bounded validation receipt

The approved self-attributing runner was used with a fresh external target and
state directory, `CARGO_NET_OFFLINE=true`, `CARGO_BUILD_JOBS=1`,
`RUST_TEST_THREADS=1`, `CARGO_INCREMENTAL=0`, `LITCHI_MEM_MAX=20G`, and the
18/19/20 GB target guards.  Runner source was
`f5f3f8d9fab8896f9ab545ffc5f167300212fefd`; the runner file was
`/home/zhuhe/code/litchi-worktrees/uo2-gate-runner-self-attribution/docs/unified-ops/results/0362ul29/tooling/run_safe_gate.py`, SHA-256
`1904177040fc7c523582433028d9bba56efc30fe0d9d96beb554a2fab343691e`.  The
approved workspace lock was `Cargo.lock`, SHA-256
`4100b1fbb272dd24842f33a2d6f353a3f3a97dfa7e435c4649a0e1a029bfaea1`.

The final source-equality run used fresh state and target directories retained
outside the repository at
`(deleted-coord-artifact)` and
`(deleted-coord-artifact)`.
All seven commands were `isolated=False` as permitted for this targeted cohort,
and every command exited zero:

```text
doc-raw-limit-tests                0  52.489018 s  331609673 B  27fd9478866b733c4e255a066a33b971ea804f65b21f13286ac1f00040617536
ole-custom-xml-fidelity            0  14.574248 s  429833527 B  14fcbac015cc6224f9e6b85c4783691a01d26784d805a1e74876ab500590fd39
ole-custom-xml-limits               0   1.929009 s  429833527 B  6ec781c1f6f80b0ef4136799c386e91b2e6cb43b8b983706eaec378a28a80fce
doc-all-lib-tests                  0  16.607896 s  429833527 B  c4a32c72bdc91eaff59ad9ad52c1cd366c81fc7c29bc9aaaeb0847e2aebb4a60
ole-common-all-lib-tests           0   3.767713 s  429833527 B  1fd19fda424340aaa346ec2c3eb3be8747458da41ee54f2ff34d82cb4ac5e241
doc-ole-crypto-check-all-targets  0  38.389855 s  518418681 B  283319d1157ec43a2e791be0e2a2ce8fa4ecc0c2dba42458210b048b800567e8
doc-ole-crypto-clippy-all-targets 0  34.258220 s  540630995 B  4bcf9e3f9c242a6053c6a58c87bc1a58db552f745d04bff5be51b4f81f1bc88a
```

The raw-limit selector ran 4 tests, the OLE fidelity selector ran 2 tests, and
the OLE limit selector ran 1 test; all passed.  The full DOC lib suite completed
1,248 test cases with 1,246 passing and 2 ignored; the full OLE-common lib suite ran
238 tests, all passing.  The check and clippy commands completed without a
test harness.  The largest sampled target was 540,630,995 bytes, below the
19,000,000,000-byte live and 20,000,000,000-byte final guards.  The runner's
final source state matched the recorded commit and clean status after every
row.

Superseded diagnostic attempts remain outside the repository.  Their exact
source tips, terminal row, log path, and digest are retained here:

```text
1678388376  doc-raw-limit-tests                E101  .../scoped-doc-ole-1792889306615922011/01-doc-raw-limit-tests-attempt-scoped-1792889307590458863-713348-1.log  4b153044c92c983ac7f72939a97594ee185f859e86ed1068b2deef0c66317f3d1
e1919df387  doc-raw-limit-tests                E101  .../scoped-doc-ole-1792889368824554574/01-doc-raw-limit-tests-attempt-scoped-1792889369783785620-735839-1.log  4453f9b0a37109c4f105dd3be9e9c6e57cc32a3da302d27eb6ef6791beab52d1
480dbb1f61  doc-ole-crypto-clippy-all-targets  E101  .../scoped-doc-ole-1792889442082080390/07-doc-ole-crypto-clippy-all-targets-attempt-scoped-1792889572458519879-766840-7.log  8c7e8bd5521e56be88242bf6b7bc3d71fd8a69eb8ff4bc545b6a3cba92c95af8
44fdb94bbc  doc-ole-crypto-clippy-all-targets  E101  .../scoped-doc-ole-1792889601364250786/07-doc-ole-crypto-clippy-all-targets-attempt-scoped-1792889735834071751-825976-7.log  43a643b462ba4812e3573198180e4f1f2fd56b8f2dcd3a397176e5b80c0c8101
```

The first failed source had the missing `BytesStartExt` import and temporary
UTF-16 byte-array borrows; the second then exposed the DOC toolbar input
lifetime.  The latter two were Rust 1.95 clippy diagnostics fixed by the final
source tip.  The common prefix for the abbreviated paths above is
`(deleted-coord-artifact)`.

The broad `clippy-leaves` gate was previously attempted on source tip
`cf5226b3eaeb6014c8b38592d6c58d124568ba9b`.  It exited 101 after 531.436594 s
at 625554465 B because
`litchi-drawingml` has seven unrelated `clippy::all` diagnostics; the complete
receipt is
`28-clippy-leaves-attempt-1792884788814310629-3205308.log`, SHA-256
`e23c8ba246fcbb020bf7c7316829c0908aa80b560b65a6832b66c228b955a8ee`.
The DOC/OLE/crypto all-target clippy command above passed independently.

No ABBA, native, or performance result is asserted; the validation only proves
the scoped correctness, limit, compile, and clippy checks listed above.  The
coordinator must run the workspace/API gates after integrating this leaf.

## Preserved predecessor rerun

After the shared `ControlHeader::encoded_len` repair (`75841b8fbd`), the
approved runner was rerun with fresh state and target directories:

```text
root   /home/zhuhe/code/litchi-worktrees/uo2-cfb-doc-olerawcodec
state  (deleted-coord-artifact)
target (deleted-coord-artifact)
command python3 run_safe_gate.py --only fmt,tests-leaves
fmt    exit 0, 29.3 s, log SHA-256 e3b0c44298fc1c149afbf4c8996fb92427ae41e4649b934ca495991b7852b855
tests  exit 1, 484.874877 s, target 1,144,826,751 B
```

The complete `tests-leaves` receipt is
`(deleted-coord-artifact)/30-tests-leaves-attempt-1792887298593559226-4032954.log`,
SHA-256 `c62c50e538fbd555edc485edb870de60351f7fcc6f405a82eed63e2a2b4a6703`.
The DOC and OLE package chunks compiled and passed in that predefined step;
the terminal failure was step 85, the unrelated `litchi-drawingml` doctest
`chart/mod.rs` (`?` could not convert to `std::io::Error`).  The failed run is
preserved and no result from it is attributed to the follow-up source repair.

## Coordinator integration review

Atomic leaf `d6cc88337c523c0d40970cdaa24c5fcc9ad7e347` differs from
validated `983ab2cfb2d95b71562babdc468bcae46873da89` only in this record.
The delta from source-reviewed `16783883765a930a13b00f7bbb578e2850718370`
contains only the required trait import, explicit input lifetime, valid local
UTF-16 byte-array borrow, equivalent decimal parsing and Clippy annotations.
Root independently verifies every final command log and Cargo trace digest,
unchanged lock identities, clean validated source, target below 541 MB and
empty owned-process lists in all terminal snapshots.

The final external `state.json` is 114,111 bytes, SHA-256
`5abb01d848563f8faab6705cdf0dcd9ac921b6db9a8668f275dff40cc6a2a4fb`.
All seven rows pass. Source references retain the tested tree and each failed
diagnostic tree. This merge accepts the scoped correctness validation; final
composed workspace and performance acceptance remain required.
