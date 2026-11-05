# 0097 — XLSX cell writes and native transaction history

XLSX now participates in `File::set`, closure edits, and atomic or best-effort
batches for direct cell paths. `value` accepts literal text, booleans, signed
64-bit integers and finite numbers; `formula` accepts native formula text
without a leading `=`; `clear` accepts true. Null explicitly clears value or
formula. Conflicting setters are refused before native mutation. The shared
property registry admits access, coercion and null policy. Native constructor
errors retain their typed cause and are classified as authored-value errors.

`ValueType::Scalar` describes the typed union without inferring a string's
kind. Wide integer inputs outside the scalar range are refused without rounding.
Capability projections advertise only the implemented writes; `value` and
`clear` are write-only and existing XLSB read declarations are unchanged.

Each operation edits the promoted native owner and directly adopts its commit.
There is no serialize/reopen boundary. Transactions retain native patches and
both immutable endpoints, preserving source-bound forward and inverse replay.
The full admission/edit/commit sequence checks source freshness. Failed
mutations poison an atomic closure; ordinary best-effort item failures roll
back that item. Semantic operation and retained-snapshot limits still apply.
The retained-data charge includes endpoints and patch buffers, but is not a
complete allocator-footprint or lazily derived-cache bound.

Native `Workbook::same_content` compares OPC part payloads, content types and
relationships, plus effective retained content-type and relationship manifests,
to recognize unchanged final snapshots without ZIP serialization. Unknown
producer markup participates in the comparison; a retained/authored manifest
mismatch conservatively counts as changed.
Equal authorized archive digests are a fast path; different ZIP framing falls
through to graph comparison. Opaque non-part bytes are proved through a shared
immutable source archive; unshared opaque state and custom parts remain typed
comparison refusals. Native patch endpoint accessors and no-op comparison expose
these facts without a facade-owned document model. Final no-ops keep the exact
input handle and publication bytes.

Source and eager tests cover all scalar kinds, formula/clear/null, literal
number/boolean/formula-looking strings, invalid XML controls, property errors,
exact no-ops, rollback, history limits, replay/inverse, foreign-owner rejection,
source freshness, source retention, and path-origin publication. Encrypted
owners refuse edits, including empty transactions. File-specific verb reports
return the same protection refusal for ordinary output and edits. Opaque OPC parts and non-part
ZIP members survive cell edits and explicit signature removal.

Query/range writes, structural commands, native command coalescing, OfficeCLI
string inference and extended clear modes remain discoverable gaps. Managed
open/publication and encrypted-profile-preserving edits remain open. No
performance improvement is claimed. Other format writes, Python, CLI and the
full goal's acceptance requirements remain active.

Validation:

```sh
CARGO_NET_OFFLINE=true cargo test -p litchi-xlsx
CARGO_NET_OFFLINE=true cargo test -p litchi-xlsx --lib workbook::tests::
CARGO_NET_OFFLINE=true cargo test -p litchi-ops --no-default-features --features xlsx,encryption
CARGO_NET_OFFLINE=true cargo test -p litchi-ops --no-default-features --features docx,xlsx,encryption
CARGO_NET_OFFLINE=true cargo test -p litchi-ops --no-default-features
CARGO_NET_OFFLINE=true cargo test -p litchi-ops --no-default-features --features xlsx --test capabilities --test xlsx_cell_writes --test xlsx_unified_save
CARGO_NET_OFFLINE=true cargo clippy -p litchi-ops -p litchi-xlsx -p litchi-opc --no-default-features --features litchi-ops/docx,litchi-ops/xlsx,litchi-ops/encryption --all-targets -- -D warnings
python3 tools/officecli_capability_delta.py --officecli /home/zhuhe/code/litchi/3rdparty/OfficeCLI --check
python3 -m unittest discover -s tools -p test_officecli_capability_delta.py
```

The full native suite passed 2,233 tests/doctests. After review strengthened
manifest comparison, all 20 targeted native workbook tests passed, including
producer-markup regressions. The XLSX/encryption suite passed 582 tests/doctests;
the final combined suite passed 841, including the later opaque-member,
round-trip no-op and signed-encryption capability regressions. The final minimal
suite passed 419; all 26 final XLSX-only integration tests passed. Warning-denied Clippy and all 14 capability-delta tests passed.
Capability goldens and the generated delta were updated from production tables.

The full native suite requires the existing LibreOffice fixture
`3rdparty/libreoffice-core/sc/qa/unit/data/xlsx/tdf169496_hidden_graphic.xlsx`.
It was temporarily copied from the base checkout and removed after validation;
no fixture or binary is added by this record.

Review found and resolved native-error classification, retained manifest
comparison, encrypted-owner capability reporting, and retention-budget
documentation issues. Formatting, diff hygiene, rewritten-base ancestry, and
artifact-size checks passed. Both required staged checks printed nothing.
Git staging still fails on the read-only common Git directory's `index.lock`,
so no commit was created. Temporary logs, fixture copy, generated root lockfile,
and build outputs are removed after verification. No temporary worktree was
created.
