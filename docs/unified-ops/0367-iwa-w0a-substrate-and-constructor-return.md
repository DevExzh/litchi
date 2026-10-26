# 0367 — W0-A: substrate admission fixes, raw physical access, and the IWA constructor return

- Status: landed on branch `uo2/iwa-w0a` (cut from trunk `feat/unified-ops-api`
  at `9c53797811`; trunk `1e78bb87b6` merged).
- Date: 2026-10-26.
- Scope: (1) the owner-priority return of
  [0361fda](0361fda-iwork-constructor-performance-return.md): IWA metadata
  constructors and retained-header parse/validate/encoded-length/encode;
  (2) wave W0-A of [0363](0363-iwa-monolith-exit-plan.md): admission and
  serialization fixes ported from the donor branches (§7) and the host's raw
  physical access moved to `litchi-iwa-archive::unstable` (owner decision D5).
- `performance_claim`: callgrind instruction counts and allocation counts,
  matched against baseline `ddf2ac8ae2` (constructor routes) and trunk
  `1e78bb87b6` (owner routes). Native timing remains unqualified: the host
  never had a quiet window, so every ABBA here is a loaded observation.

## 1. Constructor and retained-header return (0361fda)

### 1.1 What was repeated, and what changed

The returned candidate re-ran work that its inputs had already proven:

1. **Limits were revalidated on every call.** `Limits` is valid by
   construction: its fields are private and the only constructors are
   `Default` and the `with_*` methods, each of which returns
   `Limits::validate()`. The 31 internal `limits.validate()?` calls in the
   parser, constructor and serializer entry points are removed and the
   invariant is documented on the type. `Limits::validate` stays public.
2. **The default-profile certificates did per-call work or returned the whole
   record through the cold path.** `FieldPath::new` and `FieldInfo::new`
   now make one comparison against `DEFAULT_FIELD_PATH_COMPONENTS` and call
   the exact admission out of line only for its error, so the admitted record
   is built once in the caller's slot. The default profile is rebuilt inside
   the cold function, so a caller with a constant profile does not
   materialize the 88-byte `Limits` per call. The certificate bounds are
   compile-time assertions against the default ceilings.
3. **The raw `ArchiveInfo::new(Vec<MessageInfo>)` certificate compared the
   exact version values.** Inline version values do not change any admitted
   resource, so the certificate now reads each public field once (eight
   lengths/options, the spill flag and the length ceiling) with one combined
   branch per record. Records outside that shape, and every strict profile,
   take the exact admission with its original error precedence; a new test
   checks the certificate against the exact path for arbitrary inline values
   and every excluded field.
4. **The admitted collection did not carry its admission to the
   constructor.** A `MessageInfoCollection` filled only by minimal
   default-profile pushes now publishes without another arithmetic pass; the
   collection stores the default profile as a unit variant
   (`CollectionProfile::Default`) instead of a copied `Limits`, and
   `push_new` writes the conventional record in place.
5. **The retained-header comparison called `memcmp` per field.** The
   canonical-header probe compares each emitted key, varint or packed payload
   with the retained canonical bytes. Those slices are 1–10 bytes, and the
   libc call cost a mispredicted branch per message: retained validation
   measured 1.09× baseline wall time at 0.72× its instructions (on trunk as
   well). A branch-free OR-reduction for slices of at most 16 bytes brings
   validation to 0.59× instructions, the baseline's simulated 13
   mispredictions per operation, and 0.40× baseline wall time.

Two correctness repairs came out of this work:

- **Exact-ceiling refusal of a retained header (regression on trunk).**
  `retained_unknown_header_obeys_stricter_field_limits` failed on trunk: the
  cached admission totals refused correctly but reported the total field
  count (observed 3) where fresh ingress reports the first charge over the
  ceiling (observed 2). A refusal by the cache now replays the bounded wire
  preflight (cold path), so the retained header refuses with exactly the
  error its parse would.
- **Quadratic duplicate-identifier scan in the encoded-length preflight**
  (the donor's `ffda361b38`/`29070e91fb` concern, still present on trunk's
  `encoded_archive_len_preflight`). The first duplicate position is located
  before the ordered scan, pairwise for up to 32 objects and by sorting one
  fallibly reserved `(identifier, position)` vector otherwise; the reported
  object and the precedence of earlier errors are unchanged (new test against
  the ordered scan at 4, 40 and 200 objects).

No public signature changed, so no consumer needed migration; explicit
admission policies, exact-ceiling errors, unknown retained header bytes and
output bytes are preserved (the probe's eight output files are
byte-identical between baseline and candidate, §1.2).

### 1.2 Matched instruction counts (15 routes)

Probe [`constructor-probe.rs`](results/0367/constructor-probe.rs) (one source;
feature `before` adapts only the constructor signatures), fat LTO,
`panic = "abort"`, rustc 1.95.0. `Ir/op = (Ir(12,000) − Ir(2,000)) / 10,000`,
eight messages, script [`constructor-callgrind.sh`](results/0367/constructor-callgrind.sh).
Every row has equal output SHA-256 and checksum on both endpoints. Candidate
`f2e5e78985`.

| Route | Base Ir/op | Candidate Ir/op | Ratio |
|---|---:|---:|---:|
| field_path | 36.00 | 40.00 | 1.111 |
| field_info | 65.00 | 73.00 | 1.123 |
| message_info | 173.00 | 102.01 | 0.590 |
| archive_info (raw `Vec`) | 44.01 | 276.00 | 6.272 |
| archive_object | 14,094.99 | 1,722.01 | 0.122 |
| archive_object_reject | 3,476.00 | 328.00 | 0.094 |
| parse | 54,106.04 | 46,471.01 | 0.859 |
| validate | 9,487.00 | 3,165.00 | 0.334 |
| encode | 38,220.79 | 27,507.00 | 0.720 |
| retained_parse | 90,300.64 | 86,017.01 | 0.953 |
| retained_validate | 9,493.00 | 5,602.00 | 0.590 |
| retained_encoded_len | 14,719.00 | 5,870.00 | 0.399 |
| retained_encode | 46,974.02 | 30,242.00 | 0.644 |
| retained_validate_tight | 9,564.03 | 5,625.00 | 0.588 |
| retained_encode_tight | 46,155.31 | 30,299.00 | 0.657 |

The 0361fda return listed field paths 2.36×, field information 3.00×,
message information 2.65×, archive information 67.2×, retained
parse/validate/encoded-length/encode 1.141/2.075/1.170/1.246× and the
exact-ceiling validate/encode 4.706/2.312×. Twelve routes are now at or below
the baseline. The three above 1.05 are routes where the baseline did
strictly less work because it skipped a required check (§1.3).

Lifecycle routes (the same probe, construction included in every
iteration):

| Route | Messages 0 / 1 / 8 / 32 / 128 (ratio) |
|---|---|
| raw `Vec` assembly + `ArchiveInfo::new` | 1.038 / 0.992 / 0.848 / 0.390 / 0.243 |
| `MessageInfoCollection::push_new` + `ArchiveInfo::new` | 1.026 / 1.033 / 0.933 / 0.456 / 0.295 |
| `ArchiveObject::new` + serialize | 0.562 (8 messages) |
| `push(MessageInfo)` from a prebuilt `Vec` (double buffer) | 1.124 / 1.321 / 1.649 / 0.950 / 0.672 |

The `push(MessageInfo)` lifecycle copies an existing `Vec` record by record
into a second collection; a caller holding a `Vec` uses the raw constructor
(0.85× at eight messages), and no workspace consumer uses the collection.

### 1.3 Routes whose baseline skipped a required check

The baseline constructors were unchecked moves. Contract 0351 requires each
raw constructor to take a limit profile and refuse typed `Limit` errors, so
each needs at least one comparison and a fallible return:

- **field_path (+4 Ir/op), field_info (+8 Ir/op).** A diagnostic build of the
  candidate with the checks deleted but the same `Result` API (`floor` in the
  evidence; not a candidate) measures 35 and 76 Ir/op. The candidate's
  remaining cost is the one comparison with the component ceiling and the
  caller's `Result` discriminant test; `field_info` is below that floor. This
  is the cheapest correct implementation of the required check.
- **archive_info, raw `Vec` (6.27× at eight messages).** The input is a
  caller-owned, mutable `Vec<MessageInfo>`; admission must read every
  record's variable-size fields. The cost is `85 + 23.07 × messages` Ir
  (0, 1, 8, 32, 128 messages: 85, 115, 276, 828, 3,036), and the floor build
  without the scan is a constant 47. The 23-instruction loop body is one load
  per public field combined without a branch per field. Callers that own the
  construction use `MessageInfoCollection::push_new`, whose certified
  publication is constant. The raw lifecycle is at or below the baseline
  from one message and the `push_new` lifecycle from eight; at zero and one
  message they are 1.026–1.038× (24–37 Ir of fixed collection setup) (§1.2).

### 1.4 Allocations

Counting allocator, 100 operations, eight messages; retained bytes are zero
on every row.

| Route | Base bytes / calls / peak | Candidate bytes / calls / peak |
|---|---|---|
| field_path, field_info, archive_info (raw), retained_encoded_len | 0 / 0 / 0 | 0 / 0 / 0 |
| message_info | 1,200 / 100 / 12 | 0 / 0 / 0 |
| archive_object | 566,400 / 2,700 / 5,664 | 166,400 / 100 / 1,664 |
| archive_object_reject | 148,000 / 1,000 / 1,480 | 6,400 / 100 / 64 |
| parse | 2,286,200 / 7,800 / 14,708 | 1,886,200 / 6,000 / 10,612 |
| validate | 395,600 / 1,900 / 3,956 | 5,200 / 100 / 52 |
| encode | 2,016,900 / 5,700 / 14,069 | 1,616,900 / 3,100 / 13,973 |
| retained_parse | 2,337,600 / 8,100 / 15,176 | 1,968,200 / 6,400 / 11,280 |
| retained_validate (and _tight) | 395,600 / 1,900 / 3,956 | 5,200 / 100 / 52 |
| retained_encode (and _tight) | 2,042,600 / 5,700 / 14,326 | 1,642,600 / 3,100 / 14,230 |
| raw lifecycle, push_new lifecycle | 176,100 / 1,000 / 1,760 | 166,500 / 200 / 1,664 |

No route allocates more than the baseline except the two collection
diagnostics, which build a second vector by design.

### 1.5 Native observations (unqualified)

`perf` is blocked and no quiet host window existed (load 8–16 on 64 CPUs).
Tolerant ABBA ([runner](results/0367/constructor-tolerant-abba.py), CPU 29,
no clean block) medians: field_path 1.107, field_info 1.083, message_info
0.897, archive_info raw 3.364, archive_object 0.131, reject 0.063, parse
0.825, validate 0.272, encode 0.675, retained_parse 0.898, retained_validate
0.402 (trunk 1.093), retained_encoded_len 0.233, retained_encode 0.524,
retained_validate_tight 0.396 (trunk 1.098), retained_encode_tight 0.542,
push_new lifecycle 0.917 (trunk 1.448).

One divergence is reported, not hidden: the raw `Vec` lifecycle measures
1.324 in time at 0.848 in instructions. It predates this work (trunk 1.299;
the pre-SmallVec v24 candidate 1.908, v24 SmallVec 1.360). The floor build
without the record scan measures 1.133, the same as the vector construction
and drop alone (1.135), so roughly 0.13 belongs to the `SmallVec`
`MessageInfo` layout and 0.19 to the required scan. It is a lifecycle
diagnostic outside the 15 routes; the follow-up is a native profile once a
quiet window exists.

### 1.6 Owner routes and the 5-byte peaks

The 0361fd owner probe, built with [`owner-build.sh`](results/0367/owner-build.sh)
against trunk `1e78bb87b6` and this branch; per-op = `(N5 − N1) / 4`; output
bytes and digests identical on every route.

| Route | Ir/op trunk → candidate | Ratio | Heap bytes/op | Peak bytes |
|---|---|---:|---|---|
| numbers_headers | 18,394,551 → 18,282,280 | 0.9939 | 3,381,115 → 3,378,555 | 2,292,727 = |
| keynote_media | 80,458,623 → 80,299,327 | 0.9980 | 6,544,004 → 6,541,444 | 5,826,644 = |
| pages_footnote_insert | 1,475,287 → 1,457,324 | 0.9878 | 142,869 = | 31,187 = |
| pages_section_name | 18,084,948 → 17,930,929 | 0.9915 | 3,359,928 → 3,360,792 | 2,192,626 = |

`pages_section_name` requests 864 more bytes per operation (one extra
allocation): the sorted identifier index of §1.1 for its archives of more
than 32 objects. Tolerant ABBA medians: 0.993, 1.023, 0.992, 0.997.

**The 5-byte peaks are attributed.** In both 0361fd routes the only
allocation site whose bytes at the global peak differ is `owner_probe::main`
(+5 B): the probe's `args: Vec<String>` holds argv[0], and
`build-candidate/owner-probe` is 27 bytes against `build-base/owner-probe`'s
22. Rerunning the same protected 0361fd binaries (`851a090f…`, `47579bd2…`)
as `b/owner-probe` and `c/owner-probe` gives equal peaks: 32,675 and 32,675
(footnote insertion), 5,909,132 and 5,909,132 (Keynote media). The library
peaks were never different. Every measurement here uses equal-length binary
paths.

## 2. W0-A substrate (0363 §7)

### 2.1 Ported

| Donor | Port |
|---|---|
| `uo2/iwa-wire-budget-fix-20261024` `5753ea5058` | `overlay_singular_wire_fields` charges both sources before materializing field vectors and the merged output before allocating it; batched nested rewrites charge edit-trie construction, including sorted-insertion shifts (`6359749bf6`) |
| `8d6b0104d9`, `27375f4568` (parity part) | selected IWA members and the legacy `Index.zip` need matching local/central names, flags and compression before any payload read; logical metadata, semantic-supported and semantic metadata authorities also require matching flags (`be3df8c376`) |
| `b17932c50e` (classifier) | application classification keeps borrowed slots 1..=15 plus a duplicate bit set in one visit instead of a field vector per scan, with the old error precedence; allocator-observed test ported (`1a73fb9f59`). New `wire::visit_wire_fields` applies exactly `parse_wire_fields`' rules without collecting (`0db613b7b6`) |
| `27375f4568` (descriptors) | `CandidateEntry<'data>` is bound to its source (compile-fail doc test) and `admit_entry` refuses foreign or forged descriptors before any charge, looking the central record up directly (`ZipSliceArchive::central_record_at`); Deflate admission charges the decoder scratch envelope (`DEFLATE_READ_SCRATCH_*`) and metadata traversal (`a81938aa97`, `b3152ee670`) |
| `c09d393684` (archive part) | `CandidateEntry::is_operation_storage` by exact portable basename (`a96a631f9a`) |
| `ffda361b38`, `29070e91fb` (duplicate scan) | see §1.1 (`f2e5e78985`) |

### 2.2 Raw physical access (D5)

`litchi_iwa_archive::unstable` now owns `entry_names`, `iwa_entry_names`,
`entry`, `normalize_entry_name`, `parse_component`/`cached_component` with
a typed `ComponentReadError`, and `count_message_types`,
`most_common_message_type` and `message_type_summary` (`38f35f9f27`). The
host's `Package::{entry_names, iwa_entry_names, entry, archive}` and
`Document::stats` with `DocumentStats::{most_common_message_type,
message_type_summary}` delegate to them until W3; each typed refusal maps to
the exact host error it returned before, and the Snappy profile is still
derived only for an existing, non-operation-log member (`6e4a5ffcc9`).

Parity (`6a307a3693`): read-only Tier A cases on native-saved Numbers, Pages
and Keynote fixtures compare every host read with the same read through the
unstable functions over a package state built from the leaf `SourceCatalog`
(`crates/litchi-iwa/tests/parity/physical_cases.rs`). The eleven ledger rows
are `migrated` with the Numbers fixture's native receipt as evidence. These
routes write nothing, so owner decision D1 adds no macOS queue row. The
header-parity refusals of §2.1 change no output for a native file (Apple's
writers emit matching local and central headers); they refuse only
inconsistent inputs and need no queue row either.

### 2.3 Not ported (follow-ups)

| Donor | Reason |
|---|---|
| `3d58ce1219` reader names in one sorted pass | superseded: trunk's reader index is a `HashMap`, so the quadratic scan does not exist |
| `72e23f9c15` owned component extraction | already on trunk |
| `29070e91fb` serialization resource reports (core, Snappy, protos) | the donor's allocation census models its own encode path (metadata clone, compatibility and Buffa projections); trunk's retained-provenance encoder allocates differently, so the report must be re-derived and allocator-verified for trunk |
| `6ab75a4562`, `ab12a24fee` writer metadata admission | new report APIs whose consumers are the W1 format ledgers; port with the first consumer |
| `f779cd0fac`, `ea881e16db` catalog storage preflight | trunk has its own census (`fb039e1f7a`, `cafcb5b186`); the donor's logical/nested-token extensions conflict across five files and need a separate port |
| `a4082e7efc` identity-clone admission | conflicts heavily with trunk's core `archive.rs`; its wire part equals the ported budget fix plus a `_with_limits` parser (banned name) |
| `8f04527c92` fuzz-target ownership | moves files out of `crates/litchi/fuzz`, outside W0-A's paths |
| `5818f7613f` wire preflight views | consumed only by the `numbers-wire` formula graph donor (W0-B) |

## 3. Gates

All commands ran in the worktree with the W0-A env; each exit status was
checked.

| Gate | Exit | Result |
|---|---|---|
| `cargo fmt --all --check` | 0 | |
| `cargo clippy` `-p` core, common, archive, detect, soapberry-zip, litchi-iwa, pages, keynote, numbers, litchi-ops, litchi, litchi-py `--all-targets --all-features -D warnings` | 0 | |
| `cargo clippy -p litchi-ops --no-default-features` none, pages, keynote, numbers, all three | 0 ×5 | |
| `cargo test` core / common / archive / detect / soapberry-zip `--all-features` | 0 ×5 | 126 / 227 / 213 / 43 / 663 passed |
| `cargo test` the nine other IWA substrate crates and `litchi-numbers-wire` | 0 | 1,411 passed, 219 ignored |
| `cargo test` pages / keynote / numbers `--all-features` | 0 ×3 | 716 / 1,174 / 1,099 passed |
| `cargo test -p litchi-iwa --all-features` (incl. `parity`) | 0 | 1,807 passed |
| `cargo test -p litchi-ops --all-features`: lib + six iWork targets + capabilities + registry; the other 159 targets in five chunks; doc | 0 ×7 | 1,110 + 1,397 + 217 passed; goldens unchanged |
| capabilities/registry with every feature but `encryption` | 0 | 14 passed, goldens unchanged |
| `cargo test -p litchi --all-features` | 0 | 209 passed |
| `RUSTDOCFLAGS=-Dwarnings cargo doc --no-deps --all-features` (six touched crates) | 0 | |
| `RUSTC_BOOTSTRAP=1 tools/check_api_glossary.py`, 19 profiles (soapberry-zip, every iwa-*, numbers-*, ops-iwork, facade-iwork, pages, keynote-*) | 0 ×19 | |
| `check_crate_boundaries.py`, `non_iwork_gate.py verify` | 0, 0 | |
| tool unit tests (five modules + `test_iwa_exit_ledger`) | 0 | 1,289 tests |
| `officecli_capability_delta.py --check`, `officecli_corpus.py --check`, `check_feature_matrix.py`, `iwa_exit_ledger.py --check`, `check_iwork_leaf_public_api.py` | 0 ×5 | |
| `unified_bench_matrix.py --check` | 0 | after regenerating (stale on trunk after 0369) and an `evidence_without_rows` entry for 0367 |
| `crates/pyo3-litchi/tests/run_tests.sh` | 0 | 301 tests |
| `tools/perf-baseline`: `cargo check --locked --all-targets` | 0 | |

## 4. Evidence

Committed (small): [`constructor-evidence.json`](results/0367/constructor-evidence.json)
(every callgrind row, floor rows, allocation rows, ABBA blocks, binary and
lock hashes), [`owner-routes.json`](results/0367/owner-routes.json), the probe
source and the three scripts. Binaries, captures and source snapshots stay
outside the repository under
`/home/zhuhe/code/litchi-worktrees/.coord/iwa-w0a/{ctor,owner}`; the JSON
files record their paths and SHA-256 (candidate callgrind binary `cce3f2f5…`,
baseline `e82c777b…`).
