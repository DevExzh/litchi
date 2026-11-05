# 0311 — Remaining single-op misses: the OPC double audit, the DOC adapter, the ODP and ODT wall-time puzzles

- Status: implemented on `uo2/perf-opc`, cut from `feat/unified-ops-api`
  at `1178eb235c`; trunk `a41adc99cf` (record 0309) merged (`4171af6671`;
  the merge touches no file this record changes), then `8f5569d0d8`
  (`PROGRESS.md` only, `359704aab0`).
- Date: 2026-10-16
- Inputs: record 0305 §4.5 (DOCX 1,000-paragraph set 1.56×, add 1.46× on
  the default source-backed owner), record 0308 §7 (DOC unified/leaf 1.56;
  ODP `Scene::parse` +3.4 % / +6.6 % wall with −0.14 % instructions),
  record 0304 §5 (ODT plain-paragraph set +3.4–4.7 % wall with +0.22 %
  instructions).
- Scope: `litchi-opc` (`source_backed.rs`, `source_backed/revision.rs`),
  `litchi-docx` (`document/transaction.rs`, `source_backed{,/owner}.rs`),
  `litchi-doc` (`package/{model,codec,edit}.rs`), new isolated examples in
  `litchi-odp` and `litchi-odt`, and new legs/scenarios (new names only) in
  `litchi-ops` `examples/{docx_large_bench,office_gaps_bench}.rs`. No
  adapter under `litchi-ops/src/format/` needed a change.
- `performance_claim`: §3. `profiling` builds (release with debug info,
  stripped for timing) of
  `cargo build --profile profiling -p litchi-ops --features docx,doc,odp,xls,rtf,odt,pptx,ppt --example <name>`,
  pinned with `taskset -c 30`, processes B A A B. The base build is this
  branch with `crates/{litchi-opc,litchi-docx,litchi-doc}/src` reverted to
  `1178eb235c`, so it has the new legs. Instruction and cache counts come
  from callgrind with `--cache-sim=yes --I1=32768,8,64 --D1=32768,8,64`
  (`perf` is blocked, `perf_event_paranoid=4`). Evidence:
  [results/0311/](results/0311/): `measure.sh`, `docx_cg.sh`,
  `summary.txt`, the raw harness outputs, `docx-cachesim.txt` and
  `binaries.sha256`. Host: a shared 32-core EPYC with load 5–6.

## 1. Summary

| Item | Before | After | Status |
|---|---:|---:|---|
| DOCX 1,000 ¶ text set, default owner / leaf `replace_paragraph_text` | 1.57 (1,206–1,208 µs) | **0.54** (410–416 µs) | **fixed** |
| DOCX 1,000 ¶ add, default owner / leaf `insert_paragraph` | 1.47 (1,641–1,646 µs) | **0.40–0.47** (447–517 µs) | **fixed** |
| the same add / the eager owner's add | 3.89 | 1.05–1.22 | residual, explained in §4.1 |
| DOCX set on a handle already committed from the default owner | 1,906–1,907 µs (2.48× leaf) | 417–419 µs (0.55×) | **fixed** (new leg) |
| DOC text set / the `Package::edit` leaf it wraps | 1.025–1.031 | 1.029–1.040 | within budget, before and after |
| DOC text set / the body-text snapshot leaf (`doc_text_set`) | 1.69–1.73 | 1.12–1.27 | improved; the residual is the readable package the leaf does not produce (§4.2) |
| ODP `Scene::parse` (0308) | +3.4 % / +6.6 % in `office_gaps_bench` | isolated: −0.5 to −1.7 %; aligned: +0.3 % | **not a regression**: code layout (§2.4) |
| ODT plain-paragraph set (0304) | +3.4–4.7 % in `text_set_bench` | isolated: −1.5 to −2.5 %; aligned: −0.5 % | **not a regression**: code layout (§2.5) |

All output stays byte-identical. The query-corpus differential, corpus
fidelity, the checklist goldens, the capability and registry goldens
(both feature sets) and the OfficeCLI corpus pass unchanged. Nothing was
re-blessed (§5).

## 2. Findings and changes

### 2.1 DOCX on the default owner: litchi-opc audited the same pair twice

Callgrind of the base `auto_set` leg, at 28.7 M instructions per call
against 17.1 M for the leaf:

| pass | per call |
|---|---:|
| the commit's source gate, `verify_source(base)` (`publication_accepts_preserved_xml`) | 9.1 M |
| `PreparedTopology::apply_xml_bytes` → `validate_source_part_replacement_xml` (original + replacement pair) | 9.56 M |
| the same candidate's `prepare_topology` → `prepare_topology_state` → `audit_source_part_replacement_xml`: **the same pair again** | 9.56 M |
| `Docx` resolution of `/body/p[501]` (`load_body`) | 1.3 M |

The unmanaged source-backed commit publishes through `apply_xml_bytes`, not
`apply_source_xml`. Record 0305 named `apply_source_xml`, but that is the
managed route, and it re-audits nothing because its `SourceXmlPart` proof
is checked by address. `apply_xml_bytes` audited the pair. It then put the
bytes into an ordinary plan replacement, and publishing that plan audited
the pair a second time. The commit had also audited the base: the eager
route gets a `VerifiedSource` proof from that pass, but the identity-bound
route discarded it.

Changes:

- **litchi-opc** (`83fabde8bd`):
  - New `PreparedTopology::apply_verified_xml_bytes(partname, expected,
    replacement: xml_minifier::audit::VerifiedSource)`. The token is
    `xml-minifier`'s existing proof type, and only running the audit can
    create one. Its constructors are `VerifiedSource::verify` and
    `verify_replacement`. The proof replaces auditing the replacement only
    when `covers` accepts the exact allocation under the default limits.
    A proof made under any other limits is ignored, and those bytes are
    audited in full. `apply_xml_bytes` keeps its signature and contract.
  - Both entry points share `audit_physical_replacement_xml`. It keeps the
    verdict and error of the pair audit: the source audit of the original,
    then that of the replacement.
  - The verdict on the physical original is remembered per Part and source
    revision, on unmanaged packages only, since managed packages retain no
    unreserved memory. The bytes of one revision cannot change, and a
    failure is never remembered. So an original that fails is audited, and
    reported, on every call.
  - Once the original has passed, an unproven replacement gets its own
    `verify_source` instead of the pair. An insertion has no provable
    window, so before this the pair was two full scans.
  - `apply_xml_bytes` passes the proof its own audit makes into the plan
    (`SourceTopologyPlan::try_replace_part_verified`, crate-private), so
    publishing the candidate no longer audits the pair a second time.
  - Tests: `prepared_topology.rs` checks that proven and audited
    replacements publish identical bytes, fresh and chained; that a stale
    precondition is refused; that a proof under other limits changes
    nothing; that a failing original is refused with the same error, twice,
    even with a proof; and that a remembered original still refuses every
    invalid unproven replacement with a fresh package's error. Lib test
    `physical_replacement_audit_remembers_only_passing_originals`.
- **litchi-docx** (`ac692e5e2f`):
  - `commit_preserving_unmodified` gives a source-identity snapshot the same
    pair audit as an identity-free one
    (`VerifiedSource::verify_replacement`). Its verdict on the base is the
    historical gate's `verify_source`, so the route cannot change. The
    commit therefore carries the proof.
  - `source_backed::Package::apply` publishes through
    `apply_verified_xml_bytes` when the target has a proof, and through
    `apply_xml_bytes` otherwise.
  - Test `a_source_with_identity_is_proven_for_the_source_backed_writer`
    (formerly `…_keeps_the_historical_gate_and_builds_no_proof`) checks the
    proof and the historical route's bytes.

### 2.2 DOCX: the effective snapshot memo

On a handle committed from a source-backed owner, an edit goes through
`effective_document_snapshot`. That function copied the effective bytes and
scanned them, and the commit's `apply` scanned the same revision again. The
unmanaged branch now serves and fills the owner's `main_snapshot` memo, with
the 0305 rule: identity matches and bytes are equal (`ac692e5e2f`). Test
`committed_handles_share_one_effective_transaction_snapshot` checks four
things: the memo holds the effective revision, it never answers for the
physical bytes, a chain from the memo publishes the bytes of a chain from
fresh scans, and the memo starts empty on each committed owner. The
`committed_set` leg went from 42.96 M to 10.22 M instructions per call.

### 2.3 DOC: `Package::edit` copied and re-parsed the file on every edit

Callgrind of `doc_text_set` (base) put the unified set at 1.76 M
instructions per call, against 1.13 M for the body-text snapshot leaf. The
difference, per call:

- `Doc::native_edit` → `Package::edit_with_options` copied the whole file
  (`Package::to_bytes`) and opened a fresh `body_text::Snapshot` from it:
  0.51 M plus the copy. The leaf leg opens its snapshot once.
- The package commit (`commit_with_package`) parses the published bytes
  into the `Document` the new package reads: 0.59 M. The body-text leaf's
  commit makes only the next body snapshot: 0.51 M.
- The adapter's own work (resolution, coercion, admission) was under
  0.01 M.

So the adapter adds almost nothing over the leaf API it wraps. The new
`doc_package_text_set` scenario shows this: unified against
`Package::edit` + `replace_paragraph` + `commit` was 1.025–1.031 before
any change. The cost sat in the leaf hook.

`litchi-doc` `Package` now keeps the last body snapshot it opened, with the
`body_text::OpenOptions` it was opened under, shared by clones
(`0a2fbbc845`). An edit under equal options starts its transaction from
that snapshot. It first checks that the retained source still has its
captured version, which is the check the copy made. A snapshot is
immutable, and the memo covers only the bytes of the `Arc<SharedOleFile>`
it shares with its clones. Test
`edits_share_the_body_snapshot_opened_under_their_options` checks four
things: the memo is reused, including from a clone; other options open
their own snapshot; the memo's bytes are the package's; and a fresh open
publishes the same bytes. The cost is the memory of one retained snapshot
per opened package, which the body-text leaf holds too.

### 2.4 ODP `Scene::parse`: code layout, not work

The 0308 numbers compared two different `office_gaps_bench` binaries. To
separate code from layout, `crates/litchi-odp/examples/scene_parse_bench.rs`
times the same scene alone in its own binary (`6cb7e3c926`). It was built
twice from this branch: once as is, and once with `crates/litchi-odp/src`
at the pre-0308 trunk `6c8d517eff`, so nothing else differs.

- **Wall, isolated** (B A A B × 3): before 442.0–447.6 µs, now
  438.2–443.8 µs. The new code is 0.5–1.7 % *faster*.
- **Instructions** (callgrind, per parse): 9.669 M → 9.628 M (−0.43 %).
  The styles subtree is now opaque to the scanner.
- **Cache simulation**: D1 read and write misses are unchanged (7.6 K and
  6.5 K). Simulated L1i misses rise from 40.6 K to 48.0 K. Inlining the new
  `open_style_element` made `Scanner::open` larger, and `open` runs for
  every element. The wall time does not follow this. Marking the helper
  `#[inline(never)]` cut the simulated misses to 45.8 K but was slower in
  wall time (448–451 µs against 442–444 µs). It was rejected.
- **Layout control**: both trees rebuilt with 64-byte function and
  block alignment (`-C llvm-args=-align-all-functions=6
  -C llvm-args=-align-all-nofallthru-blocks=5`): before 440.4–443.9 µs,
  now 443.6–446.2 µs. That is within 0.5 %.
- **Inside the large binary**: the same scene code runs at 482–488 µs in
  both this record's base and after `office_gaps_bench` builds, whose scene
  code is identical, against about 443 µs isolated. Placement inside the
  large binary alone moves this loop by about 9 %, which is more than the
  differences 0308 saw.

Verdict: no added work, and no regression once layout is held fixed. No
change was made.

### 2.5 ODT plain-paragraph set (0304): the same method, the same answer

`crates/litchi-odt/examples/plain_set_bench.rs` times the leaf leg of
0304's `odt_plain` alone: `Edit::set_text` + `commit` on the middle
paragraph of 1,000 (`6cb7e3c926`). It was built from this branch as is,
and with 0304's two `litchi-odt` commits (`59e515dc72`, `c8b3585325`)
reversed. No other commit has touched `litchi-odt/src` since.

- Wall, isolated (B A A B × 3): before 427.7–436.7 µs, now 419.0–424.7 µs
  (now 1.5–2.5 % faster).
- Aligned builds: before 418.2–420.6 µs, now 416.2–417.7 µs (−0.5 %).
- Callgrind per call: 12.423 M → 12.403 M instructions (−0.17 %); L1i
  misses 35.9 K → 36.2 K; D1 misses unchanged.

Verdict: the +3.4–4.7 % came from the layout of the `text_set_bench`
binary. The 0304 code is not slower. No change was made.

## 3. Measurements

### 3.1 DOCX, `docx_large_bench` (µs per call; B A A B)

| leg | base 1 / 2 | after 1 / 2 |
|---|---:|---:|
| leaf text set | 769.9 / 769.2 | 766.8 / 756.3 |
| unified set, default owner | 1,205.9 / 1,208.0 | **415.9 / 410.3** |
| unified set, eager owner | 771.6 / 769.4 | 761.6 / 820.9 |
| unified set, committed handle (new) | 1,907.3 / 1,906.4 | **418.5 / 416.6** |
| leaf `insert_paragraph` | 1,116.7 / 1,120.4 | 1,106.1 / 1,105.3 |
| unified add, default owner | 1,646.1 / 1,640.6 | **446.5 / 516.7** |
| unified add, eager owner | 422.9 / 420.1 | 424.0 / 423.9 |
| unified add, committed handle (new) | 1,581.2 / 1,583.7 | 459.8 / 531.2 |

Callgrind per call ((15 − 5 loops) / 10, `docx-cachesim.txt`):

| leg | base Ir | after Ir | after L1i misses | after D1 read misses |
|---|---:|---:|---:|---:|
| leaf set | 17.06 M | 17.05 M | 2,620 | 23,102 |
| default-owner set | 28.66 M | **9.99 M** | 4,845 | 22,458 |
| committed set | 42.96 M | **10.22 M** | 3,444 | 41,635 |
| eager add | 8.08 M | 8.09 M | 2,939 | 12,818 |
| default-owner add | 38.22 M | **9.98 M** | 5,274 | 16,808 |

The default-owner set now costs less than the leaf. The leaf leg scans its
document snapshot on every edit (7.15 M), while the source-backed owner
serves it from the 0305 memo. Its commit pays one pair audit (9.56 M),
which the leaf's commit pays too. Its overlay pays nothing more.

### 3.2 DOC, `office_gaps_bench` (in-process ABBA; µs)

| scenario | run | leaf | unified | ratio |
|---|---|---:|---:|---:|
| `doc_text_set` (body-text leaf) | base 1 / 2 | 57.8 / 58.1 | 98.0 / 100.6 | 1.694 / 1.732 |
| | after 1 / 2 | 57.8 / 55.0 | 73.4 / 61.4 | 1.269 / 1.118 |
| `doc_package_text_set` (`Package::edit` leaf, new) | base 1 / 2 | 92.8 / 98.6 | 95.7 / 101.0 | 1.031 / 1.025 |
| | after 1 / 2 | 72.7 / 74.3 | 74.8 / 77.2 | 1.029 / 1.040 |

Instructions per call (callgrind, `doc_text_set`): unified 1.759 M →
1.199 M; body-text leaf 1.126 M. The `Package::edit` leaf itself got 22–25 %
faster, from the same memo.

## 4. Residuals

### 4.1 DOCX add against the eager owner: 1.05–1.22×

The unified add on the default owner is 0.40–0.47× the leaf
`insert_paragraph` leg, so the record's target is met. Against the eager
owner's add, it costs 9.98 M against 8.09 M instructions per call.

An insertion commits on the preserved-markup route, which compacts nothing
and so audits nothing in the transaction. The candidate therefore carries
no proof. The source overlay must audit it before the revision exists, and
with the original's verdict remembered it pays one full `verify_source` of
the replacement (9.1 M). The eager owner's commit stores the copied bytes
unaudited, and its package writer audits them when it writes the package.

The overlay cost could move to publish time only by deferring the overlay's
validation. That is a contract change to `litchi-opc`'s prepared topology,
not a performance fix. Two narrower routes also stay closed:

- A window proof for a pure insertion would need `xml-minifier` support,
  which is outside this record's scope.
- A proof minted by the insert's commit would cost the same audit, just
  moved into the commit.

### 4.2 DOC against the body-text leaf: 1.12–1.27× wall, 1.065× instructions

Unified against the `Package::edit` leaf it wraps: 1.03–1.04×, inside the
single-op budget. The body-text leaf's commit publishes only the next body
snapshot. The package commit also parses the published file into the
`Document` the returned package reads: `Document::from_ole_with_options`
0.24 M, the package adoption 0.08 M, the revision editor's reopen 0.07 M.
This parse is the commit's validation of the output, so skipping it would
weaken the commit. The wall-time gap is larger than the instruction gap
because of L1i misses: 12.4 K against 10.8 K per call, from the second
parser's code.

### 4.3 What stays layout-sensitive

The ODP scene parse moves by about 9 % with its placement in a large
binary, and the ODT set by a few percent. Before/after comparisons of these
loops across two large example binaries cannot separate a 3–6 % change from
layout. The isolated examples of §2.4–2.5 are the comparison to use.

## 5. Gates

Every command below ran on the merged tree (`4171af6671` + this record),
one cargo command at a time, with the build environment and the
memory-capped cargo shim of the rules, and test executables pruned between
chunks. Every exit status was checked.

| Command | Exit |
|---|---:|
| `cargo fmt --all --check` | 0 |
| `cargo clippy -p litchi-opc -p litchi-docx -p litchi-doc -p litchi-odp -p litchi-odt -p litchi-ops -p litchi -p litchi-py --all-targets --all-features -- -D warnings` | 0 |
| `cargo clippy -p litchi-ops --no-default-features [--features F] --all-targets -- -D warnings`, F = none, docx, doc, odp, odt | 0 each |
| `cargo test --all-features` of `litchi-opc` (1,182), `litchi-docx` (2,400), `litchi-doc` (1,481), `litchi-odp` (439), `litchi-odt` (1,122) | 0 each |
| `cargo test -p litchi-ops --all-features`: `--lib` (918), `--doc` (210), the 137 test targets in seven groups of up to 20 (1,248) | 0 each |
| `cargo test -p litchi-ops --no-default-features --features` (every feature but `encryption`) `--test capabilities --test registry` (not re-blessed) | 0 |
| `cargo test -p litchi --all-features` (230) | 0 |
| `cargo test -p litchi-cli --all-features --test officecli_corpus` (not re-blessed) | 0 |
| `RUSTDOCFLAGS=-Dwarnings cargo doc --no-deps --all-features -p litchi-opc -p litchi-docx -p litchi-doc -p litchi-odp -p litchi-odt -p litchi-ops` | 0 |
| `python3 tools/check_crate_boundaries.py` | 0 |
| `python3 tools/non_iwork_gate.py verify` | 0 |
| `RUSTC_BOOTSTRAP=1 python3 tools/check_api_glossary.py --target-dir target/glossary` (40 profiles, 2,306 bans, 227 families) | 0 |
| `python3 tools/test_no_3rdparty_references.py` | 0 |
| `crates/pyo3-litchi/tests/run_tests.sh` (279 tests) | 0 |
| `cd tools/perf-baseline && cargo check --locked --all-targets` | 0 |

No capability, registry or OfficeCLI-visible behaviour changed, so neither
the delta table nor the OfficeCLI corpus was regenerated. The one new
public item, `litchi_opc::PreparedTopology::apply_verified_xml_bytes`, is a
low-level OPC method outside the unified vocabulary, and the glossary gate
passes. No glossary row changes.

## 6. Not claimed

- That a source-backed DOCX add matches the eager owner's add (§4.1).
- That the DOC set matches the body-text snapshot leaf (§4.2). The leaf it
  wraps is matched.
- Any wall-time comparison of the ODP or ODT loops across two large
  binaries.
- Any change of a frozen route name. `docx_large_bench` gained the
  `committed_*` legs and `office_gaps_bench` gained `doc_package_text_set`.
