# 0368a — W0-B follow-up: the drawable and cell-tile regressions, the map-aware owner allocator, the theme reference encoder

- Status: landed on branch `uo2/iwa-w0b` (trunk `f35f1c9f55`, the 0368
  merge, merged first).
- Date: 2026-10-26.
- Input: coordinator instruction after the 0368 merge — fix the iWork
  regressions 0368 listed instead of carrying them.
- `performance_claim`: every route 0368 measured is now within 5% of the host
  or cheaper (table in §5).

## 1. Drawable decode (was +205%)

`shape_properties_codec::decode_drawable_archive` ran the rewrite scan: it
reserved a field table sized to the source length for every message, kept
all tables, and then decoded the source twice more through the Buffa lazy
views as a cross-check. The read path now streams: `stream_drawable`,
`stream_geometry` and `stream_fixed_pair` visit each field once through the
same strict framing (`stream_message`, which `parse_message` now wraps) and
keep only the five known slots. Framing, canonical keys and lengths,
duplicate known fields, finiteness, non-negative sizes and UTF-8 are checked
exactly as before, every field and byte is still charged to the caller's
`DecodeOptions`, and the decode allocates nothing (its `DecodeReport` now
reports zero allocations and scratch bytes). `read_varint` takes a one-byte
fast path. The rewrite path keeps its field tables and the generated
cross-check; `buffa_traversal_limits_are_charged_before_projection` now
asserts the precharge on `prepare_drawable_archive_rewrite` and that the read
succeeds under the same eight-allocation budget.

## 2. Cell-tile remap (was +35.5%)

- `formula_graph::Remap::new` now checks the three maps once (sorted, no
  zero, no repeated source or target) and returns `Result<Remap, Error>`;
  the rewrites no longer re-validate (and re-sort a copy of) every map per
  call. This also helps the owner rewrite.
- `remap_cell_record_tile` no longer re-decodes its result in release
  builds: the rewrite replaces only owner-identifier varints that the strict
  source decode admitted, with checked non-zero `u32` targets, so the
  result's structure is the source's. Debug builds still re-decode.

## 3. The map-aware formula-owner allocator

`litchi_numbers_wire::formula_graph::next_owner_id(engine, owners, &Limits)`
returns one above the largest internal owner identifier in use: every owner
payload's field 2 and every owner-map entry and unregistered identifier of
the engine's tracker. The host's `next_internal_owner_id` counts owner
payloads only.

**Host defect (recorded, not reproduced by the leaf).** On
`chart-arrangement-native.numbers` the host's `duplicate_table` registers an
internal identifier the owner map already holds (a repeated map entry, a
corrupt engine). On `shared-model-native.numbers` the host's duplicate route
refuses earlier (its table-appearance read); its clone primitive collides in
the same way.

**Documented intentional difference.**
- `crates/litchi-iwa/tests/parity/numbers_owner_allocation.rs` (parity
  target): on both fixtures the leaf allocation lies above every mapped
  identifier; on chart-arrangement every host duplicate that succeeds leaves
  a repeated map entry below the leaf's allocation; on shared-model the host
  duplicate refuses. It is not a Tier A/B case — the outputs differ by
  design.
- `formula_clone/exit_parity.rs::leaf_allocator_avoids_the_host_owner_map_collision`:
  on every Numbers fixture the leaf and host allocators agree except on these
  two, where the leaf registration is valid and, under the host's colliding
  allocation, the leaf refuses (`InvalidRemap`). The Tier A rewrite parity
  cases now allocate with `next_owner_id`.
- `docs/unified-ops/0363-mac-queue.md`: one W0-B row (the allocator, adopted
  by the W1-A duplicate route; output differs from the host).

## 4. `ThemeReferenceSnapshot::encode`

In `theme_codec` and `keynote_theme_codec`, `encode` returned `Vec<u8>` and
panicked on a zero identifier, with `encode_checked` as the fallible twin.
`encode` now returns `Result<Vec<u8>, DecodeError>` (a typed invalid-source
error for identifier 0); `encode_checked` is deleted and banned in the
`iwa-protos-wire-codecs` glossary profile.

## 5. Performance

Callgrind, release, CPU 3 pinned, Ir, ABBA probes of 0368 (each route run
twice).

| Route | 0368 host | 0368 leaf | 0368 Δ | 0368a host | 0368a leaf | 0368a Δ |
|---|---:|---:|---:|---:|---:|---:|
| formula-graph clone (36 families) | 110,336,488 | 90,720,201 | −17.8% | 109,809,428 | 87,026,639 | **−20.7%** |
| · owners (792) | 83,524,376 | 69,034,032 | −17.3% | 82,806,394 | 66,391,141 | −19.8% |
| · engine registration (72) | 17,821,424 | 18,376,322 | +3.1% | 17,965,064 | 18,267,328 | +1.7% |
| · cell tiles (162) | 2,349,699 | 3,182,678 | +35.5% | 2,392,855 | 2,071,964 | **−13.4%** |
| drawable decode (315 drawables) | 1,773,702 | 5,406,830 | +205% | 1,687,092 | 1,440,494 | **−14.6%** |
| theme decode (68 themes) | 11,149,785 | 11,500,972 | +3.1% | 10,969,483 | 11,373,032 | +3.7% |

The leaf probe now builds one checked `Remap` per family, as the host builds
its maps once. Host counts differ slightly between the two sessions (heap
layout of the ABBA run); each Δ compares legs of the same run.

## 6. Verification

| Command | Exit |
|---|---|
| `cargo fmt --all --check` | 0 |
| `cargo clippy -p litchi-iwa-protos -p litchi-numbers-wire -p litchi-iwa-common -p litchi-numbers -p litchi-pages -p litchi-keynote -p litchi-iwa -p litchi-ops -p litchi -p litchi-py --all-targets --all-features -- -D warnings` | 0 |
| `cargo clippy -p litchi-ops --no-default-features [--features numbers\|pages\|keynote] --all-targets -- -D warnings` | 0 ×4 |
| `cargo test --all-features`: protos 1,084, numbers-wire 209, iwa-common 253, numbers 1,073, pages 716, keynote 1,174, litchi-iwa 1,811 (parity target 18), litchi 209 | 0 each |
| `cargo test -p litchi-ops --all-features` in chunks | §6.1 |
| remaining gates | §6.1 |

### 6.1 Remaining gates

| Command | Exit |
|---|---|
| `cargo test -p litchi-ops --all-features`: `--lib` (1,022), `--doc` (217), nine `--test` groups (1,485) | 0 each |
| `RUSTDOCFLAGS=-Dwarnings cargo doc --no-deps --all-features -p litchi-iwa-common -p litchi-numbers -p litchi-numbers-wire -p litchi-iwa-protos` | 0 |
| `RUSTC_BOOTSTRAP=1 python3 tools/check_api_glossary.py` (71 profiles, 3,909 banned names) | 0 |
| `python3 tools/check_crate_boundaries.py`, `non_iwork_gate.py verify`, `iwa_exit_ledger.py --check` | 0, 0, 0 |
| tool unit tests (1,289, incl. `tools.test_iwa_exit_ledger`) | 0 |
| `officecli_capability_delta.py --check`, `officecli_corpus.py --check`, `unified_bench_matrix.py --check` (after regenerating) | 0, 0, 0 |
| `python3 tools/check_feature_matrix.py` (31 legs) | 0 |
| `crates/pyo3-litchi/tests/run_tests.sh` | 0 |
| `cd tools/perf-baseline && cargo check --locked --all-targets` | 0 |

Glossary: `next_owner_id` gains a `required_signatures` row; `encode_checked`
is banned in `theme_codec` and `keynote_theme_codec`; the 0368 rows note the
checked `Remap::new` and the fallible reference encoder. No capability,
corpus or crate-edge change.
