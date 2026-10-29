# W2-B Pages chart-font facade (draft)

## Scope

This slice adds selector-first Pages body-chart font identity/face-trait and
point-size routes over the shared `litchi_iwa_style::chart_font` graph engine.
The public surface is `Package::{body_chart_font, edit_body_chart_font,
apply_body_chart_font}` and the corresponding `_font_size` methods. Native
object identifiers remain private to the Pages package adapter.

The implementation owns exact immutable snapshots, semantic `Edit` verbs,
exact-source `Patch`/`inverse`, no-op preservation, candidate reopen/readback,
ZIP/component/object locality proofs, and typed graph/wire/archive limits. The
engine limits are derived from the opened package's physical ceilings; this
facade does not replace caller/package admission with a payload-only
`Limits::for_source` profile.

## Proof status

Host-authored source-built Pages inputs from `chart_style_cases::input` select
the newly added Column2d chart at body-chart position zero. The facade proof
covers direct host readback before the leaf read, font set/reset, point-size
set/reset, preservation of the other property, exact inverse/no-op bytes, and
candidate semantic readback. The existing native fixture's mixed-font and
mixed-size chart remains a typed refusal control.

The six facade routes remain host-only in the exit ledger until the final
integrated performance qualification and native mutation evidence. The Mac
queue remains pending; synthetic/self-round-trip evidence does not claim
native acceptance.

The final scoped proof also compares host and facade reads after host-authored
font/size setup, host and facade set/reset readback, and the preserved
unmodified property. A stale exact-source patch is refused, an out-of-range
selector is a typed `ChartNotFound`, and a tightened source ceiling returns a
typed finite resource limit before publication. The mutation proof compares
ordered ZIP member names and confirms both changed and untouched members; the
production locality verifier additionally checks untouched objects and raw ZIP
records. Its focused unit proof injects distinct unknown fields into one
candidate-changed and one untouched source object, checks both survive the
candidate, and proves that a negative unknown-field mutation on the untouched
object is rejected by that production verifier.

## Gaps and boundaries

- No native E3/E4 evidence is claimed; the routes remain queued in
  `docs/unified-ops/0363-mac-queue.md`.
- Full release/ABBA facade measurements are still required on the final
  integrated candidate. Existing engine probes are not facade qualification.
- Unified `litchi-ops` dispatch/schema/capability registration is the separate
  D7 follow-up.
- Caption, legend, value-axis and lifecycle families are outside this slice.

## Verification

The final branch source was checked with:

- `cargo fmt --all`
- `cargo check -p litchi-pages --lib --tests -p litchi-iwa --test parity --jobs 2`
- `cargo clippy -p litchi-pages --lib --tests --all-features --jobs 2 -- -D warnings`
- `cargo doc -p litchi-pages --no-deps --all-features`
- `RUSTC_BOOTSTRAP=1 python3 tools/check_api_glossary.py --profile pages --target-dir /tmp/iwa-w2b-pages-glossary-Q7b8Be`
- `python3 tools/iwa_exit_ledger.py --check`
- `python3 tools/check_crate_boundaries.py`
- `python3 -m unittest tools.test_gate_runner`
- `cargo check --locked --manifest-path crates/litchi-pages/fuzz/Cargo.toml --bins`
- `cargo test -p litchi-pages --lib body_chart_font -- --nocapture`
- `cargo test -p litchi-iwa --test parity chart_font_cases:: -- --nocapture`
- `cargo test -p litchi-pages --lib production_locality_preserves_unknown_changed_and_untouched_objects -- --nocapture`

The Pages facade controls and the existing host-vs-leaf chart-font cases pass;
the seven release/cost probes remain intentionally ignored for the
coordinator's final CPU slot. The shared substrate and shape/adapter checks
are recorded in their respective records; the coordinator must rerun the
complete gate on one integrated commit.

An initial source-bound release probe binary was built but not executed; it
was superseded by the final CPU5 source-bound build below:

- source `e481893f62ead345a501e240a2529a746029f32d`
- command `cargo test -p litchi-iwa --test parity --release --no-run --jobs 2`
- external immutable binary `/tmp/iwa-w2b-pages-facade-release-Fqz8mQ/parity-bf6f552e55aaee94`
- 40,183,512 bytes, SHA-256
  `89254437719e28057c018bf7ee237381b10bb9713ded92554d29bd40fa7ad474`

No raw captures or binaries were added to the repository. The final CPU5
capture was instruction-count-only and is retained externally below.

### CPU5 instruction qualification

After pinning the probe helper symbols with `#[inline(never)]`, the final
source-bound binary was rebuilt and archived before execution:

- prior source/binary: `e481893f62ead345a501e240a2529a746029f32d`,
  `pages-parity-0ce5798ee9`, SHA-256
  `867984d1aebcdcffd4c917db550a00922439b83cb3978fcb85d4f10cbb0b4f86`
- measured source: `c3dd90d7dc7de91326200471729b3b8e12a5da42`
- measured binary: `/home/zhuhe/code/litchi-worktrees/.coord/fullmeasuredSHA/bin/pages-parity-c3dd90d7dc`
- measured binary: 40,122,072 bytes, SHA-256
  `7c1ccade1640baea94f1f327a718158f65da2e15500c6fbbb662cb1aaa35d7ac`
- build: `cargo test -p litchi-iwa --test parity --release --no-run --jobs 2`
- capture root: `/home/zhuhe/code/litchi-worktrees/.coord/fullmeasuredSHA/pages-cpu5-c3dd90d7dc`
- manifest: `manifest.tsv`; `taskset -c 5 valgrind --tool=callgrind`
  captured each exact ignored test twice, with the matching host or leaf
  symbol toggle. Every capture had exit 0, one matching test pass, and
  positive symbol Ir.

| Route | Host Ir | Leaf Ir | Leaf minus host |
| --- | ---: | ---: | ---: |
| font read | 1,658,570,789 | 34,487,704 | −97.920637% |
| font set | 2,149,725,478 | 206,192,863 | −90.408410% |
| font reset | 1,914,342,996 | 203,436,287 | −89.373049% |
| size read | 1,659,387,223 | 33,822,994 | −97.961718% |
| size set | 2,042,812,143 | 207,761,686 | −89.829624% |
| size reset | 1,861,018,959 | 202,191,477 | −89.135442% |

The six tests execute explicit host/leaf/leaf/host ABBA calls, with host
setup outside the measured route. This is instruction-count evidence only;
there is no wall-time claim. The six ledger rows remain host-only pending the
coordinator's integrated final qualification and native queue policy.

No raw captures or binaries were added.
