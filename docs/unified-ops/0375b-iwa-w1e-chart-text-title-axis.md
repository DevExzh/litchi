# 0375b — W1-E slice: shared chart title and axis-title payload engine

- Status: implementation slice on uo2/iwa-w1e; scoped branch gates pass,
  while coordinator review, native evidence and final integrated gates are
  pending.
- Date: 2026-10-28.
- Parent: 0375a (0375a-iwa-w1e-style-variation-write-engine.md), plan
  0363 (0363-iwa-monolith-exit-plan.md), parity harness and ledger
  0369 (0369-iwa-exit-parity-harness-and-ledger.md), and the native queue
  0363-mac-queue.md.
- Scope: title and category/value-axis title only. Caption authoring remains
  host-only and is explicitly deferred; no Keynote ledger row is flipped.

This is a thin vertical slice. litchi-iwa-style::chart_text owns the bounded
generated-extension algorithm. Numbers and Pages retain selector resolution,
reference census, graph ownership, archive publication and candidate reopen.
No native object identifier, UUID, archive name, ZIP path or generated
protobuf value crosses a supported leaf signature.

## 1. Shared engine

ChartTextKind::{Title, Axis(Axis)} selects the generated scalar family.
decode borrows the visible UTF-8 value from a complete chart payload; rewrite
applies a presence-preserving set/clear to field 10,000 while retaining
unknown outer and generated-extension spans. The engine validates a unique
canonical length-delimited extension before dispatching to the strict
title/axis codecs.

Limits is finite for input bytes, fields, output bytes, nesting, work,
selected text bytes, allocations, retained bytes and scratch bytes.
Report records those costs and whether the selected payload changed.
Error::{InvalidSource, Limit, Allocation} is terminal and is mapped by each
leaf into its own Error/LimitKind; codec/archive failures are not flattened
into a successful host fallback. Axis rewrites use the generated prepared
rewrite path, including candidate readback before publication.

The shared graph's semantic section/storage count and retained-text ceilings
remain typed through the existing Pages body-family adapters as
`PayloadItems` and `TextBytes`; they are never relabeled as wire work. The
package-error mapper and the chart facade each have focused low-limit
regressions for these dimensions.

The engine tests cover:

- title set, exact no-op, unknown outer/extension preservation and readback;
- category-axis locality with value-axis preservation;
- duplicate-extension refusal and a typed low text-byte refusal.

## 2. Leaf facades and contract

The Numbers destinations are:

- Package::{sheet_chart_title, edit_sheet_chart_title,
  apply_sheet_chart_title};
- Package::{sheet_chart_axis_title, edit_sheet_chart_axis_title,
  apply_sheet_chart_axis_title}.

The Pages destinations are:

- Package::{body_chart_title, edit_body_chart_title,
  apply_body_chart_title};
- Package::{body_chart_axis_title, edit_body_chart_axis_title,
  apply_body_chart_axis_title}.

Each route has immutable source snapshots, exact-source patches,
Patch::inverse, exact no-op behavior, bounded rewrite limits, selector-first
refusals, candidate reopen/readback and untouched-object locality. Caption
facades from the earlier donor are removed from this slice rather than
published as a second orchestration path. Existing Keynote title, axis-title
and caption implementations remain unchanged.

## 3. Host differential vectors

crates/litchi-iwa/tests/parity/chart_text_cases.rs is the host-backed
differential oracle. It covers title and both axis selections for Numbers and
Pages, set/read/remove sequences, exact inverse restoration, strict leaf
reopen, and semantic host-vs-leaf reads. Numbers uses
test-data/iwork/numbers/chart-data-native.numbers and adds the Column2d chart
through the host before serializing the differential input. Pages uses a
source-built Draft with a host-added body chart: the tracked native Pages
chart-data fixture currently contains a malformed cross-component chart
reference for host discovery, so it is retained for the native queue but not
used as a false differential vector. The host remains authoritative; the
cases are declared Evidence::Queued until native application evidence is
recorded.

The intentional boundary is caption authoring: the host rows remain
host-only until a later shared caption graph slice supplies the actual
bounded graph algorithm and native evidence.

## 4. Performance qualification

`crates/litchi-iwa/tests/parity/chart_text_cost.rs` contains eighteen
source-bound host/leaf ABBA tests: title, category-axis and value-axis, each
split into read, set and clear for Numbers and Pages. The named helpers are
full facade routes. Read includes open, selector resolution and readback; set
and clear include commit, candidate reopen/readback and serialization. Each
axis test selects exactly one axis, so category and value regressions cannot
hide in an aggregate route. The ignored test is run with `--exact`, limiting
callgrind collection to the selected test/helper route.

The following pre-integration release callgrind run was pinned to CPU 3. `Ir`
is the positive instruction count from `callgrind_annotate --inclusive=yes`;
delta is `(leaf - host) / host`. Every independently measured route is below
the +5% slower limit. This is instruction-cost evidence, not wall-clock
qualification.

| Format | Exact test | Host Ir | Leaf Ir | Delta | Check |
|---|---|---:|---:|---:|---|
| Numbers | `numbers_chart_title_read_cost_abba` | 69,594,721 | 33,920,586 | -51.26% | pass |
| Numbers | `numbers_chart_title_set_cost_abba` | 151,030,837 | 72,948,282 | -51.70% | pass |
| Numbers | `numbers_chart_title_clear_cost_abba` | 151,062,590 | 72,945,210 | -51.71% | pass |
| Numbers | `numbers_chart_category_axis_read_cost_abba` | 69,576,145 | 33,916,792 | -51.25% | pass |
| Numbers | `numbers_chart_category_axis_set_cost_abba` | 151,026,347 | 72,960,615 | -51.69% | pass |
| Numbers | `numbers_chart_category_axis_clear_cost_abba` | 151,006,075 | 72,983,662 | -51.67% | pass |
| Numbers | `numbers_chart_value_axis_read_cost_abba` | 69,601,895 | 33,927,383 | -51.26% | pass |
| Numbers | `numbers_chart_value_axis_set_cost_abba` | 151,008,131 | 72,956,836 | -51.69% | pass |
| Numbers | `numbers_chart_value_axis_clear_cost_abba` | 151,033,507 | 72,978,293 | -51.68% | pass |
| Pages | `pages_chart_title_read_cost_abba` | 53,270,374 | 2,153,037 | -95.96% | pass |
| Pages | `pages_chart_title_set_cost_abba` | 108,932,574 | 4,871,364 | -95.53% | pass |
| Pages | `pages_chart_title_clear_cost_abba` | 109,021,251 | 4,885,791 | -95.52% | pass |
| Pages | `pages_chart_category_axis_read_cost_abba` | 53,362,096 | 2,173,101 | -95.93% | pass |
| Pages | `pages_chart_category_axis_set_cost_abba` | 108,990,982 | 4,886,349 | -95.52% | pass |
| Pages | `pages_chart_category_axis_clear_cost_abba` | 108,718,221 | 4,906,575 | -95.49% | pass |
| Pages | `pages_chart_value_axis_read_cost_abba` | 53,282,918 | 2,153,406 | -95.96% | pass |
| Pages | `pages_chart_value_axis_set_cost_abba` | 108,896,978 | 4,890,805 | -95.51% | pass |
| Pages | `pages_chart_value_axis_clear_cost_abba` | 108,991,420 | 4,919,619 | -95.49% | pass |

Reproduction receipt (all commands run from the checkout; raw captures remain
outside it):

```sh
git rev-parse HEAD
# eca2ea2c06
env CARGO_BUILD_JOBS=4 CARGO_INCREMENTAL=0 CARGO_PROFILE_RELEASE_DEBUG=0 \
  CARGO_NET_OFFLINE=true LITCHI_MEM_MAX=20G \
  cargo test --release --offline -p litchi-iwa --test parity chart_text_cost --no-run
BIN=target/release/deps/parity-25144fe4e83e134e
```

The exact collection loop used for all eighteen receipts was:

```sh
for SPEC in \
  numbers_chart_title_read_cost_abba:numbers_title_read_host:numbers_title_read_leaf \
  numbers_chart_title_set_cost_abba:numbers_title_set_host:numbers_title_set_leaf \
  numbers_chart_title_clear_cost_abba:numbers_title_clear_host:numbers_title_clear_leaf \
  numbers_chart_category_axis_read_cost_abba:numbers_axis_read_host:numbers_axis_read_leaf \
  numbers_chart_category_axis_set_cost_abba:numbers_axis_set_host:numbers_axis_set_leaf \
  numbers_chart_category_axis_clear_cost_abba:numbers_axis_clear_host:numbers_axis_clear_leaf \
  numbers_chart_value_axis_read_cost_abba:numbers_axis_read_host:numbers_axis_read_leaf \
  numbers_chart_value_axis_set_cost_abba:numbers_axis_set_host:numbers_axis_set_leaf \
  numbers_chart_value_axis_clear_cost_abba:numbers_axis_clear_host:numbers_axis_clear_leaf \
  pages_chart_title_read_cost_abba:pages_title_read_host:pages_title_read_leaf \
  pages_chart_title_set_cost_abba:pages_title_set_host:pages_title_set_leaf \
  pages_chart_title_clear_cost_abba:pages_title_clear_host:pages_title_clear_leaf \
  pages_chart_category_axis_read_cost_abba:pages_axis_read_host:pages_axis_read_leaf \
  pages_chart_category_axis_set_cost_abba:pages_axis_set_host:pages_axis_set_leaf \
  pages_chart_category_axis_clear_cost_abba:pages_axis_clear_host:pages_axis_clear_leaf \
  pages_chart_value_axis_read_cost_abba:pages_axis_read_host:pages_axis_read_leaf \
  pages_chart_value_axis_set_cost_abba:pages_axis_set_host:pages_axis_set_leaf \
  pages_chart_value_axis_clear_cost_abba:pages_axis_clear_host:pages_axis_clear_leaf
do
  IFS=: read -r ROUTE HOST LEAF <<< "$SPEC"
  taskset -c 3 valgrind --tool=callgrind --quiet \
    --callgrind-out-file="/tmp/litchi-w1e-chart-text-20261028-${ROUTE}-host.callgrind" \
    "--toggle-collect=*chart_text_cost::${HOST}*" "$BIN" \
    "chart_text_cost::${ROUTE}" --ignored --exact --test-threads=1 >/dev/null 2>/dev/null
  taskset -c 3 valgrind --tool=callgrind --quiet \
    --callgrind-out-file="/tmp/litchi-w1e-chart-text-20261028-${ROUTE}-leaf.callgrind" \
    "--toggle-collect=*chart_text_cost::${LEAF}*" "$BIN" \
    "chart_text_cost::${ROUTE}" --ignored --exact --test-threads=1 >/dev/null 2>/dev/null
done
```

The release binary is `target/release/deps/parity-25144fe4e83e134e`,
39,688,944 bytes, SHA-256
`26883da1062e498de6461693d8c9a3c350990a21d125fe587ce1713eff56c9cb`.
The source commit and binary hash must be re-recorded after integration.

Raw captures are all under `/tmp` and are not repository artifacts. Each line
is `filename size SHA-256`:

```text
litchi-w1e-chart-text-20261028-numbers_chart_title_read_cost_abba-host.callgrind 90916 7a4d7ad1e3cdb72af6e6afea3ad3338d62595e4ee180514b15388763efdc3e0d
litchi-w1e-chart-text-20261028-numbers_chart_title_read_cost_abba-leaf.callgrind 82190 c8c0cf4f4e1bacc0c07805598f8c166f14ec43fa8e341a4e80fb68cc1e45a1ee
litchi-w1e-chart-text-20261028-numbers_chart_title_set_cost_abba-host.callgrind 108889 e00d086e5b5f04f24e3e2a51439a1e72bfd5b6fd0059cca16c10cb9355608aa0
litchi-w1e-chart-text-20261028-numbers_chart_title_set_cost_abba-leaf.callgrind 106744 ab632a81ab715562ece99fcafa84c8e4f68bc891a94a3704beeeee43f60f966e
litchi-w1e-chart-text-20261028-numbers_chart_title_clear_cost_abba-host.callgrind 108593 295bb21066d429c3045acef2b41643add29a2c7a3081b294a48b48fa747d88dc
litchi-w1e-chart-text-20261028-numbers_chart_title_clear_cost_abba-leaf.callgrind 106119 77147c3925671f13402554cb9a7d1ebe930075865730c584b55d2a07b708d687
litchi-w1e-chart-text-20261028-numbers_chart_category_axis_read_cost_abba-host.callgrind 90625 b671e67e3ad8dd31d186cf278573680addb5fa21bf37392139d5480b55598589
litchi-w1e-chart-text-20261028-numbers_chart_category_axis_read_cost_abba-leaf.callgrind 82390 623cee9b7b3ab3a50da5cfe64f04d068b0d2ef7827ff6bfd5dbfd8894528a058
litchi-w1e-chart-text-20261028-numbers_chart_category_axis_set_cost_abba-host.callgrind 109268 b6cb30ae3785c07cb6e95081b2bd3980dceea8d31d9fee6080648539584cd4cf
litchi-w1e-chart-text-20261028-numbers_chart_category_axis_set_cost_abba-leaf.callgrind 106888 0e7a76a2484e01fbfa7b32f90de8c44d252a58d38019537e13184731b584a821
litchi-w1e-chart-text-20261028-numbers_chart_category_axis_clear_cost_abba-host.callgrind 109141 79f79f68d2fbe93d2b40d1574f0d560f6bcb525b8e92ae7b7823e3cca76ac676
litchi-w1e-chart-text-20261028-numbers_chart_category_axis_clear_cost_abba-leaf.callgrind 106303 b25702413ea6ca254ef1e1bab5ae972be96848fa0b14acd4b7bd6ed8fdefe14e
litchi-w1e-chart-text-20261028-numbers_chart_value_axis_read_cost_abba-host.callgrind 90405 b17721f4e20a106fb268588fae33ce67758562b629f99b9252d42f0e2ad62400
litchi-w1e-chart-text-20261028-numbers_chart_value_axis_read_cost_abba-leaf.callgrind 82212 92f447e2c409a013adc467f3cc15bcd52ba23abeb1e3720c6e15d17541f077bc
litchi-w1e-chart-text-20261028-numbers_chart_value_axis_set_cost_abba-host.callgrind 109064 1fe11579f92c2a80f33f4aaa7b1af3712cfcede49b6234bbf67c84edb10a143a
litchi-w1e-chart-text-20261028-numbers_chart_value_axis_set_cost_abba-leaf.callgrind 106780 f70f8baa0d67fa290600d9fdfbb7f132cb21728fb3bf875a37c114b3a3aa7cad
litchi-w1e-chart-text-20261028-numbers_chart_value_axis_clear_cost_abba-host.callgrind 109402 8bdfeb908a75a65f51207330fd4d5aa3cd924a3bb446c3b3037817e7b362c6b9
litchi-w1e-chart-text-20261028-numbers_chart_value_axis_clear_cost_abba-leaf.callgrind 106441 15e7aa8e263329c018262080be52afe5c10e5e2bd40d2b53f1cde5661fa8f1fc
litchi-w1e-chart-text-20261028-pages_chart_title_read_cost_abba-host.callgrind 95593 f6fa4072b69a529297e8e3afbe4706d4c3787f7fc65a27ffd9554b56030fdf29
litchi-w1e-chart-text-20261028-pages_chart_title_read_cost_abba-leaf.callgrind 67552 3dcd648895188b3f307ad57b0d33331ef4e330d0f580675eeb976e221c105486
litchi-w1e-chart-text-20261028-pages_chart_title_set_cost_abba-host.callgrind 113597 2559fbed09214d8f5aa84c451bd0da38f296bfc4c7d365658689d354b7b8539d
litchi-w1e-chart-text-20261028-pages_chart_title_set_cost_abba-leaf.callgrind 91709 1ba9adfed22a8de00670ffec025aa7f63cd328b8b20693407642be8202f54ca7
litchi-w1e-chart-text-20261028-pages_chart_title_clear_cost_abba-host.callgrind 113031 0b41f69e29962c172d1b0c851a68b736b2569f812db1fa54f7c5a26f214511ee
litchi-w1e-chart-text-20261028-pages_chart_title_clear_cost_abba-leaf.callgrind 90995 2c3c3fe05dce858f994d4c3147c880117b48b936b78c4c52c70346bab99a0280
litchi-w1e-chart-text-20261028-pages_chart_category_axis_read_cost_abba-host.callgrind 95090 5c87767d5ab69469216f3cdda65e94133d2a180242a4a2f89b0d13fb68edf078
litchi-w1e-chart-text-20261028-pages_chart_category_axis_read_cost_abba-leaf.callgrind 67726 e038bbc8eb9a9a48ca6a07caca996e90a91d094c2dda7b00f9a255b912a0f013
litchi-w1e-chart-text-20261028-pages_chart_category_axis_set_cost_abba-host.callgrind 113602 62916b2bc40020ef6ae22300ee7c63749222f2e385b74ac2d00dc78d47a2324e
litchi-w1e-chart-text-20261028-pages_chart_category_axis_set_cost_abba-leaf.callgrind 91730 57ffd56ee952dea8be839f4586762cee1f9b081033b11423e7d9778e4fd6f318
litchi-w1e-chart-text-20261028-pages_chart_category_axis_clear_cost_abba-host.callgrind 113731 ccaab9fc0a380bc71b44384f84fdac839b8339bc971ca136e3ee63602a4e6053
litchi-w1e-chart-text-20261028-pages_chart_category_axis_clear_cost_abba-leaf.callgrind 90851 e50384b37cfb279e63a67e1f22ad00fda57ef76c8bbc43b05c726abd8cf02de6
litchi-w1e-chart-text-20261028-pages_chart_value_axis_read_cost_abba-host.callgrind 95049 9d4cf43278a4e65b1b892b23807a8f05fe41ea7d9a430067c192fe32896aed1f
litchi-w1e-chart-text-20261028-pages_chart_value_axis_read_cost_abba-leaf.callgrind 67524 4b2618153a951e3a0473d26d7302400b55e2219c3651c2299741adfb26de8a92
litchi-w1e-chart-text-20261028-pages_chart_value_axis_set_cost_abba-host.callgrind 113593 fe87d009a4edadf4eccaf6ebad8b8d65841db24112782c179fdee47d348f72b7
litchi-w1e-chart-text-20261028-pages_chart_value_axis_set_cost_abba-leaf.callgrind 91715 49b8333ba027431f0f7f3b84fd4b6ff82e329f875a37ab4a99eedb7ef0ddd5b0
litchi-w1e-chart-text-20261028-pages_chart_value_axis_clear_cost_abba-host.callgrind 113399 18ce0bdd3b6c40fc8d2267a5c474b7b64fe91026d7302f529590226cfd347353
litchi-w1e-chart-text-20261028-pages_chart_value_axis_clear_cost_abba-leaf.callgrind 90451 3628a27b3ae92bc9d2941571cba4e5948aaef1616d6ac0606fab2783c81ffb9b
```

These counts are branch diagnostics only: later merges invalidate them, no
wall-clock result is claimed, and final integrated release-mode, pinned-CPU
qualification remains required. The final coordinator run must keep all
eighteen routes independent rather than reintroducing an aggregate axis or
format total.

### 4.1 Post-merge integrated-source receipt

After the frozen coordinator trunk `39e4bf101ca51f269e9352f8fed868cb84799113`
was merged without conflicts, the accepted source was measured from the clean
merge commit `4aa57b3c19c24ce625f31cdd5489ff0a1d39bc68`. This is a separate
receipt update; it does not amend or alter the measured source commit. The
release binary is
`/home/zhuhe/code/litchi-worktrees/uo2-iwa-w1e/target/release/deps/parity-25144fe4e83e134e`,
39,688,944 bytes, SHA-256
`26883da1062e498de6461693d8c9a3c350990a21d125fe587ce1713eff56c9cb`.

All eighteen routes were rerun independently on CPU 3 with the ABBA helper
sequence already encoded by each ignored test. The route process was
fail-closed: it required a zero Valgrind/test exit, exactly one selected
`--exact` test pass, a `1 passed; 0 failed; 0 ignored` summary, successful
annotation, and a positive `Ir` for the selected helper. The integrated values
are:

| Format | Exact test | Host Ir | Leaf Ir | Delta | Check |
|---|---|---:|---:|---:|---|
| Numbers | `numbers_chart_title_read_cost_abba` | 69,581,656 | 33,924,667 | -51.24% | pass |
| Numbers | `numbers_chart_title_set_cost_abba` | 151,034,431 | 72,929,450 | -51.71% | pass |
| Numbers | `numbers_chart_title_clear_cost_abba` | 151,045,301 | 72,939,949 | -51.71% | pass |
| Numbers | `numbers_chart_category_axis_read_cost_abba` | 69,589,373 | 33,908,327 | -51.27% | pass |
| Numbers | `numbers_chart_category_axis_set_cost_abba` | 151,021,222 | 72,975,760 | -51.68% | pass |
| Numbers | `numbers_chart_category_axis_clear_cost_abba` | 151,030,859 | 72,974,591 | -51.68% | pass |
| Numbers | `numbers_chart_value_axis_read_cost_abba` | 69,569,524 | 33,915,891 | -51.25% | pass |
| Numbers | `numbers_chart_value_axis_set_cost_abba` | 151,043,932 | 72,993,476 | -51.67% | pass |
| Numbers | `numbers_chart_value_axis_clear_cost_abba` | 151,054,313 | 72,977,238 | -51.69% | pass |
| Pages | `pages_chart_title_read_cost_abba` | 53,367,762 | 2,167,655 | -95.94% | pass |
| Pages | `pages_chart_title_set_cost_abba` | 108,902,248 | 4,865,970 | -95.53% | pass |
| Pages | `pages_chart_title_clear_cost_abba` | 108,930,205 | 4,903,084 | -95.50% | pass |
| Pages | `pages_chart_category_axis_read_cost_abba` | 53,306,309 | 2,170,575 | -95.93% | pass |
| Pages | `pages_chart_category_axis_set_cost_abba` | 108,962,545 | 4,883,069 | -95.52% | pass |
| Pages | `pages_chart_category_axis_clear_cost_abba` | 108,871,297 | 4,917,896 | -95.48% | pass |
| Pages | `pages_chart_value_axis_read_cost_abba` | 53,280,581 | 2,173,875 | -95.92% | pass |
| Pages | `pages_chart_value_axis_set_cost_abba` | 108,792,690 | 4,883,510 | -95.51% | pass |
| Pages | `pages_chart_value_axis_clear_cost_abba` | 108,815,600 | 4,913,829 | -95.48% | pass |

The complete machine-readable receipt, including all 18 route specifications,
36 callgrind artifact sizes and full hashes, the full source and binary
identities, and the exact build/probe/annotation templates is
`docs/unified-ops/results/0375b/integrated-callgrind.json`. Reproduction used:

```sh
env CARGO_BUILD_JOBS=4 CARGO_INCREMENTAL=0 CARGO_PROFILE_RELEASE_DEBUG=0 \
  CARGO_NET_OFFLINE=true LITCHI_MEM_MAX=20G \
  cargo test --release --offline -p litchi-iwa --test parity chart_text_cost --no-run

taskset -c 3 valgrind --tool=callgrind --quiet \
  --callgrind-out-file="$EVIDENCE_DIR/captures/${ROUTE}-${SIDE}.callgrind" \
  "--toggle-collect=*chart_text_cost::${HELPER}*" "$PARITY_BIN" \
  "chart_text_cost::${ROUTE}" --ignored --exact --test-threads=1 \
  >"$EVIDENCE_DIR/captures/${ROUTE}-${SIDE}.stdout.log" \
  2>"$EVIDENCE_DIR/captures/${ROUTE}-${SIDE}.stderr.log"

callgrind_annotate --inclusive=yes \
  "$EVIDENCE_DIR/captures/${ROUTE}-${SIDE}.callgrind" \
  >"$EVIDENCE_DIR/captures/${ROUTE}-${SIDE}.annotate.log" \
  2>"$EVIDENCE_DIR/captures/${ROUTE}-${SIDE}.annotate.stderr.log"
```

The external evidence directory is
`/tmp/litchi-w1e-chart-text-integrated-20261028/`. Its 186-file manifest
including build, route stdout/stderr, callgrind, and annotation streams is
24,706 bytes with SHA-256
`3ece5ae7f940c51510f1e12c7b59e8d85059a00a72ae5baadb2a5286ee9a6751`.
No raw captures or binaries were added to the repository, and no wall-clock
result is claimed.

### 5.1 Current-source gate recheck

The required source checks were rerun after the receipt commit, at checked
HEAD `b571642f01ea30f13d21e99e3716c4160789d123` whose code is the measured
merge source `4aa57b3c19c24ce625f31cdd5489ff0a1d39bc68`. All commands exited
zero: Numbers and Pages all-feature hidden-item rustdoc JSON; Numbers minimal,
Numbers full and Pages glossary profiles; the exit ledger and its unit suite;
the crate-boundary checker and its unit suite; the no-`3rdparty` unit suite;
and locked-offline `--all-targets --bins` checks for both affected fuzz
workspaces. The exact commands, result paths and external log identities are
in `docs/unified-ops/results/0375b/final-source-gates.json`. The 24-log
manifest is `/tmp/litchi-w1e-final-gates-20261028/log-manifest.tsv`,
2,244 bytes, SHA-256
`1d0e7b87c3e58fc11a1b4dcd4fa196bb3d7aad5292b35443e1bf86d9eae7eaea`.

### 5.2 Test-only proof recheck

The post-queue proof audit found that the original read vectors compared leaf
reads from host- and leaf-mutated outputs, but did not directly compare the
host reader with the leaf reader on one common source. It also found that the
remove helpers authored a temporary value and then cleared it independently,
instead of proving a clear on the caller's common populated input.

`chart_text_cases.rs` now closes those gaps without changing the runtime
facades or the cost-probe routes:

- `run_numbers_reads` and `run_pages_reads` call the host `Draft` reader and
  leaf `Package` reader directly on the same host-authored, populated chart
  input for title, category-axis and value-axis selections. They assert equal
  values and byte-identical source preservation. The negative-control test
  exercises every selected family plus a present value and must reject a
  deliberately mismatched read.
- Each Numbers and Pages remove route now only clears the caller-provided
  common input. The leaf route applies that clear commit's inverse, and the
  removal cases use Tier B so the archive comparator checks untouched ZIP
  members and IWA objects in addition to the exact inverse restoration.

This is a proof/test-only append. `chart_text_cost.rs` and the eighteen-route
measurement are unchanged; `integrated-callgrind.json` remains bound to the
clean measured source commit
`4aa57b3c19c24ce625f31cdd5489ff0a1d39bc68`. No new timing claim was made;
no raw captures or binaries were added to the repository. Focused
verification was rerun with rustfmt, the
`chart_text_cases::` parity filter and the parity-test Clippy profile; logs
remain outside the checkout under `/tmp/litchi-w1e-chart-text-proof-20261028/`.

## 5. Native evidence and verification

No native-authored read receipt proves a Litchi-mutated title or axis title
survives native open/save/close/reopen. The six Numbers rows and six Pages
rows are therefore queued in 0363, with pending results. The ledger and
queue must not be treated as W3-ready until the coordinator records native
evidence, final integrated parity and final performance.

Scoped verification on this worktree includes rustfmt, targeted Clippy,
targeted rustdoc, the shared/Numbers/Pages unit tests, six host differential
sequence tests (three operations per format), eighteen ignored full-facade
cost-probe tests, facade-level low-limit candidate-budget tests, JSON
validation, API glossary registration, ledger and boundary checks,
no-3rdparty checks, affected fuzz-lock cargo checks and git diff --check.
These checks are source/branch evidence only: all twelve native queue rows
remain pending, and the complete 53-row gate plus final performance
qualification must run on the integrated trunk commit.
