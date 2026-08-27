# Root iWork fuzzing

`parse_iwork` owns fuzz coverage for the supported `litchi::iwork` byte
coordinator. It exercises bounded package admission, format dispatch, all
three semantic projections, and the archive-free snapshot facade. The fuzz
package deliberately depends only on the root crate with the `iwork` feature;
it must not acquire a dependency on the legacy `litchi-iwa` migration host or
on an internal archive, protobuf, Buffa, or concrete-format crate.

`keynote_slide_text` is the focused title/body robustness target. It first
offers arbitrary bytes to the bounded Keynote package ingress, then uses those
same bytes as commands against the repository's native `basic.key` seed. This
second path bypasses the low survival rate of CRC-protected ZIP mutation while
still exercising both placeholder roles, semantic selectors, UTF-16 boundary
validation, set/clear/replace/insert/delete/no-op staging, exact-source patch
application, inversion, and content-redacted errors. It never writes a package
to disk.

`keynote_show_settings` is the focused presentation-settings target. It offers
arbitrary bytes to bounded Keynote package ingress and reuses a fixed prefix
for no-op, playback-only, slide-number/size rendering, and combined commands
against the native `basic.key` seed. The playback-only changed command asserts
that public commit diagnostics report no deleted root previews. It covers
strict reads, exact-source patch application and conflicts, inversion, typed
limits, content-redacted failures, public in-memory `write_to` verification,
and exact byte restoration without writing a package to disk.

`keynote_slide_transition` is the focused selector-first transition target.
It offers arbitrary bytes to bounded Keynote ingress and interprets a finite
prefix as no-op, set, or clear commands for the first slide in native
`basic.key`. It covers strict reads, diagnostics, unrelated show-settings
locality, exact-source application and conflicts, inversion, typed limits,
redaction, and exact restoration using only the public package writer.

`keynote_slide_table_headers` is the focused selector-first slide-table
header/footer target. It offers arbitrary bytes to bounded Keynote ingress and
interprets a finite command prefix against tiny source-built valid and locked
packages. It covers positional and name selectors, presence-preserving
seven-field no-op/set transactions, locked or invalid atomic failures,
exact-source apply/conflict/inverse replay, candidate readback, and bounded
package/semantic limits. The target uses no private native fixture and never
writes a package to disk; the checked-in package corpora and command recipes
under `corpus/keynote_slide_table_headers/` are deliberately small.

`keynote_chart_title` is the focused selector-first chart-title target. It
drives tiny source-built packages through positional and exact-name chart
selectors, visible-empty and hidden-stale title states, set/clear/no-op
commands (including an exact source-byte no-op for hidden/absent clears), exact
patch application and inverse replay, duplicate-name ambiguity, and
source/candidate publication checks. Separate malformed-parent
and malformed-wire packages must fail closed without changing their source.
The checked-in seeds are small ZIP/IWA packages under
`corpus/keynote_chart_title/`; they contain no native fixture data. Arbitrary
bytes still exercise bounded Keynote ingress before each transaction batch.

`numbers_table_lock` is the focused interactive table-lock target. It offers
arbitrary bytes to checked Numbers package ingress and also interprets them as
bounded selector and lock-state commands against the native `basic.numbers`
seed. It covers name and index selectors, exact no-op and changed commits,
exact-source patch conflicts, inversion, and content-redacted errors without
writing a package to disk.

`numbers_table_headers` is the focused table header/footer target. It offers
arbitrary bytes to bounded Numbers ingress and reuses a fixed prefix for exact
no-op or combined count/freeze/repeat changes on the first table in native
`basic.numbers`. It covers index and name reads, typed selection and limit
errors, diagnostics, exact-source application and conflicts, inversion,
redaction, and byte-exact restoration through public in-memory `write_to`.

`numbers_names` is the focused atomic sheet/table-names target. It offers
arbitrary bytes to bounded Numbers package ingress and reuses a fixed command
prefix for no-op, sheet, table, and combined renames against the native
`basic.numbers` seed. It covers selector staging, Unicode names, typed finite
limits, content-redacted failures, exact-source patch application and
conflicts, inversion, and byte-exact restoration through public in-memory
`write_to` without writing a package to disk.

`numbers_formula_cells` is the focused formula-authoring target. It offers
arbitrary bytes to bounded Numbers ingress and reuses the native
`basic.numbers` seed for formula, cache, clear, duplicate, mixed-cell, and
cycle batches. The transaction path covers every public authorable function
name, local row/column references, bounded deep formula trees near the
aggregate work ceiling, exact-source patch application and conflicts,
inversion, and content-redacted failures.

`numbers_table_appearance` is the focused selector-first table-appearance
lifecycle target. It offers arbitrary bytes to bounded Numbers ingress and
reuses the native `basic.numbers` seed for no-op, reset, toggle, and complete
appearance replacements. It checks archive-free row-banding, row-sizing, and
all five gridline regions; exact package-byte no-op behavior; source-bound
forward apply, patch conflicts, inverse restoration, candidate readback, and
content-redacted selector/ingress failures. It never exposes native style
objects, identifiers, protobuf payloads, or archive names.

`numbers_table_cell_pop_up_menu` is the focused selector-first Pop-Up Menu
cell-format lifecycle target. It offers arbitrary bytes to bounded Numbers
ingress and reuses the native `basic.numbers` seed for checked cell-position
commands. When the source admits the control-cell graph, the harness covers
automatic/no-op reads, creation, Some-to-Some replacement, model reuse across
two cells, reset and final model cull, exact-source patch application and
conflicts, inverse restoration, candidate readback, and redacted selector or
ingress errors. Commands use only `SheetSelector`, `TableSelector`, and
`CellPosition`; native IDs, BNC bytes, archive names, and generated types do
not cross the fuzz target boundary. The checked-in seeds under
`corpus/numbers_table_cell_pop_up_menu/` are command recipes, not native
package bytes.

`numbers_table_cell_control` is the unified selector-first cell-control
lifecycle target. It drives Checkbox, StarRating, Slider, Stepper, and
Pop-Up Menu values through `Package::{table_cell_control_format,
edit_table_cell_control_format,apply_table_cell_control_format}`. When a
source admits the rooted control graph it exercises no-op and Some-to-Some
reads, cross-kind replacement, clear/reset, exact inverse/apply/conflict,
candidate readback, and source-byte atomicity. It also probes invalid ranges,
all selector forms, bounded ingress limits, and redacted errors against the
native Numbers seed. In addition to `basic.numbers`, the target embeds the
checked-in `split_component_source.numbers` recipe. That source places the
rooted model, tile, format/control lists, and Pop-Up model in separate current
members; a successful changed transaction must report every touched member,
preserve its metadata token/external-edge closure, and replay exact forward,
inverse, and conflict patches. The split path also probes shared refcount
clears, physical alias/edge failures supplied by mutated ZIP inputs, and
required-minus-one ingress limits before candidate publication. Unsupported
or malformed split ownership is expected to fail closed with the original
bytes unchanged. The recipes under `corpus/numbers_table_cell_control/` are
command bytes except for the explicitly named native split source; no native
IDs, BNC payloads, archive names, or generated values cross this target.

`numbers_table_cell_comment_reply` is the selector-first direct-reply
lifecycle target. Its first byte is a bounded command prefix (unless the
input already starts with a ZIP local header); the remaining bytes are
offered to bounded Numbers package ingress. Any admitted package is probed
for a rooted cell-comment thread and then exercised by reply ordinal and A1
address. The collection edit covers
append, duplicate-text ordinal replacement, and removal, while the direct
add/set/remove and A1 conveniences cover the same transitions. Successful
commits are reopened, applied, conflicted, inverted, and checked for exact
source restoration; malformed, missing-root, stale-ordinal, selector, and
limit paths must remain source-atomic. The target does not embed an
unverified native reply artifact: command seeds are deliberately small, and a
valid reply-bearing package can be supplied as a fuzz input when available.
The target uses only `SheetSelector`, `TableSelector`, `CellPosition`, and
`CommentReplyIndex`; native IDs, comment-storage payloads, archive names, and
generated types do not cross the fuzz boundary.

`numbers_table_sort_order` is the focused selector-first persisted table-sort
configuration target. It offers arbitrary bytes to bounded Numbers ingress
and reuses native `basic.numbers` for no-op, set, clear/reset, exact patch apply,
conflict, inverse, candidate readback, and failed-commit source-atomicity
commands. It deliberately does not execute physical row sorting; that path
rewrites tiles, formulas, comments, and view state and remains a separate
host operation. The checked-in recipes under
`corpus/numbers_table_sort_order/` are command inputs, not native package
copies.


`pages_page_layout` is the focused Pages document-layout target. It offers
arbitrary bytes to checked Pages package ingress and reuses them as bounded
layout commands against the native `basic.pages` seed. It covers public layout
reads and validation, exact no-op and changed commits, exact-source patch
conflicts, inversion, content-redacted failures, and exact restoration without
writing a package to disk.

`pages_document_settings` is the focused combined Document and Footnotes
formatter target. It offers arbitrary bytes to checked Pages package ingress
and reuses a fixed prefix for bounded option and footnote commands against the
native `basic.pages` seed. It covers strict reads, exact no-ops, combined
changes, exact-source conflicts, inversion, typed limits, content-redacted
failures, and exact restoration without writing a package to disk.

`pages_section_pagination` is the focused section-pagination target. It offers
arbitrary bytes to checked Pages package ingress and reuses them as bounded
selector and pagination commands against the native `basic.pages` seed. It
covers presence-preserving start/numbering/page values, canonical and invalid
enum aliases, exact no-op and changed commits, exact-source patch conflicts,
inverse replay, failed-commit source atomicity, tightened output/resource
profiles, and content-redacted failures without writing a package to disk.

`pages_body_footnote_lifecycle` is the focused selector-first body-footnote
graph lifecycle target. It offers arbitrary bytes to bounded Pages ingress and
reuses the same input as bounded text/custom-marker commands against native
`basic.pages`. When the fixed source accepts an insertion, the harness checks
source immutability, semantic readback, exact no-op diagnostics, exact-source
patch application/conflicts, inverse restoration, and complete graph removal.
Existing valid notes are also exercised through selector-based removal; typed
structural-marker, missing-selector, malformed-ingress, and physical input
limit failures remain atomic. The target never writes a package to disk.

The checked-in seeds under `corpus/pages_body_footnote_lifecycle/` are small
`hex:` command recipes. The native package is embedded from the repository's
`basic.pages` fixture, so this corpus is isolated from the low-level protobuf
body/footnote corpora.

`pages_header_footer_lifecycle` is the focused selector-first header/footer
text target. It offers arbitrary bytes to bounded Pages ingress and reuses a
fixed command prefix for typed section/template/role/slot selection, no-op,
set, clear, UTF-16 replacement, alias readback, exact patch application,
conflicts, inverse restoration, redacted selector errors, and input limits.
The target leaves the package source immutable on every failed staging or
commit path and never writes a package to disk. The checked-in seeds under
`corpus/pages_header_footer_lifecycle/` are small `hex:` command recipes;
they are not native Pages package copies.

`parse_iwork` uses tighter limits than the public defaults: 2 MiB of source
bytes, 512 package entries, 8 MiB per expanded entry and decoded IWA item,
32 MiB aggregate expanded bytes, 4,096 values of each semantic collection,
and 4 MiB of retained text. Keep its `-max_len` aligned with the 2 MiB source
ceiling so oversized mutations do not consume fuzzing time.

`keynote_slide_text` uses a narrower physical profile for its arbitrary-byte
path: 1 MiB of source bytes, 256 package entries, 2 MiB per expanded entry and
decoded IWA item, and 8 MiB aggregate expanded bytes. Its native `basic.key`
seed is embedded in the harness from the hash-verified source below; fuzzer
input supplies only bounded transaction commands, of which at most 1 KiB
becomes replacement text. Keep this target's `-max_len` at 4 KiB to
concentrate effort on deep-message operations.

`keynote_show_settings` reuses the same finite Keynote physical and semantic
profile. Settings commands consume only a fixed prefix; keep `-max_len` at 512
bytes so arbitrary ingress remains active while every input also reaches the
fixed native transaction.

`keynote_slide_transition` also reuses this finite Keynote profile. Selector,
transition, and validation commands consume at most 512 input bytes; keep
`-max_len` at 512 so malformed ingress remains bounded while every input also
reaches the native transition transaction.

`keynote_slide_table_headers` uses the same finite Keynote physical and
semantic profile. Header commands consume only a fixed seven-field prefix;
keep `-max_len` at 4 KiB so malformed ingress and both source-built package
variants receive every command stream.

`keynote_chart_title` uses the same finite Keynote physical and semantic
profile. Chart-title command bytes consume at most 1 KiB; keep `-max_len` at
1 KiB so malformed ingress and every source-built chart transaction receive
the same input.

`numbers_table_lock` accepts at most 512 KiB of source bytes, 128 package
entries, 1 MiB per expanded entry and decoded IWA item, and 4 MiB aggregate
expanded bytes. Its semantic profile admits at most 4,096 objects, 128 sheets,
512 tables, 8,192 references, 65,536 materialized cells, and 512 KiB of
retained text. Fuzzer-derived selector names are limited to 512 input bytes;
keep `-max_len` at 1 KiB so most work reaches the fixed native seed.

`numbers_table_headers` reuses the same finite Numbers physical and semantic
profile. Header commands consume only a small fixed prefix; keep `-max_len` at
512 bytes so arbitrary ingress stays bounded while every input also reaches
the native transaction.

`numbers_names` uses the same finite Numbers physical and semantic profile.
Fuzzer-derived names are decoded lossily as UTF-8, reject NUL, and consume at
most 256 input bytes; keep `-max_len` at 1 KiB so malformed ingress and native
name transactions both receive every input.

`numbers_formula_cells` uses the same finite Numbers physical and semantic
profile, with 8 KiB formula-render work and depth 32. Its stress command builds
at most a 5,461-node bounded formula tree; keep `-max_len` at 1 KiB so
arbitrary ingress remains bounded while native transaction commands receive
every input.

`numbers_table_appearance` uses the same finite Numbers physical and semantic
profile. The command prefix is at most eight bytes; keep `-max_len` at 512 bytes
so arbitrary ingress remains bounded while every input also reaches the native
appearance transaction.

`numbers_table_cell_pop_up_menu` uses the same finite Numbers physical and
semantic profile. Popup commands consume a fixed prefix and construct at most
three validated menu items of 4 KiB each; keep `-max_len` at 1 KiB so malformed
ingress and native create/reuse/reset transactions both receive every input.

`numbers_table_sort_order` uses the same finite Numbers physical and semantic
profile. Its command prefix is bounded and only changes persisted field 44;
keep `-max_len` at 1 KiB so malformed ingress and native set/clear commands
both receive every input. Physical row execution remains outside this target.

`pages_body_table_sort_order` is the focused selector-first persisted sort
configuration target for a Pages body table. It offers arbitrary bytes to
bounded Pages ingress and reuses the same input as bounded
`BodyTableSelector` commands against the repository's Pages seed. It covers
read, set/clear/reset, exact-source publication and conflicts, inverse replay,
candidate readback, and source-atomic limit failures. It deliberately does
not execute physical row sorting. The checked-in recipes under
`corpus/pages_body_table_sort_order/` are command inputs, not native Pages
package copies.

`keynote_movie_playback` is the focused selector-first movie-playback target.
It offers arbitrary bytes to bounded Keynote ingress and reuses the same
input as bounded `SlideSelector`/`MovieSelector` commands. When the source
contains an existing file-backed movie, it covers exact no-op and playback
replacement, optional-field clearing through a complete semantic value,
candidate readback, exact patch application/conflict, inverse restoration,
and source atomicity. Movie creation/removal, media replacement, geometry,
and other graph operations remain outside this target. The checked-in seeds
under `corpus/keynote_movie_playback/` are command recipes, not native
Keynote packages.

The target uses the finite Keynote physical profile above; keep `-max_len` at
1 KiB so malformed ingress and the fixed movie transaction both receive every
input. The tracked `basic.key` seed is embedded only as a valid package
fallback; a file-backed movie-bearing seed can be supplied by a future native
campaign without changing the public harness boundary.

`keynote_slide_movie_geometry` is the focused selector-first geometry target
for an existing file-backed movie. It offers arbitrary bytes to bounded
Keynote ingress and reuses the same bytes as typed `SlideSelector` and
`MovieSelector` commands against the checked-in Keynote seed. When a source
contains an admitted movie geometry, it covers position/size reads, no-op and
replacement commits, exact patch application and conflicts, inverse
restoration, candidate readback, selector failures, and source-byte
atomicity. Native flags, angles, movie graph records, and identifiers remain
opaque; movie creation/removal, media replacement, and cross-component writes
are outside this target.

The target uses the finite Keynote physical and semantic profile used by the
other focused Keynote targets; keep `-max_len` at 1 KiB so malformed ingress
and the fixed native transaction both receive every input. Recipes under
`corpus/keynote_slide_movie_geometry/` are command bytes, not native Keynote
packages.


`pages_page_layout` accepts at most 256 KiB of source bytes, 128 package
entries, 1 MiB per expanded entry and decoded IWA item, and 4 MiB aggregate
expanded bytes. Layout commands consume only a small fixed prefix; keep
`-max_len` at 512 bytes to retain malformed-ingress mutation while ensuring
every input also reaches the fixed native seed.

`pages_document_settings` reuses the same finite Pages physical profile.
Settings commands consume only a fixed prefix, so keep `-max_len` at 512 bytes
to combine malformed-ingress mutations with deterministic native transaction
coverage.

`pages_section_pagination` reuses the same finite Pages physical profile.
Pagination commands consume only a fixed prefix, so keep `-max_len` at 512
bytes to combine malformed-ingress mutations with deterministic native
transaction and resource-limit coverage.

All targets currently share this package's single `litchi` dependency with
the `iwork` feature. Cargo unifies dependency features for the package, so the
focused Keynote, Numbers, and Pages binaries compile the complete root iWork
feature set rather than isolated concrete-format graphs. The focused targets'
source-level imports remain confined to the relevant public facade, but true
dependency isolation requires separate owner-specific fuzz packages or
dependencies.

## Native seed provenance

The seed documents already live in `test-data/iwork`; do not check in a second
binary copy. They were created, saved, closed, and reopened with the native
macOS applications, and their visible content is documented in
`test-data/iwork/README.md`.

| Corpus name | Source | Bytes | SHA-256 |
| --- | --- | ---: | --- |
| `basic.pages` | `test-data/iwork/pages/basic.pages` | 96,417 | `21107bc9323fba6f1589152454c0b0b0cc8e239313c6a369bc4a891116601b42` |
| `basic.numbers` | `test-data/iwork/numbers/basic.numbers` | 136,357 | `f225d5b1cd59e9da454f91a96fe8f81154bc31037c10029230e75d49b45fb693` |
| `basic.key` | `test-data/iwork/keynote/basic.key` | 500,058 | `3a3d07476b45b6e543bcfba75fe38a245434176dcb3565e34570b817708b9f42` |

From this directory, prepare the ignored local corpus and verify its hashes:

```sh
mkdir -p corpus/parse_iwork
cp ../../../test-data/iwork/pages/basic.pages corpus/parse_iwork/
cp ../../../test-data/iwork/numbers/basic.numbers corpus/parse_iwork/
cp ../../../test-data/iwork/keynote/basic.key corpus/parse_iwork/
shasum -a 256 corpus/parse_iwork/basic.pages \
  corpus/parse_iwork/basic.numbers corpus/parse_iwork/basic.key
```

Run the target with an address sanitizer and explicit process ceilings:

```sh
cargo +nightly fuzz run parse_iwork corpus/parse_iwork -- \
  -max_len=2097152 -timeout=10 -rss_limit_mb=2048
```

Run the focused Keynote target without a checked-in duplicate corpus:

```sh
cargo +nightly fuzz run keynote_slide_text -- \
  -max_len=4096 -timeout=10 -rss_limit_mb=2048
```

Run the focused Keynote show-settings target:

```sh
cargo +nightly fuzz run keynote_show_settings -- \
  -max_len=512 -timeout=10 -rss_limit_mb=2048
```

Run the focused Keynote slide-transition target:

```sh
cargo +nightly fuzz run keynote_slide_transition -- \
  -max_len=512 -timeout=10 -rss_limit_mb=2048
```

Run the focused Keynote slide-table-header target with its source-built
packages and command recipes:

```sh
cargo +nightly fuzz run keynote_slide_table_headers \
  corpus/keynote_slide_table_headers -- \
  -max_len=4096 -timeout=10 -rss_limit_mb=2048
```

Run the focused chart-title target with its reviewable package seeds:

```sh
cargo +nightly fuzz run keynote_chart_title \
  corpus/keynote_chart_title -- \
  -max_len=4096 -timeout=10 -rss_limit_mb=2048
```

The `cargo +nightly fuzz run` commands are required for sanitizer-instrumented
coverage. A stable `cargo run --bin <target> -- -runs=...` invocation is only
a control-flow smoke test; on platforms without linked sanitizer runtimes it
may print missing-symbol warnings and is not sanitizer evidence.

Run the focused Numbers target without a checked-in duplicate corpus:

```sh
cargo +nightly fuzz run numbers_table_lock -- \
  -max_len=1024 -timeout=10 -rss_limit_mb=2048
```

Run the focused Numbers table-header target:

```sh
cargo +nightly fuzz run numbers_table_headers -- \
  -max_len=512 -timeout=10 -rss_limit_mb=2048
```

Run the focused Numbers names target without a checked-in duplicate corpus:

```sh
cargo +nightly fuzz run numbers_names -- \
  -max_len=1024 -timeout=10 -rss_limit_mb=2048
```

Run the focused Numbers formula-authoring target:

```sh
cargo +nightly fuzz run numbers_formula_cells -- \
  -max_len=1024 -timeout=10 -rss_limit_mb=2048
```

Run the focused Numbers table-appearance target with its checked-in command
seeds:

```sh
cargo +nightly fuzz run numbers_table_appearance \
  corpus/numbers_table_appearance -- \
  -max_len=512 -timeout=10 -rss_limit_mb=2048
```

Run the focused Numbers Pop-Up Menu lifecycle target with its command seeds:

```sh
cargo +nightly fuzz run numbers_table_cell_pop_up_menu \
  corpus/numbers_table_cell_pop_up_menu -- \
  -max_len=1024 -timeout=10 -rss_limit_mb=2048
```

Run the unified Numbers cell-control target with its command seeds:

```sh
cargo +nightly fuzz run numbers_table_cell_control \
  corpus/numbers_table_cell_control -- \
  -max_len=1024 -timeout=10 -rss_limit_mb=2048
```

The `split_component_write_commands.hex` and
`split_component_clear_refcount.hex` recipes select the multi-member write
and clear/refcount branches; the native `split_component_source.numbers` is
embedded by the target and is therefore exercised on every command campaign.

Run the focused Numbers direct-comment-reply lifecycle target with its
bounded command seeds:

```sh
cargo +nightly fuzz run numbers_table_cell_comment_reply \
  corpus/numbers_table_cell_comment_reply -- \
  -max_len=1024 -timeout=10 -rss_limit_mb=2048
```

Run the focused Numbers persisted-sort target with its command seeds:

```sh
cargo +nightly fuzz run numbers_table_sort_order \
  corpus/numbers_table_sort_order -- \
  -max_len=1024 -timeout=10 -rss_limit_mb=2048
```


Run the focused Pages target without a checked-in duplicate corpus:

```sh
cargo +nightly fuzz run pages_page_layout -- \
  -max_len=512 -timeout=10 -rss_limit_mb=2048
```

Run the focused Pages document-settings target:

```sh
cargo +nightly fuzz run pages_document_settings -- \
  -max_len=512 -timeout=10 -rss_limit_mb=2048
```

Run the focused Pages section-pagination target with its checked-in command
seeds:

```sh
cargo +nightly fuzz run pages_section_pagination corpus/pages_section_pagination -- \
  -max_len=512 -timeout=10 -rss_limit_mb=2048
```

The native packages currently store ZIP members with CRC protection. Most
arbitrary payload mutations therefore fail during physical validation, which
is appropriate for the root ingress target. Valid deep-message mutation and
format-specific behavioral invariants belong in focused format-owner fuzz
targets rather than weakening package validation here.

## 2026-08-08 bounded sanitizer evidence

The three native seeds above were copied to a private temporary corpus and
their documented SHA-256 values were verified before this run. With
`cargo-fuzz 0.13.2` and `rustc 1.99.0-nightly (1a98b1e13 2026-08-07)`, the
root target completed this bounded AddressSanitizer/libFuzzer campaign:

```sh
cargo +nightly fuzz run parse_iwork /private/tmp/litchi-iwork-root-fuzz-1JVxsJ/corpus -- \
  -max_total_time=60 -max_len=2097152 -timeout=10 -rss_limit_mb=2048
```

It executed 152,219 inputs in 61 seconds, ending at coverage 7,454, feature
count 12,062, a 249-input / 69 MiB corpus, and 566 MiB RSS. There was no crash,
timeout, or out-of-memory finding. This records one bounded root-ingress
sanitizer run only; it is not evidence of exhaustive fuzzing or of the focused
deep-message campaigns still required from the format owners.
