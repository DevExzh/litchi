# Native iWork fixtures

These compact fixtures were created with the native macOS iWork applications
on 2026-08-07 and are intentionally checked in as package-level compatibility
fixtures. Their canonical single-file package hashes are:

| File | SHA-256 |
| --- | --- |
| `pages/basic.pages` | `21107bc9323fba6f1589152454c0b0b0cc8e239313c6a369bc4a891116601b42` |
| `numbers/basic.numbers` | `f225d5b1cd59e9da454f91a96fe8f81154bc31037c10029230e75d49b45fb693` |
| `keynote/basic.key` | `3a3d07476b45b6e543bcfba75fe38a245434176dcb3565e34570b817708b9f42` |

| File | Expected semantic content |
| --- | --- |
| `pages/basic.pages` | `Litchi native Pages fixture`, `Buffa lazy-view migration verification`, `2026-08-07` |
| `numbers/basic.numbers` | Cell `B2`: `Litchi native Numbers fixture`; cell `B3`: number `42` |
| `keynote/basic.key` | One slide containing `Litchi native Keynote fixture`, `Buffa lazy-view migration verification`, and `2026-08-07` |

Each document was saved, closed, and reopened in Pages, Numbers, or Keynote.
The application accessibility tree confirmed the expected content after the
reopen and no repair prompt was shown. The format-crate integration tests also
open the packages from disk and from borrowed archive bytes.

## Native package-directory oracles

On 2026-08-08, disposable copies of all three canonical files were opened in
their native applications and converted with **File → Advanced → Change File
Type → Package**. Each resulting directory was saved, closed, reopened from
the application's Recents view, checked for the same visible semantics, and
closed without another save. No application displayed a repair or conversion
prompt. Those app-authored directory artifacts are checked in under
`directory/`; `directory/MANIFEST.sha256` records all 46 regular members.
They contain no symbolic links or special nodes.

| Directory artifact | Regular files | Disk usage | `Index.zip` SHA-256 | Reopened tree digest |
| --- | ---: | ---: | --- | --- |
| `directory/pages/basic.pages` | 7 | 116 KiB | `c1d5ed1626ac8652ee5b2fe36f9d2224a39fb690f5ce91fc8c66a8bfe7951665` | `908f121ebeb8f53a5ca5c4b7793a3c0896d2a73d4f1f6fc5c0d96105ff154ffb` |
| `directory/numbers/basic.numbers` | 7 | 152 KiB | `5c62c8553bfcc0891866663fe474cc758bc92f7c91665ecfc6da775ce50f504a` | `5d280f54308435d099d9a3ef4a6a31b5de74c5086130d908ff80957d1fd36f98` |
| `directory/keynote/basic.key` | 32 | 552 KiB | `c0e3039a597723abd32d2f7f2f9b2fa96826bff17fc2fa1d715e7402ae575acd` | `298190c828b5be72cd06de47f5a8c95995e043f73aab915777ca6cd5e1725c77` |

The tree digest is the SHA-256 of the sorted per-file SHA-256 manifest captured
after native reopen; it is provenance evidence, not a package-format checksum.
The read-only root directory adapter deliberately freezes only `Index.zip` or
loose `Index/` semantics. It does not claim exact ZIP provenance, editing, or
preservation of `Metadata/`, `Data/`, previews, or unknown sidecars.

## Pages body-table native profile (2026-09-05)

`pages/body-table-visible.pages` was authored through Computer Use in Pages
14.4 from Blank with one Plain table (five rows, four columns). Its body marker is
`Pages hidden-axis native oracle — 2026-09-05`; the table has no user-hidden
rows or columns. SHA-256:
`7af8179b1174c39d35d4f483c65a86e3801fcea9fba2123fbbf75861af1b3b8d`.

Computer Use verification used disposable copies: Pages saved, closed, and
reopened the visible table without repair UI, and the focused package save
path produced byte-identical output for the visible-table/no-op operation. The
checked-in fixture was restored to the SHA-256 above after those UI checks;
native close/resave normalization of a disposable copy is not a hidden-axis
parity result.

This fixture verifies package ingress, exact no-op output, and the admitted
native visible-profile read/no-op path. Apple's table-info version is
`[1, 0, 5]`, while its table-model and type-4008 formula-owner versions are
`[3, 2, 10]`; the indexed current profile remains separately qualified by
`[1, 0, 5]`. The native type-4008 formula owner carries `owner_kind = 1`, and
the native producer omits some aggregate and FieldInfo declarations that the
indexed profile requires. The native profile admits only this exact
6000/6001 role pair and validates any declarations that are present.
The native hidden-state envelope is ownerful with one state, but both its row
and column state lists are empty, so the native read returns
`HiddenAxes::empty()`. An exact empty edit is a byte-identical no-op; a changed
hidden-axis request is refused as `UnsupportedDependency` before publication.
This is bounded native visible-profile read/no-op evidence, not evidence of a
Litchi visibility mutation or native hidden-axis mutation parity. Host
compatibility remains for changed native edits and unsupported producer shapes.

A separate disposable copy was edited in Pages by entering
`Native profile read back` in cell A2. Pages saved, closed, and reopened that
copy with the body marker, the five-by-four table, and the new cell text still
visible and without a repair prompt. Its post-close SHA-256 was
`5094270c73ea9a2eec6f6d5d12d8ad388787d8f2a24d77504ad905794e14be65`; this is
native authoring/save/reopen evidence for the visible profile only. It does
not demonstrate a Litchi mutation, hidden-axis parsing, or hidden-axis native
mutation parity.

## Keynote table-model discovery profile (2026-09-05)

`keynote/table-discovery.key` is a disposable copy of `keynote/basic.key`
edited through Computer Use in Keynote 14.4. A Plain 5-by-4 table was added
with `Buffa discovery` in cell A1; the original title/body text remained
visible. Keynote saved, closed, and reopened the copy without a repair prompt.
The checked-in artifact has SHA-256
`d01742f1dea413581e34469babe0df64b5d46fd2399198c4783149a8019667a7`.
This fixture supplies native producer input for the bounded Keynote
table-model discovery reader. It does not certify physical table sorting,
Litchi table mutation, native byte parity, or broader Keynote table support.

## Pages Unicode body and section discovery (2026-09-05)

`pages/body-sections-unicode.pages` was created through Computer Use in Pages
14.4 using Blank. The first section contains `Pages borrowed body 😀`, a
paragraph break, and `First section marker.`. **Insert → Section Break**
created a second section containing `Second section marker — end.`. Pages
saved, closed, and reopened the document with both pages and their markers
intact and no repair prompt, then closed it. The artifact's SHA-256 is
`8d26202b6a184c92e8c232d4886ba64ce6efb929bfe30818e506a24b041c6a22`.

The fixture verifies native body/section discovery across a UTF-16 surrogate
pair and exact source preservation. It is native authoring and read evidence;
it does not qualify a Litchi section mutation or hidden-axis operation.

## Numbers Custom Number source (2026-09-06)

`numbers/custom-number-native.numbers` was authored through Computer Use in
Numbers 14.4 from Blank. Cell B2 contains the numeric value `42` and uses the
native Custom Number format `Native Grouped`, created with the default integer
token (`#,###`). Cell C2 contains `Native Custom marker`. Numbers saved,
closed, and reopened the exact file path without a repair prompt. The Cell
inspector confirmed `Native Grouped` and sample `42` after reopen; the marker
remained visible. The document was then closed. SHA-256:
`dc804f72667d8544f3209437232c70c3934b28ab86a9bd6718109f44a1cea342`.

This source records native authoring and save/reopen behavior. It does not by
itself qualify a Rust-generated Custom-format mutation.

The focused Numbers API then replaced B2's format with `Rust Grouped`
(`#,##0.00`). Numbers opened the candidate with B2 displayed as `42.00` and
the Cell inspector naming `Rust Grouped`. After changing C2 to
`Native Custom marker saved`, Numbers saved, closed, and reopened the exact
candidate path with the format and marker intact and no repair prompt.
`numbers/custom-number-native-resaved.numbers` retains that native-resaved
candidate. SHA-256:
`af1ccfdb5dfc1f28a4a8bce2daafd0e567c494ae5434935e361420edb22a859a`.

The focused API also cleared B2's Custom format from the native-resaved
candidate. Numbers displayed `42` with `Automatic` in the Cell inspector.
After changing C2 to `Native Custom clear saved`, Numbers saved, closed, and
reopened the exact clear candidate with the value, format, and marker intact.
The temporary native-resaved clear artifact had SHA-256
`432538800987ae78ee5a0d06a164d6f97abfa900bcbf4eb6dc73cbf2dda8b577`.
Both focused mutations verified semantic reopen, byte-exact no-op behavior,
and exact inverse restoration before native verification. This qualifies
these Custom Number replacement and clear cases; the separate Custom Text and
Custom DateTime records below cover those operations, while broader native
format parity remains outside this evidence.

## Numbers Custom Text source (2026-09-06)

`numbers/custom-text-native.numbers` was authored through Computer Use in
Numbers 14.4 from Blank. B2 contains the text `Orchid` with the native Custom
Text format `Native Label`, whose prefix is `Native [` and suffix is `]`.
C2 contains `Native Text marker`. Numbers saved, closed, and reopened the
exact path without a repair prompt; B2 displayed `Native [Orchid]`, and the
Cell inspector showed `Native Label` with the same sample. The document was
then closed. SHA-256:
`77f43ffada11cb5267aaf5f208cfeb1de07bcb944053c8a6679ca014352767e1`.

`numbers/custom-text-native-unicode.numbers` was derived from that native
source in Numbers by renaming the format to `Native Unicode` and adding an
emoji before the prefix. Numbers saved, closed, and reopened the exact path
with `😀Native [Orchid]`, the unchanged C2 marker, and `Native Unicode` in
the Cell inspector. SHA-256:
`12c0b819002b5175cc91d15012cb3fc73c5acfd786ca769a93cd30c3d01a6c26`.
The native pattern's cached `index_from_right_last_integer` is `9` in the
first fixture and `11` in this Unicode fixture, matching the final UTF-16
code-unit index of each pattern, including the cell-text token. The emoji
distinguishes UTF-16 units from UTF-8 bytes and Unicode scalar counts.

The focused Numbers API replaced B2's native format with `Rust Label`, using
prefix `Rust <` and suffix `>`. Numbers opened the result with `Rust <Orchid>`
and `Rust Label` in the Cell inspector. After changing C2 to
`Native Text marker saved`, Numbers saved, closed, and reopened the exact
candidate path with the value, format, and marker intact and no repair prompt.
`numbers/custom-text-native-resaved.numbers` retains that native-resaved
candidate. SHA-256:
`8bb5fdc9e7c9325a2cccb0d76ba06a1cf2e88bb5ff046d2014aea3d65ff80729`.

The focused API then cleared B2's format from that native-resaved candidate.
Numbers displayed `Orchid` and `Automatic` in the Cell inspector. After
changing C2 to `Native Text clear saved`, Numbers saved, closed, and reopened
the exact clear candidate with the value, format, and marker intact and no
repair prompt. The temporary native-resaved clear artifact had SHA-256
`0e9ed395470b2579846fcb8493f0ab58f51d2d41f4b720846a61b2a7fabcad37`.
Both Rust mutations verified semantic reopen, byte-exact no-op behavior,
and exact inverse restoration before native verification. Rust also read the
native-resaved clear result and confirmed that another clear was an exact
no-op. This qualifies these Custom Text replacement and clear cases; the
separate Custom DateTime record below covers that operation, while broader
native format parity remains outside this evidence.

## Numbers Custom DateTime source (2026-09-06)

`numbers/custom-datetime-native.numbers` was authored through Computer Use in
Numbers 14.4 from Blank. B2 stores January 2, 2024 at 15:04:05 with the native
Custom Date & Time format `Native Calendar` (`MMM d, y`); C2 contains
`Native DateTime marker`. Numbers saved, closed, and reopened the exact path
without a repair prompt. B2 displayed `Jan 2, 2024`, while the Cell inspector
retained the actual value `1/2/2024 3:04:05 PM` and `Native Calendar`. SHA-256:
`1ca74fb3e89ee2cdc913ad0a125f8f8118f14d684d9633e9038024e4c3d8ac27`.

The focused Numbers API replaced B2's format with `Rust Calendar`
(`yyyy-MM-dd HH:mm:ss`). Numbers displayed `2024-01-02 15:04:05` and confirmed
`Rust Calendar` in the Cell inspector. After changing C2 to
`Native DateTime marker saved`, Numbers saved, closed, and reopened the exact
candidate path with the value, format, and marker intact and no repair prompt.
`numbers/custom-datetime-native-resaved.numbers` retains this artifact. SHA-256:
`8e7843c6d68ea5088a39c28dc517c552747d6231029e67d4e66a237975e24987`.

The focused API cleared B2's Custom format from the native-resaved candidate.
Numbers displayed `1/2/24 3:04 PM` with `Automatic` in the Cell inspector and
the full actual value `1/2/2024 3:04:05 PM`. After changing C2 to
`Native DateTime clear saved`, save, close, and exact-path reopen preserved
these values without a repair prompt. The temporary native-resaved clear
artifact had SHA-256
`78c2a0a9fbfcc8bc5fcc0615c3430e684109985b3280bf4c0878a4ff165ee758`.
Both Rust mutations verified semantic reopen, three-component locality,
byte-exact no-op behavior, and exact inverse restoration. Rust also read the
native-resaved clear result and confirmed another clear was an exact no-op.
This qualifies these Custom DateTime replacement and clear cases; broader
native format parity remains open.


## Numbers Duration source (2026-09-06)

`numbers/duration-native.numbers` was authored through Computer Use in Numbers
14.4 from Blank. B2 contains `1h 23m 45s` with explicit `Duration`, abbreviated
style, and automatic units spanning hours through seconds. C2 contains
`Native Duration marker`. Numbers saved, closed, and reopened the exact path
with the value, marker, and Duration inspector settings intact, without a
repair prompt. SHA-256:
`d146050a7bff23622cb3acb759d53446e81d55673b806362f7233e02ca01b10d`.

The focused Numbers API changed B2 to colon style with custom hours-through-
seconds units. Numbers displayed `1:23:45`; the Cell inspector showed
`Duration`, `Custom Units`, and style `0:00:00`. After changing C2 to
`Native Duration marker saved`, save, close, and exact-path reopen preserved
all these values without repair. `numbers/duration-native-resaved.numbers`
retains that artifact. SHA-256:
`656da1ea88e5afc2d9c7ec450714a9f30397ac30c8ebce4ef9dad92c7f78b50d`.

The focused API then cleared the explicit Duration format. Numbers displayed
`1h 23m 45s` with `Automatic` in the Cell inspector. After C2 changed to
`Native Duration clear saved`, save, close, and exact-path reopen again
preserved the value, marker, and Automatic setting without repair. The
closed temporary native-resaved clear artifact had SHA-256
`c5c1343cf29e685dc017dd9924c082c9ed351b12aa9561158a73c54095a39938`.
Rust read that result and verified another clear was an exact no-op.

Both Rust mutations verified semantic reopen, exact no-op and inverse
restoration, unchanged stored Duration and cached scalar, and locality to the
format-list and tile components. These observations qualify this Duration
replacement and clear profile; broader native unit/style parity remains open.

## Numbers audio playback source (2026-09-06)

[`numbers/audio-playback-native.numbers`](numbers/audio-playback-native.numbers)
was authored through Computer Use in Numbers 14.4 from Blank. B2 contains
`Native playback marker`, and the existing
`test-data/poi/test-data/slideshow/ringin.wav` asset was inserted. The source
UI showed loop `None` and volume `1`. Rust set `Repeat` and `0.5`; Numbers UI
displayed `Loop` with typed semantic `Repeat` and volume `0.5` after
save/close and exact-path reopen, with B2 changed to `Native playback marker saved`.
The source SHA-256 is
`eaf23bbf21364715be43a1857b8065e7061bafb17ea4c8c99f8b5f5f2766e4ed`.

[`numbers/audio-playback-native-resaved.numbers`](numbers/audio-playback-native-resaved.numbers)
retains the native-resaved candidate. Its SHA-256 is
`eff3046e643345c417e048c9ac743383ab4c539fcb6a0d5bb8dc00acaf8756aa`.
Rust changed only `Index/Document.iwa`; the audio asset was unchanged and the
repeated setter was an exact no-op after the first write. This qualifies the
recorded Numbers playback profile only.

## Pages audio playback source (2026-09-06)

[`pages/audio-playback-native.pages`](pages/audio-playback-native.pages) has
SHA-256
`f6078e869651c75a689c569dbc3c9c8517896fb59214fc401ebf19108ec7e301`.
It contains `Native Pages playback marker` and the same `ringin.wav` asset.
The source reopened with loop `None` and volume `1`. Rust `Repeat` plus `0.5`
passes the new profile, changes only `Document.iwa`, and has an exact no-op
afterward. After save, close, and exact-path reopen, Pages displayed `Loop`
with typed semantic `Repeat` and volume `0.5`, retained `Saved Native Pages
playback marker`, and showed no repair prompt. The copied resaved fixture is
[`pages/audio-playback-native-resaved.pages`](pages/audio-playback-native-resaved.pages)
with SHA-256
`62e47a4ec6a2d9667b85ea836b2a9bf2bb1dab1118815eed320c9ed3cdf5c102`.

## Image adjustment native sources (2026-09-06)

Computer Use authored each source with `abstract1.jpg`, saved it, closed it,
and reopened the exact path without a repair prompt. The source wire omits
exposure and saturation (`None`) and carries explicit `EnhanceDisabled`; the
applications displayed `0%`, `0%`, Enhance off, and Advanced sharpness `25%`.

- [`keynote/image-adjustments-native.key`](keynote/image-adjustments-native.key)
  contains the `Native image adjustment marker` title. SHA-256:
  `f307ef4e215f18f198a81bfa7160ecece59d77d0fdbd069b85a4642a6156f3a8`.
- [`numbers/image-adjustments-native.numbers`](numbers/image-adjustments-native.numbers)
  contains the same marker in B2. SHA-256:
  `653771a15c7a3de99b949b3c0fa628b5c0aa2734d557098ff3a55f75de635d03`.
- [`pages/image-adjustments-native.pages`](pages/image-adjustments-native.pages)
  contains `Native Pages image adjustment marker` in the body. SHA-256:
  `eaf32fb9385d065aeda1c77ad7212ab7bc4962fa17c3c0a4a4c99815819d584f`.

The focused Rust candidates set exposure to `0.25` and saturation to `-0.2`,
retain `EnhanceDisabled`, change only the selected IWA component, and become
exact no-ops when repeated. Each application then displayed `25%`, `-20%`,
Enhance off, and Advanced sharpness `25%` after a forced native text edit,
save, close, and exact-path reopen; the controls, sharpness, and text
persisted without repair.

The copied resaved candidates are:

- [`keynote/image-adjustments-native-resaved.key`](keynote/image-adjustments-native-resaved.key)
  retains the title `Native image adjustment marker saved`. SHA-256:
  `63e7d13124a72579521047bc5d2740506ebc3b6378fe6a901d346ed53670a658`.
- [`numbers/image-adjustments-native-resaved.numbers`](numbers/image-adjustments-native-resaved.numbers)
  retains an extended B2 marker after native UI text append; this record does
  not claim an exact single saved suffix. SHA-256:
  `18e00c8ba0e88318c9f298d74af97da756e3cc4150934aa995d84f2972dc52a6`.
- [`pages/image-adjustments-native-resaved.pages`](pages/image-adjustments-native-resaved.pages)
  contains `Saved Native Pages image adjustment marker[image]` in the body.
  SHA-256:
  `e6736692dc9cbebcc66fc584843da09ae790e6a67261ef4e44052bd47235185c`.

All three focused library suites pass. Native integration remains in progress,
and full workspace hooks remain pending the root task.

## Direct Numbers and Pages image transactions (2026-09-06)

The focused Numbers package now exposes a selector-first sheet-image
transaction. Using the existing
[`numbers/image-adjustments-native.numbers`](numbers/image-adjustments-native.numbers)
source, the Rust probe set exposure to `0.5`, saturation to `-0.4`, and
`EnhanceDisabled`, then verified an exact repeated no-op and inverse
restoration with only `Index/Document.iwa` changed. Numbers saved, closed, and
reopened the exact path without repair, displaying `50%`, `-40%`, Enhance off,
and Advanced sharpness `25%`; B2 persisted exactly as
`Native image adjustment marker direct saved` after the native text edit.

The copied resaved fixture is
[`numbers/image-adjustments-direct-resaved.numbers`](numbers/image-adjustments-direct-resaved.numbers)
with SHA-256
`75f9b78875dd7450f9c1ea89763017d000efc29358f89bd625fa6d976026f7d0`.

Pages now routes body-image mutation through a selector-first transaction. Its
Rust probe verified an exact no-op and inverse with only `Document.iwa` changed.
Pages saved, closed, and reopened the exact path without repair at `50%`,
`-40%`, Enhance off, and Advanced sharpness `25%`; the body retained exactly
`Direct saved Native Pages image adjustment marker` followed by U+FFFC, with
one 6.5 × 4.9 inch inline image. The copied resaved fixture is
[`pages/image-adjustments-direct-resaved.pages`](pages/image-adjustments-direct-resaved.pages)
with SHA-256
`0ae76968fd3a73edc50b1f52d84aeba237dd1a05317a1d16b0382607dd2a4452`.

The Keynote preview retirement was deferred under the native verification
gate. Changing the checked native slide from Title to Title Only crashed
Keynote 14.4 twice, while an unchanged copy opened successfully. The focused
preview invalidator reproduced the legacy node payload and metadata bytes;
the existing layout reassignment path needs a separate native compatibility
fix before this retirement can proceed.

Native integration passes 8 Numbers and 11 Pages tests (19 total). Focused
checks pass 7 Numbers and 6 Pages cases (13 total), including inclusive and
one-under parse and prepared-rewrite limits, archive cloning, and header
staging.

## Keynote native metadata reconciliation (2026-09-06)

The deferred Keynote crash was traced to stale metadata slide dependencies when
a layout change with no layout-owned media changed `Title` to `Title Only`.
Final archive `MessageInfo` and `FieldInfo` references now reconcile style and
template edges: obsolete unmarked component-only template links are removed
only if the final archive no longer references the old component; weak,
versioned, and unknown references are preserved.

The corrected candidate was generated after the host preview helper was
deleted and its call sites were wired to the focused hidden bridge. It passed
in `/Applications/Keynote 14.4`: a forced title edit, save, close, and
exact-path reopen left the `Title Only` title visible and the body hidden,
preserved the 1024 × 768 image centered 14.4% from the top, and retained the
marker `Saved layout Native image adjustment marker` without repair or crash.
The closed and copied fixture is
[`keynote/slide-layout-native-resaved.key`](keynote/slide-layout-native-resaved.key)
with SHA-256
`fcd7e2337cb6dd098fdeeccc7e088557591dfb7d7e401dee6266cdf67775ca27`.

The focused preview bridge passes 23 tests, the final metadata-removal suite
passes four tests including weak-reference preservation, and the host layout
unit suite passes 14 tests covering movie/live-video/image materialization and
transactional-negative cases. The `native_layout_graph` integration suite
passes two tests covering the generated candidate's exact raw image and
metadata regression plus the actual native-resaved fixture. The host still
owns graph selection and layout mutation; this does not claim broad layout-API
retirement or monolith completion. Boundary verification passes 912 cases; the
scanner remains at 64 packages, 238 declarations, and 11 ordered debts.
Normal commit hooks enforce workspace formatting, lint, library/integration,
and documentation tests. The default-feature Keynote library check passes.

## Numbers table relocation source (2026-09-06)

The 92-line host `numbers/editor/table_move.rs` module is deleted. Populated
sheet duplication now retains the exact focused public `move_table` route
alongside a source-built compatibility route used only by the duplication
helper, with a known clone identity. The package `replace_archive` path keeps
the reassembled exact source, preserving source provenance and readback; exact
input is not normalized into compatibility output. Another-sheet clone-name
collisions and legitimate rooted table-index changes during multi-table
duplication are fixed. The hidden compatibility entry is restricted to
`internal-iwork-source`.

Computer Use authored and reopened the exact path in Numbers 14.4 build
7043.0.93. The source
[`numbers/table-relocation-native.numbers`](numbers/table-relocation-native.numbers)
has SHA-256
`8aaddd0615b93dacb5e1b13a2a4ca888307b98d65935a4d2cdeb6d01e61fc81b`.
`Sheet 1/Table 1` is 22 × 7 with `Native image adjustment marker` in B2 and
a 600 × 450 image at X 57.6, Y 8.6. `Sheet 2/Destination table` is 10 × 5
with its `Destination stays` B2 marker at X 800, Y 28.

The focused name-selector move `Table 1` → `Sheet 2` succeeded. After a
forced B2 edit to `Native relocation saved marker`, native save, close, and
exact-path reopen preserved both markers without repair: Sheet 1 retained only
the image, while Sheet 2 retained its existing table and the moved table at
X 0, Y 28.3. The copied resaved fixture is
[`numbers/table-relocation-native-resaved.numbers`](numbers/table-relocation-native-resaved.numbers)
with SHA-256
`bd17f1ec7554a3da1050e56db0d8f5da276efdd7d2a4e46f7197c601cb7ffdbf`.
Scoped validation passes 18 tests: four existing media-duplication library
tests, four new host duplication integration tests, two default-feature native
relocation tests, and eight all-feature focused relocation integration tests.
Workspace formatting passes. Graph creation remains host-owned, and this
record does not claim full relocation, native-producer parity, or monolith
completion. Boundary verification passes 915 tests; the scanner passes with
64 packages, 238 declarations, and 11 ordered debts. Normal commit hooks
enforce workspace formatting, lint, library/integration, and documentation
tests.

## Keynote chart caption source (2026-09-06)

The 95-line host `caption.rs` wrapper and obsolete 339-line chart-creation
example are removed. The regression now calls the focused Package caption API
while preserving coverage of host chart duplication and removal. The new
`edit_chart_caption` example supports semantic selectors and set/clear edits.
Existing captions can retain a style in an external stylesheet: type and header
checks remain strict, and graph verification proves the style's component and
content unchanged within the work and reference budgets. Text replacement now
verifies native preview-header pruning through the focused preview owner's
exact forward/inverse delta proof. The shared movie caption/title verifier
also charges that proof to its transaction budget.

Computer Use authored the native source
[`keynote/chart-caption-native.key`](keynote/chart-caption-native.key) in
Keynote 14.4, saved, closed, and reopened the exact path. Its SHA-256 is
`afef5f17f13c4001b01ad1c83ccc7051d11c846fc437728a47d7b1828b537cb1`.
Rows are `Region 1` and `Region 2`, columns are April through July, and the
values are `[[17, 26, 53, 96], [55, 43, 70, 58]]`. Focused set to
`Focused chart caption — 北区` rendered all eight values in Keynote. After a
forced native edit to `Native saved caption — 北区`, save, close, and exact-path
reopen preserved the caption and values without repair. The copied resaved
fixture is
[`keynote/chart-caption-native-resaved.key`](keynote/chart-caption-native-resaved.key)
with SHA-256
`58bf1fc5646e88600745d4c1afa9a10d87304de5d8090f6d4ad8afc23ac5800a`.
The clear candidate also opened natively with no caption and all values
intact.

Scoped validation passes 108 tests: four caption unit tests, 23 focused chart
caption tests, four native fixture tests, 16 movie caption tests, 10 movie title
tests, and 51 host chart tests. Boundary verification passes 915 tests; the
scanner passes with 64 packages, 238 internal dependency declarations, and 11
explicit debt items. Normal commit hooks enforce workspace formatting, lint,
library/integration, and documentation tests. Chart creation, duplication, and
removal remain host-owned; this slice does not complete monolith retirement.

## Keynote chart titles source (2026-09-06)

The 108-line host `title.rs` and 130-line host `axis.rs` modules are deleted.
The `focused_chart_catalog` helper moves into the parent chart module so
host listing and selection by title continue to serve chart creation,
duplication, and removal. Eighteen obsolete examples are removed. The focused
`edit_chart_title` example supports chart, category-axis, and value-axis titles
through semantic selectors and set/clear edits.

Computer Use authored the native source
[`keynote/chart-titles-native.key`](keynote/chart-titles-native.key) in
Keynote 14.4, saved, closed, and reopened the exact path. Its SHA-256 is
`101893df3949129530b3d05ab9ba13b15cdad06258e567754733aec4995706ee`.
The chart title is `Native chart title — 北区`, the value title is
`Native revenue — 元`, the category title is `Native months — 月`, and the
caption marker is `Native chart caption marker`. Rows are `Region 1` and
`Region 2`, columns are April through July, and the values are
`[[17, 26, 53, 96], [55, 43, 70, 58]]`; all eight values survived exact-path
reopen.

Selected chart-title reads now use canonical headers with zero/absent mediator
guards. Both selection paths perform a global inbound non-style metadata
ownership census. Exact selected before/after payloads are retained for patch,
inverse, and locality checks. Changed chart-title publication rejects unknown
reference metadata; reads and exact no-ops retain opaque headers, preserving
other chart owners' contracts. Stylesheet admission accepts one exact stylesheet wire plus the aggregate
registry allowance. Full codec work is precharged; this does not claim
peak-memory or full allocator proof. Forward root-preview invalidation and
inverse/double-inverse diagnostics are correct for chart and axis titles.

Focused tests pass 10 title and 16 axis cases. The native chart candidate
opened with `Focused chart title — 北区`, `Focused revenue — 元`, and
`Focused months — 月`; axes, caption, and all eight values were preserved.
`PackageMetadata` may use a generic preferred locator alongside an exact
explicit/effective archive locator; matching now uses that exact physical
locator while retaining current-component and UUID/reference authority checks.
After a forced native edit to
`Native saved chart title — 北区`, save, close, and exact-path reopen confirmed
the native chart title, both focused axes, caption, and all eight values without
repair. The copied resaved fixture is
[`keynote/chart-titles-native-resaved.key`](keynote/chart-titles-native-resaved.key)
with SHA-256
`a12c1bd02ed5316699beef3e6cbe6422359032916848beb370efc3eae5bf0cba`.
The clear candidate opened with no chart title, Y-axis `Numeric`, X-axis
`Categorical`, and no custom names; `Native chart caption marker` and all eight
values remained intact without repair. Scoped validation passes 284 focused
Rust tests with all features: 231 Keynote library, four native fixture, 16 axis,
23 caption, and 10 title tests, plus 51 host chart tests. Boundary verification passes 915 tests; the
scanner passes with 64 packages, 238 internal dependency declarations, and 11
explicit debt items. Workspace formatting and lint pass. Normal commit hooks
enforce workspace library/integration and documentation tests. Chart creation,
duplication, and removal remain host-owned; this slice does not complete
monolith retirement.

## Image-adjustment retirement across Numbers, Pages, and Keynote (2026-09-06)

Six raw-ID host sheet-image adjustment read/set methods and three dead writers
are removed. Unused focused hidden write bridges are removed as well. Required
decode bridges remain for graph reads, while host graph CRUD and info paths are
preserved. The obsolete image-creator examples
`create_numbers_image.rs`, `create_pages_image.rs`, and
`create_keynote_image.rs` are deleted. The focused surface includes three
public semantic CLI examples,
and Numbers includes default-feature native integration coverage in
`crates/litchi-numbers/tests/sheet_image_adjustments_native.rs`.

The existing
[`numbers/image-adjustments-direct-resaved.numbers`](numbers/image-adjustments-direct-resaved.numbers)
fixture remains the confirmed Numbers baseline: exposure `50%`, saturation
`-40%`, Enhance off, Advanced sharpness `25%`, and the direct-saved B2 marker.
The retirement candidates for Numbers, Pages, and Keynote set exposure `-0.2`,
saturation `+0.35`, and `EnhanceEnabled`. Each application opened in version
14.4 after a forced native description edit, save, close, and exact-path
reopen; the controls persisted. Numbers and Pages confirmed Advanced
sharpness `25%`, and the original B2, body, and slide-title markers persisted.
Numbers retained `Focused Numbers image adjustment retirement saved`; Keynote
and Pages retained the corresponding focused retirement descriptions, with
Pages retaining its literal trailing tab.

The copied native resaved fixtures are:

- [`numbers/image-adjustments-retirement-resaved.numbers`](numbers/image-adjustments-retirement-resaved.numbers) — SHA-256 `4be76e8e71c8ac461c7c289f42f526f1fe0df73e96a54e49168c7156681f61c9`.
- [`pages/image-adjustments-retirement-resaved.pages`](pages/image-adjustments-retirement-resaved.pages) — SHA-256 `b3f6a7f281433732f7f9ec490d3d4bc684038d74235b203f2ae20d5f3b21b7bd`.
- [`keynote/image-adjustments-retirement-resaved.key`](keynote/image-adjustments-retirement-resaved.key) — SHA-256 `254ea31b758362dd292dcb79e659879f8ac2c2cab38e2a0b2e5090ce243da877`.

Focused Numbers validation covers the payload-only parent-ownership contract:
the exact rooted sheet owns the payload, any declared parent attribution is
validated, duplicates, wrong paths, and cross-kind references are rejected,
and a parent backlink is not required in aggregate or field metadata. The
data reference occurs exactly once in aggregate metadata; path 11 is optional.
Native fixtures show parent reference `904475` absent from metadata while data
reference `16` appears once. Selected merge/diff validation preserves
aggregate-only metadata. Scoped validation passes in the three focused native
integration files: Keynote 5 tests, Numbers 8 tests, and Pages 12 tests (25
tests across those files). The Numbers default-feature
`sheet_image_adjustments_native` suite passes 4 tests, and the focused Numbers
`image_adjustments` library suite passes 12 tests. Boundary verification
passes 921 tests; the scanner passes with 64 packages, 238 internal
dependency declarations, and 11 explicit debt items.

## Keynote chart-arrangement host boundary (2026-09-06)

Keynote removes two selector-first host wrappers for chart arrangement. These
were semantic wrappers rather than raw-ID methods; the internal read-batch,
listing, and lifecycle helpers remain. The dead handoff writer is removed.
The broader Pages and Numbers chart-arrangement graph and lifecycle remain
host-owned; focused flag ownership is recorded in the follow-up below. The
focused CLI retains its enhanced no-clobber and forward-conflict behavior.

Selected-graph discovery, codec work, rewrite, and locality validation charge
against one shared resource ledger. Graph allocations have a separate finite
allowance; the existing limit of 64 logical rewrite/staging allocations
remains. Selected locality stream buffers are precharged from parsed-stream
bounds. Per-payload field ceilings remain in force; graph field visits consume
the shared work budget without consuming the selected codec's field allowance.
Full semantic validation and selected slide projection remain separately
bounded; physical Snappy ceilings still apply.

Computer Use opened the arrangement candidate in Keynote 14.4 with the chart
lock and `Constrain proportions` flags enabled. After selecting the active
document from the Window menu, native unlock/relock forced a dirty save; close
and exact-path reopen retained both flags, the original chart title
`Native saved chart title — 北区`, and `Native chart caption marker`. The
checked-in native fixture is
[`keynote/chart-arrangement-retirement-resaved.key`](keynote/chart-arrangement-retirement-resaved.key),
510,176 bytes, with SHA-256
`e82fa36d0a041eeaea5ea9a41ed1b2107351ab9bc67d7888d86995bcbec16e98`.

The focused CLI read both flags, produced an exact no-op with zero touched
components and no full reparse, and preserved bytes on the existing-output
refusal path. A reset probe opened the resaved oracle with both flags false;
that reset observation was not saved or reopened as a reset candidate. Focused
validation passes 18 tests; the all-feature native integration suite covers 4
targeted tests. The host arrangement/CRUD suite passes 6/6, including the
32-chart changed-listing/exact-inverse case and Numbers/Pages CRUD cases. One
private actual-candidate verification passed exact-work and one-under replays.
Boundary verification passes 922 tests; the scanner passes with 64 packages,
238 internal dependency declarations, and 11 explicit debt items.

## Focused Numbers and Pages chart-arrangement owners (2026-09-06)

The focused Numbers and Pages packages now share the archive-free
`ChartArrangement` value for the two existing Arrange-panel controls: chart
lock and aspect-ratio constraint. Numbers resolves a `SheetSelector` plus
`ChartSelector`; Pages resolves a semantic `BodyChartSelector`. Both owners
keep native identifiers and graph payloads private, use the strict preflight
and neutral lazy Buffa [`chart_arrangement_codec`](../../crates/litchi-iwa-protos/src/chart_arrangement_codec.rs),
and bound source-preserving no-op, inverse, candidate-reopen, and locality
checks with their format-specific budgets. The neutral codec's nine focused
tests pass. Broader chart data, geometry, lifecycle, and host graph behavior
remain outside these bounded owners; package ingress remains separately
bounded, and these owner budgets do not cover every semantic constructor
allocation. No exit-gate or host-debt claim follows.

Computer Use verified the native Numbers baseline
[`numbers/chart-arrangement-native.numbers`](numbers/chart-arrangement-native.numbers)
in Numbers 14.4 after authoring, saving, closing, and exact-path reopening
without repair. It is 128,644 bytes with SHA-256
`54b74c114e05884aa68fc028fb033d0b8562d7646faa0fdc6d280ba97f2055a9` and has
title `Numbers Arrange native chart`; `Sheet 1`/`Table 1` is a 3-by-5 chart
with rows `North`/`South`, columns April through July, and values
`[[17, 26, 53, 96], [55, 43, 70, 58]]`. Both flags are false and the package
has no data assets. This is native baseline evidence only; it does not promote
changed Numbers mutation acceptance.

The focused Numbers candidate set both flags true. Numbers 14.4 opened it
without repair showing `Locked`, `Constrain proportions=1`, and an enabled
Unlock control. Root clicked native Unlock, verified `Constrain proportions=1`
in Arrange, clicked Lock, saved with Cmd-S, closed the actual window, and
reopened the exact disk path. All 15 table cells and all eight chart values,
the unchanged title, `Locked=true`, `Constrain proportions=true`, and the
enabled Unlock control were preserved before the document was closed. The
checked-in resaved fixture is
[`numbers/chart-arrangement-retirement-resaved.numbers`](numbers/chart-arrangement-retirement-resaved.numbers)
with 128,604 bytes and SHA-256
`0c7a58e7c81d5f7994eac004c572cb2a6665381b4ecde8abb3b5869194a1575a`.
Native auto-parsed month headers carried Apple-second values
`[796694400, 799286400, 801964800, 804556800]` and displayed April through
July. The standalone native test oracle verifies those exact dates after editing,
reopening, reset, and exact inverse restoration.

Computer Use also verified the Pages baseline
[`pages/chart-arrangement-native.pages`](pages/chart-arrangement-native.pages)
in Pages 14.4 from a Blank word-processing document after saving, closing, and
reopening the exact path without repair. It is 118,207 bytes with SHA-256
`dccfa28a337babf786bed9433805311843184b0c8d7f3161d8795fe0047faf42` and has
body marker `Pages Arrange native body marker` followed by the native U+FFFC
chart anchor. The chart title is `Pages Arrange native chart`, with the
default 2D Column layout, rows `Region 1`/`Region 2`, columns April through
July, and values `[[17, 26, 53, 96], [55, 43, 70, 58]]`; `Move with Text` is
`1` and `Constrain proportions` is `0`. Both Lock and Unlock controls were
disabled for this text-anchored chart, and the package has no data assets.
This is native baseline evidence only: the both-true changed-candidate gate
has now been exercised for the focused Pages operation, while interactive Lock
button parity remains unclaimed.

The focused Pages candidate set both flags true while retaining `Move with
Text=1`. Pages opened it without repair with all eight values unchanged. Root
inserted `Native saved: ` into the body to force a dirty document; Cmd-S, the
actual close button, and exact-path disk reopen preserved that prefix, the body
U+FFFC anchor, title, grid, `Locked=true`, and `Constrain proportions=true`.
The checked-in resaved fixture is
[`pages/chart-arrangement-retirement-resaved.pages`](pages/chart-arrangement-retirement-resaved.pages)
with 119,482 bytes and SHA-256
`5b8d18a8b02bb6471da1fc20d02139f8d18ebe44e421f8fd8b17e841efb76986`.
Both Lock and Unlock buttons remained unavailable for the body anchor, while
the persisted Locked state was respected because all chart edit controls were
disabled; this is persisted-state evidence, not interactive button parity.
Focused integration coverage passes 25 tests: Numbers has 7 synthetic and 4
native tests, while Pages has 10 synthetic and 4 native tests. Both native
resaved fixtures, strict Date headers, body/U+FFFC/grid preservation, and
metadata-ownership regressions are included. The Pages private ZIP-mask unit
passes 1 test. The Keynote shared-rename regression set passes 22 tests;
boundary verification passes 927 tests.

## Numbers Custom-format raw-ID retirement (2026-09-06)

The three production `NumbersEditor` Custom conveniences
(`table_cell_custom_format`, `set_table_cell_custom_format`, and
`reset_table_cell_custom_format`) were removed. Existing-cell Custom reads
and edits now use the focused selector-first `litchi_numbers::Package` APIs;
native IDs, registry UUIDs, and format-list keys remain private. The obsolete
`crates/litchi-iwa/examples/create_iwork_table_number_formats.rs` example was
also deleted, leaving no legacy example entry point.

The focused CLI changed one existing Custom Number cell to the name
`Focused Custom Retirement` and pattern `#,##0.000`. Exact no-op,
candidate-reopen, inverse, three-component locality, and full-reparse checks
passed. Numbers 14.4 opened the candidate without repair with B2 displaying
`42.000` and the same name in the inspector. After C2 was changed to
`Focused Custom retirement saved`, native save, close, and exact-path reopen
preserved the formatted value, underlying value `42`, inspector name, and
marker. The checked-in native-resaved fixture is
[`numbers/custom-retirement-resaved.numbers`](numbers/custom-retirement-resaved.numbers),
137,073 bytes, SHA-256
`d7636fa0e468696b1a0d5f3ec06cc98e26eeeedb734d60f157e3ef7c1ebd227d`.

The focused reset candidate passed the same exact-source checks. Numbers
reopened it without repair with B2 displaying `42` and `Automatic`; after C2
was changed to `Focused Custom clear saved`, native save, close, and exact-path
reopen preserved the value, Automatic state, and marker. The checked-in
clear fixture is
[`numbers/custom-retirement-cleared.numbers`](numbers/custom-retirement-cleared.numbers),
136,435 bytes, SHA-256
`f854090055544d4f59cf9023c95def9a303134b7a8c3f8ff17242c34fbe9fc65`.
The final focused reread of that fixture was an exact byte-equal reset no-op
(`changed=false`, zero touched components, `full_reparse=false`).

The existing Custom Number, Custom Text (including the UTF-16 cache profile),
and Custom Date & Time records above remain operation-specific native
evidence. Generic source-built or cross-format `DataFormat::Custom`
compatibility and private attached Pages/Keynote table adapters remain
host-owned; broader Custom-format authoring and monolith exit gates remain
open.

## Keynote slide-media replacement baseline and assets (2026-09-06)

[`keynote/media-replacement-native.key`](keynote/media-replacement-native.key)
is a native Keynote 14.4 source for the focused existing slide-media data
owner. The closed artifact is 752,060 bytes with SHA-256
`f5763984974612f078486cb2310f408cbcb7cd06ef6adca494aae19eaf62609d`.
Computer Use saved it, actually closed the document, and reopened the exact
path without an error. The selected slide retains the title
`Keynote media replacement native` and body marker
`Media ownership marker — 北区`. Its source-order media are Audio, Audio,
File, File.

The two audio controls share one 192,044-byte WAV materialized record (the
source record is SHA-256
`8cee734c146cbe9dbbd961a44aa577eaaa4170dab46de0bf60eb8d3f1cb46c70`)
and the two file movies share one 34,651-byte self-authored MJPEG MOV record
and one 4,408-byte native poster record. The MOV SHA-256 is
`470ea8ba876c8ee4f7d50fda3e482b0ee5da7211be5eab52f37809b7f5c2e7dc`; the
poster SHA-256 is
`5b38397c34eb2e9cf810f7269306b17da3ec6298f8c1b2e9f20eab65fc56a6d5`.
Both movie captions are `Shared native movie caption`, and their 320×180
frames are positioned at (200,700) and (700,700). Native playback retains a
0–2 second trim, volume 1, loop disabled, start on click, and play across
slides.

The replacement inputs are locally authored and checked in under
[`keynote/media-replacement-assets/`](keynote/media-replacement-assets/):

| Asset | Bytes | SHA-256 | Role |
| --- | ---: | --- | --- |
| [`striped.mov`](keynote/media-replacement-assets/striped.mov) | 87,997 | `e4b92ec504d372e2330b37d4dcd0441d2990096f41f880343795c45508a0c997` | File-movie content candidate |
| [`tone.wav`](keynote/media-replacement-assets/tone.wav) | 192,044 | `27f1b6add1bd884455da7a36244d92fc26b81b622ce1671d5cb77bebffeb135a` | Audio content candidate |
| [`poster.png`](keynote/media-replacement-assets/poster.png) | 1,698 | `d594d8f832d3ab38d03e0d06316d2fa5e401217e50e21e01814e593d752da3e8` | Movie poster candidate |

The MOV was authored locally from CoreGraphics/ImageIO JPEG frames and
QuickTime atoms, the WAV is local PCM, and the poster is local PNG. No vendor
video resource is retained in this baseline or replacement set.

The combined pre-native candidate applied `tone.wav` to audio position 0,
`striped.mov` to file-movie content position 2, and `poster.png` to that
movie's poster. It is 742,670 bytes with SHA-256
`3d6cae99029d0e32ed090576b3a1df3a73885e75066975c234157467a22af600`.
The shared audio and movie records remained shared across their two source
occurrences, and all three selected ZIP payloads matched the replacement
assets byte-for-byte.

Keynote 14.4 opened the candidate without repair or error. After playback, the
body marker was changed to `Saved media ownership marker — 北区`; native save,
actual close, and exact-path reopen succeeded. The native-resaved artifact is
[`keynote/media-replacement-retirement-resaved.key`](keynote/media-replacement-retirement-resaved.key),
807,293 bytes with SHA-256
`31524a79cdb6e5c421fd59953225490f4f0a51bf15d879ad84676dc57dbdc152`.
All four media objects, both shared captions, both 320×180 movie geometries,
and the 0–2 second, volume-1, loop-off, click-start, across-slides playback
settings survived without repair or error.

Keynote displayed the video-derived first frame for the replacement movie.
The arbitrary `poster.png` bytes were retained byte-for-byte as the selected
poster record, so this verifies poster storage and metadata preservation rather
than a visual match for an arbitrary PNG preview. Strict Litchi reread of the
native-resaved artifact passed six shared-record reads and six exact no-op
writes, establishing operation-specific E4 for the three replacement paths.
The focused suite reports 21 owner cases plus 2 budget cases (23 total), while
neutral metadata validation reports 15 module and 18 integration cases (33
total); full boundary verification reports 933 policy units across 64 crates,
238 internal edges, and 11 explicit debt items. A bounded 256-run ASAN smoke
passed with harness assertions covering all three changed paths and their exact
inverses; this is bounded E1 fuzz evidence, not exhaustive coverage. Previous
full hook runs passed in commits `f57f5e74c` and `f4b4865ae`; those repository
checks are prerequisites, not lifecycle mutation evidence. The separate
caption mutation owner refuses this profile, so this record does not claim
native geometry or caption mutation parity.

## Keynote media lifecycle native oracles and remaining owner boundary (2026-09-06)

Two permanent native-authored Keynote 14.4 oracles record expected lifecycle
behavior. [`keynote/media-lifecycle-duplicate-native.key`](keynote/media-lifecycle-duplicate-native.key)
is 752,730 bytes with SHA-256
`e9a9d2779749252861a9fc592a3af5020b92ca3cb9adf14a03d155606bb7c276`; it has
five media objects, a duplicate offset by `(10,10)`, and an appended caption.
[`keynote/media-lifecycle-remove-native.key`](keynote/media-lifecycle-remove-native.key)
is 697,237 bytes with SHA-256
`88813924afc1566dcffa92a8a59ab33e65b919bcef83dc3352cab3782bf95d13`; removing
the final movie leaves two audio objects and culls the movie/poster assets.
Both artifacts were saved, actually closed, and reopened from their exact
paths in Keynote 14.4.

These are native-authored expected-behavior oracles only. They do not certify a
Litchi-mutated candidate or establish focused mutation E3/E4. The neutral
prerequisite implementation is now landed: the core clone path preserves raw
headers and source bytes, stages source-atomic remaps, sorts exact remap scratch
deterministically, and requires explicit consistent remaps for known self-references. The neutral map decoder
uses a bounded two-pass traversal, one temporary allocation for nonempty O(n log n)
ordering, lazy Buffa views with canonical `int32` fields, and reference-parity
checks. The metadata codec atomically removes final owners and their `DataInfo`
records. The native baseline test removes four owner references and two
`DataInfo` records. Mapped or ambiguous records and surviving empty or
versioned component references refuse removal before publication.

The existing focused media-replacement closure delegates map decoding to the
neutral reader and has an unknown-extension read-preservation regression. Core
validation passes 6 cases, neutral map validation passes 22 unit plus 18
integration cases (40 total), Keynote validation passes 22 replacement plus 4
native-oracle integration cases (26 total), and boundary verification passes
936 units. Workspace strict linting passes, and both clone/remap and metadata fuzz
targets pass 256 AddressSanitizer runs. The full boundary scanner passes (64 packages, 238 internal dependency
declarations, 11 explicit debts). Commit `55499d9c5` passed normal formatting,
manifest sorting, strict workspace lint, all-feature workspace library and
integration tests, and documentation tests. At this Sep 6 snapshot the
selector-level lifecycle owner and further host-route retirement remained
pending; the Sep 7 section below supersedes that status.

## Keynote audio lifecycle native oracles (2026-09-06)

The audio branches of the lifecycle matrix have two additional permanent
Keynote 14.4 oracles. [`keynote/media-lifecycle-audio-duplicate-native.key`](keynote/media-lifecycle-audio-duplicate-native.key)
is 752,241 bytes with SHA-256
`4613f7a275388a1d053407f849a3845de2789acdfe5c36139892d20a17e57c3c`; it
contains three audio objects and two captioned movie objects, and retains the
WAV, movie-content, and poster data records. [`keynote/media-lifecycle-audio-remove-native.key`](keynote/media-lifecycle-audio-remove-native.key)
is 559,085 bytes with SHA-256
`052e6389af8719e2e6ffadf2d5fbae0c5275983b94c5ce9efd811f59cf1c1bfb`; it
contains no audio objects and retains two captioned movies, removes the WAV
record (`DataInfo` 9075), and retains movie-content/poster records (`DataInfo`
9085 and 9086). Both permanent artifacts were authored, saved, actually
closed, and reopened from their exact paths in Keynote 14.4.

The first-audio-removal intermediate snapshot retained the shared WAV through
the remaining audio occurrence. It is a temporary, uncommitted 751,692-byte
diagnostic artifact with SHA-256
`7e39344f54222786ae3241ef4f14ebcefdfc96bb0c54a7c8401348a71318f36e`; it did
not receive the native close/reopen gate. These artifacts are native-authored
expected-behavior oracles only. They are not Litchi-mutated candidates and do
not establish focused lifecycle E3/E4. At this Sep 6 snapshot the
selector-level lifecycle owner and its wire, clone-payload, and
metadata-adapter implementation were still in progress; the Sep 7 owner and E4
section below supersedes that pending status.

## Keynote focused media lifecycle owner and native E4 receipts (2026-09-07)

The focused `litchi-keynote` owner now exposes
`Package::{duplicate_slide_media, remove_slide_media}` and typed movie/audio
aliases `duplicate_slide_movie`, `duplicate_slide_audio`, `remove_slide_movie`,
and `remove_slide_audio`. `SlideSelector` plus source-order `MovieSelector`
keep native identifiers, UUIDs, component paths, `DataInfo` keys, and ZIP
members private. The returned `SlideMediaLifecyclePatch` is bound to the exact
source, supports `inverse()` and `is_noop()`, and is replayed through
`Package::apply_slide_media_lifecycle` only after bounded candidate readback.

The owner clones or removes the selected slide/build/build-chunk graph,
preserves shared materialized data until its final owner, and reclaims final
data records only after a package-wide ownership census. One `LifecycleBudget`
is shared across graph selection, clone-payload work, metadata transitions,
lazy wire decoding, archive encoding, ZIP reassembly, and candidate readback.
The lazy
[`KNMediaLifecycleArchive.proto`](../../crates/litchi-iwa-protos/src/buffa-projections/KNMediaLifecycleArchive.proto)
projection is 957 bytes and contains five messages; its borrowed snapshots and
source-preserving rewrites avoid generated repeated views.

Computer Use verified six operation-specific native E4 profiles from temporary
candidates in `/private/tmp/litchi-media-lifecycle-20260906r/candidates`. Each
was opened at its exact path in Keynote 14.4, saved with Cmd-S, actually closed
until the theme chooser appeared, reopened at its exact path, checked for
expected media counts, text, and captions without alerts, and closed again.
The native-saved receipts below remain temporary copies in
`/private/tmp/litchi-media-lifecycle-20260906r/native-saved`; none were copied
into the permanent fixture set:

| Native-saved receipt | Expected profile | Bytes | SHA-256 |
| --- | --- | ---: | --- |
| `focused-native-duplicate-movie.key` | 3 movies, 2 audio | 752,707 | `bb2f3d0318a311ba0cfdf4fa1d8b88488f8e0172ba02fda73d689ea15ae0344a` |
| `focused-native-duplicate-audio.key` | 2 movies, 3 audio | 752,189 | `39eb5ba491acbdf2bc7295a2c549520ed538e1ca73f5b5e109b0cb74d587be38` |
| `focused-native-remove-movie-shared.key` | 1 movie, 2 audio | 745,166 | `16d1241c50fb4ad093e10c2befebac1303a9442e4efda22ef6355c6d2b933c91` |
| `focused-native-remove-audio-shared.key` | 2 movies, 1 audio | 751,698 | `452af453a6b453aa454319e1033a1b5be0389a8954c245d2b789bee310c97170` |
| `focused-native-remove-movie-final.key` | 0 movies, 2 audio | 698,799 | `9e3b40157610ded1f618d64112c4e964c6f03089a5ceabb9fb7899eca5320b99` |
| `focused-native-remove-audio-final.key` | 2 movies, 0 audio | 559,060 | `1477986335aca0e093b404a259bba9d866d7083a101aeeb1b2b1509982f03ece` |

The lifecycle integration target passes 21 cases: the original 18 plus three
header-reference refusal regressions. With
`LITCHI_KEYNOTE_MEDIA_LIFECYCLE_NATIVE_SAVED_DIR` enabled, the direct
integration test `native_saved_candidates_are_read_back_without_rewriting_them`
strictly rereads all six native-saved candidates without rewriting them. Four
focused native lifecycle tests pass for the six exports, establishing
operation-specific lifecycle E4 evidence only.

The neutral identity codec passes 45 unit cases, including the versioned
component accounting regression; the new lazy lifecycle codec passes 10 cases;
and strict protobuf and Keynote Clippy pass. The private lifecycle unit slice
passes 11 cases, while existing replacement validation passes 28 cases (22
replacement, 2 budget, and 4 native-oracle cases). The hardening fixes cover
the actual `ShapeInfoArchive` reference edge, exact versioned-component work
charging, source-relative core-header precharge with atomic byte/event
accounting, and serialization precharge before candidate allocation.
Regenerating all six focused candidate inputs produces bytes exactly equal to
the native-verified originals. Fresh-export equality and strict native-saved
readback both pass. The bounded lifecycle AddressSanitizer campaign passed 256 runs with no
findings (244 MiB peak RSS); boundary verification passes 942 unit tests.
Implementation commit `182f31566` passed all normal repository hooks:
formatting, manifest sorting, workspace library Clippy, all-feature
library/integration tests, and doctests, and no host lifecycle route was retired.

The six temporary focused candidates and native-saved copies were removed
after verification. The isolated sanitizer target was cleaned with
`cargo clean` (2.1 GiB reclaimed); permanent audio oracle fixtures remain
checked in.

## Keynote lifecycle cache-hardening scope (2026-09-07)

This extends the verified lifecycle baseline `ca9b9f5cd`; no host route is
retired. Typed movie/audio calls resolve complete source-order media positions,
including live-video and placeholder siblings, and report `KindMismatch`
before metadata rewriting. Selected direct comments return
`UnsupportedComment` until reply/author ownership has a dedicated transaction.
Comments on unselected media remain preserved.

A changed build topology invalidates only the selected `SlideNode` caches:
fields 15, 20, 22, and 23 are removed; fields 26 and 27 become `u32::MAX`.
Already-invalidated nodes keep their exact component bytes. Other node fields
and ZIP members remain preserved, and inverse patches restore the original
cache bytes. Five Keynote-authored source/duplicate/removal fixtures establish
the invalidated baseline state. Keynote may recompute a cache during save;
that native saved state is checked separately from the transaction output.
The sixth scalar-only Buffa projection keeps five generated files totaling
191,918 bytes under a 192 KiB bound, with no generated repeated views.

Source-built interoperability covers mixed movie/audio order, multiple builds,
titles/captions, duplicate/shared/final-owner removal, host reopen, data
retention/culling, exact inverse, and unselected-comment preservation. A narrow
Movie → CaptionInfo → ShapeStyle witness admits the producer's transitive
style references without ignoring unexplained header edges. Playback UUIDs
and registry UUIDs are separate identity domains. Clone identities use a
deterministic hash with collision checks instead of XOR, which collided for
source-built label graphs. Chunk UUID occurrences must agree within a build.

The focused allocation watermark remains monotonic: removals retain
`last_object_identifier`, and duplicates allocate above it. The host's
trailing-suffix release remains separate legacy behavior. Source-built
interoperability is E1 evidence; comment ownership and the remaining host
compatibility decisions still gate ADR 0028 deletion.

Targeted verification passed: 27 focused lifecycle integration tests, three
source-built interoperability tests, 16 neutral lifecycle codec tests, strict
Keynote/protos Clippy, 942 boundary unit tests, and the full scanner (64
packages, 238 internal edges, 11 explicit debts). The final cache candidate
also passed the environment-enabled native-saved strict readback test. Implementation commit `fe4592c68` passed the normal workspace hooks:
formatting, lint policy, all-feature library/integration tests, and
documentation tests. The codec source-tracking guard also covers the
nested node-cache module path. This follow-up retires no host routes.

## Keynote comment duplication evidence (2026-09-07)

The focused owner implements selected direct comment/reply duplication through
the neutral lazy batch codec and a private comment-graph/author dependency
witness. Before the admission hardening, the focused lifecycle library slice
passed 28 cases and the integration target passed its 27 existing cases plus
11 comment-duplication cases. The default integration run does not enable
native saved-candidate readback, while the explicit native environment run
passed one strict reread. These counts are pre-hardening; the current final
rerun remains pending.

The native baseline comment root is on movie `2653286`, with storage
`2653723`, author `2653721`, and `externalAnnotationAuthorStorage` `2652381`.
Native Cmd-D produced movie `2653814` and storage `2653826`, kept the storage
UUID byte-exact, and shared the author records. After native save, actual
close, and exact-path reopen, the duplicate contained three movies and two
audio objects with comments on both copies; the baseline contained two movies
and two audio objects with its original comment.

The temporary native removal oracle culled the comment root while retaining
the author and author-storage records. It removed only component external edge
`(2652150, 2652381, 2653721)`, changing the external set from 700 to 699;
the duplicate external set remained 700. Permanent fixtures:

- [`media-comments-baseline-native.key`](keynote/media-comments-baseline-native.key) — SHA-256 `69d493b183308b6a0f336b6a944b2d78fe8439a40648a0dcb72ff1e171380fff`
- [`media-comments-duplicate-native.key`](keynote/media-comments-duplicate-native.key) — SHA-256 `8c04282f5877ae671c808cb6a35022a3c41432459fb2bc0e43a22378d3c59cb6`

The temporary removal oracle hash is
`1d0897ebd1f79b5e5ee3a33ce7f034f69baec9bac05be6cddad73d3dd828d509` and is
not a permanent fixture.

The exported focused candidate
`focused-media-comment-duplicate-movie.key` changed from SHA-256
`d3c839320da017b48e79627d50f48eace995978c2344b3be8fb7739711ae1869` to
`8c3ed9b005a5bc280d9d9859d4c4b0dae71c958ec04dbc2de486e6d77a0472a8` after
Keynote Cmd-S, actual close to the theme chooser, and exact-path reopen. It
opened with three movies and two audio objects, retained two comments, and
showed no alert. The explicit strict reread verified the new comment IDs and
preserved storage UUIDs and authors. This focused native duplicate leaf remains
verified; native reply and comment removal remain unverified.

No native reply was authored because the prototype Keynote reply popover was
not operable; synthetic and source-built reply graphs remain separate
evidence. The five host cases pass, covering positive selected movie/audio
reply cloning plus the legacy host reply-duplicate/remove refusal oracle.
Comment removal remains `UnsupportedComment` and no host route has been
retired.

The current extension rejects unknown raw-header references and unknown or
deprecated direct-comment references before publication. New selected-comment
admission is fail-closed when opaque unknown fields in the comment root, date,
or UUID could hide a reference back to the source graph. Admission requires a
strict known comment envelope and an exact header-edge census. The neutral
codec still preserves unknown bytes, and untouched or unselected comments
remain preserved. Full-inspection budgets are shared across graph, author,
header, codec, and metadata walks. Boundary verification previously passed
943 unit cases and `litchi-iwa-protos --lib` previously passed 846 cases,
including corrected mixed-UUID shrink/reply-varint-growth exact raw/Buffa
parity; those are pre-hardening counts and the current final rerun is pending.
The full scanner, strict Clippy, and normal hooks remain pending. No cleanup
result is claimed here. The next goal is author culling plus metadata removal.

Final selected-comment admission verification passes: the complete Keynote
library has 258 passing tests, with 27 existing media lifecycle and 12 comment
integration tests also passing. The explicit native-saved readback ran with its
path supplied. The final generated duplicate has the same SHA-256
`d3c839320da017b48e79627d50f48eace995978c2344b3be8fb7739711ae1869` as the
artifact verified in Keynote before native save. Strict Clippy for both changed
production libraries and workspace formatting pass. Boundary unit verification
passes 943 cases. Full boundary scanning and normal commit hooks are pending
the final receipt.

Final commit and cleanup receipt: `ce6ab2e43` passed the normal pre-commit
hooks: workspace formatting, all-feature production lint, all-feature workspace
library/integration tests, and workspace documentation tests. The full boundary
scanner passed for 64 packages and 238 internal dependency declarations, with
11 explicit migration-debt items. This supersedes the pending final-verification
notes above. No further host route was retired.

After verification, the owned temporary directory
`/private/tmp/litchi-media-comments-20260907t` was removed. `cargo clean` removed
14,129 files and reclaimed 9.8 GiB; free disk space was approximately 62 GiB.
The two checked-in native comment fixtures remain. Native reply authoring and
selected commented-media removal remain outside this completed duplication slice.
