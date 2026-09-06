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
these Custom Number replacement and clear cases; Custom Text, Custom DateTime,
and broader native format parity remain outside this evidence.

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
no-op. This qualifies these Custom Text replacement and clear cases; Custom
DateTime and broader native format parity remain outside this evidence.

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
