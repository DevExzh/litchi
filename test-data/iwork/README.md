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
