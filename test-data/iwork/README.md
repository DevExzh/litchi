# Native iWork fixtures

These packages were made for this repository. No upstream project's file is
copied here except the two inserted media assets named below, and the license
is the repository's.

Unless a section says otherwise, each fixture was authored or edited in the
native macOS application (Keynote, Numbers, or Pages 14.4), saved, actually
closed, and reopened from its exact path without a repair or conversion
prompt. "Focused" fixtures were first produced by the named `litchi-keynote`,
`litchi-numbers`, or `litchi-pages` transaction and then opened, re-saved, and
reopened in the native application; the checked-in bytes are the native-saved
result. Fixtures not described here are described by the `*-receipt.json`
file next to them.

Inserted media come from other repository fixtures:

- `ringin.wav`: [`test-data/poi/test-data/slideshow/ringin.wav`](../poi/test-data/slideshow/ringin.wav),
  from Apache POI (`test-data/slideshow/ringin.wav`, Apache-2.0; see
  [`../poi/PROVENANCE.md`](../poi/PROVENANCE.md)).
- `abstract1.jpg`: [`test-data/images/jpg/abstract1.jpg`](../images/jpg/abstract1.jpg).

## Basic packages

| File | Expected semantic content |
| --- | --- |
| `pages/basic.pages` | `Litchi native Pages fixture`, `Buffa lazy-view migration verification`, `2026-08-07` |
| `numbers/basic.numbers` | Cell `B2`: `Litchi native Numbers fixture`; cell `B3`: number `42` |
| `keynote/basic.key` | One slide containing `Litchi native Keynote fixture`, `Buffa lazy-view migration verification`, and `2026-08-07` |

## Package directories

`directory/` holds the three basic packages converted in their applications
with **File → Advanced → Change File Type → Package**, saved, reopened from
Recents, and closed without another save. `directory/MANIFEST.sha256` lists
every regular member; there are no symbolic links or special nodes. The
directory reader freezes only `Index.zip` or loose `Index/` semantics; these
fixtures are not exact-ZIP sources for editing.

## Pages

- `pages/body-table-visible.pages`: Blank document with body marker
  `Pages hidden-axis native oracle — 2026-09-05` and one Plain 5 × 4 table
  with no hidden rows or columns. Its native hidden-state envelope has one
  state with empty row and column lists, so it reads as no hidden axes; the
  table-info version is `[1, 0, 5]` and the table-model and formula-owner
  versions are `[3, 2, 10]`.
- `pages/body-sections-unicode.pages`: Blank document whose first section
  contains `Pages borrowed body 😀`, a paragraph break, and
  `First section marker.`; **Insert → Section Break** added a second section
  containing `Second section marker — end.`.
- `pages/audio-playback-native.pages`: body `Native Pages playback marker` and
  the inserted `ringin.wav`, loop `None`, volume `1`.
  `pages/audio-playback-native-resaved.pages` is the focused change to `Repeat`
  (shown as `Loop`) and volume `0.5`, with the body edited to
  `Saved Native Pages playback marker` before the native save.
- `pages/image-adjustments-native.pages`: body
  `Native Pages image adjustment marker` with the inserted `abstract1.jpg`;
  exposure and saturation absent, Enhance explicitly disabled, sharpness 25%.
  Focused native-saved variants:
  - `image-adjustments-native-resaved.pages`: exposure `0.25`, saturation
    `-0.2`; body `Saved Native Pages image adjustment marker` plus the image
    anchor.
  - `image-adjustments-direct-resaved.pages`: exposure `0.5`, saturation
    `-0.4` through the body-image transaction; body
    `Direct saved Native Pages image adjustment marker` followed by U+FFFC and
    one 6.5 × 4.9 inch inline image.
  - `image-adjustments-retirement-resaved.pages`: exposure `-0.2`, saturation
    `+0.35`, Enhance enabled; the image description ends with a literal tab.
- `pages/chart-arrangement-native.pages`: Blank document with body marker
  `Pages Arrange native body marker` followed by a text-anchored 2D Column
  chart `Pages Arrange native chart` (rows `Region 1`/`Region 2`, columns
  April–July, values `[[17, 26, 53, 96], [55, 43, 70, 58]]`), `Move with Text`
  on, Lock and Constrain proportions off.
  `pages/chart-arrangement-retirement-resaved.pages` is the focused change of
  both flags to true, with `Native saved: ` inserted into the body before the
  native save.

## Numbers

- `numbers/custom-number-native.numbers`: Blank document; B2 holds `42` with
  the Custom Number format `Native Grouped` (`#,###`), C2 holds
  `Native Custom marker`. `custom-number-native-resaved.numbers` is the
  focused replacement with `Rust Grouped` (`#,##0.00`, shown as `42.00`), C2
  edited to `Native Custom marker saved` before the native save.
- `numbers/custom-text-native.numbers`: B2 holds `Orchid` with the Custom Text
  format `Native Label` (prefix `Native [`, suffix `]`), C2 holds
  `Native Text marker`. `custom-text-native-unicode.numbers` renames the
  format in Numbers to `Native Unicode` and puts an emoji before the prefix;
  the cached `index_from_right_last_integer` is `9` and `11` respectively, the
  final UTF-16 index of each pattern. `custom-text-native-resaved.numbers` is
  the focused replacement with `Rust Label` (prefix `Rust <`, suffix `>`), C2
  edited to `Native Text marker saved`.
- `numbers/custom-datetime-native.numbers`: B2 holds 2024-01-02 15:04:05 with
  the Custom Date & Time format `Native Calendar` (`MMM d, y`), C2 holds
  `Native DateTime marker`. `custom-datetime-native-resaved.numbers` is the
  focused replacement with `Rust Calendar` (`yyyy-MM-dd HH:mm:ss`), C2 edited
  to `Native DateTime marker saved`.
- `numbers/custom-retirement-resaved.numbers`: focused change of
  `custom-number-native` B2 to the Custom format `Focused Custom Retirement`
  (`#,##0.000`, shown as `42.000`), C2 edited to
  `Focused Custom retirement saved`. `custom-retirement-cleared.numbers` is
  the focused reset to Automatic, C2 edited to `Focused Custom clear saved`.
- `numbers/duration-native.numbers`: B2 holds `1h 23m 45s` with an explicit
  abbreviated Duration format and automatic hours-through-seconds units, C2
  holds `Native Duration marker`. `duration-native-resaved.numbers` is the
  focused change to colon style with custom hours-through-seconds units
  (shown as `1:23:45`), C2 edited to `Native Duration marker saved`.
- `numbers/audio-playback-native.numbers`: B2 holds
  `Native playback marker` and the sheet holds the inserted `ringin.wav`, loop
  `None`, volume `1`. `audio-playback-native-resaved.numbers` is the focused
  change to `Repeat` and volume `0.5`, B2 edited to
  `Native playback marker saved`.
- `numbers/image-adjustments-native.numbers`: B2 holds
  `Native image adjustment marker` and the sheet holds the inserted
  `abstract1.jpg`, with the same adjustment state as the Pages source.
  Focused native-saved variants: `image-adjustments-native-resaved.numbers`
  (exposure `0.25`, saturation `-0.2`; B2 extended in the native UI),
  `image-adjustments-direct-resaved.numbers` (exposure `0.5`, saturation
  `-0.4` through the sheet-image transaction; B2
  `Native image adjustment marker direct saved`), and
  `image-adjustments-retirement-resaved.numbers` (exposure `-0.2`, saturation
  `+0.35`, Enhance enabled; description
  `Focused Numbers image adjustment retirement saved`).
- `numbers/table-relocation-native.numbers`: `Sheet 1/Table 1` is 22 × 7 with
  `Native image adjustment marker` in B2 and a 600 × 450 image at (57.6, 8.6);
  `Sheet 2/Destination table` is 10 × 5 with `Destination stays` in B2 at
  (800, 28). `table-relocation-native-resaved.numbers` is the focused move of
  `Table 1` to `Sheet 2`, B2 edited to `Native relocation saved marker`; the
  moved table sits at (0, 28.3) and Sheet 1 keeps only the image.
- `numbers/chart-arrangement-native.numbers`: title
  `Numbers Arrange native chart`; `Sheet 1/Table 1` is 3 × 5 with rows
  `North`/`South`, columns April–July (auto-parsed dates), and values
  `[[17, 26, 53, 96], [55, 43, 70, 58]]`; both Arrange flags off and no data
  assets. `chart-arrangement-retirement-resaved.numbers` is the focused change
  of both flags to true, re-locked through native Unlock/Lock before the save.

## Keynote

- `keynote/table-discovery.key`: copy of `basic.key` with a Plain 5 × 4 table
  holding `Buffa discovery` in A1.
- `keynote/image-adjustments-native.key`: title
  `Native image adjustment marker` and the inserted `abstract1.jpg`, with the
  same adjustment state as the Pages source.
  `image-adjustments-native-resaved.key` is the focused exposure `0.25`,
  saturation `-0.2` change with the title edited to
  `Native image adjustment marker saved`;
  `image-adjustments-retirement-resaved.key` sets exposure `-0.2`, saturation
  `+0.35`, and Enhance enabled.
- `keynote/slide-layout-native-resaved.key`: focused layout change from
  `Title` to `Title Only` on the image-adjustment slide, title edited to
  `Saved layout Native image adjustment marker`; the 1024 × 768 image stays
  centred 14.4% from the top. It guards stale metadata slide dependencies,
  which made Keynote crash on this layout change.
- `keynote/chart-caption-native.key`: chart with rows `Region 1`/`Region 2`,
  columns April–July, and values `[[17, 26, 53, 96], [55, 43, 70, 58]]`.
  `chart-caption-native-resaved.key` is the focused caption
  `Focused chart caption — 北区`, edited natively to
  `Native saved caption — 北区`.
- `keynote/chart-titles-native.key`: the same chart with title
  `Native chart title — 北区`, value-axis title `Native revenue — 元`,
  category-axis title `Native months — 月`, and caption
  `Native chart caption marker`. `chart-titles-native-resaved.key` sets
  focused titles (`Focused chart title — 北区`, `Focused revenue — 元`,
  `Focused months — 月`), with the chart title edited natively to
  `Native saved chart title — 北区`.
- `keynote/chart-arrangement-retirement-resaved.key`: focused change of the
  chart lock and `Constrain proportions` flags to true on the chart-titles
  deck, re-locked through native unlock/relock before the save.

### Slide media

- `keynote/media-replacement-native.key`: one slide titled
  `Keynote media replacement native` with body
  `Media ownership marker — 北区` and media in source order Audio, Audio,
  File, File. The two audio controls share one WAV record; the two file movies
  share one locally authored MJPEG MOV record and one poster record, both
  captioned `Shared native movie caption`, 320 × 180 at (200, 700) and
  (700, 700), with a 0–2 second trim, volume 1, no loop, start on click, and
  play across slides.
- `keynote/media-replacement-assets/`: locally authored replacement inputs —
  `striped.mov` (MJPEG from CoreGraphics/ImageIO JPEG frames and QuickTime
  atoms), `tone.wav` (PCM), and `poster.png`. No vendor media is included.
- `keynote/media-replacement-retirement-resaved.key`: focused replacement of
  audio 0 with `tone.wav`, file movie 2 with `striped.mov`, and that movie's
  poster with `poster.png`; body edited to
  `Saved media ownership marker — 北区`.
- Native lifecycle oracles authored from `media-replacement-native.key`:
  `media-lifecycle-duplicate-native.key` (a movie duplicated, offset
  (10, 10), with an appended caption; five media),
  `media-lifecycle-remove-native.key` (final movie removed; two audio objects
  remain and the movie and poster data are culled),
  `media-lifecycle-audio-duplicate-native.key` (three audio, two captioned
  movies), and `media-lifecycle-audio-remove-native.key` (no audio; the WAV
  record is removed and the movie and poster records remain).
- `keynote/media-comments-baseline-native.key`: the media deck with a comment
  on one movie. `media-comments-duplicate-native.key` duplicates that movie
  with Cmd-D (three movies, two audio, comments on both copies; storage UUID
  kept, author records shared).
- `keynote/media-comments-shared-removal-native.key` (2 movies, 2 audio,
  1 comment; the duplicate removed) and
  `media-comments-final-removal-native.key` (1 movie, 2 audio, no comments;
  the original removed): focused removals of a commented movie. Author and
  author-storage records stay; only the final component external edge is
  removed.
- `keynote/media-watermark-removal-native.key` (2 movies, 2 audio,
  1 comment after removal) and `media-watermark-reduplicate-native.key`
  (3 movies, 2 audio, 2 comments after duplicating again): focused removal
  that lowers `last_object_identifier`, and the next duplicate allocating
  above the metadata maximum.
- `keynote/media-audio-comment-duplicate-native.key` (3 audio, 2 movies) and
  `media-audio-comment-removal-native.key` (2 audio, 2 movies): focused
  duplication and removal of an audio object that received a root comment
  and reply programmatically; the existing movie comment is kept.
- `keynote/media-audio-position-ui-native.key`: Audio A moved natively in
  Arrange from (960.0, 539.4513549804688) to (1120.0, 420.0).
  `media-audio-position-focused-native.key` is the focused move of Audio A to
  (123.25, 456.5).
- Focused media properties on the comments baseline:
  `media-properties-audio-focused-native.key` (Audio A locked, aspect locked,
  linked to `https://example.test/audio-a`, description
  “Audio A — accessible 北区”) and `media-properties-file-focused-native.key`
  (file movie 2, `https://example.test/file-a`, “File A — accessible 北区”).
- `keynote/slide-audio-creation-focused-native.key`: focused creation of an
  audio control from a 1,644-byte mono 8 kHz 16-bit PCM WAV (800 samples,
  each `(sample % 50 - 25) * 400`), 0.100-second trim, start on click, at
  (120.5, 240.25); five media and five builds.
- `keynote/slide-movie-creation-source-native.key`: native reference that
  inserts the baseline's movie through Media > Choose (320 × 180, two-second
  trim, start on click); Keynote reuses the content and poster records and
  adds a movie-start build with a 0.5-second animation.
  Focused creations: `slide-movie-creation-focused-native.key` (reused video
  and poster, 320 × 180 at (120.5, 240.25), two-second trim) and
  `slide-movie-creation-fresh-focused-native.key` (new video and poster — the
  baseline movie with an extra ISO-BMFF `free` atom and the poster with an
  extra PNG `tEXt` chunk — 640 × 360 at (321, 42), 1.25-second trim).
- `keynote/slide-movie-geometry-source-native.key`: movie 4 of the fresh
  creation fixture rotated 27.5° and reflected horizontally in Arrange, at
  320 × 180. Focused geometry: `slide-movie-geometry-focused-native.key`
  ((450.5, 120.25), 320 × 180, 27.5°, horizontal reflection) and
  `slide-movie-geometry-vertical-focused-native.key` ((420, 200), 480 × 270,
  207.5°, no reflection).
- `keynote/media-properties-placeholder-native.key`: movie 4 of the fresh
  creation fixture turned into a media placeholder with **Format → Advanced →
  Define as Media Placeholder** and described as
  “Native movie placeholder — accessible 北区”; its content is
  `Data/fresh-content-9094.mov` and its poster `Data/fresh-content-9095.png`.
- `keynote/drawable-comments-source-native.key`: the media comments deck
  plus a 100 × 100 square commented “Native square comment — 北区”. The
  focused drawable-comment fixtures derived from it are described in
  [`keynote/drawable-comments-cross-component-native/README.md`](keynote/drawable-comments-cross-component-native/README.md).

## Litchi-built fixtures

[`keynote/multi-slide.key`](keynote/multi-slide.key) is a five-slide deck
built from `keynote/basic.key` with the public slide-creation, slide-text,
notes, and skip transactions; its steps and content are in
[`keynote/multi-slide.md`](keynote/multi-slide.md).
