# `multi-slide.key` provenance

A five-slide Keynote package **built by Litchi** (record 0298) from
[`basic.key`](basic.key) (SHA-256
`3a3d07476b45b6e543bcfba75fe38a245434176dcb3565e34570b817708b9f42`, the
repository's native single-slide fixture; see `../README.md`). No upstream
project or third-party content is involved; the license is the repository's.

- File: `multi-slide.key`, 463,067 bytes.
- SHA-256: `9f33d0ff0110fcb173fac2fe666eedcf830f5620fe1d7130159d35ac36aca616`.
- Generator: `build_multi_slide()` in
  `crates/litchi-keynote/tests/slide_creation.rs`, public transactions only.
  Regenerate with
  `LITCHI_BLESS=1 cargo test -p litchi-keynote --test slide_creation multi_slide_fixture_is_reproducible`.
  The test `multi_slide_fixture_is_reproducible` rebuilds the package and
  compares its semantic content (slide count, titles, bodies, notes, skipped
  flags) with the committed file; the build is deterministic, so a rebuild
  currently reproduces the SHA-256 above byte for byte.

Generating steps:

1. `edit_slide_creation`: add slides from the layouts `Title & Bullets`,
   `Section` and `Title Only` at positions 1, 2 and 3 (one commit).
2. `edit_slide_text`: set the titles of slides 0–3 and the body of slide 1.
3. `edit_slide_creation`: duplicate slide 1 to position 4; then set its title
   and body.
4. `edit_slide_notes`: set speaker notes on slides 0 and 1.
5. `edit().skip_slide`: hide slide 2.

Content:

| Slide | Layout | Title | Body | Notes | Hidden |
| --- | --- | --- | --- | --- | --- |
| 0 | `Title` (source slide) | Litchi multi-slide fixture | Buffa lazy-view migration verification | Welcome everyone. | no |
| 1 | `Title & Bullets` | Agenda | First topic / Second topic / Third topic | Walk through the agenda. | no |
| 2 | `Section` | Hidden section | (empty) | (empty) | yes |
| 3 | `Title Only` | Closing | (empty) | (empty) | no |
| 4 | duplicate of slide 1 (made before its notes were set) | Agenda (continued) | Fourth topic / Fifth topic | (empty) | no |

Structural edits delete the root previews (`preview*.jpg`), so the package
has none.
