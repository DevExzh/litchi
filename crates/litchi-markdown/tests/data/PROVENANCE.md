# Markdown test-data provenance

The release gate is entirely offline. Corpus counts and hashes are asserted by
`tests/normative.rs`; updates require an intentional file and provenance change.

## CommonMark 0.31.2

- Upstream: <https://spec.commonmark.org/0.31.2/spec.json>
- Version/date: CommonMark 0.31.2, published 2024-01-28.
- Vendored file: `tests/corpus/commonmark-0.31.2/spec.json`.
- Examples: all 652 normative examples.
- SHA-256: `d431b29d97b6f73e69d547109cf5081578fac931e72afe95639ebe766c1b2a20`.
- License: Creative Commons Attribution-ShareAlike 4.0 International,
  reproduced as `LICENSE-CC-BY-SA-4.0.txt` beside the corpus, SHA-256
  `28a9529c7d0bb4dc51f4bf5c116a3d16ef247a052f7591466768ddf563fd1cf5`.
- Attribution: “CommonMark Spec,” John MacFarlane and contributors.

Reproduction command:

```sh
curl -fsSL https://spec.commonmark.org/0.31.2/spec.json \
  > tests/corpus/commonmark-0.31.2/spec.json
```

## GitHub Flavored Markdown

- Upstream: <https://github.com/github/cmark-gfm>.
- Pin: tag `0.29.0.gfm.13`, peeled commit
  `587a12bb54d95ac37241377e6ddc93ea0e45439b`.
- Source: `test/spec.txt`, SHA-256
  `7d8e5814befec287ac116786d81ff14e0adc9b13295b4494649e995408fd871c`.
- Generator: `test/spec_tests.py`, SHA-256
  `8f2c8ad3f819922b2cf95557b9af91a3c8513b6e7e8a8461a6fb632d59be616e`.
- Vendored file: `tests/corpus/gfm-0.29.0.gfm.13/spec.json`.
- Examples: all 670 extracted examples, including all 22 examples explicitly
  marked with the table, strikethrough, autolink, or tag-filter extensions.
- SHA-256: `89cfcb21173de246f141ef6832395b74d45a23595ddf65bf6ffb0334d3e7c651`.
- Specification license: Creative Commons Attribution-ShareAlike 4.0
  International, reproduced beside the corpus, SHA-256
  `28a9529c7d0bb4dc51f4bf5c116a3d16ef247a052f7591466768ddf563fd1cf5`.
- Extractor/repository license: BSD 2-Clause; the pinned upstream `COPYING` is
  reproduced as `COPYING-BSD-2-Clause.txt` beside the corpus, SHA-256
  `c22e885f33b821bddb24cf007145e5540655b6c0f403e49e6c76a93c28e6d9a9`.
- Attribution: GitHub's `cmark-gfm` contributors; the GFM specification is
  based on the CommonMark specification by John MacFarlane.

All 670 examples gate exact-source parsing, range integrity, deterministic
reparse, and reversible editing. Expected HTML is asserted for the eight table
and two strikethrough examples supported by the current parser dependency. The
eleven extended-autolink examples and one tag-filter example remain in the
parse/edit gate but are not claimed as semantic rendering conformance because
`pulldown-cmark 0.13.4` does not implement those extensions. The remaining 648
examples describe the older CommonMark 0.29 baseline, so their rendering is
superseded by the complete CommonMark 0.31.2 gate rather than used to require
obsolete parsing behavior.

Reproduction commands:

```sh
git clone --depth 1 --branch 0.29.0.gfm.13 \
  https://github.com/github/cmark-gfm.git cmark-gfm
python3 cmark-gfm/test/spec_tests.py \
  --spec cmark-gfm/test/spec.txt --dump-tests \
  > tests/corpus/gfm-0.29.0.gfm.13/spec.json
```

The `roundtrip-*.md` documents are project-authored integration fixtures under
the repository's Apache-2.0 license.
