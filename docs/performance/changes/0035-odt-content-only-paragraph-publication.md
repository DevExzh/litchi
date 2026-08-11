# Change 0035: ODT content-only paragraph publication

## Scope

`Operation::ReplaceParagraph` now publishes the regenerated `content.xml`
through the accepted common ODF content-replacement primitive. All other ODT
operations, including structural, metadata, style, image, embedded-resource,
and mixed transactions, retain the established package rebuild path.

The optimization removes inflation, copying, and recompression of eligible
unchanged package members. It does not skip the compact-XML audit, package-size
check, immutable snapshot construction, complete ODT reopen, semantic readback,
source check, or reversible patch behavior. Unsupported ZIP layouts,
signatures, encryption, and size-bearing manifest entries keep the common
publisher's existing safe fallback or refusal behavior.

## Directional measurement

The unchanged release harness case `odt_semantic_one_edit_save`, shape `large`
(10,000 generated paragraphs), was run with 10 warmups and 50 samples per
state on the same working tree and host. The pre-change binary was frozen
before rebuilding the post-change binary.

| State | p50 | mean | p95 | p99 |
| --- | ---: | ---: | ---: | ---: |
| Before | 19.506 ms | 19.792 ms | 21.419 ms | 22.753 ms |
| After | 15.003 ms | 15.115 ms | 16.173 ms | 18.605 ms |

The p50 delta is -23.08% and the mean delta is -23.63%. This is directional
single-order evidence, not an accepted ABBA performance claim. A media-rich
ODT corpus, matched reverse-order runs, allocation/RSS profiles, and retained
raw JSON remain follow-up evidence.

## Preservation and correctness gates

The new integration regression builds an ODT with a 1 MiB opaque media member,
replaces one paragraph through the public transaction API, and proves:

- `content.xml` changes while `mimetype`, `styles.xml`, `meta.xml`,
  `META-INF/manifest.xml`, and the media member retain raw local/central ZIP
  records (apart from the required relocated local-header offset field);
- the media payload and semantic paragraph read back exactly;
- forward patch replay reproduces the candidate; and
- the inverse restores the complete source artifact byte-for-byte.

Executed gates:

- all 23 `packaged_transactions` integration tests passed;
- `cargo clippy -p litchi-odt --lib -- -D warnings` passed; and
- package-scoped formatting passed.

## Remaining work

- Add an opt-in deterministic ODT corpus with multiple large incompressible
  `Pictures/` members and retain ABBA, allocation, peak-heap, RSS, and counter
  evidence.
- Consider other content-only operations only after each operation proves its
  dependency closure; do not infer that insertion/removal or mixed edits are
  eligible from this paragraph-replacement result.
- Keep real-producer fallback coverage: a regenerated `content.xml` that cannot
  satisfy the checked one-splice publication proof must rebuild rather than
  weakening preservation validation.
