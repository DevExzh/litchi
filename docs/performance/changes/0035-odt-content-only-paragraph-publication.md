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

## Matched media-rich measurement

The opt-in `odt_media_paragraph_edit_save` case fixes a 16,786,287-byte ODT
with 200 paragraphs and eight deterministic incompressible 2 MiB `Pictures/`
members. It replaces paragraph 100, materializes the candidate in the timed
region, then performs untimed full paragraph, media, manifest, deterministic
output, patch replay, inverse, stale-source, and raw unchanged-member checks.
The corpus SHA-256 is
`097b17ebc8fe811888eba6a9a7d118bc94bb99651880782ccecedf93c003b5ed`.

Matched ABBA release runs used CPU 2, 10 warmups, and 100 samples per leg on
the same host. The pooled distributions contain 200 samples per state.

| State | p50 | mean | p95 | p99 |
| --- | ---: | ---: | ---: | ---: |
| Before | 249.177 ms | 250.159 ms | 258.311 ms | 276.830 ms |
| After | 11.001 ms | 10.937 ms | 11.811 ms | 12.663 ms |

The p50 delta is **-95.58%**, mean **-95.63%**, p95 **-95.43%**, and p99
**-95.43%**. Heaptrack allocation calls fall from 20,458 to 19,085
(-6.71%) and temporary allocations from 3,855 to 3,344 (-13.26%); its
106.03 MB peak heap is flat. Independent GNU Time ABBA maximum RSS moves from
111,480 KiB in both before legs to 110,764/110,888 KiB after (-0.59% by
pooled mean).

Whole-process counter ABBA, which includes untimed corpus construction and
verification, moves cycles -68.64%, instructions -76.03%, branches -80.77%,
branch misses -90.50%, cache references -65.45%, and cache misses -20.17%.
The target raw samples are [`before A`](../results/abba-odt-media-paragraph-before-a.json),
[`after A`](../results/abba-odt-media-paragraph-after-a.json),
[`after B`](../results/abba-odt-media-paragraph-after-b.json), and
[`before B`](../results/abba-odt-media-paragraph-before-b.json). Guard,
heap, RSS, counter, and binary hashes are retained beside them and indexed by
[`odt-media-paragraph-sha256.txt`](../results/odt-media-paragraph-sha256.txt).
The frozen before/after harness binaries hash to `b7a64cd9...70053` and
`1e4807e3...86db3`, respectively.

The ordinary large ODT guard also improves: open p50/mean -1.96%/-2.53%,
no-op edit/save -13.25%/-15.89%, and one-edit/save -22.37%/-22.65%.

## Preservation and correctness gates

The integration regression builds an ODT with a 1 MiB opaque media member,
replaces one paragraph through the public transaction API, and proves:

- `content.xml` changes while `mimetype`, `styles.xml`, `meta.xml`,
  `META-INF/manifest.xml`, and the media member retain raw local/central ZIP
  records (apart from the required relocated local-header offset field);
- the media payload and semantic paragraph read back exactly;
- forward patch replay reproduces the candidate; and
- the inverse restores the complete source artifact byte-for-byte.

The common raw publisher intentionally limits regenerated `content.xml` to
16 MiB. A second regression builds valid semantic text whose XML exceeds that
limit and proves paragraph replacement returns to the established bounded ODT
package rebuild instead of exposing the narrower optimization limit to the
public transaction API.

Executed gates:

- all 24 `packaged_transactions` integration tests passed;
- all 27 standalone harness tests passed;
- complete all-feature ODT and ODF common test suites passed;
- `cargo clippy -p litchi-odt --lib -- -D warnings` passed; and
- ODF common and harness warning-denied Clippy plus formatting and diff-hygiene
  gates passed.

The workspace boundary checker still reports unrelated pre-existing
unclassified/dev-only edges, including `litchi-odc`/`litchi-odg`/`litchi-pptx`
edges. This tranche adds no dependency edge and does not claim that gate as
clean.

## Remaining work

- Consider other content-only operations only after each operation proves its
  dependency closure; do not infer that insertion/removal or mixed edits are
  eligible from this paragraph-replacement result.
- Keep real-producer fallback coverage: a regenerated `content.xml` that cannot
  satisfy the checked one-splice publication proof must rebuild rather than
  weakening preservation validation.
