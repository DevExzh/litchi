# Change 0078: source-backed XLSX sheet-protection publication

Date: 2026-08-12

Production base: `242dd8f79277f7a3f2e2e6bdd175d56809e55cf9`

Status: accepted

## Hypothesis and implementation

Worksheet protection already had a complete typed model and an exact XML
rewriter for direct `sheetProtection`, core `protectedRanges`, and Office 2010
`x14:protectedRanges`. Publishing that small semantic change still required an
eager OPC conversion, however, which inflated and recompressed every media
Part. The existing accepted XLSX worksheet-source closure and one-Part OPC
publisher can avoid that work without widening protection semantics.

`litchi_xlsx::sheet_protection::SourceBackedEditor` now owns one immutable
positional source. Its snapshot resolves one existing normal worksheet and
binds all state that can alter the interpretation or ownership of that Part:

- exact workbook URI, content type and XML;
- the unique package `officeDocument` relationship;
- selected workbook-sheet name, position and relationship;
- exact worksheet URI, content type and XML; and
- the complete sorted outbound worksheet relationship set.

The isolated editor stages one complete `Metadata` value. `set` validates the
whole value before adoption, `update` edits a clone atomically, and `clear`
removes both sheet protection and protected ranges. Commit applies the existing
byte-preserving protection rewriter, reparses the complete result, and exposes
an exact reversible source-specific patch. Patch application and source-backed
publication recapture the complete retained closure before changing the one
worksheet payload.

The capability does not verify passwords, change worksheet relationships,
select chartsheets, add/remove Parts, edit cells/formulas/styles, or change
workbook topology. Protection selected through MCE cannot be changed
byte-exactly and is refused. Stale/foreign source state, relationship
retargeting, source-version changes, changed signed sources, invalid metadata,
OPC limits, unsupported physical layouts and partial sinks retain typed errors.
An exact no-op, including on a signed input, reproduces the complete source
artifact. No dependency, runtime, unsafe code, global cache, or iWork/IWA code
was added.

## Matched corpus and protocol

Both controls use the fixed XLSX media corpus with one workbook, one ordinary
worksheet, one DrawingML drawing and eight referenced deterministic
incompressible 2 MiB PNG Parts. It has 12 ordinary Parts, 17 ZIP members,
16,782,412 logical Part bytes and a 16,786,830-byte archive with SHA-256
`c11a9424accfc6ce56e4deb6ecb18a2142d2f0076395018ef00ba93897049f7c`.

Both paths create the same complete protection state on `Sheet1`: a legacy
sheet verifier and operation locks, one core protected range with a legacy
verifier, and one Office 2010 protected range with a strong SHA-512 verifier
descriptor. The eager control converts the positional package into complete
OPC ownership, applies the same typed `replace_protection` codec, and publishes
through the ordinary writer. The candidate performs one guarded source-backed
transaction and consumes the one-Part overlay publisher.

Both emit the same 16,787,247-byte artifact with SHA-256
`76b3a48e9682479a7dc75b047e37f221c52dc6953c10168e20995adbf992b8f3`.
Complete typed XLSX reopen, calculation-metadata checks, package topology,
relationships, content types, all untouched Part payloads, eight media
payloads, output hashing, source counters and sink bounds remain outside each
timed interval.

The eager control was frozen before the production source-backed modules were
added. Its binary SHA-256 is
`d4d88556fddf6225944c2313113f04c62a5fd34e3eb2ded36a229ab80f5d1e12`;
the measured candidate binary SHA-256 is
`ff053c52889f79172a240c72aa16dd9e58bf5a21ec7cf5b464bec89c24c9a75a`.
The final rebuilt binary, after documentation-only source comments and focused
test additions, is
`98e9b1dad757a477388e301937719dd39db0ada8d182851b40409c0bafc3342d`.

The formal retained ABBA pair on CPU 2 ran eager A, source-backed A,
source-backed B, eager B. Each process used ten warmups and 100 samples, for
200 samples per state. Two additional direction-balanced pairs used the same
protocol; the candidate won all six legs. The formal raw reports are
[`before A`](../results/abba-xlsx-sheet-protection-before-a.json),
[`after A`](../results/abba-xlsx-sheet-protection-after-a.json),
[`after B`](../results/abba-xlsx-sheet-protection-after-b.json), and
[`before B`](../results/abba-xlsx-sheet-protection-before-b.json). Aggregated
counters and supplementary-leg evidence are in the
[`measurement summary`](../results/xlsx-sheet-protection-publication-summary.json).

## Results

| Metric | Eager control | Source-backed | Delta |
|---|---:|---:|---:|
| pooled samples | 200 | 200 | — |
| p50 | 221.877 ms | 4.982 ms | **-97.75% (44.54x)** |
| mean | 222.778 ms | 5.010 ms | **-97.75% (44.46x)** |
| p95 | 229.135 ms | 5.324 ms | **-97.68% (43.04x)** |
| p99 | 234.878 ms | 5.510 ms | **-97.65% (42.63x)** |
| semantic Part materializations | 12 | 2 | -83.33% |
| output bytes | 16,787,247 | 16,787,247 | exact |
| sequential writes | 630 | 547 | -13.17% |
| largest write | 32,768 B | 32,768 B | bounded |

Across all three balanced pairs, mean leg p50 is 221.597 ms eager versus
4.392 ms source-backed (-98.02%). The candidate reads the workbook and selected
worksheet semantic payloads; the other ten Parts remain compressed and are
copied into the sequential output.

## Allocation, counters and memory

One-sample Heaptrack attribution covers the whole process, including fixed
corpus construction and untimed complete verification. Allocation calls are
13,524 eager versus 13,893 source-backed (+2.73%, within the 5% regression
ceiling); peak heap is 152.84 versus 152.81 MiB (-0.02%). Uninstrumented GNU
Time maximum RSS is 141,528 versus 141,980 KiB (+0.32%, flat). The retained
corpus and complete output dominate both peak measurements.

Five `perf stat` repeats per state used two warmups and ten samples:

| Counter | Eager | Source-backed | Delta |
|---|---:|---:|---:|
| cycles | 19.435 billion | 5.518 billion | -71.61% |
| instructions | 48.096 billion | 10.641 billion | -77.87% |
| branches | 8.155 billion | 1.384 billion | -83.03% |
| branch misses | 176.618 million | 14.483 million | -91.80% |
| cache references | 1.367 billion | 403.860 million | -70.46% |
| cache misses | 29.160 million | 20.571 million | -29.46% |

The latency, materialization and instruction reductions clear the acceptance
threshold. Allocation calls and RSS remain below the 5% regression ceiling.

## Correctness and regression closure

Focused integration tests cover add, replace, clear and exact no-op behavior;
typed core range metadata; patch replay and inverse; exact unselected payload,
content-type and relationship preservation; stale, foreign, source-version and
retargeted workbook relationships; changed outbound worksheet relationships;
signed changes and signed no-ops; MCE-selected protection; chartsheet refusal;
OPC read limits; and partial-sink failure. Existing codec tests retain strict
and transitional namespaces, strong/legacy verifier validation, protected
range uniqueness, limits and byte-preserving unrelated XML behavior.

The complete `litchi-xlsx` and performance-harness suites plus library/focused
Clippy with warnings denied are release gates. XLSX rustdoc passes with the
repository's pre-existing private/broken-link lint allowances; an unqualified
denied-warning run still reports those unrelated existing links. The
ODF-common GenericArray deprecation fix is separately revalidated under fully
denied Clippy and rustdoc warnings in this batch; this change introduces no
deprecated use.

## Alternatives retained

The audited RTF follow-up reuses sparse paragraph selection for a measured
approximately 5.9% one-percent-edit win. The strongest ODF follow-up is an ODT
count-only paragraph scan with an approximately 48% query proxy. OLE2's next
candidate needs stage attribution before implementation. All remain queued;
none has the breadth or absolute media-rich save reduction of this bounded
XLSX transaction. General XLSX cell/formula edits still require a wider
calculation and topology closure and were not approximated by this work.
