# 0712: DOCX active-offset validation and search attribution

`performance_claim: none`. This batch rules out a no-anchor validation shortcut
and identifies a bounded search experiment. Production is unchanged from 0711's
restored baseline, including the 0710 custom-properties preservation fix.

The writer strips a leading BOM, calls `alt::scan`, collects outer paragraph,
table and altChunk ranges with `active_block_ranges`, and then preserves the
original body slices. The alt scanner selects active anchor offsets itself.
The range helper separately selects active offsets for its complete outer
target set. Paragraph and table targets matter even though the writer consumes
only the helper's altChunk ranges.

## Current-source counterexamples

The source-bound release probe runs 11 deterministic synthetic documents through
public `alt::scan`, `alt::active`, and `Package::document_mut`. Its complete
report reproduces byte for byte in a second process.

| Case | Anchor scan | Empty/alt-only selection | Public mutable acquisition |
| --- | --- | --- | --- |
| Malformed AlternateContent containing non-branch children, no anchors | Empty success | Empty success | `Mce(NonConformant("non-ignorable AlternateContent child"))` |
| Unknown `MustUnderstand`, no anchors | Empty success | Empty success | `Mce(MustUnderstand("urn:unsupported"))` |
| Anchor inside paragraph, plus direct anchor | Two anchors | Two selected anchors | One preserved anchor |
| Anchor inside table, plus direct anchor | Two anchors | Two selected anchors | One preserved anchor |

The first two refusals occur at `document_mut`, after package construction
succeeds. Empty anchor metadata therefore cannot prove that active-block
selection is dispensable. The nested cases show why replacing outer target
ranges with all anchor offsets changes writer behavior.

Controls cover transitional and strict namespaces, supported Choice and
unsupported Choice/Fallback, and ordinary documents without anchors. A separate
call supplies 1,000,001 offsets and observes `active-offset count` refusal. That
limit case is a public `active` test: its small document still opens for mutable
use and does not demonstrate a million-node facade limit.

The probe's `active_full` inputs are lexical starts in fixed synthetic fixtures,
including nested starts. They are not an independent implementation of the
production range scanner. The public facade supplies the nested-suppression
observation. This probe does not certify arbitrary XML, package preservation,
native Office compatibility, or allocation-failure behavior.

## Historical instruction attribution

The analyzer binds four retained 0709 measured edit profiles, their exact raw
edges, source manifests, receipts and binary identities. Source comparison
confirms the edit path remains unchanged. These are historical Callgrind guest
instructions, not new native timings or predicted savings. Values below sum two
profiles of each corpus.

| Boundary | Generated medium Ir | NumberedList Ir |
| --- | ---: | ---: |
| `active_block_ranges`, inclusive | 3,449,750 | 2,663,144 |
| Its selected `scan_word_element_ranges` edge | 2,294,438 | 415,959 |
| Its selected `alt::active` edge | 1,068,977 | 2,245,011 |
| `active_offsets`, all incoming calls | 1,033,124 | 2,208,408 |
| Its `process_markup_compatibility` edge | Not reached | 1,704,307 |
| Its direct `find_bytes` edge | No separate edge | 237,438 |
| Its direct `memcmp` edges | 600,948 | 131,186 |

Nested inclusive rows must not be added. The shared range scanner also serves
paragraph-section validation, so its full function partition includes those
other callers; the selected edge above does not. The analyzer checks disjoint
immediate-child accounting at each owner.

Absence of a generated `find_bytes` symbol does not imply absence of search
work: the namespace-presence search is inlined there, with 42,918 direct
`memcmp` invocations across the two profiles. No MCE marked document is produced
on that path. NumberedList reaches MCE processing, which consumes 77.17% of the
active-offset owner. Its start-event handling remains the largest nested child.
Allocator/copy symbol attribution is diagnostic; it is not native allocation
request counts or copied-byte measurement.

## Next experiment and constraints

First measure replacing the existing scalar `find_bytes` window search with
the already available `memchr::memmem::find`. It must preserve exact substring
semantics, empty needles, overlaps, marker collision selection, all validation
calls, offset vectors, limits and error order. This is a smaller experiment
against work observed on both corpora; no speedup is established here.

A larger writer-side fusion of anchor parsing and range scanning remains a
separate opportunity. Its first version should retain both existing MCE calls
and their exact offset vectors. A union of anchor and outer-block offsets can
change marked-byte/count limits and includes nested anchors suppressed by the
range scanner. A fused scanner also needs to defer range errors until the
original anchor parse and anchor-only MCE selection have succeeded. Neither
fusion nor the search substitution is implemented in this packet.

ADR 0003 atomic refusal, ADR 0005 explicit resource limits and measured evidence,
and ADR 0006 preservation and validation constrain both experiments. The
existing scanners, branch choices and refusal boundaries remain authoritative.
All 33 previously read goal/ADR constraint files are hash-verified unchanged.

All six repository evidence gates pass. The probe builds with the locked
dependency graph and passes formatting; analyses replay deterministically.
There is no fresh latency, allocation, RSS, hardware-counter, cold-cache or
scaling claim. The broader non-iWork performance goal remains active.

[Evidence packet](results/change-0712/README.md).
