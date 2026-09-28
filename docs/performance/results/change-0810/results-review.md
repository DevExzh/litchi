# 0810 numerical and profile results review

This review covers the terminal 0810 packet after all 216 native, 72
allocation, and four scoped profile processes completed. The main lanes, including 18 before-only qualification reports and samples,
contain 306 reports and 6,714 samples. The four profiles add four reports and
samples, for 310 reports and 6,718 samples in the packet. The
main analysis and independent raw audit use the same 18 native cases and the
same paired block data. The aggregate validator records the exact comparison
of all 18 native p50 ratios and bootstrap endpoints, all six native block
pairs per case, and the 144 formal allocation guard rows.

## Frozen policy result

Four capture cases satisfy the frozen benefit rule of at least 3% improvement
with a bootstrap 95% high endpoint below one:

| case | paired p50 ratio | improvement | bootstrap 95% interval |
| --- | ---: | ---: | ---: |
| large/capture | 0.950252156 | 4.974784% | [0.943694980, 0.953075729] |
| vendor/capture | 0.959269905 | 4.073009% | [0.954917265, 0.966247688] |
| unicode-vendor/capture | 0.961971580 | 3.802842% | [0.958162323, 0.978841271] |
| valid-4attr/capture | 0.958714372 | 4.128563% | [0.952841875, 0.968791003] |

The benefit gate therefore passes without relying on an allocation count or a
Callgrind result. No one of the 18 native p50 rows meets the latency veto of
ratio above 1.05 with a bootstrap low endpoint above one. Every one of the
144 formal allocation comparisons—calls, allocated bytes, net live bytes, and
peak above entry across both blocks and all 18 cases—is equal before and
after, so the resource guard has no increase. The ordinary, vendor,
unicode-vendor, and valid-4attr rows showing the same direction broadens the
observed fixture coverage, but does not establish a universal workload gain or
causal attribution.

## Tail and RSS diagnostics

The native p50 spread review has no spread flag above 5%. The packet does
retain 21 review-only native spread flags: 10 p95/p99 timing spreads and 11
RSS spreads. Three conspicuous paired p99 block deltas are tiny/capture block
0 at +96.947276%, large/commit block 3 at +11.276773%, and vendor/commit
block 4 at +20.163655%. Their paired p99 medians remain below one, but their
bootstrap intervals cross one (respectively [0.980171, 1.483035],
[0.952592, 1.048934], and [0.939305, 1.110487]). These observations support
the packet's diagnostic limits and provide no tail-latency improvement claim.

Process RSS is review-only evidence. Allocation guard equality does not imply
an RSS reduction, and the packet makes no RSS, native-cycle, or phase-fraction
claim.

## Scoped profile evidence

All four Callgrind reports pass the repaired reader's exact-scope and fixture
checks. `namespace_uri_probe::capture_region_0793` has one incoming owner call
and owner self Ir of 10 in each publication. The two before owner summaries
are 549,292,102 and 549,348,645 Ir; the two after summaries are 537,148,235
and 537,183,147 Ir. The direct scanner-to-
`NsReader::process_event` edge disappears after the change while the scanner
still makes 282,612 `Reader::read_event_impl` calls. `NamespaceResolver::push`
retains self Ir 13,019,166 in every publication. Other global
`NsReader::process_event` callers remain, so this is a mechanism diagnostic
for the scoped wrapper rather than a global zero-call result.

Callgrind Ir is guest-instruction attribution. It does not measure native
latency, RSS, cycles, phase fraction, or production speedup, and it is not an
additional adoption threshold.

The offline raw-audit check also exposed a serialization-only mismatch: the
qualification output was reconstructed as tuples in the reader while JSON
stores it as a list. The reader repair emits the JSON list representation;
the retained `root-audit.json` bytes and all numeric policy values are
unchanged. This repair adds no measurement.

## Review disposition

The observed numeric values satisfy the frozen numerical policy, with zero
latency and formal allocation violations. Quality, semantic, source, and
profile checks are separate evidence and remain bound to their recorded
receipts and fixture identities. The result is limited to this PPTX fixture,
host, toolchain, and 18-row workload; it does not establish cross-format,
historical, platform-wide, or universal document improvement. Final source
retention and owned-target cleanup remain governed by the decision,
disposition, and cleanup witnesses.
