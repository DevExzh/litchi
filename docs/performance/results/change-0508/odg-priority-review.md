# ODG priority review after 0507

0502's historical comparison measured the old `f3f9221` parser against the
richer metadata parser. Retained p50 changes were +5.90–6.39% plain-small,
+5.95–6.06% plain-large, +53.40–55.21% metadata-small and +110.09–111.14%
metadata-large. Those observations remain historical evidence.

0504–0507 instead have successive matched richer-parser controls. The latest
0507 pair has 17.06–33.51% p50 improvements with no paired adverse flag above
5%. This is the causal claim supported by that batch's experiment.

The 0507 endpoint medians lie below the old 0502 control in all four rows
(approximately 52–53% on plain inputs and 13–19% on metadata inputs). This is
only a descriptive cross-capture trajectory, with no joint confidence interval.
Metadata checksums differ: old 2400/40886 versus current 2432/41014. The delta
is consistent with recognition of three 3D owners and the Scene name, eight
checksum units per page. The checksum does not certify all metadata semantics.

The old control commit does not contain the probe path, and 0502 did not bind
its complete probe source to that binary. A new old-to-current claim needs a
fresh four-corpus ABBA capture with the exact probe/manifest graft recorded for
the old checkout, complete source and executable custody for both phases,
identical corpora/toolchain/CPU/sample protocol, and separate rich semantic
checks. Profiles must distinguish whole-child setup from open-only timing.

This audit does not close the historical comparison with a new causal claim.
Current residual costs are the appropriate guide for further optimization;
0507 retains 349.4M inclusive parse_content and 153.2M scalar attribute
Callgrind instruction references on metadata-large. Broader default CRUD
coverage is now a useful next priority, alongside the unresolved 0499 local
Part-batch overhead and 0500 K1 latency/RSS follow-ups.
