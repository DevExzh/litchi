# Initial supplemental lookup run

This exact probe and capture preceded an unused-Result warning cleanup in the
warmup loop. Its production sources match the same 0690 candidate/baseline.
Only the standalone warmup statement changed; the timed loop is unchanged.
Final supplemental builds and captures are retained at the parent packet.
Original raw data, probe/lockfile, build identities and summaries remain here.
These initial results are not pooled with final supplemental results.

To rebuild this initial probe, restore its directory to the parent packet's
`lookup-probe/` location in a separate checkout of the selected production
revision. Its relative dependency paths intentionally retain that original
location; the archive itself is not a new standalone build root.
