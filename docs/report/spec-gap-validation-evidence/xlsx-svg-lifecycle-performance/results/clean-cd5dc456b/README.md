# XLSX SVG lifecycle sealed semantic capture

Review status: approved for reproducible semantic and scaffold evidence.
The retained timing and resource numbers remain provisional; this bundle is
not approved evidence for performance or optimization claims. A fresh run
with corrected host provenance and an uncertainty summary is required for
that purpose.

This capture used clean committed source
`cd5dc456b4d64a6aba6b4833a424d6bcc32e1707` and the approved lifecycle production
pin `ac288a303264f9ea0bb4081baa44031bee5b79a7`. All 19 pinned paths passed their
Git-blob checks. The native read fixture is tracked; no third-party symlink or
omitted manifest input was used.

The 69 acceptance lanes each ran in three fresh processes, with two warmups
and 20 measured samples per process: 207 receipts and 4,140 measured samples.
All 207 stderr files are empty and all GNU time sidecars report success.
The seven exploratory lanes are outside this measured acceptance capture.

The retained runner verification and independent root verification both check
4,870 manifest inputs and recompute the report from raw samples. The before
and after source manifests have SHA-256
`c347f6640a6c2bda22ed51e3b202f53eaf857f30eb915631c31836e0ad790d52`.
The measured executable's stable SHA-256 is
`0036e449333d1690260eb60a8c2447fef58476d6ea63187d377923862caf82fc`.
Its disposable Cargo target was cleaned after capture.

The measured runner had a host-metadata quoting defect: CPU model and total
memory appear as `unavailable` in the original `build-provenance.txt`.
`host-supplement-after-run.txt` records those fields separately after the run
on the same host. Original provenance was not rewritten. The runner fix is
`e7b80f056`; that later commit did not produce these measurements.

These are retained allocator-instrumented workload observations, not a
before/after improvement, scaling guarantee, native changed-save acceptance,
or rendering claim. Fixture setup and each lane's timed work follow the
requirements at the measured commit. RSS describes the whole fresh process;
requested allocations and peak-live deltas describe the instrumented sample.

To replay verification, use the measured commit's `verify.py` with `--root`
pointing to a clean checkout of that commit, `--results` pointing here,
`--report` pointing to this directory's `report.md`, and `--output` pointing
to a new review receipt. Do not rebuild or replace raw measurements to verify
the retained report.
