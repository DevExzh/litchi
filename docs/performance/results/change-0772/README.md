# 0772 evidence packet

[Report](../../0772-opc-mutated-save-current-baseline.md).

- `source.json`, `constraints.json`, `workspace-inputs.json`: captured inputs.
- `build.py`, `build.json`, `build-native.log`: fresh ordinary build.
- `capture.py`, `captures/`: three serial processes and complete reports.
- `analyze.py`, `analysis.json`: replayable descriptive statistics and flags.
- `source-review.md`: measured boundary and correction of the proposed decode ROI.
- `timed-boundary.json`, `timed-boundary.objdump.gz`: exact static call site.
- `profile.py`, `profile-qualification/`: rejected ordinary-stack qualification,
  including the first decoder failure, corrected decoded stacks and raw capture.
- `decode-profile.py`, `qualify-stacks.py`: successful offline decode and explicit
  attribution rejection. The original raw `perf.data` is archived as specified
  in `profile-storage.json`; restore it before invoking the decoder. Decoding
  again also requires a matching ELF; the build target has been cleaned.
- `cleanup.json`, `draft-disposition.json`: verified owned cleanup.

Offline result replay (no binary required):

```sh
python3 -B docs/performance/results/change-0772/analyze.py
python3 -B docs/performance/results/change-0772/qualify-stacks.py
```

The capture/build scripts guard the recorded source revision. Rebuilding requires
that revision, its recorded lock inputs and toolchain, and fresh output paths;
the scripts intentionally refuse to overwrite their existing captures.
`host-observation.json` was recorded after capture and is labeled accordingly.
No historical timing comparison or candidate acceptance is supported.
