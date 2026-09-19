# 0697 — post-0696 MCE context-ownership attribution packet

Status: completed measurement and instruction-attribution experiment.
`performance_claim: none`. This packet is bound to current HEAD
`1bf58ace2c2d69db4ab88e21f62aff5f1ba7cf1e` (`perf(ooxml): skip empty namespace
scope installation`). It measures the current shared MCE codec after 0696 so
the remaining context-owned work can be inspected with fresh isolated sequence
profiles and native instruction evidence. It does not reuse 0695 timings and
does not turn isolated sequence ratios into additive capture fractions.

The corpus and call order come from the first real count-one diagnostic capture
in change 0693. That capture is historical provenance: its temporary
instrumentation bytes and timings are never treated as current measurements.
`prepare.py` rechecks the trace files, extracts the same fourteen OOXML members,
and binds their bytes and the 0693 call topology to the current 0696 source
hashes. `topology-binding.json` records that the complete PPTX tree is unchanged
between the historical 0693 commit and the current 0696 commit; the shared MCE
codec is the source owner that changed.

The standalone `mce-attribution-0697` probe owns its input bytes, reports exact
input/output identities and full output hashes, and emits buffered per-sequence
samples. `measure.py` runs the four ordered legs for `all`, `presentation`,
`slides`, and each individual slide. `profile.py` collects startup-subtracted
hardware-counter controls and one call-graph recording. `summarize-profile.py`
retains each repeated slope. `annotate.py` and `audit-instructions.py` retain
and verify symbol assembly and `perf annotate` output for the current binary.
`count-elements.py` is an independent source-syntax census; it is not a claim
about production branch execution.

Reproduce in a disposable checkout at the recorded production revision, copying
this packet and its report into that checkout without changing HEAD. Preserve
the original receipts elsewhere: drivers overwrite results. The locked offline
build requires its dependencies in Cargo's local cache. CPU 12, sibling scratch
paths and toolchain details are host-specific; record any adaptation. Run the
following commands from this packet directory. The scripts resolve repository
paths themselves.

The ignored root `Cargo.lock` is unchanged from the retained
`../change-0696/workspace-Cargo.lock`. In a disposable checkout where that root
file is absent, restore it from that retained copy before running the drivers;
the build and cleanup receipts verify its digest.

The expected execution order is:

```text
python3 environment.py
python3 prepare.py
python3 build.py
python3 measure.py
python3 profile.py
python3 summarize-profile.py
python3 report-metrics.py
python3 count-elements.py
python3 annotate.py
python3 summarize-instructions.py
python3 audit-instructions.py
python3 validate.py
python3 audit.py
python3 cleanup.py
python3 seal.py
```

`build.py` binds every tracked Rust and manifest source, the standalone probe
tree, the workspace lockfile, and the exact current revision. The initial
locked build receipt records an initial package-name mismatch caused by the
copied 0695 lockfile; that failed attempt is retained as provenance only. The
probe package name and lockfile were corrected to 0697 before the successful
frozen build, whose receipt is the one used by later audits.

The packet owns only `litchi-target-0697` and `litchi-0697-profile` as sibling
scratch directories. `cleanup.py` removes exactly those paths and records the
unchanged workspace `Cargo.lock` digest. Raw native samples, textual counter
reports, instruction annotations, and validation receipts remain in this
packet after cleanup; the binary and raw `perf.data` remain represented by
their hashes in the receipts after scratch cleanup. `environment.py` records
the host and CPU context before measurements. No iWork material is included.
