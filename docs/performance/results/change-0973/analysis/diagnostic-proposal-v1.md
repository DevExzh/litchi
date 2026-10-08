# 0973 dynamic-bias instruction diagnostic

This packet is a diagnostic-only follow-up to the 0972 offline decoder result. It
will use the unchanged retained 0967 baseline executable and one fixed cell:
`Float/e1/default` (`FloatingPictures.doc`, one edit, default retention), with
12 serial children on CPU 31. Each child records `instructions:u` at period
1,000,000 with `dwarf,32768`, performs five warmups and 80 measured samples,
and retains the combined 85-call marker lifecycle. The measured p50/p95 scope
is the 80 measured samples only. No speed, exact timed-region, phase-fraction,
or adoption claim is admitted.

The first child remains a predeclared pilot. It must have at least 100 exact
marker samples, zero lost events, and a positive marker period share. A failed
pilot or child stops the run without retry; there is no selected retry or
subsample. Every completed child is parsed and retained with its raw report,
perf data, launch metadata, command, terminal result, no-inline and inline
text, perf header, build-ID export, and dynamic-layout report. The target path,
SHA-256, BuildID, report PID/CPU, marker DSO, and all five/80 indices are
fixed custody fields.

The decoder correction is ASLR-safe and is configured separately for every raw
child. The helper reads the exact target `PERF_RECORD_MMAP2` rows from
`perf script -D` and parses the retained executable's ELF64 `PT_LOAD` headers.
It requires the target path and BuildID on every selected row, page-aligned
MMAP2 offsets matching the PT_LOAD offsets, and one consistent load bias across
all mapped load segments. It computes the corrected DWFL base as
`mmap_start - page_down(PT_LOAD.p_vaddr)`. It computes the perf comparator
start from the executable PT_LOAD mapping as `exec_mmap_start -
page_down(MMAP2.pgoff)`. Both values are written to a child layout receipt and
passed to the offline shim through explicit environment variables; no ASLR
address is stored in source. Ambiguous, missing, or inconsistent rows fail
that child before decoded results are admitted.

The shim interposes only `dwfl_report_elf` and `dwfl_module_info` for the exact
retained target. The six-argument ABI is bound to the extracted official
`libdwfl.h`; `fd` and `add_p_vaddr` pass through unchanged. In corrected mode,
only the exact target's fifth-argument base is changed to the child-derived
ELF load bias. Only a direct `/usr/bin/perf` caller of `dwfl_module_info` for
the exact target may receive the child-derived perf comparator start. A
no-op mode forwards all values unchanged and must produce byte-identical
perf-script stdout/stderr to stock in the retained replay control. Non-target
DSOs and BuildID aliases are not rewritten.

The runner requires an explicit plan and explicit root GO binding. It sets
`LC_ALL=C`, `LANG=C`, and `TZ=UTC`, serializes children, uses exclusive output
creation, kills and reaps a timed-out process group, and never retries a
failed child. The runner does not invoke Cargo or build the retained binary.
Before root GO only harmless parser, layout, wrapper, and synthetic controls
may run; no Office workload or new profile capture is permitted.

The reducer revalidates every raw artifact hash and reparses every child's
no-inline report with the fixed PID/CPU/event/period contract. It rejects
missing, duplicate, or reordered packet identities, layout custody drift,
marker DSO drift, lost events, malformed samples, and failed correctness
oracles. It reports all 12 children and all owner associations, treating a
missing owner as unknown/zero coverage only in the descriptive association
calculation. Any marker-period share is raw decoder output, not a timed-region
fraction. The authoritative outputs remain external under
`/home/zhuhe/litchi-performance-0973`; no repository evidence files are
created.
