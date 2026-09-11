# XLSB `themeFamily` performance evidence

This directory owns a bounded, process-isolated profile for the XLSB host integration of the shared DrawingML `themeFamily` API. The harness is under `harness/`; raw process JSON and `/usr/bin/time -v` output are under `results/`; `report.md` is generated only after a complete final run.

The matrix uses the small native `test-data/ooxml/xlsb/date.xlsb` workbook and an in-memory opaque-extension-heavy derivative. The derivative keeps the native family attributes and adds 96 valid DrawingML `a:ext` children inside the family namespace `extLst`, with vendor attributes, comments, and a distinct foreign namespace declaration plus nested payload for every child. The harness independently checks that shape before timing, so no generated XML fixture is retained.

The lanes are:

- `codec_read`: shared DrawingML base Theme codec component reference.
- `metadata_read`: eager XLSB open, Theme read, and family discovery.
- `source_read`: source-backed open with counted logical `ReadAt` calls and bytes.
- `family_clone`: clone of the parsed shared Family value, with source-pointer sharing checked after timing.
- `noop`: prepared eager Theme no-op commit/publication, with source-pointer sharing checked after timing.
- `add`, `update`, `remove`: prepared host family edit, forward publication, and exact inverse publication.
- `base_edit`: prepared host base Theme edit and inverse, as a component reference for read-regression context.

Each lane runs three fresh processes with three warm-ups and thirty measured samples per process. The process-local allocator observer keeps direct allocation bytes separate from successful realloc old/new sizes. Requested allocation bytes are direct sizes plus realloc new sizes; the live-balance equation is checked per sample, including an 8-byte to 32-byte realloc self-test. Peak live is reported as the incremental peak above the timed closure's live-before baseline. `/usr/bin/time` RSS is whole-process RSS, not per-operation memory.

The metadata input clone is prepared before the timed/allocation region. No large Theme or family bytes are hashed in a no-op or clone timed region; source identity checks and semantic/preservation/inverse checks run afterward. Raw JSON records SHA-256 identities for the generated package, Theme, and Family bytes; the independent verifier requires those identities and sizes to match across every process and operation for each fixture. The report makes no broad speedup or regression claim. A before/after control build would be required for that comparison.

Run the final profile from the repository root with the reviewed source frozen:

```sh
WARMUP=3 SAMPLES=30 PROCESSES=3 PERF=1 \
  docs/report/spec-gap-validation-evidence/xlsb-theme-family/performance/run_profile.sh
```

`PERF=1` retains one optional `perf stat` attempt; an unavailable or restricted hardware counter does not invalidate the allocator, timing, source-manifest, or semantic gates. The runner records exact commands, Rust toolchain and host information, binary SHA-256, a Cargo package source manifest before and after build/run, relevant source hashes, and all cleanup decisions. It removes only its own temporary directory and a newly created default profile target; a pre-existing target is preserved.
