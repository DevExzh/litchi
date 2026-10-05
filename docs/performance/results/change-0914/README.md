# 0914 DOCX section-layout baseline

This packet records a descriptive baseline for
`docx_source_backed_existing_section_layout_edit_save`. It contains four
fresh pinned children, each with 30 untimed warmups and 200 retained samples,
on CPU 12. The timed value is the harness's six-phase sum (preparation,
source-backed open, section query, edit, commit, and sequential publication).
GNU `time -v` values are setup-inclusive whole-child process observations.
There is no candidate comparison or speedup claim.

The accepted source is revision
`a4ca85e28cdea151176534b7f8dc912253a1ca1e`. The external archive is
`/home/zhuhe/litchi-performance-0914/source-HEAD-code.tar`, 171120640 bytes,
SHA-256
`7b77f1875adcdc3c822351b11dcf621f4ebff501fe501cab80fd4fb3f9a59571`.
The external release binary is
`/home/zhuhe/litchi-performance-0914/target/release/litchi-perf-baseline`,
69940744 bytes, SHA-256
`fe56e5192c1198e831583e870beb5e1ba4731fb2077faaffdfc46044bb0ed2db`.
The frozen source, protocol, build, child, and process receipts are copied
under `receipts/`; the four raw reports are under `reports/`. The archive and
binary remain outside the repository.

`receipts/rust-files.json` records the frozen Rust/toolchain configuration and
source file hashes; `source-tree-verify.json` records the independent archive
and extracted-tree check.

The retained reduction is `reduced.json` (27697 bytes, SHA-256
`770bc343ce6ab351cfd86bd433657c32be40f5196accf9995c3f1061308ca058`). Full
external-custody verification is:

```sh
python3 docs/performance/results/change-0914/capture.py verify
```

To regenerate a separate reduction without replacing the retained file, use
the command recorded in `reduced.json`, which writes
`/home/zhuhe/litchi-performance-0914/reduced-replay.json`. The retained raw
reports support external-free statistical replay:

```sh
python3 -B docs/performance/results/change-0914/replay.py
```

This verifies the retained receipts and statistics without inspecting the
external binary or source archive. `custody-rejections.json` records the
omitted-fixture build attempt and the earlier child cohort rejected for receipt
drift. `cleanup.json` identifies the removed temporary trees and retained
external artifacts. Source/archive regeneration is described in the
[change record](../../0914-docx-section-layout-baseline.md).
