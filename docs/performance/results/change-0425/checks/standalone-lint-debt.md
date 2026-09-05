# 0425 standalone lint debt

The standalone performance harness reached Clippy and failed with 29 unique
diagnostics. The [receipt](perf-baseline-strict.json) records Rust 1.98.1,
revision `340cc91ae2bdec338dfe7682b4b5d8c219a2d288`, four Cargo jobs, and exit
101 from 15:19:51.846 to 15:20:13.504 UTC. The raw [log](perf-baseline-strict.log.gz)
is retained.

| Category | Count | Locations |
| --- | ---: | --- |
| `redundant-field-names` | 1 | `filesystem.rs:1884` |
| `manual-is-multiple-of` | 7 | `docx_story_hyperlinks.rs:132,141,145,153`; `lib.rs:453` (two expressions), `25344` |
| `field-reassign-with-default` | 1 | `docx_story_hyperlinks.rs:557` |
| `needless-lifetimes` | 1 | `filesystem.rs:1518` |
| `collapsible-if` | 8 | `filesystem.rs:4323,5368,5393`; `parallel_metrics.rs:251,263`; `lib.rs:15197,15204,51623` |
| `manual-contains` | 1 | `filesystem.rs:5623` |
| `obfuscated-if-else` | 3 | `xls_numeric.rs:801,804`; `lib.rs:19846` |
| `useless-conversion` | 2 | `lib.rs:15239,56338` |
| `single-range-in-vec-init` | 1 | `lib.rs:25351` |
| `needless-range-loop` | 2 | `lib.rs:52736,52755` |
| `unnecessary-map-or` | 1 | `lib.rs:56366` |
| `needless-update` | 1 | `operation_metrics.rs:1455` |

The standalone harness working-tree diff at capture time contained only
`tools/perf-baseline/src/lib.rs`. Its five changed CFB chunk views were:
`cfb_workbook_chain` directory/name chunks, `cfb_directory_entry_name`,
`cfb_opaque_payload_ranges`, and the Unicode sheet-name path in
`parse_xls_bound_sheets`; one additional hunk only wrapped an import. Every
`lib.rs` finding above is outside those changed hunk ranges. The other finding
files were unchanged. Thus this is an honest partial strict result: the leaf
chunk migration is not implicated by the remaining harness diagnostics, while
the modified CFB views pass their subsequent eight-test focused XLS harness
check (`perf-xls-tests.json`).

The receipt's `source_unchanged: true` means the source did not change during
the command. It does not mean the harness was the original Git blob: the
receipt's `tools/perf-baseline/src/lib.rs` hash is
`0ba81aa389eac1c97a179c2770c45f2aef9ece4cd979270a2857229a42aab9db`, while
the `340cc91ae` blob is `db2b572189f03bb216ac74c4b070aa6c737cfb817ed48756b1196e0287a2d05b`.
The diff above identifies the already-present five CFB conversions and keeps
the lint attribution separate from them.

The native-resave command did not reach compilation. Its [receipt](native-resave-strict.json)
and [log](native-resave-strict.log.gz) record exit 101 from 15:20:22.710 to
15:20:22.887 UTC because Cargo tried to update
`tools/native-resave/Cargo.lock` under `--locked`. No native-resave Clippy
finding can be inferred from this failure. The lock remained unchanged:

```text
tools/native-resave/Cargo.lock
e8d4bed6c73e1b1944858a0328222824bd194414347bbf3e83e6899dbcdeae48
```

That SHA-256 is identical in the receipt before/after snapshots, the working
file, and the `340cc91ae` blob. Do not remove `--locked` or update the lock as
part of this debt note. The standalone strict gate remains partial; no blanket
lint allowance or performance claim is introduced.
