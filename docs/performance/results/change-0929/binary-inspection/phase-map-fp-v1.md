# 0929 FP static phase map

The machine-readable map is [`phase-map-fp-v1.json`](phase-map-fp-v1.json), SHA-256 `0b2aa67f12291083b5e60676ca79c82c6956b2e51e49f2054646198395ae42e6`.
It binds the corrected non-stripped diagnostic ELF only:

- ELF: `/tmp/litchi-performance-0929/target-fp/release/litchi-perf-baseline`
- ELF SHA-256: `c195a1d94acb7b240ca86ad32c13a9fd8b4c7b3affe085122970fe0f75c2169e`
- Build ID: `1fab480cdc3541389487fa4e9fd5d3aebd240e89`
- Source `tools/perf-baseline/src/lib.rs` SHA-256: `7904a6161e7b40a675d61acd9867b31b81a47f049b1ebb5a9979778bf14a373a`

All PCs are static ELF virtual addresses. Add the capture's measured PIE load bias to obtain runtime PCs. `return_pc` is the instruction after the call; `return_minus_one` is the final call byte for tools using that convention.

| scope | call | return | return-1 | source / target |
|---|---:|---:|---:|---|
| expected snapshot | `0xf47137` | `0xf4713d` | `0xf4713c` | `lib.rs:50757`, `Package::section_layout_snapshot` |
| expected publication | `0xf4783f` | `0xf47844` | `0xf47843` | `lib.rs:50775`, expected publication clone `0x16bbce0` |
| timed query | `0xf49831` | `0xf49837` | `0xf49836` | `lib.rs:51009`, `Package::section_layout_snapshot` |
| timed publication | `0xf49cb8` | `0xf49cbd` | `0xf49cbc` | `lib.rs:51034`, timed publication clone `0x16bc8f0` |
| output verification | `0xf4a494` | `0xf4a499` | `0xf4a498` | `lib.rs:51123`, `verify_docx_section_layout_output` |
| inverse publication | `0xf4a622` | `0xf4a627` | `0xf4a626` | `lib.rs:51128`, inverse clone `0x16bd500` |

The runner, query, publication clones, verification, and `sha256_hex` all begin with `55 48 89 e5` (`push rbp; mov rsp,rbp`), proving the requested frame-pointer prologue. This map proves static callsites; phase attribution still requires a valid FP ancestry and the capture-specific load bias.

## Regeneration commands

```sh
BIN=/tmp/litchi-performance-0929/target-fp/release/litchi-perf-baseline
nm -C --defined-only "$BIN" | rg 'run_docx_source_backed_existing_section_layout_edit_save|main_section_layout_snapshot|section_layout_snapshot|publish_section_layout_commit_to_stream|publish_section_layout_inverse_to_stream|verify_docx_section_layout_output|sha256_hex' > /tmp/litchi-performance-0929/phase-symbols.txt
objdump -drwC --start-address=0xf46b40 --stop-address=0xf4f460 "$BIN" > /tmp/litchi-performance-0929/runner-fp-disassembly.txt
readelf --debug-dump=decodedline "$BIN" | rg -n -B2 -A3 'lib.rs[[:space:]]+(50742|50757|50774|50775|50969|50971|51006|51008|51009|51011|51013|51025|51027|51030|51033|51034|51036|51123|51128)[[:space:]]' > /tmp/litchi-performance-0929/debugline-phase-focused.txt
addr2line -e "$BIN" -f -C 0xf49831 0xf49837 0xf49cb8 0xf49cbd 0xf4783f 0xf47844 0xf4a494 0xf4a499 0xf4a622 0xf4a627
```
