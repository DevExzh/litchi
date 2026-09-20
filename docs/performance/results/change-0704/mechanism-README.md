# 0704 default-MCE call-count mechanism

This separate instrumented executable reuses the 0703 default-wrapper trace
format: exact raw/output SHA-256, lengths, output ownership/capacity, pointer
identity and explicit phase-file labels. It has no timers. The shared codec
patch is temporary; `mechanism/build.py` verifies the complete 603-file final
candidate census, builds in the root's serial lane, and restores the exact
codec in `finally`. Native timing executables are separate frozen files.

```sh
python3 docs/performance/results/change-0704/mechanism/build.py
python3 docs/performance/results/change-0704/mechanism/run.py
python3 docs/performance/results/change-0704/mechanism/audit.py
```

Twelve fresh processes cover real/generated inputs, no-op/one/two edits and
two repeats. Capture calls remain 18 real / 19 generated. Changed real commits
call MCE 6 times after one edit and 7 after two edits; generated changed
commits remain 19. Both no-op commits call it zero times. Setup and verify
phases are explicitly excluded from those counts. Every trace call succeeds,
source/output hashes bind exact bytes, and repeats agree after excluding
allocation addresses. The diagnostic covers this default wrapper only, not
all semantic readers or custom profiles. The public observer is the authority
for actual retained charge; trace capacity sums include presentation output
and are not retained-table memory accounting.

The first build had an incorrect relative dependency path after relocation;
its receipt and failure log are preserved in `mechanism/preflight-build-01/`.
The manifest was fixed and rebuilt. The audit's generated-no-op expectation
was corrected to zero, and JSON list normalization was fixed; raw traces were
unchanged. The successful final audit independently checks the counts and
source restoration. No failed build supplied timing data.
