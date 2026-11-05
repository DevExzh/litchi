# XLS F3 role-split command correction

The paired packet is preserved at
`sheet-native-readiness-472b-f3-role-split.json`. A corrected packet was
published at
`(deleted-coord-artifact)/sheet-native-readiness-472b-f3-role-split-corrected.json`
with SHA-256
`2a2b61ef7a5d0a13489bd4285ffe4b600afbd7ebed2223581229c625ea477e03`.

All three command strings were normalized so JSON decoding contains actual
shell quotes around `$CPU`, never a literal backslash-quote. A source-only
render test substituted `CPU=30`, ran `shlex.split(posix=True)`, and passed:

- baseline reuse: `taskset -c 30`;
- candidate build: `taskset -c 30`;
- qualification: `taskset -c 30` and `--cpu 30`.

The corrected packet also records complete structured argv/env launch facts,
including the fresh candidate target/TMP and reused baseline target. No source,
candidate, baseline, binary, Cargo target, or historical packet was changed;
no Cargo/native/qualification execution occurred.
