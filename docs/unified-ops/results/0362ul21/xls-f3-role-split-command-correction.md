# XLS F3 role-split command correction

The paired packet is preserved at
`sheet-native-readiness-472b-f3-role-split.json`. A corrected packet was
published at
`/home/zhuhe/code/litchi-worktrees/.coord/sheet-preparation-691941/sheet-native-readiness-472b-f3-role-split-corrected.json`
with SHA-256
`e329ba6f93abf8a82a76511c8e6eed3b2df48534be6009272764bd77d66fab80`.

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
