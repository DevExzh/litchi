# XLS F3 role-split command correction

The paired packet is preserved at
`sheet-native-readiness-472b-f3-role-split.json`. A corrected packet was
published at
`/home/zhuhe/code/litchi-worktrees/.coord/sheet-preparation-691941/sheet-native-readiness-472b-f3-role-split-corrected.json`
with SHA-256
`812b98b78d56b626d5a310ba07fe55f2e08a39e3bede81f1b10f69b55a91d08b`.

All three command strings were normalized so JSON decoding contains actual
shell quotes around `$CPU`, never a literal backslash-quote. A source-only
render test substituted `CPU=30`, ran `shlex.split(posix=True)`, and passed:

- baseline reuse: `taskset -c 30`;
- candidate build: `taskset -c 30`;
- qualification: `taskset -c 30` and `--cpu 30`.

The corrected packet also records structured argv/env facts. No source,
candidate, baseline, binary, Cargo target, or historical packet was changed;
no Cargo/native/qualification execution occurred.
