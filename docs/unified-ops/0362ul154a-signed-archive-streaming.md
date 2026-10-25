# 0362ul154a — bounded signed source-archive verification

The signed source generator in coordinator record 152a reached the 4 GiB
memory cap while validating a 6,646,016,000-byte tar. The terminal journal
proves the OOM; source inspection identifies whole-artifact and whole-member
reads in the protocol, without attributing the kill to a precise Python
allocation frame.

`prepare_signed_native_inputs_v2.py::_artifact_identity` now counts bytes and
computes SHA-256 using 1 MiB reads. `_archive_tree_identity` similarly streams
each member's Git-blob SHA-1 and ignored-file SHA-256/byte count. The Git blob
header uses the tar's declared member length. Existing exact tree, regular
file, mode, duplicate, link, path, ignored-file size and digest checks remain.
The identity dictionaries and refusal boundaries are unchanged.

The policy tests migrate with this repair. A 192 MiB sparse artifact verifies
under a 128 MiB address-space cap, and a member-reader test rejects unbounded
reads while checking the expected Git and SHA-256 identities. Independent
review also compared successful and wrong-tree behavior with the previous
implementation. The integrated policy and signed-consumer suites pass all
25 tests:

```sh
PYTHONDONTWRITEBYTECODE=1 python3 -B -m unittest \
  tools.test_signed_native_policy_v2 tools.test_signed_native_v2_consumer
```

The external log is
`/home/zhuhe/code/litchi-worktrees/.coord/protocol-P-streaming-154a-20261025/focused-tests.log`,
124 bytes, SHA-256
`a30729f89e79713a7da5fcd1c31463cb6a770c6a0a6737331e0471994f48e2af`.
This is a source-only protocol check; it does not claim native, corpus,
performance or full-workspace qualification.

All candidate bindings remain pending in P. The completed current-native
binding for C=`f8cd398a770f7363f628be40edb3ce4219566c16` is retained as
verified evidence for that source, but the final contract audit requires a
corrected production candidate before binding and acceptance. The failed
signed v4 tar and its failure receipt remain preserved and unverified.
Final generation must verify the exact corrected candidate through this
bounded implementation; an archive from the previous source cannot stand
in for it. Nothing is pushed.
