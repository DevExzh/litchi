# 0362ul26a — generic compile-only child lifecycle

Date: 2026-10-24. This is a follow-up to
[0362ul26](0362ul26-compile-only-build-overlap-guard.md). It preserves the
accepted lease/phase contract and records the generic-runner lifecycle repair
and the three races found during independent rereview.

## Gap and repaired scope

The generic runner started each child in a new process group and checked that
group plus the direct descendant tree after the leader exited. A helper could
call `setsid`, outlive its leader, and be reparented to PID 1 before that scan.
The old check then reported an empty tree without proving that the helper was
gone. A process-group-empty result is not a process-tree ownership proof.

`tools/native_qualification_runner.py` now enables Linux
`PR_SET_CHILD_SUBREAPER` before spawning a command. While the leader runs, it
records every newly owned process with PID, kernel start time, parent/group,
executable, cwd and argv. The same start-time-bound identities are retained
when Linux reparents escaped descendants to the runner. Cleanup signals only a
recorded PID whose current start time still matches, reaps owned children when
possible, escalates from TERM to KILL within a bounded window, and never
signals a reused or pre-existing PID. An escaped descendant remains a failed
event even when cleanup succeeds; survivors and PID-reuse observations are
retained in the event journal. Non-Linux hosts fail closed because this runner
has no equivalent ownership primitive there.

The rereview fixes close three ownership races. The leader subtree is walked
only when the leader PID and captured kernel start time still match; a reused
leader and its current descendants are rejected. A pre-existing runner child
and its current subtree are removed from the candidate set before any new
identity is admitted, including when PID reuse is discovered after the first
leader-tree scan. During both bounded cleanup phases, and before each final
KILL signal, ownership is recaptured and newly observed identities are
signalled and reaped. The final survivor snapshot therefore remains a bounded
failure proof rather than relying on one stale process-tree scan.

The focused tests use a committed fake-Cargo tree only. One fake command
creates a real `setsid` child and exits immediately; the test verifies the
reparented child is recorded, killed/reaped, and reported as a lifecycle
failure. Additional synthetic tests reject a reused leader tree, discard a
reused pre-existing subtree before admission, and recapture a child forked
after the first final-KILL snapshot. A separate test verifies that a changed
start time prevents signalling a reused PID. Existing overlap and peer-lease
tests remain unchanged.

## Checks

No Cargo, Rust, native qualification, timing, allocation, ABBA, or benchmark
command was run. The exact source-only checks and reduced logs are outside the
repository under
`/home/zhuhe/code/litchi-worktrees/.coord/cfb-generic-lifecycle-20261024/`:

| Check | Exit | Log (bytes, SHA-256) |
| --- | ---: | --- |
| `python3 -m py_compile tools/native_qualification_runner.py tools/test_native_qualification_runner.py` | 0 | `pycompile.log` (0, `e3b0c44298fc1c149afbf4c8996fb92427ae41e4649b934ca495991b7852b855`) |
| `python3 -m unittest -q tools.test_native_qualification_runner` | 0; 40 passed | `native-runner-tests.log` (333, `af9fd67e3be1ae29bddd5a145aff01ed29922d3ac675befd6b19c1f097afc40c`) |
| `python3 -m unittest -q tools.test_drawingml_raw_qualification tools.test_legacy_native_qualification` | 0; 31 passed | `paired-runner-tests.log` (98, `014164901c3dffbd4186ebb368e8c20db9b505ad16afff9ae84d302b1379bf68`) |

The initial repaired source files at this leaf were:

| Path | Bytes | SHA-256 |
| --- | ---: | --- |
| `tools/native_qualification_runner.py` | 188,946 | `38c221982e201fd5e6b88c94f7af8fabb53bd0d71a63c104d2f692434d6a1a56` |
| `tools/test_native_qualification_runner.py` | 73,806 | `8a278c5b4bc3733de2e9620b3ee8efe12a33205394eb709bd1c4701243d41f89` |

The follow-up source-only rerun used the same commands and is retained under
`/home/zhuhe/code/litchi-worktrees/.coord/cfb-generic-lifecycle-20261024-r2/`:
`pycompile.log` (0 bytes, SHA-256
`e3b0c44298fc1c149afbf4c8996fb92427ae41e4649b934ca495991b7852b855`),
`native-runner-tests.log` (333 bytes, SHA-256
`07a7924c6d51d2f57916921f096a3f656c1d1094824378cf456f74e2520f6c67`), and
`paired-runner-tests.log` (98 bytes, SHA-256
`df92f1da813f79675b00846421966de998dfd837b423a276a44f425ef55baa3b`).

The rereview-fix source-only rerun used the same commands and is retained under
`/home/zhuhe/code/litchi-worktrees/.coord/cfb-generic-lifecycle-20261024-r3/`:

| Check | Exit | Log (bytes, SHA-256) |
| --- | ---: | --- |
| `python3 -m py_compile tools/native_qualification_runner.py tools/test_native_qualification_runner.py` | 0 | `pycompile.log` (0, `e3b0c44298fc1c149afbf4c8996fb92427ae41e4649b934ca495991b7852b855`) |
| `python3 -m unittest -q tools.test_native_qualification_runner` | 0; 43 passed | `native-runner-tests.log` (333, `4dd7bfe7b0ea852c29c0173bf46cb0e8ef80f37f7b02b474504f6cd5a4d21dd9`) |
| `python3 -m unittest -q tools.test_drawingml_raw_qualification tools.test_legacy_native_qualification` | 0; 31 passed | `paired-runner-tests.log` (98, `046de6d4d238de31888c5a2a2816070524f563c84a9c813ed0e25d2aafde2084`) |

The current source files at this leaf are:

| Path | Bytes | SHA-256 |
| --- | ---: | --- |
| `tools/native_qualification_runner.py` | 192,850 | `6cafe7f8ca6db768ded17c9e6bf715926039b4574e8fea0463b89e9d05bb8e2e` |
| `tools/test_native_qualification_runner.py` | 78,837 | `8319172e4c4f95afa919141e4f994b10bd35b161b1bf22456ca935137ad66ad7` |

This record makes no correctness, performance, RSS, target-size, or native
qualification acceptance claim. Compile-cohort permits remain paused until
this leaf is independently reviewed and integrated.
