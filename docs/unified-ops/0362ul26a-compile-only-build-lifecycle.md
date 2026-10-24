# 0362ul26a — generic compile-only child lifecycle

Date: 2026-10-24. This is a follow-up to
[0362ul26](0362ul26-compile-only-build-overlap-guard.md). It preserves the
accepted lease/phase contract and repairs a concrete generic-runner lifecycle
gap found during the first compile-cohort review.

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

The focused tests use a committed fake-Cargo tree only. One fake command
creates a real `setsid` child and exits immediately; the test verifies the
reparented child is recorded, killed/reaped, and reported as a lifecycle
failure. A second test verifies a changed start time prevents signalling a
reused PID. Existing overlap and peer-lease tests remain unchanged.

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

The repaired source files at this leaf are:

| Path | Bytes | SHA-256 |
| --- | ---: | --- |
| `tools/native_qualification_runner.py` | 188,569 | `babfc45c1ea2ded9c78f7c3ea0caebca6da9b8c2c976b271ef0eb74af41914f2` |
| `tools/test_native_qualification_runner.py` | 73,612 | `dbcdf40e65091d4fbc0c1a72678f0dfb5c05acd4b7e7969e078230e0fa42c95b` |

This record makes no correctness, performance, RSS, target-size, or native
qualification acceptance claim. Compile-cohort permits remain paused until
this leaf is independently reviewed and integrated.
