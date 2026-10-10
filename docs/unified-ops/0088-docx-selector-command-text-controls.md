# 0088 — DOCX paragraph selector command text controls

Command text policy now runs inside the shared mutation planner after resolving
the complete target set. Homogeneous supported paragraph selectors accept the
same CRLF/CR/LF/VT/tab input behavior as direct paragraph command setters. The
full unified operations goal remains active.

## Result

The transaction coordinator has one internal set implementation for native and
command policy. Public `Edit::set` keeps native input semantics. Command dispatch
selects command policy without changing limits, poisoning, formatting barriers,
per-item isolation or commit retention. The DOCX planner resolves targets once,
refuses controls if any target is not a paragraph, proves write selections and
then normalizes inputs before shared property admission and native mutation.

This removes the earlier direct-path pre-query and adds selector support without
a second query pass. Eligible text-only paragraph fanout uses the native bulk
route from 0087. Predicates observe prior successful changes. Mixed paragraph/run
selections are refused before staging; empty matches, unknown predicates and
unscoped selectors retain their normal typed diagnostics. The physical-source
fence encloses command planning and execution just as it encloses native sets.

Run addition still uses the same input normalizer. Paragraph addition and run
setters with controls remain unsupported. There is no new public policy type,
selector grammar or document model. The capability reason and generated delta
now include homogeneous paragraph selectors in the supported input context.

## Validation

New eager/source and atomic/best-effort tests cover paragraph selector fanout,
CRLF/CR/VT normalization, tabs, a preceding formatting barrier, repeated no-ops,
exact patch inversion and unchanged source bytes. A selector whose predicate
matches a preceding text edit confirms sequential execution. Negative tests
check bare selector refusal, zero matches, unknown predicate diagnostics and
mixed paragraph/run rejection, all with original item indices. Earlier direct
paragraph tests now use a mixed selector for their refusal case because a
homogeneous paragraph selector is supported.

The command fixture uses an explicit comparison predicate: an unfiltered or
bare-token path remains governed by the existing path-versus-query classifier.
The parser was not changed to broaden mutation scope.

Reproduce with:

```sh
CARGO_NET_OFFLINE=true cargo test -p litchi-ops --no-default-features --features docx
CARGO_NET_OFFLINE=true cargo test -p litchi-ops --no-default-features
CARGO_NET_OFFLINE=true cargo test -p litchi-ops --no-default-features --features docx --test docx_batches
CARGO_NET_OFFLINE=true cargo clippy -p litchi-ops --no-default-features --features docx --all-targets -- -D warnings
python3 tools/officecli_capability_delta.py --officecli /home/zhuhe/code/litchi/3rdparty/OfficeCLI --check
python3 -m unittest discover -s tools -p test_officecli_capability_delta.py
```

Paragraph splitting, literal run-set policy, full text-output parity, broader
command-family coalescing, remaining formats, Python and paired performance
acceptance remain required. This checkpoint makes no timing or allocation claim.

Final validation passed 690 DOCX and 418 minimal tests/doctests, all 31 batch
tests and 14 delta tests. All-target Clippy passed with warnings denied. The
generated delta check, scoped formatting and diff hygiene passed. The rewritten
base remains an ancestor of HEAD. Both required staged evidence checks printed
nothing.

Staging and committing remain blocked by the read-only linked Git metadata at
`/home/zhuhe/code/litchi/.git/worktrees/unified-ops-api/index.lock`. No commit
was created; changes remain in the checkout. Checkpoint scratch files, generated
lockfiles, Python cache and build output were cleaned. No temporary worktree
was created; earlier benchmark evidence and its intentional lockfile remain.
