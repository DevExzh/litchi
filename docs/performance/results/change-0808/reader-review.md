# 0808 terminal reader review

The terminal validator replays successfully after cleanup:

```text
python3 -B docs/performance/results/change-0808/validate_early_stop.py \
  --check --require-cleanup
0808 early-stop validation PASS: 18 qualification reports, 3 probe gates, gate-4 stop
```

This is an offline custody read. It does not run Cargo, a workload, a
capture, or a profiler.

## Findings

The validator checks the complete before-only qualification: 18 reports and
18 samples, allocation-binary identity, the sealed non-timing oracle fields,
and no imported timing. It checks all three probe-quality receipts and their
exact commands, including the 36-test result. The first three production
quality commands are checked byte-for-byte, as is the gate-four Clippy
command; the receipt sequence is exactly three passes followed by exit 101.

The production test log is independently summed across 85 result groups to
1,241 passed, zero failed, and three ignored. Both the candidate quality log
and the post-restoration baseline control must report the exact three
locations `464:14`, `538:10`, and `557:10`.

The diagnostic assertion uses the right evidence boundary. Each raw log has
three full `.err().expect()` error headers, while `clippy::err-expect` and
`clippy::err_expect` occur once as lint-note/help text. The validator matches
the three source-location headers and the three full diagnostic headers; it
does not incorrectly require three occurrences of the lint name.

The source chain is intact: the current tracked production census matches
the 9,196-file before manifest and the disposition's restored-source witness.
The cleanup witness records the three before binaries and the removed owned
target, and the validator accepts the packet after that removal. The
candidate allowlist, application witness, rejected disposition, retained
qualification/probe evidence, and failed quality logs all remain bound to the
same packet.

There is no reader finding that changes the early-stop disposition. No after
build, paired native/allocation lane, or profile lane was started, so the
packet makes no performance or adoption claim.

The remaining custody step is packet sealing. Before finalizing the archive,
`seal_packet.py` must seal the six report documents and pass both staged-index
and `HEAD` checks. The six document inputs are
`0808-pptx-direct-event-handling.md`, `BASELINE.md`, `CRUD_COVERAGE.md`,
`GOAL_AUDIT.md`, `HOTSPOTS.md`, and `REPORT.md`. Until those write/index/HEAD checks complete,
the validation pass is terminal evidence but not the final Git custody seal.
