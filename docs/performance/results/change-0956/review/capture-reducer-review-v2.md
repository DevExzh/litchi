# 0956 v3 capture/reducer prospective review

Reviewed the frozen v3 tools and draft plan without starting a build, capture, reducer workload, or benchmark.

- `tools/capture_0956_v3.py`: SHA-256 `26501b64194cdfca85b6178d37f6a603bcc2802dcfe80501c27c174913aad5b9`, 29,886 bytes
- `tools/reduce_0956_v3.py`: SHA-256 `2948909058a8d65a4ac8f691b3b2fc3063829ec18338eab9da0bddb612d6192c`, 20,953 bytes
- `tools/controls_0956_v3.py`: SHA-256 `c21fb8e5271ce7a304b7ee9fff40920c94d72fc60ec6cc5a5159f99355013fc8`, 9,458 bytes
- `plan-v3-draft.json`: SHA-256 `0bd934ff12ad92755ad4c482cefd995436d9578d7ece4cf5b08066c37d3a3d72`, 64,394 bytes
- `review/tooling-receipt-v5.json`: SHA-256 `cb133a3eac8deb8731db79bfa0ee718c700de8b84ed825823c08817aeef8fb20`

## v2 blocker closure

The v3 reducer now closes the earlier child identity gap. For every expected row it checks the exact row object and index, derives and checks the expected argv/cwd/environment/deadline/output paths, checks requested and observed affinity, and requires the validation row to match. It also reads `capture-start.json` and `capture-terminal.json` and binds their plan, runner, reducer, controls, and source-manifest descriptors. The nine report-contract checks include the exact timing/setup/verification/performance-claim strings, sink capacity, and sink maximum-write bound. The two synthetic custody controls reject a changed schedule row and a changed post-terminal descriptor.

The six-block estimator remains correct. ABBA and BAAB blocks form adjacent A/B pairs at positions 0/1 and 2/3, treatment/serial orientation is explicit, and the six block geometric means are the bootstrap unit. Nearest-rank p95 and p99 over 200 samples select the 190th and 198th ordered samples respectively. Zero denominators remain fail-closed. No formal workload has started.

## Remaining custody condition

The v3 checks are internally consistent but are not, by themselves, a tamper-evident packet manifest. `capture_0956_v3.py` seals each `terminal.json` before parsing, then creates `post-terminal-artifacts.json`; the outer terminal records only the outer start/plan/runner descriptors. The reducer recomputes current hashes and compares them to the current post/validation records, but no immutable capture-time record anchors each child terminal, post-terminal list, stdout, stderr, time file, report, or validation file. A coordinated edit to (for example) `time-v.txt` and its post-terminal descriptor would therefore remain self-consistent and could change the process metric.

Before relying on formal ratios, the root-side final custody step must produce and retain an immutable manifest of the completed packet (outer records plus every child start, terminal, validation, post-terminal list, stdout, stderr, time file, and report), and the reduction receipt must bind to that manifest. This can be an external wrapper/receipt; it need not require another runner revision. The v3 controls currently exercise one-file mutations, not coordinated file-plus-descriptor mutation.

The v3 capture imports the legacy v2 validator, and read-only reduction intentionally does not require `status: authorized` or rehash the binary. The final root preflight must therefore bind the transitive `capture_0956_v2.py` descriptor, the authorized plan status, and the binary/source/helper descriptors before launch and again at reduction. The v5 receipt explicitly says root owns this condition.

## Ratio-field interpretation

The v3 reducer labels RSS as `rss_review_metric` and `rss_is_review_only`, and the result remains `descriptive_review_only`, which correctly prevents an allocation/adoption claim. However, `aggregate_point_over_1_05` still computes `any(publication_point, total_point, rss_point)`, so an RSS ratio can set that aggregate boolean. Keep this as an explicitly diagnostic aggregate or split it into separate latency and RSS flags in the final interpretation; do not treat it as the p50/p95 latency gate. The plan's declared p95 tail threshold and p99 review-only status remain clear, and process metrics are separately labeled whole-process `/usr/bin/time -v` measurements.

Conditional disposition: the v3 runner/reducer is ready for formal launch once the root-side transitive helper/authorized-plan binding and immutable completed-packet manifest are in place. No performance or adoption conclusion is supported by this review.
