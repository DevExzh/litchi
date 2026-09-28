# 0807 — current PPTX capture event processing

The fresh profile keeps the per-slide notes scanner as the main observed capture
path and identifies namespace-aware event processing as the largest sampled
leaf in both native repeats. The next experiment should isolate event handoff
work while preserving the existing parser and namespace checks. No production
change or speedup is adopted in this batch. The rejected 0806 iterator candidate
remains rejected.

The [packet](results/change-0807/README.md) retains the source review, frozen
plans, commands, raw profiles, disassembly and independent replay scripts.

## Scope and controls

Production is unchanged at `3624e73236`; all 9,196 source files match the 0806
before manifest. All 35 previously read architecture, goal and taxonomy inputs
are unchanged. The exact 0791 Rust probe is reused, retaining the original
0784 wrapper and 0780 fixture marker. Older packets supply exact fixture and
semantic identities only; historical timings are not pooled or compared.

The three warm, generated, in-memory shapes are tiny (3 slides × 4 shapes),
medium (12×8) and large (100×100). The timed wrapper contains only public
opened-presentation capture. Fixture setup, serialization, readback and retained
owner destruction remain outside the operation clock. This is supporting
PPTX category-15 evidence, not native Office, cold-source, remote/range,
concurrent or comprehensive CRUD coverage.

Root ran all builds and captures serially on the recorded EPYC 9R45 host,
pinning workloads to CPU 12. The ordinary control and profiling-wrapper
binaries share production source; the native stack diagnostic has a separate
frame-pointer build. No heavy offline replay ran during capture.

| Lane | Reports | Measured outputs | Purpose |
| --- | ---: | ---: | --- |
| Six alternating control/wrapper blocks × three shapes | 36 | 1,080 | Current operation latency and wrapper perturbation |
| Two ordered Callgrind passes × three shapes | 6 | 6 | Exact wrapper-scoped guest instruction work |
| Two large frame-pointer native profiles | 2 | 200 | Exact-owner sampled ancestry |

All **44 reports and 1,286 measured outputs** pass exact source, output,
reopened text and semantic checks. Native controls use 30 samples after three
warmups; Callgrind uses one sample without warmup. Native profiles use 100
samples after three warmups, `cycles:u` at 499 Hz. Sampled stacks include
warmups. None of these lanes measures a candidate optimization.

## Native wrapper perturbation

The table shows medians of six per-process nearest-rank p50 values. Ratios
are medians of paired block ratios, with 10,000 bootstrap resamples, seed
807080 and zero-based interval endpoints 250/9749.

| Shape | Control p50 ms | Wrapper p50 ms | Wrapper/control ratio | Bootstrap interval |
| --- | ---: | ---: | ---: | --- |
| tiny | 0.237846 | 0.237867 | 0.999729 | 0.995409–1.005036 |
| medium | 0.465063 | 0.462687 | 0.995492 | 0.989065–0.999238 |
| large | 19.502578 | 18.844866 | 0.967152 | 0.962525–0.972282 |

The large wrapper leg is 3.285% faster in this control comparison. This is
code-generation/wrapper perturbation, not a production improvement; it prevents
interchanging the profile binary with the ordinary baseline. All six >5%
process-spread flags remain visible: tiny/medium control p99 and whole-process
RSS for both legs of tiny/medium. No samples were excluded or retried.

Whole-process GNU-time RSS is retained as an accounting observation with the
0789/0790 limitations. No allocation, residency, peak-memory or I/O improvement
is claimed. Full p50/p95/p99, mean, RSS and spreads remain in
[native-analysis.json](results/change-0807/native-analysis.json).

## Guest instruction attribution

All six regions have exactly one positive numbered dump, one owner invocation
and an empty termination dump. Every region total equals the sum of all
function self costs and also owner self plus its immediate child cost. Nested
inclusive rows overlap and must not be added as independent shares.

| Shape | Pass 0 Ir | Pass 1 Ir |
| --- | ---: | ---: |
| tiny | 7,756,233 | 7,758,712 |
| medium | 13,725,991 | 13,726,162 |
| large | 549,233,425 | 549,297,169 |

Large pass 0 enters `scan_processed_xml` 104 times and `inspect_element`
181,678 times. The inspector calls the checked attribute iterator 152,410
times; its inclusive edge costs 37,176,356 guest instructions. The inspector's
own self cost is 29,165,311. The scanner calls `NsReader::process_event`
282,612 times at an inclusive edge cost of 46,146,741. These are local-work
observations, not native CPU-cycle fractions or predicted savings.

Software SHA is still the largest guest self row, while native samples name
hardware SHA. The guest ranking does not justify removing fingerprint checks.
The main-presentation notes loader has one incoming call and 3,027,999 total
self-plus-child guest instructions in pass 0; its duplicate presentation scan
is therefore a lower-priority hypothesis than the per-slide path.

## Native sampled ancestry

Both primary and independent raw readers require exactly one public
`Package::opened_presentation_with_limits` owner in each qualified stack.
Nested rows below overlap.

| Observation | Repeat 0 | Repeat 1 |
| --- | ---: | ---: |
| All process stacks | 3,164 | 3,169 |
| Exact capture-owner stacks | 1,036 | 1,033 |
| Notes scanner nested under owner | 896 | 898 |
| Element inspector nested under owner | 223 | 244 |
| Checked attribute iterator nested under owner | 107 | 99 |
| quick-xml attribute iterator nested under owner | 108 | 118 |
| Notes presentation-index loader nested under owner | 10 | 10 |
| Fingerprint nested under owner | 95 | 96 |
| Unresolved interior frame | 1 | 1 |

The largest sampled leaf is `NsReader::process_event`, with 189/190 stacks.
Other leaves include the scanner itself (157/136), hardware SHA (91/93), and
UTF-8 validation (64/67). These are observed samples, not instruction costs or
phase percentages. Each repeat has an unresolved interior frame, so the frozen
qualification explicitly forbids native phase-fraction claims. All unqualified
stacks, periods and compressed original perf data remain retained.

## Exploratory event handoff follow-up

After the planned captures completed, root retained the exact frame-pointer
binary's symbol table and disassembly for the leading event leaf. This is a
post-capture exploratory follow-up, not a changed measurement plan. The
function separately receives and returns `Result<Event>` and contains payload
copies, tag dispatch and the namespace-resolver call. An independent replayable
offset census maps all 189/190 sampled leaf IPs to its retained instructions.

Offsets `0x43`, `0xbe` and `0x118` have 54/67/45 observations in repeat 0 and
47/55/59 in repeat 1. The corresponding instructions occur around event
payload movement. Sampling skid, pipeline effects and the separate frame-pointer
build prevent assigning those observations a causal cost or predicted gain.
See [event-sample-offsets.json](results/change-0807/event-sample-offsets.json)
and [disassembly](results/change-0807/event-assembly.txt).

The [next-step review](results/change-0807/next-step-review.md) selects a
bounded hypothesis: let the notes scan consume the parser event
directly while delegating namespace push, pop and resolution to the same
quick-xml owner. It must preserve namespace-error priority, pending scope pops,
empty/end events, duplicate and malformed attributes, limits, root/name
precedence, and Transitional/Strict retry masking. The unchanged buffered
scanner oracle and opened-capture refusal tests are required before any fresh
paired workflow trial. No manual namespace-parser replacement or validation
bypass is justified by this evidence.

The [source review](results/change-0807/source-review.md) distinguishes the
mandatory root preflight from the optional full notes proof. Removing only the
proof moves required work into the notes fallback and may repeat MCE processing;
it does not inherently change refusal order or remove validation. Removing the
classification from both sites would weaken the contract. Name/presentation
pass fusion remains a separate hypothesis needing byte, retry and refusal
evidence outside this packet.

## Verification and limits

Fresh ordinary, wrapper and frame-pointer release builds and probe formatting
pass. Each build retains the same sixteen individual unused-helper warnings
and aggregate warning line as 0791. Production Rust is unchanged; no fresh
production-suite, fuzz, native Office or cross-platform test claim is made.

The independent reader review identified an incomplete quality-receipt check.
The final validator now requires its exact formatting command, source descriptor,
probe census and ordered timestamps alongside the successful exit and log.
All main analyzers, raw frame/function-row audits and the exploratory offset
census replay after cleanup. The three copied executable identities are recorded
before the owned target is removed. Final sealing covers the packet and all six
documents; staged and committed blob inventories are checked separately.

No public API, production dependency, unsafe policy, source limit, namespace
contract, publication behavior or retained cache changes. ADR ownership and
preservation constraints remain intact. OLE2/OOXML work continues, ODF remains
deferred, iWork is excluded, and the broader performance goal is incomplete.
