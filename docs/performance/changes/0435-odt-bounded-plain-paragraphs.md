# 0435: bounded fresh ODT plain paragraphs

The new ODT API consumes ordered plain paragraphs once and publishes to a
caller-owned sequential sink with an explicit paragraph XML window. It avoids
a document-sized paragraph model, content XML buffer, and output archive vector.
This is an optional bounded-memory creation path with a substantial measured
CPU cost; the existing Builder remains available.

The common writer adds typed authored-XML slice publication for the exact
Builder styles/meta defaults, including their comment, and typed generated-XML
limit attribution through reader adapters. ODT owns paragraph grammar,
whitespace encoding, and explicit CR refusal. Paragraph/text/XML/output limits,
hierarchical execution budgets, cancellation, nested producer/sink errors, and
acknowledged partial output are tested. No dependency or unsafe code was added.
The modeled provider Memory reservation excludes ZIP/auditor allocator peaks.
Provider retention is bounded for a lazy source; caller-owned iterator and sink
storage remain outside that claim. The measured source generates one String at
a time and the sink discards acknowledged output.

Baseline: `7c25f299d19d9e02ed1fe7c6b36a126165b474fc`.
Candidate: `6fc3179646bf812e76845636d9cbec5bf30b278b`.
The [bundle](../results/change-0435/README.md) retains 36 formal reports,
1,080 samples, and six whole-process profiles on AMD EPYC 9R45, Linux
7.0.0-1011-aws, Rust 1.98.1, CPU 2, one worker, system allocator, release
builds with frame pointers and debug level 1. Each report has 30 samples and
three warmups; two repeats run the roles forward and then in reverse.

## Measured API tradeoff

Normal p50 milliseconds, R1 / R2, compare the candidate's two APIs:

| Paragraphs | Buffered | Streaming | Streaming / buffered |
| ---: | ---: | ---: | ---: |
| 64 | 0.091695 / 0.091000 | 0.169836 / 0.168446 | 1.852 / 1.851 |
| 8,192 | 4.147957 / 4.624400 | 13.664514 / 13.606935 | 3.294 / 2.942 |
| 32,768 | 16.343746 / 16.546601 | 54.773431 / 55.738948 | 3.351 / 3.369 |

Allocator region peak above operation entry, in bytes, is identical across
all 30 samples in each repeat:

| Paragraphs | Buffered | Streaming | Change |
| ---: | ---: | ---: | ---: |
| 64 | 458,819 | 420,091 | −8.441% |
| 8,192 | 5,616,425 | 420,091 | −92.520% |
| 32,768 | 22,450,985 | 420,091 | −98.129% |

At 32,768 paragraphs, allocation calls fall from 450,741 to 204,977 and
requested allocation bytes from 51,449,840 to 7,121,306 (−86.159%). All
allocator operation live-byte deltas are zero. The measured XML window is
4,096 bytes; the 420,091-byte allocator peak is a separate observation.
Whole-process GNU time RSS stays between 84,537,344 and 84,729,856 bytes
across the entire matrix. This batch demonstrates no RSS reduction or bounded
whole-process memory claim.

All 38 flags among 156 matched metric comparisons remain in `summary.json`.
Every streaming latency/paragraph-throughput comparison triggers review.
The same-API buffered control also has a medium R2 p50 regression of +7.983%,
with p95/p99/mean/throughput flags, and a tiny R1 reported process high-water
RSS flag of +5.201%. Streaming has two additional reported process high-water
RSS flags: normal tiny R1 +5.408% and allocator tiny R2 +6.177%. These are
separate from GNU time whole-process maximum RSS. Ten repeat-drift flags cover
baseline large latency/throughput and candidate buffered medium latency/throughput.
No average conceals these results. Confidence intervals resample within-process
samples; two process repeats do not establish a general speedup.

## Profiles and decision

The streaming whole-process record attributes 45.74% self samples to
`ExecutionContext::consume`, 10.28% to SHA-256, and 6.39% to the XML auditor.
Whole-process instructions are 14.794 / 14.825 / 35.991 billion for
before-buffered / after-buffered / after-streaming. These profiles include
setup, warmups, hashing, and the untimed oracle, so they do not isolate a causal
operation-only cost. All three records report zero lost samples and retain
6 / 7 / 8 addr2line warnings. L1 readings are zero observations; LLC was not
collected. See `formal-profile-summary.json` for every counter and its scope.

Retain the optional API as a measured bounded sequential-publication enabler:
large-case allocator peak and requested bytes are materially lower. Its CPU
tradeoff is explicit. The next measured hypothesis is bounded ordinary-text
Work batching within ODT, preserving scalar cancellation and exact limit refusal,
using this committed streaming selector as the baseline. ODP fresh creation
follows that review. There is no blanket latency, scaling, physical-copy, native
compatibility, or program-wide 10x claim.

## Validation and remaining scope

The final ODT/common release suite passes 1,445 tests with one existing ignored
test; the standalone harness passes 315 with one existing ignored test. Scoped
production Clippy, warnings-denied documentation, owned-file format, and crate
boundaries pass. The pre-existing common large-enum allowance remains explicit.
Harness strict Clippy retains exactly 29 existing diagnostics, with none in the
new helper or changed harness lines. Failed attempts remain in the bundle.

Per-report and outer verification require paragraph semantic equality, exact
styles/meta, five-member topology, manifest bindings, and deterministic output
within each role. Buffered before/after archive and content bytes also match.
Cross-API content XML and ZIP framing may differ. Eight preparatory oracle
mutations passed. Copied verification and eight mutation probes pass both
before and after cleanup. Five task scratch directories totaling 1,823,520,117
bytes were removed; both shared targets and the user goal were preserved. A native runtime was unavailable; fixture tests are a separate evidence
class. Fresh plaintext creation does not close existing-document append, Part
addition, repackaging, rich authoring, native breadth, cold/range I/O, or scaling.
The original non-iWork goal remains open.
