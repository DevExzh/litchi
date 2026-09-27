# 0783 independent results review

This review recomputes the phase shares and paired timing ratios from the raw
reports under `native/`. It did not run `analyze.py`, Cargo, a native capture,
or a profiler, and it writes no other packet artifact.

## Raw evidence checks

`native/complete.json` reports 36 processes and `native/receipts.json` has 36
successful receipts: six alternating blocks × three shapes × two legs. The 36
JSON reports contain 1,080 measured samples (30 per process), including 540
phase samples. Every phase sample has all five phase metrics, and the integer
sum of those metrics equals `elapsed_ns` for all 540 samples. This is the
exact-sum check, rather than an assumption based on rounded percentages.

All 1,080 samples report `semantic_check=true`, `reopened=true`, and
`marker_matches=true`; each output identity exactly matches its readback byte
count and SHA-256. Source, semantic-text, and output identities are stable
within each of the tiny, medium, and large shapes. The counts and identities
come directly from `native/{0..5}-{tiny,medium,large}-{control,phases}.json`.

## Independent arithmetic

For each phase process, I calculated each sample's `phase_i / elapsed_ns`,
took the nearest-rank median of its 30 fractions, and then took the median of
the six process medians. Therefore the displayed shares are robust summaries
of per-sample fractions and do not have to sum to exactly 100%.

| shape | capture | stage edit | commit | apply | serialize | share sum |
| --- | ---: | ---: | ---: | ---: | ---: | ---: |
| tiny | 17.233% | 4.399% | 15.233% | 2.806% | 60.115% | 99.786% |
| medium | 23.908% | 4.840% | 14.778% | 2.749% | 53.454% | 99.729% |
| large | 69.051% | 2.671% | 4.290% | 0.631% | 23.320% | 99.963% |

The phase/control ratio is the phase process's nearest-rank p50 of 30
`elapsed_ns` values divided by the paired control process's p50. The six block
ratios are summarized by their median. The intervals below are a 10,000
resample bootstrap of that six-ratio median, using plan seed `783078`.

| shape | median phase/control ratio | median change | 95% bootstrap interval | six-block change range |
| --- | ---: | ---: | ---: | ---: |
| tiny | 0.990512 | −0.949% | [0.989173, 0.996081] | −1.123% to −0.305% |
| medium | 0.990573 | −0.943% | [0.989369, 0.992642] | −1.069% to −0.725% |
| large | 0.963472 | −3.653% | [0.958465, 0.964104] | −4.510% to −3.578% |

The downward phase/control ratios, especially the consistent large-shape
shift, do not establish negligible instrumentation perturbation, negative
clock overhead, or a production improvement. They can include code-generation
and measurement effects. The phase shares are therefore current-workflow
ranking evidence, not unperturbed CPU fractions and not an explanation of the
0780 before/after result. No cause is inferred from whole-process counters.

## Next bounded diagnostic

The large generated corpus is capture dominated (69.051%), with serialization
second (23.320%). The next operation-local CPU profile should target the
current `Package::opened_presentation` capture on the same 100 × 100 generated
corpus, with package ingress and readback kept outside the profiled region.
The public dispatch is
`crates/litchi-pptx/src/package/model.rs:243-268`; the operation delegates to
`capture_internal` at `crates/litchi-pptx/src/opened/model.rs:730-876`.

A feasible profile can use the existing release probe route on CPU 12 with
three warmups and a small fixed number of current-source measured captures,
then separate integer samples for these existing seams: root/catalog setup
(`opened/model.rs:738-747`), slide/MCE/notes-root capture
(`presentation/package.rs:65-135`), slide identity and relationship checks
(`opened/model.rs:755-817`), notes graph/index completion
(`opened/model.rs:818-827`), and owned clone/revision/digest/memo finalization
(`opened/model.rs:828-876`). Keep the profile sink fixed and numeric, avoid
per-part callbacks or a second fingerprint, and pair it with an unobserved
control build. This ranks the capture's operation-local work without claiming
that any one seam caused the historical regression or justifies a production
change.

The current phase split cannot locate where the 0780 before/after difference
entered. A historical attribution run would require its own matched plan and
both historical legs; it is not supplied by this current-source packet.
