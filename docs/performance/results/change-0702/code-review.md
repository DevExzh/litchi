# Change 0702 code review — borrowed inherited namespace views after marker-search isolation

Review status: complete source/lifetime/assembly/semantic/resource/timing/gate
review; retention recommended with an explicit early-name first-pair tail
limitation. The candidate packet passed the required gates and independent
follow-up audit; the scoped recommendation below does not make a universal
performance claim.

## Scope reviewed

The baseline is the retained 0701 source at revision
`72d6f4500d906b9a5d6d4a0e571b0bb840ab1e68`. The candidate patch currently
recorded in `candidate-production.patch` and `source-diff.patch` changes only
`crates/litchi-ooxml-common/src/mce/codec.rs`. It carries the 0701 private
`#[inline(never)] contains_mce_namespace` marker helper unchanged and applies
the 0698 borrowed `Inherited` view to the newer processor.

The intended production diff is:

```rust
struct Inherited<'a> {
    ns: Option<&'a Arc<NamespaceLayer>>,
    emitted: Option<&'a Arc<NamespaceLayer>>,
}
```

The patch removes the temporary owner clones at the three `start` paths,
passes borrowed options to `hoists` and `for_each_hoisted`, and keeps the
owned result of `Inherited::after`. It also makes the existing mutable
AlternateContent bookkeeping produce an `alt` value before constructing the
borrowed view. The focused test source is unchanged, so the candidate and
baseline bind the same 104-test source witness. The candidate codec hash is
`a5b5b0aca3ec5a392bc7ae1ea6ca482653bd0a72cb9bdd4c30b8ea9faff87bee`; the
baseline focused codec hash is
`084e65b63c7d2597c262d15401a69442a6b0efc5d927290345899ac1354dbd6b`; the
shared test hash is
`14879f98747f30b78b4eca43a3657a9f74245e172a0f5b5cb71154b54f51df91`.

No Cargo or lockfile change, import or dependency addition, public item,
unsafe code, global state, cache, parser or frame representation rewrite,
streaming change, limit movement, or active-offset change is present in the
recorded patch. The static review is limited to the source currently frozen;
assembly has now been checked against the frozen native post-images below.

## Lifetime and ownership review

The source ordering satisfies the central lifetime condition. The candidate
does not build `Inherited` from `c.ns.head`, which is the child context whose
namespace head may be replaced by local declarations. Each view instead reads
the parent frame:

```rust
let parent = st.last();
let inherited = Inherited {
    ns: parent.and_then(|f| f.ctx.ns.head.as_ref()),
    emitted: parent.and_then(|f| f.emitted_ns.as_ref()),
};
```

The opaque and ordinary paths place this view and all uses of it in a block;
the block constructs a complete `Frame` before `close` receives a mutable
`st`. The AlternateContent path first stores its selection as `alt`, ending
the `st.last_mut()` borrow, then uses a separate block for the borrowed view
and owned frame. The candidate therefore cannot pass an `Inherited` reference
into `close`, which may reserve and push into the frame vector.

`Inherited::after` remains the ownership boundary. An emitted element clones
`ctx.ns.head` for the child frame. A skipped, unwrapped, or branch element
clones the parent's nearest emitted boundary with `self.emitted.cloned()`.
The candidate does not make `Frame.emitted_ns`, `Ctx`, `NamespaceLayer`, or
the directive layer borrowed. The temporary references cannot escape the
`start` call.

The candidate compiles with warnings denied, and both frozen focused receipts
report 104 passed, 0 failed, and 181 filtered tests. The source review finds
no semantic blocker in the lifetime rewrite. The compiler result confirms that
the explicit block boundaries are sufficient for the inferred lifetimes;
generated control flow and performance consequences are reviewed below.

## Error and behavior ordering

The candidate delays only the side-effect-free observation of parent owner
references. It retains this order:

- attribute decoding and XML syntax errors;
- parent `Ctx` clone and local namespace parsing, duplicate checks, binding
  limits, and `xmlns=""` handling;
- element QName expansion;
- opaque detection and validation;
- MCE directive parsing, target checks, and limits;
- local directive-layer construction;
- AlternateContent attribute and branch validation, including choice/fallback
  counts and selection/refusal precedence;
- compatibility filtering and hoisted declaration emission;
- multiple-root checking;
- owned `Frame` construction through `after`; and
- `close` and its empty-element or stack behavior.

The `alt` restructuring returns the same `(active, mode)` values as the old
branch, including `Mode::Skip` for an ignored AlternateContent child. It does
not move a fallible check after frame construction or suppress a report
increment. `write_start` still receives the same pre-local inherited scope
and still sees the same `Arc` pointer identities. The 0701 marker predicate,
input-before-output limits, marker-free borrow, parser route, and malformed
byte behavior are outside this diff and remain baseline behavior.

## Focused semantic coverage

The existing 104-test `mce::` source includes the three retained 0698
regression witnesses:

- `inherited_scope_survives_selected_alt_and_unwrapped_branch_rebindings`;
- `opaque_and_skipped_scopes_keep_the_nearest_emitted_namespace`; and
- `nested_context_errors_keep_first_refusal_precedence`.

Together with the surrounding MCE tests, they cover selected and skipped
branches, unwrapped and opaque descendants, empty/default namespace resets,
nearest emitted boundaries, prefix shadowing and hoisting, empty elements,
limits, and typed refusal order. The three 0701 marker tests remain in the
same source and cover the independent byte-dispatch boundary. Since this
candidate changes no test source, no additional focused test is requested by
the static review. The shared oracle must still compare exact output,
ownership, reports, and typed/debug errors for the full corpus.

## Resource and architecture review

The temporary view has constant-size references and retains no state after
`start`. The candidate adds no allocation, source owner, cache, lock,
execution context, worker, or budget dimension. It keeps the existing source,
output, depth, namespace, directive, and frame-stack bounds. The potential
resource effect is limited to transient `Arc` owner operations and generated
code shape; identical allocation counters would be expected but must be
measured rather than assumed.

The helper remains private to `litchi-ooxml-common` and does not expose archive,
physical package, relationship, snapshot, edit, patch, or facade state. No
accepted ADR boundary moves. The candidate uses the existing `Arc` and does
not require unsafe code or a new dependency.

## Static cost and risk assessment

The intended benefit is removal of two temporary `Arc::clone` increments and
their matching drops for each processed element view. Those operations do not
normally allocate, so this hypothesis must not be described as an allocation
or RSS optimization. The parent `Ctx` owner and the child frame's required
owner remain, as does the 0701 marker helper's scan cost.

The main risks are generated code, register pressure, helper/frame
reservation, and reachable marker-positive refusal paths. The prior 0698
candidate had a repeatable `early-name-error` regression; 0702 must rerun
that exact case and the complete refusal matrix on the 0701 baseline. The
candidate can also change code layout around `process_markup_compatibility`
and the outlined marker helper even though its source is untouched. No local
symbol or destructor change can establish an end-to-end gain by itself.

## Frozen assembly review

The frozen native post-images keep the processor and marker helper at the same
addresses and symbol sizes. The `start` image is smaller, while its local stack
reservation is larger:

| symbol or image metric | baseline | candidate | change |
| --- | ---: | ---: | ---: |
| `process_markup_compatibility` | 7,183 B at `0x170650` | 7,183 B at `0x170650` | 0 B |
| `contains_mce_namespace` | 170 B at `0x1705a0` | 170 B at `0x1705a0` | 0 B |
| `start` | 18,599 B at `0x173bf0` | 18,192 B at `0x173bf0` | −407 B |
| `start` local reservation | `0x598` | `0x618` | +128 B |
| native `.text` | 2,598,022 B | 2,597,494 B | −528 B |
| native `.data` | 60,872 B | 60,872 B | 0 B |
| native `.bss` | 1,488 B | 2,032 B | +544 B |

The baseline contains a 118-byte `drop_in_place<Inherited>` symbol; the
candidate has no such destructor. The 118-byte `drop_in_place<Ctx>` and
231-byte `drop_in_place<Frame>` cleanup symbols remain. This supports the
intended owner boundary, but the additional 128 bytes are a per-call frame
reservation and do not establish a whole-process stack or RSS result.

The source-mapped `start` disassembly has nine lock-increment and fourteen
lock-decrement instruction sites in the baseline, versus six and eight in the
candidate. These are static instruction-site counts, not dynamic per-element
owner-operation counts; shared control flow means the three fewer increment
sites and six fewer decrement sites cannot be read as a direct count of the
two logical temporary owner pairs.

The processor and helper have the same placement and size, but their complete
disassemblies are not byte-identical: the frozen comparison reports 119
processor lines and three helper lines with differences. The assembly result
therefore supports the source review without proving that all generated code
costs are unchanged; timing, counters, and the required follow-up decide
whether the stack and layout trade is acceptable.

## Initial frozen measurements

The exact oracle completed 1,920 invocations across 192 inputs and five
profiles with zero mismatches. The focused source receipts remain 104 passed,
zero failed, and 181 filtered on both builds. The initial native total medians
use the candidate `b0/a2` and `b1/a3` pairs against their matching baseline
legs:

| workflow / input | paired change |
| --- | ---: |
| one-real | −3.58% / −4.50% |
| one-control | −0.77% / −3.05% |
| one-generated | −0.95% / −1.87% |
| one-notes-poi | +1.88% / −2.77% |
| one-notes-lo | −1.78% / −3.27% |
| noop-real | −7.10% / −8.34% |
| noop-control | −3.93% / −4.15% |
| noop-generated | −1.48% / −2.74% |
| noop-notes-poi | −2.85% / −0.97% |
| noop-notes-lo | −1.74% / −3.31% |
| two-real | −6.66% / −5.27% |
| two-control | −3.30% / −1.73% |
| two-generated | +2.12% / −2.15% |

The old `early-name-error` refusal cost does not repeat: its paired medians
are −1.99% and −0.32%. The marked late refusal cases improve by roughly
3.08%–3.43% in both legs. The initial refusal review still has one material
new trigger: `slide-raw-overlimit-16m-to-64m` is +6.66% / +4.50% at p50
(about +10 us), with the first pair also +10.03% at p95. It must be resolved
by the long follow-up before retention; the follow-up result is recorded
below. The native trigger receipt also keeps
15 phase or tail observations visible; the largest are one-real clone p99
+41.99% (+7,580 ns) and two-generated total p99 +14.40% (+231,701 ns), so
the favorable total medians do not erase those diagnostics.

The marker controls show a small marked-input p50 increase of 4.65% (+20 ns)
in both candidate pairs. The first pair has mean +10.46% (+45.8 ns) and p99
+6.12% (+30 ns), while the second pair's mean is −7.28%. The root-comment
control improves by 6.38% and 4.26% at p50. The shared synthetic oracle
controls improve about 3.6%–11.0% at p50. Declaration controls improve in
each pair, and the real controls are near-flat for DOCX (+0.18% / +1.28%)
while XLSX and PPTX generally improve. The PPTX opaque-many second pair has
tail triggers of p95 +39.47% and p99 +9.77%, which remain in the retained
trigger record alongside the follow-up results.

Allocation comparisons are identical for all 78 process/phase pairs,
including call counts, requested bytes, reallocations, and peak live bytes.
The profile summary reports per-open capture cycles −3.96%, instructions
−0.13%, branch misses +0.98%, and cache misses +0.69%; whole-child peak RSS
is 5,560 KiB to 5,636 KiB (+76 KiB). These are diagnostics rather than an
allocation or RSS claim. All seven integration and six evidence gates pass;
the independent follow-up and its audit are recorded in the final disposition.

## Evidence protocol and audit record

The packet provides and independently audits:

1. Formatted baseline and candidate source maps and exact codec diff, with no
   test, Cargo, lockfile, dependency, or unrelated source change.
2. Passing 104-test focused receipts for both source states.
3. Exact 192-input × 5-profile oracle parity for bytes, lengths, `Cow`, full
   `Report`, and typed/debug errors.
4. Fresh marker controls, 13-workflow native timing, ten-case refusal timing,
   allocation comparisons, profile counters, binary sizes, processor
   assembly, and stack diagnostics.
5. All seven integration and six evidence gates, completed before the longer
   follow-up.
6. The independent six-case native and ten-case refusal follow-up with four
   ABBA legs, 300 samples, ten warmups, raw distributions, and a passing
   recomputing audit.

Every approximately greater-than-5% latency/throughput or peak-RSS trigger
must remain visible with its absolute duration and baseline drift. The review
must specifically decide whether the previous early-name refusal cost repeats,
whether real edit/no-op/two-target work has a useful net result, and whether
any generated-code or stack cost is justified. No geometric mean or favorable
control can erase an individual trigger.

## Conditional disposition

The candidate is eligible for measurement after the static source pass. It is
eligible for retention only if exact parity holds and the complete evidence
shows a credible representative end-to-end benefit without a recurring
refusal or real-workflow regression beyond the review threshold. If the old
early-name regression repeats, if another material refusal cost appears, if
resource/code growth lacks compensation, or if any gate fails, production
must remain the 0701 codec. The measured candidate witness and semantic tests
may be retained as evidence, but no production performance claim follows.

The conditional rule above is applied to the completed evidence in the final
section below.

## Final disposition

The required evidence is admissible. The focused receipts pass 104 tests on
both source states, the exact oracle reports zero mismatches across 1,920
invocations, all seven integration and six evidence gates pass, the marker
control audit passes, and the independent follow-up audit passes 28 native
rows, 40 refusal rows, and 156 comparisons (8,400 native and 12,000 refusal
samples). The pre-cleanup packet audit also passes its source, probe, build,
assembly, raw-matrix, deterministic-summary, and cleanup checks.

The longer native follow-up confirms a useful but scoped common-work signal:

| workflow / input | follow-up paired total p50 change |
| --- | ---: |
| one-real | −3.718% / −4.770% |
| noop-real | −4.431% / −4.290% |
| two-real | −3.743% / −1.154% |
| two-generated | −1.568% / −1.235% |
| one-control | +1.323% / −5.210% |
| one-generated | −0.451% / −0.127% |
| one-notes-poi | −2.016% / −3.118% |

The native follow-up has four clone p99 triggers: one-generated first pair
+24.98% (+3,890 ns), two-generated first and second pairs +50.87%
(+6,710 ns) and +14.78% (+1,890 ns), and two-real second pair +14.47%
(+2,790 ns). The complete 24-row trigger receipt retains those extrema and
all other phase observations. The total p50 real-work results remain below
the review threshold in both pairs for one-real, noop-real, and two-real;
the two-real second pair is a smaller −1.154% gain.

The old refusal blocker does not recur as a median regression. In the
independent follow-up, `early-name-error` is +0.752% / −0.214% at p50 and
the over-limit case is +0.420% / −0.822%. The MCE-positive late-root and
late-missing-relationship cases improve in both pairs. One first candidate
early-name leg still has a material tail trigger: mean +5.135% (+963 ns),
p95 +55.05% (+10.35 µs), and p99 +15.40% (+3.92 µs). Its raw distribution
contains 27 of 300 samples at or above 25 µs, clustered in the first 25
timed samples plus two later samples; the other candidate leg and both
baseline legs have only 3–6 such samples. This review records the tail as an
observed leg-specific cost, makes no attribution to noise or warmup, and does
not discard those samples. The required four-leg follow-up is sufficient to
close this experiment; a future change to this refusal path should repeat
this exact case and preserve the trigger as a guard.

The candidate has identical allocation counters and requested bytes across
all 78 allocation comparisons. Its native text is 528 bytes smaller and its
`start` reservation is 128 bytes larger; `.bss` grows by 544 bytes, while the
profile diagnostic peak RSS grows by 76 KiB. These costs are visible and
bounded, and the source retains no allocation, cache, or owner beyond the
existing frame boundary.

**Recommendation: retain the candidate codec change with the stated scope.**
The borrowed view passes exact semantic and repository evidence, improves the
three common real-work totals in both paired follow-up legs, clears the prior
early-name median blocker and the initial over-limit trigger, and leaves no
allocation-count gain to overstate. Retention is limited to this private MCE
codec representation, the measured host and workflow mix, and the recorded
first-pair early-name tail. It makes no claim of universal parser, Office,
RSS, cold-cache, or parallel-throughput improvement.
