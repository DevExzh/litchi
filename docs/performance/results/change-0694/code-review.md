# Change 0694 code review — lazy expanded MCE element names

The reviewed candidate is the change from baseline `829bed696` in
`crates/litchi-ooxml-common/src/mce/codec.rs`, with its focused MCE tests in
`crates/litchi-ooxml-common/src/mce/tests.rs`. I read the 0694 design packet,
`docs/GOAL.md`, and the accepted API, snapshot, performance, validation,
topology, and migration ADRs. This review is read-only with respect to Rust
production code and does not rely on a Cargo or performance run performed by
the reviewer.

## Review result

No production correctness blocker was found. The candidate preserves the
existing expanded-name resolution and validation decisions while avoiding the
owned `Name` on the common path where no extension or ignorable-element lookup
needs it. The focused source tests cover the new extension and preservation
branches, and the parent reports the corrected 85-case focused MCE run as
passing.

## What is sound in the candidate

* `expand_parts(q, &c.ns, true)` uses the same resolver and the same position
  as the old `expand`. The local name is borrowed from the current
  `BytesStart`; the namespace is borrowed from the cloned, local namespace
  context. Both are consumed before the event buffer is cleared, and neither
  borrow is stored in a `Frame`, `Ctx`, report, or output. The optional `Name`
  owns its strings when an owned value is actually required.

* QName, namespace, directive, `MustUnderstand`, `AlternateContent`, and
  resource checks remain in their existing order. Resolving the element QName
  still precedes the inherited-opaque early return, so malformed or unbound
  element names retain their old typed refusal even when their parent is
  opaque. The only moved operation is an infallible temporary `String` pair;
  moving it after directive validation cannot replace a typed validation error.

* Extension lookup retains the exact `HashSet<Name>` equality semantics. The
  `then` closure allocates a `Name` only when `caps.extensions` is nonempty;
  it is not the eager `then_some` form. `contains` therefore sees the same
  expanded namespace and local-name values as before. With an empty extension
  set, the only remaining `Name` construction is the lazy
  `get_or_insert_with` in the ignorable-element branch, where preservation or
  process-content matching genuinely needs the owned key.

* Inherited opaque elements still validate their own QName and then use the
  existing unfiltered writer and `Mode::Emit` path. Unselected or inactive
  alternate-content branches still go through the same current-element
  validation and selection logic; the borrowed slices only replace string
  storage and do not change the active flag, report counters, or frame mode.

* The added tests exercise prefixed and default-namespace extension names,
  matching and mismatching extension profiles, inherited opaque descendants
  containing MCE markup, namespace rebinding for preservation directives, and
  QName-versus-directive refusal precedence. These cases directly cover the
  lifetime, HashSet, opaque, and error-ordering seams introduced by the diff.

## Findings and coverage disposition

There is no required production fix from this review. One small focused test
gap identified during the initial review combined an inherited opaque parent
with an invalid descendant QName while an extension profile was enabled. The
frozen independent oracle now closes that gap with its synthetic opaque-invalid-
QName case: all five capability profiles preserve the exact typed
`NonConformant` refusal for `bad:too:many`.

The extension-profile tests also provide the useful negative control that a
nonmatching, nonempty extension set retains baseline MCE selection. No
additional API, allocation-failure, or unsafe-code requirement follows from
this private ownership change.

## Secondary capability-profile controls

The supplemental oracle control matrix contains 10,800 source-bound samples:
ordinary and one-opaque-subtree XML, each under the empty profile, one explicit
opaque name, and 4,096 explicit names, with four baseline legs and two
candidate legs. All profile/output identities match between binaries. With the empty
profile, the ordinary input improves by 7.20% and 7.42% at p50. When the
explicit profile matches the opaque subtree, p50 improves by 9.31% and 9.71%
for one name, and by 8.37% and 6.12% for the 4,096-name stress profile.

The bounded cost is the ordinary input under a nonmatching explicit profile:
one explicit name changes p50 by +2.39% and +6.11% (about 6--15 microseconds
on this 14 KiB, 1,000-element input), while 4,096 names changes it by +3.27%
and +5.36%. The large set is a stress control rather than a production
configuration. The current production call sites that build nonempty sets use
one name in conditional formatting, data validation, and sheet protection,
two names in calculation properties, and two names in PPTX media parsing.
Those owners either target a matching extension subtree or run on a focused
metadata/media part; the XLSX worksheet consumers are the relevant case where
a large ordinary sheet could repeatedly pay the nonmatching lookup cost.

Disposition: retain the candidate and treat the nonmatching explicit-profile
cost as a path-scoped tradeoff. The default profile and matching opaque paths
show the intended allocation removal, and the synthetic +6.11% trigger alone
does not outweigh the 4--5% end-to-end PPTX gain. The real-part
controls below bound that choice. The 4,096-name stress profile remains a
stress control rather than a reason to gate the default-path optimization.
The measured custom-profile cost is recorded as a path-scoped caveat; a
follow-up heterogeneous borrowed lookup for `HashSet<Name>` can address it if
future product measurements show that path is material.

## Real-part consumer controls

The follow-up oracle covers 16,200 source-bound samples from three marker-
bearing parts: DOCX `word/numbering.xml` (56,460 bytes), XLSX
`xl/worksheets/sheet1.xml` (19,886 bytes), and PPTX `ppt/slides/slide1.xml`
(2,906 bytes). Each part was run under the baseline, one-name, and
4,096-name profiles with 300 samples on each of four baseline and two
candidate legs. The source, output, typed report, and exit-status identity
receipts agree across the paired binaries.

For the default profile, the two paired p50 changes are DOCX -0.36% and
-2.56%, XLSX -1.76% and -4.37%, and PPTX -7.35% and -7.16%. The explicit
nonmatching profiles range from -1.81% to +3.97% at p50 across these real
parts; the one-name profile ranges from +1.06% to +3.80%. Thus the selected
real XLSX worksheet does not reproduce the synthetic >5% p50 trigger, and the
controls do not justify a separate full consumer feature probe for this
change.

One tail receipt needs to stay attached to the scope: the first PPTX opaque
one-name pair has mean +39.97% and p99 +15.54%, because one candidate sample
is 3,072,415 ns against that leg's 26,320 ns median. The second pair is mean
+4.25% and p99 +2.36%. The raw samples are retained in the control artifact;
they support a tail diagnostic and no host-cause attribution. Taken with the
default-profile PPTX gain and the allocation reduction, this supports
retaining the 34-line candidate while stating that its improvement is scoped
to the measured default and matching paths rather than universal.

## Measured evidence

The source-bound native probe covers 13 one/no-op/two-edit cases with four
baseline legs and two candidate legs, 100 samples and five warmups per case
and leg. The paired total p50 changes for the real corpus are -4.12% and
-5.26% for one edit, -2.58% and -3.93% for no-op, and -3.78% and -3.43% for
two edits. The marker control, generated corpus, and both notes-bearing
corpora remain within the recorded total-phase noise range; no total-phase
metric crosses the packet's 5% regression trigger. The complete phase trigger
list remains in `native-review-triggers.json`: the largest median flag is the
no-op real clone phase at +6.44% (about 1.0 microsecond), while several p99
flags are on sub-millisecond clone/apply phases. These are retained as phase
diagnostics and are not converted into an unqualified end-to-end claim.

The one-real total allocation diagnostic changes from 190,338 to 143,848
allocation calls and from 13,815,905 to 12,430,874 requested bytes; realloc
calls remain 4,870. Peak-above-start remains 459,170 bytes and net live change
remains 195,730 bytes. The native prefix profile reports instructions down
3.28% and cycles down 2.57% per open/capture; branch misses rise 3.95%, page
faults rise 10.41%, and whole-child peak RSS moves from 5,696 to 5,700 KiB.
These measurements support the scoped implementation review and preserve the
packet's `performance_claim: none`; they do not establish a universal speed,
RSS, or allocation guarantee.

The refusal comparison has one retained >5% phase-tail trigger: generated
12x8 valid p99 is +17.49% (about 112 microseconds), while the candidate and
baseline typed outcomes and source bindings are checked by the refusal
receipts. The independent common-MCE oracle covers 192 cases (36 real XML,
144 deterministic mutations, and 12 synthetic cases) across five profiles,
for 1,920 invocations with zero output, report, error, or ownership
mismatches. Final repository gates remain parent-owned release checks.
