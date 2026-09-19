# 0696 source and mechanism review

performance_claim: none

Independent source review found no correctness blocker for the codec-only guard.
The empty helper branch cannot return an error and preserves the same namespace
head/binding count. Skipping its clone-and-replace keeps inherited pointer
identity and hoisting boundaries. Nonempty lists, including `xmlns=""`, still
perform all existing namespace syntax, duplicate, binding-limit and QName checks.
The streaming implementation is distinct and remains unchanged.

Baseline assembly confirms the redundant owner operations for `Some(head)`:

- `0x17a274–0x17a277`: inlined `with_local` tests the declaration count and
  branches to `0x17a43d` when empty.
- `0x17a446`: `lock incq (%r14)` clones the inherited namespace head.
- `0x17a6f7`: `lock decq (%rax)` releases the previous equivalent `c.ns` head.
- `0x17a708`: the returned head and unchanged count are installed.

Addresses refer only to the native baseline binary bound by
`assembly/baseline.json`. The root `None` branch has no atomic refcount pair.
Other scope/inherited/frame clones remain; static assembly does not establish
dynamic frequency or an achievable end-to-end gain.

Three new behavioral tests cover inherited scopes plus empty-default reset,
exact finite namespace limits, and same-element declarations with QName/directive
error precedence. The first focused run exposed a test-bound assumption: the
implicit `xml` binding counts as one, so two declarations require a ceiling of
three. The corrected test checks ceilings three and two. The retained initial
log contains that failure; the final focused run passes all 88 MCE tests.

## Candidate assembly review

The frozen candidate (`assembly/candidate.json`, binary SHA-256
`ed34f8e4cfb72ca50a80481c7438abb6a1ff4917b10440d672c309b40483b9e1`) checks
the local declaration count at `codec.rs:781`: `0x17875b` tests the `Vec` length
and jumps to `0x17900c` for the empty case. That target sets the existing result
state and joins the common continuation; it does not call `with_local` or install
a replacement `Namespaces` value. The nonempty arm still calls
`Namespaces::with_local` at `0x1787bb`, and the existing result installation
continues through the old `c.ns` drop at `0x1787ef–0x1787f9` and stores the new
head/count at `0x17880d–0x178812`.

This removes the baseline's empty `Some(head)` pair: baseline tests the count at
`0x17a274–0x17a277`, clones the inherited head with `lock incq` at `0x17a446`,
then drops the equivalent old field with `lock decq` at `0x17a6f7`. No matching
pair occurs on the candidate start's empty branch. The standalone candidate
helper still contains its old empty clone path (`0x1707f5`, then
`0x17096a–0x17096f`); this is unreachable from the guarded start call and is
expected because `with_local` was left unchanged. Other Arc operations in
`start` remain outside this narrow branch.

The start symbol is 18,083 bytes in the candidate versus 19,108 bytes in the
baseline, while the candidate also emits a separate 1,313-byte `with_local`
symbol that was inlined into the baseline start. Thus the relevant symbol-size
sum is about 288 bytes larger in the candidate; this is a static codegen cost,
not a runtime-size or speed claim. Both start functions reserve the same
`0x598` bytes (`sub $0x598,%rsp`). The outlined helper has six saved registers
plus `sub $0x88` (0xb8 bytes before the call's return address), so the empty
path avoids that call/frame and the nonempty path retains it. These sizes and
addresses are bound to the frozen native binaries; they do not establish dynamic
frequency or end-to-end performance.

## Final disposition

Retain the guard. The declaration-control follow-up in
`oracle-declaration-control-comparisons.json` covers 24 comparison rows and 96
percentile/mean deltas across the declared 1,000-sibling redeclaration case and
the mixed inherited-child case. For the declaration-heavy pairs, median
regressions range from +2.55% to +3.86% across the three profiles; the largest
positive value in the full record is +4.09% (opaque-many, `b1/a3` mean), below
the +5% review trigger. The mixed inherited-child pairs improve by 1.87% to
3.18% on median. All 36 recorded oracle runs exit successfully, with repeated
source/output identities and matching semantic reports for each case/profile.
The seven integration checks and six evidence/repository checks also all record
exit code zero.

The final native record reports the representative real edit 6.50%–6.72%
faster, exact allocation comparisons, and no total metric above the +5%
trigger; existing real and synthetic secondary controls improve 2.15%–7.52%.
Eleven phase-tail triggers remain, with the largest at +35.88 microseconds.
Those tails and the outlined helper's modest nonempty-path cost are retained as
documented limitations. The consistent real-path gain and the absence of a
threshold breach in declaration-heavy controls support retaining this
codec-only guard.
