# Next investigation — avoid empty namespace installation

performance_claim: none

The next candidate is the unconditional `c.ns = c.ns.with_local(...)` in
`crates/litchi-ooxml-common/src/mce/codec.rs`. When the local declaration list
is empty, `with_local` immediately returns `Ok(self.clone())`; assignment drops
the previous equivalent owner. Test an empty-list guard around this call while
leaving every nonempty declaration, QName, directive and resource check in its
current order. No production edit is part of 0695.

The current isolated 18-call sequence is 2.512–2.538 ms. Its five presentation
calls separately take 83.46–84.87 µs (3.33% of the full median), while slide 11
alone takes 1.082–1.094 ms. Counter repetitions corroborate presentation's low
weight. Therefore removing a repeated presentation pass is a lower priority
than shared element work on this corpus; these are not additive capture
fractions or achievable speedups.

The fresh sampling profile assigns 32.22% self to `start`, 6.46% to `Inherited`
destruction and 3.27% to `Ctx` destruction. Inline namespace clone/drop frames
are investigation signals only: DWARF attribution is imperfect, percentages
are not additive savings, and this batch does not verify the empty branch's
exact machine instructions. Inspect the native assembly before attributing a
candidate gain to eliminated atomics.

`count-elements.py` provides an independent Expat source-syntax census of
10,072 starts across the weighted 18-call input sequence; 23 starts carry one
or more namespace declarations. This is not production branch instrumentation
or a claim about active MCE output. It shows that empty local declaration lists
are common enough to measure. The exact per-member rows are retained.

The five presentation passes arise from initial root validation, initial slide
references, the fresh view's catalog, notes root conformance and the notes graph
scan. Relevant owners are `opened/model.rs::capture_internal`,
`presentation/package.rs::capture_slides`, `parts/presentation.rs`, and
`notes/package.rs::load_snapshot` through `notes/codec.rs`. Their order and
refusal precedence remain binding if reuse is revisited.

Run a single-change A/B experiment with the current native PPTX one/no-op/two
edit probe and marker/generated/notes controls. Supplement with these isolated
member/group cases and exact current MCE output, ownership, full Report and
refusal parity under default and custom capability/limit profiles. Preserve
namespace rebinding, unbound/invalid QNames, ignored branches, opaque descendants,
hoisting, finite namespace limits and early errors. Measure allocations,
native counters, tails and code size; retain only a useful end-to-end result.

Broader changes to frame ownership or attribute normalization have larger proof
obligations and should not be bundled into the empty-list experiment. No global
state, retained cache, public API or resource-limit change is needed for the
small candidate.
