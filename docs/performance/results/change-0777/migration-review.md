# Current-source migration audit

Independent read-only audit of production source
488453934567165e39eae5adf1134a863217e26e found no migration blocker.

The historical survey has 581 sites in 15 crates. The 16th lint-enforced
crate, litchi-imgconv, has no XML attribute sites. Classification is 546
fail-fast sites (545 checked adapters and one dynamic exception), 26 unchecked
sites (25 migrated plus MCE streaming), four first-wins sites, two bounded
count sites and three helper sites. Current production has all 545 checked
migrations with matching per-crate counts; four additional checked calls are
adapter tests.

The reviewer independently inspected the 26 exceptional checked callers. Each
stops at the first iterator error through `?`, an explicit return, `all`, or a
single `next`. None continues through errors using flatten/filter_map. The
four lenient migrations retain their first-wins policy; the two count sites
retain bounded count_up_to, including the count reused for reservation.

Remaining production raw attributes calls are in the five bounded adapter
copies, the formula owner's private first-wins copy, and XLSX selected
plain_attributes. The latter dynamically selects checking and exits by the
second attribute, so it cannot reach the hash-filter path. There are 21 scoped
clippy::disallowed_methods allowances: 17 production helpers/dynamic scopes
and four test scopes. No crate-wide or module-wide allowance remains.

Raw checked/unchecked calls in DOCX transaction, PPTX notes and XLSX compact
code are test-only oracle/equivalence code. This source audit complements the
runtime equivalence and reader differential; it does not turn their fixture
coverage into exhaustive validation of every possible XML document.
