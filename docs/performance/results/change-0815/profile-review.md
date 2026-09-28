# 0815 owner-scoped Callgrind review

This is an independent offline review of the four terminal owner-scoped
Callgrind publications for the large public PPTX capture. The profile lane
used CPU 12, the `profile` binaries, one sample, no warmup, `Ir` only, and the
exact owner `namespace_uri_probe::capture_region_0793` in the orders
`before/after` and `after/before`. The lane reports guest-instruction
attribution for this wrapper; it is not a latency, RSS, native-cycle,
phase-fraction, or adoption gate.

The receipts use schema `litchi.performance.0815.callgrind-receipt.v1`, all
four commands exited zero, and all four numbered publications have exactly
one corresponding empty termination publication. The before-profile binary
is bound to SHA-256
`245f5b515d64d41cc1f389a6d4e6e79ce8c6f7af38790e39ee5040dbbef2e38a`; the
after-profile binary is bound to
`ad565a40267a02331256bd49a8debcc0a38f09f8b88337d83713674fb965c142`.

## Exact owner eligibility

The retained parser accepts a publication only when the exact owner has one
positive incoming call from `namespace_uri_probe::run_one`, that call has
count one, and the owner has one immediate outgoing call to
`litchi_pptx::package::model::Package::opened_presentation_with_limits`.
Every row satisfies those conditions. The owner self `Ir` is 10 in every
row, its immediate-child inclusive `Ir` plus that self value reconstructs the
owner inclusive summary, and whole-function self `Ir` reconstructs the
publication summary. Nested inclusive rows are retained as diagnostics and
are not added to the total.

| order | leg | numbered publication SHA-256 | owner summary Ir | scanner self Ir | scanner inclusive Ir | scanner outgoing calls | scanner → `read_event_impl` inclusive Ir | `NamespaceResolver::push` self Ir |
| --- | --- | --- | ---: | ---: | ---: | ---: | ---: | ---: |
| 0 | before | `6274a30521fb157ff4087b7071f89aafa8b5db11d28bfb2a2b57898a902df597` | 531,559,983 | 16,724,054 | 313,977,875 | 646,384 | 104,471,566 | 13,019,166 |
| 0 | after | `24ed20b4113cb75d944dec57b5fc83755b5d3ef538b33bfd7e01d46e283d8285` | 535,229,436 | 16,713,423 | 317,662,247 | 928,892 | 104,497,428 | 13,019,166 |
| 1 | after | `7447b0fb5174e5e671add4e6b54e5213726ed46d8fa1ac75495b219cdbc7f9fe` | 535,272,233 | 16,713,423 | 317,664,178 | 928,892 | 104,482,358 | 13,019,166 |
| 1 | before | `a0a86ab3a333cacd092c2e9714a78dcb78ff500101bc96892385e0112d6df464` | 531,695,209 | 16,724,054 | 314,089,381 | 646,384 | 104,574,901 | 13,019,166 |

All four reports use the sealed 0813 fixture and produce the same 215,220-byte
output with SHA-256
`9c46542b763fc4bef63dfe4336cadd2bfba2b7e7b3f18a376c3924eb5643b3e9`. The
reader verified reopening, `semantic_check`, expected and actual text,
readback bytes and hash, and the semantic text hash for every report.

## Conservation and call-edge scope

The scanner has the same retained workload call counts in all four rows:
`Reader<R>::read_event_impl` is called 282,612 times over two scanner edges,
`inspect_element` is called 181,678 times, and
`NamespaceResolver::push` is called 181,678 times. The after profiles record
a direct `drop_in_place<core::result::Result<quick_xml::events::Event,
quick_xml::errors::Error>>` call family: it has 282,612 calls in each after
row, split into 282,508 calls with inclusive `Ir` 3,672,604 and 104 calls
with inclusive `Ir` 728. The before rows instead retain 104 direct
`drop_in_place<quick_xml::events::Event>` calls with inclusive `Ir` 416; no
direct `Result<Event, Error>` drop edge is present there. This new drop edge
accounts for the scanner outgoing-call count changing from 646,384 to
928,892. It is an observed call-graph difference, not evidence of less total
work or a causal native-cycle saving.

The scanner-to-
`quick_xml::reader::ns_reader::NsReader<R>::process_event` edge is absent in
both before rows and both after rows. The generic `process_event` symbol
nevertheless has four unrelated incoming callers in each publication, with
call counts 8, 345, 400, and 300. Its global presence is therefore not a
zero-call gate, and this lane makes no global `process_event` claim. The
resolver-push self `Ir` remains exactly 13,019,166 in all four rows.

The scanner self `Ir` falls by exactly 10,631 in both counterbalanced pairs
(16,724,054 before to 16,713,423 after), while scanner inclusive `Ir` rises
in both corresponding rows. The owner summaries also rise in both pairs.
These observations include the new result-drop work and do not turn the
static removal of arm-local copies into a total-work or performance claim.

## Review result and limits

**Profile publication and attribution review: PASS.** Owner eligibility,
numbered and termination dump structure, binary and report custody, semantic
output binding, and raw Callgrind conservation checks all pass. The retained
profile result is diagnostic evidence for this exact large-capture wrapper.

Callgrind `Ir` is Valgrind guest-instruction attribution, not hardware retired
instructions, elapsed latency, RSS, a CPU or phase fraction, or a causal cost
assigned to the borrowed bindings. There is one sample per leg in each of two
orders, and the owner contains the full capture wrapper. The observed profile
values cannot satisfy the public-workflow adoption threshold or replace the
native, allocation, semantic, independent-reader, and policy checks.
