# 0813 owner-scoped Callgrind review

This is an independent review of the four terminal owner-scoped Callgrind
publications for the large public PPTX capture. The frozen profile lane used
CPU 12, the `profile` binaries, one sample, no warmup, `Ir` only, and the
exact owner `namespace_uri_probe::capture_region_0793` in the orders
`before/after` and `after/before`. The review makes no latency, RSS, native
cycle, phase-fraction, or adoption claim.

The offline reader passed both commands:

```text
python3 -B docs/performance/results/change-0813/profile_analysis.py --write
python3 -B docs/performance/results/change-0813/profile_analysis.py --check
0813 Callgrind analysis PASS
```

The reader verified four positive numbered publications and four empty
termination publications. Each positive publication uses the exact owner,
has one incoming owner call, has a positive `Ir` summary, and has one
numbered dump. The owner self `Ir` plus its immediate-child inclusive `Ir`
reconstructs the publication summary; summing whole-function self `Ir` also
reconstructs it. Nested inclusive rows are retained as diagnostics and are
not added to the total.

## Exact publications

The profile binaries are bound to the build receipts: before-profile SHA-256
`18b18523b28f39cfd0bb13516bf96fe0705e34b8b811f3ca53e35a71c54185bc` and
after-profile SHA-256
`8b2fa9c6677382a4991ac8d91d15300bc0bb361806cee844e7f04487749ff1fa`.
The four positive numbered raw publication hashes and the key attribution
fields are:

| order | leg | numbered publication SHA-256 | owner summary Ir | scanner self Ir | scanner → `read_event_impl` inclusive Ir | `NamespaceResolver::push` self Ir |
| --- | --- | --- | ---: | ---: | ---: | ---: |
| 0 | before | `0379d1a427283ca7112980fedcc298f6289e12c841189149d414264bc3aedb32` | 537,151,032 | 22,295,758 | 104,437,752 | 13,019,166 |
| 0 | after | `49140269ef8ed85da2691057bb8700c7a8732dd4aba6e630006119d3ebb452e6` | 531,661,794 | 16,724,054 | 104,478,576 | 13,019,166 |
| 1 | after | `7b3cd977f5ea3528f88489d716e8365b548c49512c382925057910d3d448ba25` | 531,570,923 | 16,724,054 | 104,478,401 | 13,019,166 |
| 1 | before | `6ef3e2fee4dd3f99969ad728853b275e323b404d75d94c18caafb2c139fba6dd` | 537,174,712 | 22,295,758 | 104,437,941 | 13,019,166 |

All four reports use the sealed 0806 fixture source hash
`9c46542b763fc4bef63dfe4336cadd2bfba2b7e7b3f18a376c3924eb5643b3e9` and
produce the same 215,220-byte output with that hash. The reader verified
`semantic_check`, reopening, expected text, actual text, readback bytes, and
readback hash for every report. The positive summaries range from
531,570,923 to 537,174,712 guest `Ir`; the range is an observation of these
four single-sample publications, not a timing distribution.

## Mechanism observations

The scanner call graph has the same call counts in every publication:

* `quick_xml::reader::Reader<R>::read_event_impl`: 282,612 calls, with two
  recorded scanner edges;
* `litchi_pptx::notes::codec::inspect_element`: 181,678 calls;
* `quick_xml::name::NamespaceResolver::push`: 181,678 calls;
* scanner outgoing call count: 646,384.

The scanner-to-`quick_xml::reader::ns_reader::NsReader<R>::process_event` edge
is absent in both before publications and both after publications. This
profile lane therefore supplies no before/after savings claim for that edge;
the edge was already absent from the retained scanner path. The generic
`NsReader::process_event` symbol still has four unrelated incoming callers in
each publication, with call counts 400, 300, 345, and 8 (1,053 total). Its
global presence is not a zero-call gate.

The two paired owner summaries show diagnostic after/before differences of
5,489,238 and 5,603,789 `Ir` (ratios 0.9897808295 and 0.9895680328). Within
the scanner, self `Ir` falls by exactly 5,571,704 in each pair, while
scanner inclusive `Ir` falls by 5,521,412 and 5,547,328. Immediate-child
inclusive `Ir` rises by 50,292 and 24,376, and the reader edge inclusive `Ir`
rises by 40,824 and 40,460. These are exact guest-instruction attribution
observations from the four publications; they do not establish that the
source rewrite caused a workflow speedup.

The identical resolver-push self cost in all four publications is useful
negative evidence for this candidate: direct `Result<Event>` matching did not
remove the retained namespace push work. The separate code-generation review
records the narrower static mechanism result: the three pre-dispatch event
result copy chains disappear from both ordinary and profile scanner symbols.

## Review result and limits

**Profile publication and attribution review: PASS.** The owner scope,
positive and terminal dump structure, source/output/semantic bindings, and
raw Callgrind parsing all pass. The profile data records unchanged event-read
and resolver-push call counts, an absent scanner-to-`NsReader::process_event`
edge in both legs, and the scanner `Ir` observations above.

Callgrind `Ir` is Valgrind guest-instruction attribution for this exact
capture wrapper. It is not hardware retired instructions, elapsed latency,
RSS, a CPU or phase fraction, or a causal cost assigned to the direct match.
There is one sample per leg in each of two orders, and the owner includes the
full large capture wrapper. The observed `Ir` differences therefore remain
diagnostic evidence alongside the exact assembly review; they cannot satisfy
the 0813 public-workflow adoption threshold or replace the native,
allocation, semantic, independent-reader, and policy checks owned by root.
