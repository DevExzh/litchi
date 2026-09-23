# 0750: publication's source XML audit refuses what XML 1.0 and Namespaces in XML refuse

Status: retained, implemented in `xml-minifier`; a correctness change (ADR 0006
compliance). `performance_claim: none` — the timings and instruction counts
below are the change's cost, reported as evidence, not registered as claims.

OLE2 and OOXML remain the active priority. ODF optimization stays deferred until
that goal completes; iWork is excluded.

Base `3174242282` (the branch tip with records 0745 and 0747); branch
`perf/0750-xml-audit-well-formedness-gaps`; commits `a07d680852` (fixture
corrections in dependent tests) and `de69fb407d` (the auditor), and the commit
that adds this record. The coordinator's task: close the well-formedness gaps
that 0747's reviewer found in the source policy of the OOXML publication XML
auditor, look for more of the same kind, keep 0747's window-proof invariant, run
a compatibility census over the repository's real OOXML parts, and measure only
to show what the change costs.

**Result.** At the base, `verify_source` and `verify_source_replacement`
(`Policy::SOURCE`) accepted every one of the reviewer's inputs, because their
checks stopped at what quick-xml's tokenizer rejects. It is the audit that
publication runs over every XML payload the eager writer publishes and over the
XML parts a source-backed save replaces. Of the 150 inputs probed for this
record, 96 were accepted at the base although none is well-formed. After this
change the audit refuses all 96 with `Error::Malformed` at the offending byte,
and the pair API returns the same error. No record tolerated any of them on
purpose. `verify`, `verify_authored` and the streaming readers are **not
changed**, because fragment callers depend on their verdicts; this is reported
under *What remains* rather than decided silently. Across the repository's 7,222
real OOXML XML members and 5,504 evidence-packet members, three loose sample
parts are newly refused, and each is genuinely malformed. No package member
changes verdict. A release differential campaign of 24 million generated pairs
found no difference between a window proof and two complete audits. The cost: a
source audit of a real part takes 8–18% more time and executes 14–24% more
instructions. The source-backed XLSX and PPTX one-edit saves are unchanged
within noise. The DOCX one-edit save is 3.90% slower (95% CI [+1.73%, +6.46%],
+68 µs at p50); it audits its whole main document twice per commit, so it pays
the new checks in full. Every output byte is unchanged.

## What was changed

Production code, all in `crates/xml-minifier`:

* `src/audit.rs`
  * `Policy` gains `well_formed`, set only by `Policy::SOURCE`. `verify`,
    `verify_authored`, `verify_reader` and `verify_authored_reader` keep
    `false` and behave as before (see *Which policy tolerates each gap*).
  * `State` gains `document_start` (no token read yet), `namespaces`
    (`Option<Namespaces>`, created by the first start tag under the source
    policy only) and a reusable `prefixed` list. `State::within(depth,
    namespaces)` replays both pieces of position state into a window.
  * `scan` takes a `Source` (input, text, BOM length, and the result of one
    character pass). Under the source policy it refuses the first illegal
    character at the token that holds it. It also checks references, `]]>` in
    text, the XML declaration's place and grammar, comments and processing
    instructions, and it closes namespace scopes at end and empty-element tags.
  * `check_start` / `check_attribute_layout`: under the source policy, element
    and attribute names are scanned once as qualified names. Values are scanned
    once for `<` and `&`, and checked further only when they contain one. Each
    attribute is passed to the namespace tracker.
  * `WindowSearch::consider` copies the bindings in scope where the window
    starts; `Window::prove` and `audit_window` replay them.
  * Documentation covers the checks `verify_source` makes and why `verify`
    and `verify_authored` do not make them. The window invariant at `scan` and
    `Policy::SOURCE` now lists the two pieces of replayed position state.
* `src/audit/wellformed.rs` (new, private): the lexical rules. These are the
  XML 1.0 Fifth Edition name classes, `scan_qname`, `check_reference`,
  `check_attribute_value`, attribute-value normalization, `check_comment`,
  `check_processing_instruction` and `check_declaration_grammar`. It also
  holds `scan_characters`, one pass over the input that finds its first
  illegal character and its first `]]>`.
* `src/audit/namespaces.rs` (new, private): the bindings in scope, a keyed
  prefix index (`RandomState`, so collisions cannot be chosen), a 32-slot
  direct-mapped lookup cache, declaration checks, scope closing, and the
  per-tag prefix and expanded-name checks.

No public item is added, removed or re-signed. `Limits`, `Resource`, `Error`,
`Report` and every signature are as they were.

Tests:

* `crates/xml-minifier/tests/well_formedness.rs` (new, 21 tests): one test per
  reviewed gap and per further family. Each asserts the exact
  `Error::Malformed` offset and detail, and that the pair API reports the same
  error. The file also has well-formed controls and a test that the authored
  and default policies keep their verdicts on fragments.
* `crates/xml-minifier/src/audit/wellformed.rs` unit tests (5), including every
  placement of `]]>` around a 64-byte block boundary.
* `crates/xml-minifier/tests/source_replacement.rs`: a window-replay test and a
  generator extension (see *The differential generator*).
* `crates/litchi-opc/tests/source_xml_census.rs` (new): the census as a test.
* Fixture corrections in five `litchi-docx` / `litchi-xlsx` tests (see
  *Fixture corrections in dependent tests*).

**Breaking changes: none in the API; a behaviour change in what is refused.**
`verify_source` and `verify_source_replacement` now refuse inputs they
accepted. So publication returns `OpcError::XmlPublication` for a member whose
XML is not well-formed where it used to write it. Three inputs that were already
refused now give an earlier or more precise error:

* `<r/>` followed by U+0001: the offset is still 4, and the detail is now
  "character U+0001 is not allowed in XML" instead of "character data outside
  the document element";
* `<r>< a/></r>`: byte 4 "invalid XML name", where it was byte 6 "attribute
  name must be followed by '='";
* `<r>a < b</r>`: byte 6 "invalid XML name", where it was the same detail as
  the previous input at byte 11.

Change 0652's first trade-off accepts such changes in this alpha library.

## Authority

ADR 0006: "Publication audits every XML member it is about to write for
encoding, well-formedness, a single document element, the absence of a DTD or
DOCTYPE, and its finite budgets". Its amendment under 0665 lists
well-formedness among the checks that stay. Accepting `<r>&;</r>` or an
undeclared prefix does not meet that sentence. This change brings the source
audit into line with it and does not change the ADR's text.

Change 0652's trade-offs set the shape of the change:

* **Trade-off 2** (correctness first) accepts the cost measured below.
* **Trade-off 3** (optimize the benign majority) is why the checks run in the
  audit's existing single pass. They add one branch-free character pass, and
  they build namespace tables only for documents that use the source policy.
  A malformed input is refused at its first defect.
* **0747's window-proof invariant** is kept. The two new checks that read
  position state are replayed into a window, not left unreplayed (see *The
  window proof*).

## Which policy tolerates each gap, and whether deliberately

Records 0654, 0665, 0677 and 0747 were read for any decision to tolerate one of
these inputs.

* **The source policy (`verify_source`, `verify_source_replacement`) tolerates
  none of them on purpose, so these gaps are fixed.**
  * 0654 describes the audit of original bytes as keeping "well-formed XML as
    quick-xml parses it", which describes the mechanism, not a decision to
    accept these inputs.
  * 0665's amendment to ADR 0006 says every non-compactness check "stays", and
    names the same list.
  * 0677 moves only byte-order-mark offsets.
  * 0747 lists the reviewer's inputs under *What remains* as pre-existing gaps
    left for a separate item. Its own differential generator already classes
    `&bogus;`, `x]]>y`, `e='<'` and an inner `<?xml …?>` among the tokens that
    are "malformed or refused by the audit".
  * No record names a real producer that needs any of these inputs, and the
    census below finds none.
* **The authored and default policies (`verify_authored`, `verify`) and the
  streaming readers have the same gaps, and their callers depend on them.** No
  record calls this deliberate, but the callers rely on it:
  * `litchi_opc::AuthoredXmlFragment` wraps markup in a synthetic element and
    audits it with `verify_authored`, then splices it into a document whose
    root declares its prefixes.
  * The ODF crates audit each streamed fragment of a generated document the
    same way. They wrap whole documents, declaration included, in the same
    kind of synthetic element, and keep test fixtures with undeclared prefixes
    and entities.

  The first implementation of this change applied every check under every
  policy. It failed 46 OOXML tests (31 of them through the fragment wrapper) and
  86 ODF tests
  ([`draft-all-policies/failures.md`](results/change-0750/draft-all-policies/failures.md)).
  The ODF crates are out of scope for this program and may not be changed, and
  an undeclared prefix in a fragment is not a defect of the fragment. So those
  policies are **not changed**; the consequence is reported under *What
  remains*.

## The gaps, and what refuses each now

Unless a row says otherwise, all four auditors accepted each example at the base.
After the change, `verify_source` refuses each with `Error::Malformed` at the
offset shown. `verify`, `verify_authored` and `verify_reader` return exactly
what they returned before. The full table of 150 probed inputs is
[`gaps/verdicts.md`](results/change-0750/gaps/verdicts.md): 96 newly refused, 0
newly accepted, 39 well-formed controls accepted on both legs, 12 refused
identically on both, and the 3 refused more precisely listed above.

| family | example | refusal (offset) |
| --- | --- | --- |
| reviewed: reference outside the root | `<r/>& ;` | character data outside the document element (4) |
| reviewed: empty reference | `<r>&;</r>` | malformed reference (3) |
| reviewed: reference that is not a name | `<r>&a b;</r>` | malformed reference (3) |
| reviewed: malformed or illegal character reference | `<r>&#xZZ;&#0;</r>` | malformed character reference (3). `&#0;`, `&#x1;`, `&#xD800;`, `&#xFFFE;` and `&#x110000;` give "character reference to a character XML does not allow" |
| reviewed: declaration not at the start | `<r><?xml version="1.0"?></r>` | an XML declaration is allowed only at the start of the document (3). Also refused after whitespace, a comment, the root, a first declaration, or a BOM and a space |
| reviewed: `]]>` in character data | `<r>]]></r>` | `']]>' is not allowed in character data` (3) |
| reviewed: undeclared prefix | `<p:r/>`, `<r p:a="1"/>` | undeclared namespace prefix (1, 3). Also refused for a binding used after its element closed |
| illegal characters | C0 controls other than tab, LF and CR, and U+FFFE/U+FFFF, in text, attribute values, comments, PIs, CDATA and outside the root | character U+XXXX is not allowed in XML (at the character). Outside the root the base already refused it, as character data outside the document element |
| lone surrogate in UTF-8 | `<r>\xED\xA0\x80</r>` | already refused as `Error::Encoding` (3); unchanged |
| `<` in an attribute value | `<r a="<"/>` | `'<' is not allowed in an attribute value` |
| bare `&` or undeclared entity in an attribute value | `<r a="a&b"/>`, `<r a="&bogus;"/>` | unterminated reference / reference to an undeclared entity |
| undeclared entities | `&nbsp;`, `&AMP;` | reference to an undeclared entity. Without a DTD, which the audit refuses, only the five predefined entities are declared |
| duplicate attributes after prefix resolution | `a:x` and `b:x`, with `a` and `b` bound to one URI | two attributes have the same namespace name and local name. Exact duplicates are still refused by quick-xml |
| `xml` and `xmlns` misuse | `xmlns:xml="urn:other"`, `xmlns:xmlns=…`, `xmlns:p="http://www.w3.org/XML/1998/namespace"`, `xmlns="http://www.w3.org/2000/xmlns/"`, `<xmlns:r/>` | the matching Namespaces in XML 1.0 diagnostic. Namespace names are compared after attribute-value normalization, so `&#110;` cannot hide one |
| prefix bound to the empty string | `xmlns:p=""` | a namespace prefix must not be undeclared (`xmlns=""` is still accepted) |
| mismatched end tags | `<r><a></b></r>` | already refused by quick-xml; unchanged, and now pinned by a test |
| comments | `<!-- a -- b -->`, `<!-- a --->` | `'--' is not allowed in a comment` / `a comment must not end with '-'` |
| names | `<1r/>`, `<r><a<b/></r>`, `< a="1"/>`, `<a/ >`, `<×/>` | invalid XML name |
| qualified names | `<a:b:c …/>`, `<r:/>`, `<:r/>`, `a:="1"`, `xmlns:=""` | name is not namespace-well-formed |
| PI targets | `<?XML x?>`, `<? x?>`, `<?1x?>`, `<?a:b x?>`, `<?xmlversion="1.0"?>` | target `xml` is reserved / invalid processing-instruction target |
| declaration grammar | `<?xml?>`, no version or a late one, `version="2.0"`, `standalone="maybe"`, wrong order, an unknown attribute, `encoding="ISO-8859-1"`, `"UTF-16"` or `"UTF8"` | the matching declaration diagnostic. The audit reads UTF-8, so any other declared encoding is refused: XML 1.0 §4.3.3 makes a mismatch fatal, OPC permits only UTF-8 and UTF-16, and UTF-16 input is not UTF-8 |

Well-formed spellings are still accepted:

* `]]&gt;`, `]] ]>`, and CDATA ending in `]]`;
* `'>'` in attribute values;
* U+0085, U+2028, DEL, private-use characters, U+FEFF and supplementary
  characters;
* Fifth Edition name characters (`é`, `中文`, `a·`);
* a prefix declared on its own element or after the attribute that uses it,
  shadowing, and aliases with distinct local names;
* `xml:` attributes and `<xml:r/>`;
* `version="1.1"`, `encoding='utf-8'`, whitespace around `=` in the
  declaration, and a BOM before the declaration.

## How the checks are made

* **One token at a time.** Names, references, attribute values, comments,
  processing instructions and the declaration's grammar are rules on the bytes
  of one token. Names are validated in the same pass that finds where they end.
  A 256-entry class table decides an ASCII name. Only a name with a non-ASCII
  byte or a defect is decided again, character by character. Attribute values
  are scanned once, as before, and checked further only when they contain `<`
  or `&`.
* **Characters and `]]>` in one pass.** Before the scan, `scan_characters`
  reads the whole input once in 64-byte blocks, branch-free, and records the
  first illegal character and the first `]]>`. The scan reports an illegal
  character at the token that holds it, so this error comes in source order
  like every other. A text token is checked against the next `]]>` at or after
  its start. Only a document that contains a `]]>` somewhere (a CDATA
  section's end, a comment) ever searches again.
* **Namespace bindings.**
  * Declarations are checked where they appear: reserved prefixes and names,
    compared after normalization, and no undeclaring.
  * Prefixed bindings are kept while their element is open.
  * Every prefix of an element or attribute name must resolve.
  * Attributes must be unique by namespace name and local name. In the common
    case, where all prefixed attributes share one binding, this is decided
    without a comparison; otherwise it is decided by sorting.
  * Lookups go through a direct-mapped cache of prefixes packed into a `u64`,
    backed by a keyed hash index. So a document cannot make lookups slower than
    constant time by piling up declarations.
  * Memory is bounded by the input and by the attribute budget, since every
    binding is an attribute. The tables exist only under the source policy.
* **The declaration's place.** `document_start` is true only before the first
  token of a complete scan.

## The window proof

0747's invariant says a source-policy check may read only the token's bytes,
the depth, the aggregate counters and two narrow uses of `spaces` and `roots`.
Any new check that reads more must be replayed into a window or must disable
window proofs. Two new checks read more, and both are replayed exactly:

* **The declaration's place.** `State::within` sets `document_start` false,
  because a window starts inside the document element. A declaration in a
  window is refused there, as the complete scan refuses it.
* **The bindings in scope.** When the window's element closes in the
  original's audit, the bindings left in scope are those of its ancestors.
  Their start tags precede the window, so they are identical in the
  replacement.
  * `WindowSearch::consider` copies these bindings, and the window scan starts
    from the copy.
  * A balanced window closes every binding it opens, so the bindings after it
    are also the original's.
  * If the copy cannot be allocated, no window is used.

Characters do not depend on position. The window's own bytes go through
`scan_characters`, and the rest of the replacement repeats bytes that the
original's pass checked. A `]]>` cannot straddle a window's edge: the window
starts with a `<`, and it ends with a `>` that closes markup.

The documentation of the invariant at `scan` and `Policy::SOURCE` now lists
both replayed pieces. `debug_check_window_proof` is unchanged and passes in
every debug test run.

Two mutation checks confirm that the replay matters:

* Setting `document_start` true in `State::within` makes the new window test
  fail in the debug cross-check (`window proof Window { original: 129..152, … }
  contradicts the complete source audit … an XML declaration is allowed only at
  the start of the document`).
* Starting windows with no bindings turns the new test's window proofs into
  complete scans and fails the generator's window floor.

Both were reverted, with the diff hash unchanged.

## Compatibility census

The census audits every XML member that publication would audit —
`is_xml_part(name, media_type)`, with the media type from each package's own
content-type map. It covers every OPC package under `test-data/`,
`docs/performance/` and `crates/*/tests/` (a ZIP holding `[Content_Types].xml`,
found by content rather than extension), plus every loose `.xml` / `.rels` file
under `test-data/ooxml/`. The base auditor and the final auditor each audited
every member under all four entry points
([`census/summary.txt`](results/change-0750/census/summary.txt); rows in
`census/census.jsonl.gz`; `scripts/census.py` and `scripts/census_summary.py`).

| corpus | files | XML members | `verify_source` accepted → refused | refused → accepted | other auditors' verdict text changed |
| --- | ---: | ---: | ---: | ---: | ---: |
| real fixtures (`test-data`: 338 packages, 81 loose parts) | 419 | 7,222 | **3** | 0 | 0 |
| evidence packets (`docs/performance/results`: 1,658 packages, many deliberately corrupted; 1,397 members unreadable) | 1,491 | 5,504 | 0 | 0 | 0 |

**No member of any package changes verdict.** The three newly refused members
are loose sample files under `test-data/ooxml/pptx/`. None of them is a
namespace-well-formed document, so each should be refused:

* `backgrounds/solid.xml`: a `<p:bg>` fragment whose `p:` and `a:` prefixes
  are declared by the enclosing slide (undeclared prefix at byte 1);
* `transitions/p14_ripple.xml`: an `<mc:AlternateContent>` fragment whose
  `p:transition` relies on the slide's `p` binding (byte 196);
* `placeholders/invalid-placeholder.xml`: a deliberately invalid slide with
  `type="&unknown;"` (reference to an undeclared entity, byte 221).

The `litchi-pptx` tests that read these files either embed them in documents
or test that they are refused, and they pass. The six members refused on both
legs keep identical text:

* three that are not UTF-8;
* the two POI entity-expansion fixtures (DOCTYPE);
* a deliberately mismatched end tag.

Among the 34 refused evidence-packet members, none changes text either. The
real corpus has at most 39 prefix declarations in scope at once (Word's document
roots).

The census is also a test. `litchi-opc`'s `source_xml_census.rs` walks
`test-data` the same way (338 packages, 7,141 package members and 81 loose
parts; 3.5 s in a debug build) and asserts that the refused members are exactly
the nine above, with their messages.

## Fixture corrections in dependent tests

Five tests built fixtures that are not well-formed and passed them to
`verify_source`: four through the eager writer, which audits every XML member
with it, and one through `litchi-docx`'s publication gate. Both now refuse them,
as they should. Each fixture was corrected without changing what its test checks
(commit `a07d680852`):

| test | fixture defect | correction |
| --- | --- | --- |
| `litchi-docx` `document::transaction::tests::the_gate_accepts_every_noncompact_spelling_and_still_refuses_structure` | `<w:p/>` inputs with `w` undeclared | declare `xmlns:w`. The refused inputs and the corpus loop are unchanged, and the loop still passes: no DOCX main document changes verdict |
| `litchi-docx` `validation::main_relationship_closure_reports_missing_and_incompatible_targets_at_main_uri` | `<w:hdr/>` with `w` undeclared | declare `xmlns:w` |
| `litchi-xlsx` `source_backed_page_setup` (7 tests share one fixture) | worksheet with `r:id` and `r` undeclared | declare `xmlns:r`, as the fixture's other worksheets already do |
| `litchi-docx` `source_backed_glossary_story_text` (2 tests) | `&unknown;` and `Alpha&#x1;`, on purpose, to test the glossary reader's refusal | the writer publishes a well-formed stand-in, and the raw archive writer puts the malformed glossary bytes back into the archive, so the reader still sees them |
| `litchi-docx` `source_backed_secondary_story_text` (1 test) | footnotes with `&unknown;`, on purpose | the same |

## The differential generator

0747's seeded differential test compares the pair API with two complete
`verify_source` calls. Its element name `x:y` was bound nowhere. Under the new
check almost every generated document would have been refused on the original
side (63% instead of 36%), and the window floor failed. The generator was
changed as follows; its assertions are unchanged:

* It declares `x` on the document element 19 times in 20, so unbound prefixes
  still occur.
* It adds namespace attributes: `xmlns:x`, `x:k`, and an alias `xmlns:z` of
  `x`'s name with `z:k`. It sometimes picks two per element, so duplicate
  expanded names occur.
* It adds newly refused tokens to the invalid pools: `&#0;`, a vertical tab,
  U+FFFF, `&bogus;` in a value, `xmlns:n=""`, `n:k`, `9="1"`, `<!--a--b-->`,
  `<?XML x?>` and `<n:e/>`.

Across six other seeds at the suite's size, two-edit windows ranged from 239 to
256, against a floor of 200.

A release-mode campaign of 8 seeds × 3,000,000 cases on the final code found no
difference in verdict, failing side or error value
([`campaign/`](results/change-0750/campaign/totals.txt)). It covered 3,765,562
window proofs (145,828 of them over two separated edits), 9,454,154 refusals on
the original and 7,239,727 on the replacement.

**The campaign found a defect in an earlier candidate of this change.** A `]]>`
whose `]]` ended one 64-byte block and whose `>` began the next was missed by
the character pass, so a complete audit accepted it. Every seed failed on it.
The pass now checks the pair that starts at a block's last byte, and it searches
the tail from one byte earlier. A unit test places `]]>` at every offset of
inputs of eight lengths around the block boundaries. It fails on the defective
pass (length 129, offset 126) and passes on the fixed one.

## Measured

Host AMD EPYC 9R45, `Linux 7.0.0-1012-aws x86_64`, shared with other agents.
Every measured process was pinned to CPU 24 with `taskset`, and no build or test
of this record ran while a timed process ran.

* **Harness.** `tools/perf-baseline`, unchanged by this record, with its
  deterministic corpora. Both legs were built by the identical command, from a
  detached worktree at the base and from the branch worktree, with rustc 1.95.0
  (pinned by `rust-toolchain.toml`).
* **Audit probe** (`probe/audit`). The same source was built against each leg's
  `xml-minifier` from outside the checkout, so with rustc 1.98.1, the host
  default, for both legs. It times the audits on five real parts from
  `test-data` and on the whole accepted census corpus
  ([`probe/parts.txt`](results/change-0750/probe/parts.txt), rebuilt
  byte-identically by `scripts/extract_parts.py`).
* **Binary SHA-256s** are in [`binaries.txt`](results/change-0750/binaries.txt).
  The final candidate's harness is `6ac6e6c9…`; the base's is `15d6f91b…`.

### Timing: the harness, ABBA

The harness ran four rounds. In each round every case ran A B B A (A = base,
B = candidate), so each leg had eight processes, paired within the round (A1
with B2, A4 with B3). The table shows:

* the median of the process p50s and p95s;
* the median paired p50 change;
* a percentile bootstrap 95% interval over the eight paired changes (20,000
  resamples, seed 750).

| case | processes × samples | before p50 ms | after p50 ms | paired p50 change | 95% CI | p95 ms, before → after | output |
| --- | --- | ---: | ---: | ---: | --- | --- | --- |
| `xlsx_source_backed_cell_values_one_edit_save` medium | 8+8 × 30 | 4.199 | 4.178 | −1.01% | [−1.61%, +2.19%] | 4.277 → 4.322 | identical |
| `xlsx_source_backed_cell_values_one_edit_save` dense-sparse | 8+8 × 20 | 26.824 | 26.792 | +0.32% | [−1.78%, +1.17%] | 27.154 → 27.059 | identical |
| `xlsx_eager_cell_values_one_edit_save` dense-sparse | 8+8 × 20 | 37.815 | 38.237 | +0.76% | [−0.08%, +2.31%] | 38.250 → 38.558 | identical |
| `docx_source_backed_one_edit_save` | 8+8 × 40 | 1.795 | 1.863 | **+3.90%** | [+1.73%, +6.46%] | 1.867 → 1.922 | identical |
| `pptx_source_backed_one_edit_save` | 8+8 × 40 | 6.731 | 6.737 | −0.11% | [−2.97%, +0.93%] | 7.164 → 7.037 | identical |
| `xls_semantic_one_edit_save` (control: OLE2, no XML audit) | 8+8 × 20 | 0.080 | 0.081 | +0.14% | [−0.33%, +0.52%] | 0.092 → 0.089 | no digest reported |

"Identical" means every process of both legs reports one output SHA-256 for the
case: the change moves no published byte.

The harness splits the source-backed XLSX interval into phases. The table shows
medians of the per-process medians, in ms:

| case | open | planning | commit | publication |
| --- | --- | --- | --- | --- |
| one-edit medium | 0.100 → 0.097 | 1.561 → 1.544 | 0.840 → 0.809 | 1.697 → 1.717 (+1.2%) |
| one-edit dense-sparse | 0.112 → 0.108 | 10.263 → 10.187 | 5.338 → 5.249 | 11.108 → 11.234 (+1.1%) |

Publication is the only phase that runs the audit. It is 20 µs and 126 µs
slower. Planning and commit code is unchanged, and their 1–2% moves are within
what this host shows for unchanged code.

### Instructions: the audit's share of each operation

The harness binaries keep symbols for the auditor's entry points. The
callgrind isolation pairs run each case at `--samples 1` and at `--samples 3`;
half the difference is one sample. For each entry point, the table below gives
its inclusive instructions per sample and the entry point's caller
([`callgrind/pairs-summary.txt`](results/change-0750/callgrind/pairs-summary.txt),
`scripts/cg_attribute.py`):

| case | entry point (caller) | before per sample | after per sample | change |
| --- | --- | ---: | ---: | ---: |
| XLSX one-edit medium | `verify_source_replacement` (`write_topology_to_stream`) | 6,837,808 | 7,585,362 | +747,554 (+10.9%) |
| XLSX one-edit dense-sparse | `verify_source_replacement` (`write_topology_to_stream`) | 48,068,156 | 53,180,802 | +5,112,647 (+10.6%) |
| DOCX one-edit | `verify_source` (`Edit::commit`: `litchi-docx`'s `publication_accepts_preserved_xml` gate) | 1,240,511 | 1,516,649 | +276,138 (+22.3%) |
| DOCX one-edit | `verify_source_replacement` (`write_single_part_overlay_to_stream`) | 1,315,528 | 1,604,206 | +288,677 (+21.9%) |
| PPTX one-edit | `verify_source_replacement` (`write_single_part_overlay_to_stream`) | 200,740 | 243,952 | +43,213 (+21.5%) |

The program totals per sample move +0.005%, +0.41%, +0.19% and −0.03% for the
four cases. The totals of unchanged code differ between processes by up to
about two million instructions per sample (hash seeds, heap layout), so the
entry-point figures are the ones that measure the audit.

The DOCX commit audits its main document twice, once in each of the two rows
above. Together the two audits add 565K instructions per commit. On the probe's
DOCX and PPTX parts, each million added instructions costs 23–32 µs. That
predicts about 13–18 µs; even at the audit's whole-audit rate it would be about
26 µs. The measured median change is 68 µs, and even the lower end of its
interval (31 µs) is above the prediction. The instruction count does not account
for the whole change. The remainder is not attributed; code layout and kernel
time for new allocations, which callgrind does not count, are candidates, not
findings.

### Unit probe: the audits on real parts

The probe ran four rounds of A B B A processes, pinned. Each process times every
case in 41 batches of about 20 ms. The table shows:

* the median of the process medians;
* the median paired change, with the same bootstrap interval as above;
* callgrind instructions per audit (the probe at N iterations minus 0
  iterations, divided by N).

| case | bytes | before µs | after µs | paired change | 95% CI | instructions per audit, before → after | change |
| --- | ---: | ---: | ---: | ---: | --- | --- | ---: |
| source, `ws-structured` (worksheet) | 1,558,391 | 4,321.5 | 4,740.8 | +9.64% | [+9.06%, +10.44%] | 99,994,825 → 114,551,990 | +14.56% |
| source, `ws-patriarch` (worksheet) | 3,382,556 | 12,429.7 | 13,653.4 | +9.86% | [+9.13%, +11.71%] | 300,237,054 → 343,249,874 | +14.33% |
| source, `docx-drawing` (main document) | 288,070 | 565.0 | 636.8 | +13.02% | [+11.93%, +13.76%] | 11,753,300 → 13,996,126 | +19.08% |
| source, `docx-table-alignment` (main document) | 40,151 | 71.7 | 82.5 | +15.27% | [+14.39%, +16.25%] | 1,627,164 → 2,023,220 | +24.34% |
| source, `pptx-slide11` (slide) | 121,498 | 307.8 | 332.9 | +8.25% | [+7.29%, +8.54%] | 7,082,244 → 8,162,344 | +15.25% |
| pair, `ws-structured` with a one-byte edit (window proof) | 1,558,391 | 4,585.0 | 5,075.4 | +10.39% | [+9.77%, +11.53%] | 105,385,698 → 120,272,179 | +14.13% |
| pair, `docx-drawing` with a one-byte edit (replacement refused, below) | 288,070 | 614.1 | 722.8 | +17.77% | [+15.96%, +18.86%] | not counted | |
| source, every accepted `test-data` member (7,213) | 36,816,208 | 95,622.0 | 108,344.1 | +13.38% | [+12.78%, +13.64%] | 2,022,238,166 → 2,415,654,489 | +19.45% |
| source, a 55-byte document | 55 | 0.23 | 0.43 | +87.81% | [+82.83%, +96.87%] | 5,107 → 9,356 | +83.20% |
| authored, a 38-byte fragment | 38 | 0.19 | 0.19 | −0.16% | [−1.69%, +0.21%] | 4,319 → 4,397 | +1.79% |
| authored, `ws-patriarch` (control) | 3,382,556 | 12,779.8 | 12,572.1 | −1.24% | [−3.34%, +0.40%] | 303,870,864 → 308,823,796 | +1.63% |

Over the whole accepted corpus the source audit moves from 385 MB/s to 340 MB/s.
It also has a fixed cost of about 200 ns per audit, which dominates only for
tiny documents. Callgrind attributes the 55-byte document's 4,249 added
instructions as follows
([`callgrind/tiny-document.txt`](results/change-0750/callgrind/tiny-document.txt)):

* about half to the character pass's byte-by-byte tail, because an input
  shorter than one 64-byte block has no full block;
* most of the rest to declaring, hashing and resolving the document's one
  prefix.

The authored audit's time is flat; its 1.6–1.8% more instructions come from the
policy branches it now passes. The probe's DOCX pair edit replaces the first byte of a
two-byte UTF-8 character, so on both legs that pair is the original's complete
audit followed by `Err(Replacement(Encoding { valid_up_to: 146044 }))`
([`probe/pair-verdicts.txt`](results/change-0750/probe/pair-verdicts.txt)).
Both legs do the same work, so the comparison stands, but the case does not
time a window proof.

### Regressions and every over-5% flag

**Adverse results reported as the change's cost:**

* the DOCX one-edit save: +3.90% (+68 µs), interval [+1.73%, +6.46%];
* every source audit on the probe: +8.25% to +17.77% in time and +14% to +24%
  in instructions;
* about +200 ns fixed per source audit;
* +1.6–1.8% instructions on the authored audit, with time flat.

No other harness case's interval excludes zero.

**Harness flags.** Of the 144 paired-process comparisons (6 cases × 8 pairs ×
p50, p95 and mean), 32 moved more than 5%, and 15 of those are adverse
([`timing/flags.json`](results/change-0750/timing/flags.json)). Only the DOCX
case has a consistent pattern: in round 2 both pairs are adverse at p50 (+5.96%,
+6.46%), and the second is also adverse at p95 and mean (+6.60%, +7.07%). The
other eleven adverse flags are isolated:

* single slow processes, which both legs have:
  * DOCX round 1: both after-leg processes (p95s of 5.080 and 4.821 ms, and a
    p50 of 2.977 ms), against before-leg processes in rounds 2 and 3 with p95s
    of 4.832 and 4.279 ms;
  * XLSX medium and dense-sparse, round 2, slot 2 (after): p95s of 7.734 and
    56.327 ms, against an eager dense-sparse before-leg p95 of 77.389 ms in the
    same round;
* eager dense-sparse, round 2: a p95 of 60.509 → 64.539 ms (+6.66%), in a round
  whose other eager flags favour the candidate by 5–27%;
* PPTX, round 2: a p50 of 6.640 → 6.988 ms (+5.24%), against five favourable
  PPTX flags of −5.96% to −15.32%.

**Probe flags.** 219 comparisons moved more than 5%
([`probe-timing/flags.json`](results/change-0750/probe-timing/flags.json)):

* all 24 comparisons (8 pairs × 3 metrics) of each of the nine source and pair
  cases are adverse, which is the cost in the table above;
* three are favourable, on the authored tiny fragment.

## What remains

* **The authored, default and streaming audits keep every gap in this
  record**, on purpose, because fragment callers rely on them (see *Which
  policy tolerates each gap*). This has two consequences for OOXML
  publication, neither changed here:
  * Some `litchi-opc` publication sites audit a *complete* authored document
    with `verify_authored` or `verify_authored_reader`, and so get only the
    tokenizer-level checks:
    * topology additions and canonical relationship XML
      (`validate_overlay_xml`);
    * a cross-package precompressed transfer's donor bytes;
    * the decoded-splice stream audit.

    The next step is a document-scoped authored audit for those sites (the
    same `well_formed` flag with the compact contract), leaving fragment
    callers on the fragment-compatible audit. It would move `litchi-opc`
    refusals and needs its own census.
  * A replacement carrying a `SourceXmlPart` skips the publication audits and
    relies on `litchi-opc`'s `validate_source_xml`, which 0747 already noted
    checks less. It was not examined here.
* **The DOCX commit audits one document twice.** On `docx_source_backed_one_edit_save`
  the main document is audited completely twice per commit:
  * by `litchi-docx`'s `publication_accepts_preserved_xml` gate, over the base
    snapshot;
  * by `litchi-opc`'s replacement pair, whose original half is a complete scan
    of the part the package holds. For a first edit these are the same bytes.

  Each audit is now 22% more instructions. Carrying the gate's verdict into the
  pair would remove one complete audit per DOCX commit, which is more than this
  change adds. It needs its own record, because it crosses crates and would
  need a proof like 0747's.
* **The fixed cost.** The character pass checks an input's last partial block
  (all of an input shorter than 64 bytes) byte by byte. Checking that tail as
  one padded block would remove about half of the roughly 200 ns fixed cost.
  This matters only to callers that audit many tiny documents, and it was not
  pursued here.
* **What the source audit still does not check.**
  * Namespace names are checked for emptiness and the two reserved names only,
    not as URI references. Mainstream parsers do not check them either.
  * A document that declares `version="1.1"` is judged by XML 1.0's rules. It
    refuses character references to C0 controls, which 1.1 allows. It accepts
    literal C1 controls (U+007F–U+0084, U+0086–U+009F), which 1.1 requires to
    be escaped.
  * Schema validity, and OPC's rules on `xml:` and `xsi:` attributes, are not
    well-formedness and are not checked.
* **A new refusal is a behaviour change.** A producer that writes any of the
  refused spellings now fails `verify_source`. The census found none among
  7,222 real members and 5,504 evidence-packet members. Packages outside this
  repository were not measured.

## What is not claimed

* No performance claim; `performance_claim: none`. The change adds work to
  every source audit, and that work is reported as a cost.
* The per-part audit costs are this host's, measured on these real parts. The
  end-to-end cases use the harness's synthetic corpora.
* The equivalence of window and complete audits rests on evidence (campaigns,
  the debug cross-check, mutation checks), not on a proof.
* No claim is made that the audit now checks everything XML 1.0 and Namespaces
  in XML 1.0 require; the exceptions are listed above.

## Verification

Gates were run in the worktree on the final candidate. That is the content of
commits `a07d680852` and `de69fb407d`: no source file changed between the gates
and the commits. Commands, tails and exit codes are in
[`gates.txt`](results/change-0750/gates.txt). Every gate exits 0:

| gate | result |
| --- | --- |
| `cargo fmt --all --check` | clean |
| `cargo check -p xml-minifier -p litchi-opc -p litchi-ooxml-common -p litchi-docx -p litchi-xlsx -p litchi-pptx -p litchi-xlsb -p litchi --all-targets --locked --offline` | clean |
| `cargo check -p litchi-odf-common -p litchi-odt -p litchi-ods -p litchi-odp -p litchi-xls -p litchi-ppt -p litchi-imgconv --all-targets --locked --offline` (compile-only: these use the auditor and are not changed) | clean |
| `cargo clippy -p xml-minifier -p litchi-opc -p litchi-docx -p litchi-xlsx --lib --no-deps --locked --offline -- -D warnings`, and the same with `--all-targets` | clean, both |
| `cargo test -p xml-minifier -p litchi-opc -p litchi-ooxml-common -p litchi-docx -p litchi-xlsx -p litchi-pptx -p litchi-xlsb --locked --offline` | 5,732 passed, 0 failed, 46 ignored |
| `cargo test -p litchi --features doc,docx,ppt,pptx,xls,xlsx,xlsb,odt --locked --offline` | 382 passed, 0 failed, 7 ignored |
| `cargo test -p litchi-odf-common -p litchi-odt -p litchi-ods -p litchi-odp -p litchi-xls -p litchi-ppt --locked --offline` (these crates are not changed; run to confirm their tests still pass) | 5,049 passed, 0 failed, 13 ignored |
| `cargo test --release -p xml-minifier --test source_replacement`, 8 seeds × 3,000,000 cases | all seeds pass |
| `RUSTDOCFLAGS="-D warnings" cargo doc -p xml-minifier -p litchi-opc -p litchi-docx -p litchi-xlsx --no-deps --locked --offline` | clean |
| `python3 tools/check_crate_boundaries.py` | pass |
| `python3 tools/non_iwork_gate.py verify` | pass |
| `python3 tools/check_perf_claims.py --registry docs/performance/claim-registry-v1.json --repo-root . --mode structural` | 10 claims validated |

The `xml-minifier` suites pass in these counts: lib 16, `attribute_cardinality`
8, `audit` 19, `ooxml_assets` 5 (1 ignored), `source_replacement` 15,
`stream_audit` 20 and `well_formedness` 21. The debug test runs exercise
`debug_check_window_proof` on every window proof. The harness is unchanged, so
its own test suite and the coverage validator were not required. No shared log,
registry or coverage index was edited. The sections ready to paste are in
[`log-sections.md`](results/change-0750/log-sections.md).

## Cleanup

After the evidence was copied into the packet, the gates passed and the two code
commits were made, the following were deleted:

* the candidate target directory (`targets/0750`, 115.1 GB);
* the before leg's target directory (`targets/0750-before`, 1.10 GB);
* the probes' target directory (`targets/0750-census-after`, 59 MB);
* the detached base worktree (`0750-before-src`, 10.2 GB, removed with
  `git worktree remove --force`, then `prune`);
* the scratch directory's contents (211 MB).

Each is verified absent in
[`cleanup.json`](results/change-0750/cleanup.json). Binary identities are in
[`binaries.txt`](results/change-0750/binaries.txt). Not kept: binaries, raw
callgrind profiles (their summaries are kept), the probe's input parts
(`scripts/extract_parts.py` rebuilds them), and build and test logs (their
totals are in `gates.txt`). The worktree and the branch are kept.
