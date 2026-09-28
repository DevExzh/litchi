# 0807 next-step review: event transport and namespace state

This is a bounded source review at HEAD `3624e73236`. It uses only the
retained 0807 artifacts and the locally installed `quick-xml 0.41.0` registry source;
it made no build, capture, replay, or production edit. The 0807 numbers below
are attribution evidence, not a native latency fraction or a speedup claim.

## What the fresh packet actually measured

The exact-owner frame census puts
`quick_xml::reader::ns_reader::NsReader<R>::process_event` at 189 and 190
leaf samples in repeats 0 and 1. The same owner frame has 1 unknown interior
sample in each repeat. Its nested counts are 896/898 for
`scan_processed_xml`, 223/244 for `inspect_element`, and 64/67 for
`core::str::converts::from_utf8`; these are overlapping ancestry counts. The
full rows are in [`root-frame-counts.json`](root-frame-counts.json). The
post-capture disassembly is tied to the separate frame-pointer binary in
[`event-assembly.json`](event-assembly.json) and shows a 589-byte
`process_event` body with event aggregate copies around its Start/Empty
branches. Sampled offsets 0x43, 0xbe, and 0x118 account for 54/67/45 samples
in repeat 0 and 47/55/59 in repeat 1. Sampling skid and code-generation
differences mean those offsets identify a shape to test, not removable
instruction cost.

The complete relevant Callgrind rows for `profiles/0-large.callgrind.1` are
in [`profile-analysis.json`](profile-analysis.json). The wrapper owner has
2 self IR, and `scan_processed_xml` has 22,308,021 self IR plus these direct
children (each child value is inclusive at that edge):

| Function | Calls from scanner edge | All-context function self IR | Scanner edge inclusive IR |
| --- | ---: | ---: | ---: |
| `inspect_element` | 181,678 | 29,165,311 | 158,607,508 |
| `Reader<R>::read_event_impl` | 282,612 | 34,766,287 | 104,438,776 |
| `NsReader<R>::process_event` | 282,612 | 12,051,464 | 46,146,741 |
| `NamespaceResolver::default` | 104 | — | 135,962 |
| `drop_in_place<NsReader<&[u8]>>` | 104 | — | 65,310 |
| `drop_in_place<Event>` | 104 | — | 416 |

Function self costs above aggregate all incoming contexts; scanner-edge calls
and inclusive costs name only that edge. They are different scopes and are
not additive. The `process_event` row has one relevant direct edge: 182,198 calls to
`NamespaceResolver::push`, with 34,941,704 inclusive IR. The push row has
13,019,166 self IR and two direct children: 276,440 calls to
`IterState::next` with 21,471,516 inclusive IR, and 1,039 calls to
`NamespaceResolver::add` with 451,022 inclusive IR. The reader row's direct
children are `memchr3_raw` (27,541,257 IR), `emit_start` (19,077,207),
`memchr2_raw` (14,251,679), `emit_end` (9,276,917), `memchr_raw` (23,032),
and `emit_question_mark` (9,210). These rows overlap when viewed through
different parents; they must not be summed into a claimed percentage.

The all-context `IterState::next` top-self row is 35,114,190 self IR plus
14,102,788 inclusive IR in `check_for_duplicates` over 281,618 calls; that
row is not limited to the resolver's 276,440 child calls. The repeat-1 large
row retains the same call counts and `process_event` self IR; its resolver
push edge is 35,057,719 inclusive IR and its `add` edge is 567,037. This
repeat agreement supports the shape of the seam, while still saying nothing
about native benefit.

Source shows that namespace scanning occurs on the Start/Empty branches.
The 182,198 push calls aggregate the process-event row, whereas 282,612
process-event calls name the scanner edge; these different scopes do not
form an event distribution. The 1,039
`add` calls show that declarations are sparse in this fixture, but they do
not measure the number of no-declaration tags or prove that the attribute
iteration can be skipped. The packet therefore supports a transport
hypothesis, not deletion of namespace scanning.

## Source semantics that the candidate must retain

The production path is [`scan_processed_xml`](../../../../crates/litchi-pptx/src/notes/codec.rs:356).
It constructs `NsReader` at line 372, sets `trim_text(false)`, reads events at
line 381, and calls `inspect_element` for Start and Empty events at lines 391
and 411. `inspect_element` receives the `NsReader` at line 448, resolves the
element at line 459, and then runs the checked-attribute path from line 477.
That checked path validates malformed and duplicate attributes, skips only
`xmlns` entries for inventory, resolves qualified attributes, unescapes their
values, and enforces the attribute budgets. It remains required.

In quick-xml 0.41.0, `NsReader::read_event_impl` calls `pop()` before the
underlying reader and then passes its `Result<Event>` to
`process_event` (`.../quick-xml-0.41.0/src/reader/ns_reader.rs:64-71`).
`process_event` (`:80-101`) does exactly four state transitions:

* Start: `NamespaceResolver::push` and return the Start event.
* Empty: `push`, set `pending_pop`, and return the Empty event.
* End: set `pending_pop` and return the End event.
* Any other result: return it unchanged.

`NamespaceResolver::push` is public at
`.../quick-xml-0.41.0/src/name.rs:704-729`. It increments
`nesting_level`, scans all attributes with duplicate checks disabled, adds
`xmlns` and `xmlns:*` bindings, enforces the declaration limit, and breaks on
an attribute iterator error. `pop` at `:762-764` removes that scope. This
scan is needed even when the later checked iterator also visits the attributes:
declarations may occur after ordinary attributes, and element/attribute
resolution must see the declarations before `inspect_element` runs. `add`
also retains reserved-prefix and namespace binding errors. Removing `push`,
calling it only when an element has a non-empty raw attribute tail, or folding
it into the checked iterator without proving error order would change
namespace resolution or refusal behavior.

## Selected next seam: direct scanner event handling

The one source seam justified by these rows is a scanner-local experiment in
`scan_processed_xml`: use `quick_xml::Reader<&[u8]>` plus a local public
`NamespaceResolver::default()`, and dispatch the event after `read_event()`
directly.
The experiment would remove the intermediate `Result<Event>` handoff through
`NsReader::process_event` and its event aggregate reconstruction while
delegating the same resolver operations. It would require adapting the
scanner-local inspector to receive the resolver (or an equivalent narrow
view); it does not require changing quick-xml or the other `NsReader`
callers. The buffered scanner at `codec.rs:851` remains the differential
oracle, and the test helper at `codec.rs:726` is a separate callsite.

The direct loop must preserve this order on every event:

1. Before reading, pop when the prior event set `pending_pop`.
2. Read with the same `Reader` configuration and map parser errors through
   `xml_error`.
3. For Start and Empty, call `resolver.push` before root/name inspection.
4. For Empty and End, set `pending_pop` exactly as `NsReader` does.
5. Keep the existing node, depth, root, DTD/PI/CDATA, EOF, relationship,
   attribute, and byte-limit branches unchanged.
6. Convert a `NamespaceError` through the same quick-xml error display path
   before `xml_error`, then compare refusal text and ordering rather than
   assuming that a direct `Display` call is identical.

This seam targets the 12,051,464 self-IR `process_event` row and its event
return shape. It does not target the 34,941,704 inclusive `push` edge or the
21,471,516 IR attribute iterator beneath it. A successful experiment would
need a paired owner-scoped profile showing what replaces those rows and a
semantic differential against the existing buffered oracle. The source tests
at [`notes/codec.rs:523`](../../../../crates/litchi-pptx/src/notes/codec.rs:523)
cover conformance retry, generic root masking, and limits; the cases at
[`notes/codec.rs:624`](../../../../crates/litchi-pptx/src/notes/codec.rs:624)
cover empty tails, declarations, malformed attributes, duplicate attributes,
unknown prefixes, and invalid UTF-8; and the opened capture matrix at
[`opened/tests.rs:329`](../../../../crates/litchi-pptx/src/opened/tests.rs:329)
through `:567` covers root/name/notes refusal order, proof fallback,
foreign inventory, and the distinct 16 MiB/64 MiB limits. Any changed
conformance, `XmlScan`, error string, or first-refusal position stops the
experiment.

The exact-empty attribute-tail shortcut is already retained; another
checked-iterator rewrite is not selected for the next experiment. The 0806 iterator candidate was
rejected and is not revived here. 0807 provides no measured savings claim for
either alternative; the direct `Reader` plus resolver loop is only the next
profiling/experiment seam.
