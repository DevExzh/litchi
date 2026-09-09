# Next measured work: integrate existing-document DOCX append

The full non-iWork performance objective remains open. This batch supplies a
bounded reader auditor and a source-checked decoded OPC insertion primitive.
The measured comparison covers XML auditing only. The existing public DOCX
append lifecycle still materializes document XML and retains its paragraph
index; the 0481 measurements remain the applicable end-to-end baseline.

Follow `../change-0481/window-contract.md` when integrating the primitive:

1. Make the DOCX source scanner retain only a finite lexical/parser window,
   insertion offset and source/candidate proof. Track the final opaque
   `sectPr` and bind the locally valid Word namespace prefix. Preserve the
   exact existing XML and reject unsupported ambiguous layouts.
2. Publish through OPC's decoded insertion plan with generic XML validation,
   explicit memory/work/I/O/output budgets and ZIP preservation. Complete
   required semantic candidate readback and transactional patch behavior.
3. Replace retained generated paragraph XML with an explicit replayable
   producer for very large append operations. The current OPC primitive owns
   one finite fragment; its capacity is charged, but this does not establish
   output-size-independent generation.
4. Add durable serialized patch/inverse handling, source-provider variants,
   short range reads, sequential sinks and atomic filesystem save evidence.
5. Repeat the 0481 named lifecycle and scaling studies with normal and
   allocator builds. Report source/fragment growth separately, all latency and
   RSS review flags, logical I/O and unchanged-member preservation. Retain
   current materialized behavior as the measured control until the replacement
   passes the full semantic and resource contract.

Broader Office CRUD coverage, bounded generation across formats, remote and
concurrent scenarios, explicit parallel scaling and remaining legacy paths
still require the program-level completion audit in `docs/GOAL.md`.
