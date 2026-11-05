# 0362yba — return the streamed ODP projection candidate

- Reviewed candidate: `f854b651a8bdcda05f496d32eca778a0bcb611f8`.
- Decision: returned for correctness, fidelity and bounded-memory repair.
- This review does not accept the earlier 0362ya heap tradeoff or claim native
  performance qualification.

The candidate replaces materialized `content.xml` input with `BufRead` for
semantic projections. That is a useful direction, but two public output
regressions prevent integration.

`Parser::write_slides_text_from_reader` creates `event_buffer` and passes it to
`read_resolved_event_into` on every iteration without clearing it. The buffer
therefore retains earlier events and grows with the complete XML document.
Clear it before each read, including iterations reached through the many
`continue` branches, and verify retained memory with a growing sequence of
bounded-size slide events. A small fixture passing the text tests does not
establish the promised streaming bound.

`SourceBackedPresentation::write_text_to` now calls
`with_verified_member_reader` directly. Its new error match lacks the
encrypted-member fallback used by `slides_uncached` and `slide_uncached`.
The former `read_content_xml` path used `get_file`, which supports the admitted
decrypted member. Preserve that supported output path, including sink
progress and before/after source-freshness checks. Also verify prologue
DOCTYPE parity: the former materialized reader neutralized that prologue
before semantic parsing.

After those repairs, compare the actual 0358a reference with the candidate,
including first and repeated slide/text queries, streaming text output,
unchanged package output and append/replay publication. The intermediate
eager `Arc<String>` comparison in 0362yb is diagnostic evidence, not a
replacement for that reference. Report total and post-open allocation costs
separately, with exact source, executable and output identities.

The repair is assigned to the ODP worker as 0362yc. The candidate remains out
of the coordinator integration branch until review and validation complete.
