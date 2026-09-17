# Integration review findings

Read-only subagents reviewed the deferred error paths and CFB layout work.
These findings directed implementation follow-ups; the final implementation
records and merged gates establish the resulting code state.

## Deferred OPC graph errors

- Force an orphan PPTX theme before linking it into an authored master.
- Stage public master/layout authoring so a late decode refusal cannot leave
  the package partially modified.
- Resolve XLSB threaded comments against the actual office-document root;
  conventional-path fallback must not swallow a present root's decode error.
- Keep allocation provenance from 0665 when integrating lazy payload storage.
- Force payloads before PPTX revision memo projection and fingerprinting.

The first three fixes are in 7bc6ebc97 (cherry-picked as 03353b254). The
provenance conflict resolution is in c06534899; the memo follow-up is
2fb80e36f. OPC and PPTX focused merged suites passed.

## CFB layout and overlay

- Reopen and read back the complete plan before sink publication.
- Normalize v3 stream size high words, including empty streams and the
  equal-length object overlay path.
- Preserve explicit zero CLSID setters and source metadata during fallback.
- Use checked arithmetic and fallible source-sized allocations.
- State whether the allocator selects all-free or released-only sectors.
- Prove raw directory metadata preservation, not just paths and CLSIDs.
- Report the observed Reuse/Rewrite timing regression honestly; sector
  placement counts do not establish a payload-copy or latency saving.

## Execution budget composition

- Cap actual per-operation task waves by granted I/O width even when a wider
  private pool already exists.
- Keep serial narrowing serial at the low-level ZIP session.
- Serialize private pool admission so concurrent calls cannot replace a live
  pool's retained worker reservation.
- Test actual simultaneous reads, not only budget counters.
- Define when cumulative CPU task admission is consumed on refused attempts.

## XML marker audit

An independent reader reviewed 0677's split-marker handling, physical offsets,
streaming/slice parity, malformed markers and publication refusal tests. No
blocker was found. The packet records the existing limit-precedence edge for
`max_bytes < 3` explicitly.

## Final merged XLS inverse regression

A read-only XLS review confirmed `f14e3c85a` is the narrow correction: the 14
mismatched bytes were outside the root storage's declared mini-stream length,
inside its retained physical sector. They were padding rather than logical
stream data. Clearing this range follows the stream-tail zero-fill contract;
true no-op publication still returns the original bytes. The existing exact
inverse assertion remains unchanged and passes, together with the new CFB
grow/shrink regression.
