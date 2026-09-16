# Change 0640: the two refusals this change unmasked

Removing the `grffldEnd` consistency refusals lets each witness fixture reach a
further refusal. Neither is in change 0640's brief and neither is changed by it.
This file freezes what a bounded read-only investigation established, so the
next batch can start from decoded bytes rather than from the error string. Both
verdicts below are **provisional**: they are this investigation's reading, not a
disposition, and each rests on a specification claim this batch could not
verify.

Nothing here is authorized by change 0640.

---

## A. `ole/doc/watermark.doc` — `bookmark ibkl values must be unique and in range`

Site: `crates/litchi-doc/src/parts/bookmarks.rs`, `BookmarksTable::parse`,
line 83:

```rust
let end_index = usize::from(read_u16(property, 0, "bookmark ibkl")?);
if end_index >= ends.len() - 1 || !used_end_indexes.insert(end_index) {
    return Err(PackageError::Corrupted(
        "bookmark ibkl values must be unique and in range".to_string(),
    ));
}
```

Decoded from the `1Table` stream (FIB index 21 `SttbfBkmk` fc=5300 lcb=386,
index 22 `PlcfBkf` fc=5686 lcb=76, index 23 `PlcfBkl` fc=5762 lcb=40):

- 9 bookmark names: `_Ref189041054`, `_Ref189041043`, `_Ref189041491`,
  `_Ref189041649`, `__RefHeading__8_476954814`, `__RefHeading__10_476954814`,
  `__RefHeading__12_476954814`, `__RefHeading__14_476954814`,
  `__RefHeading__16_476954814`
- `PlcfBkf` CPs: `[0, 0, 415, 415, 416, 425, 536, 1083, 1108, 1466]`
- `FBKF.ibkl` in order: `[0, 1, 4, 4, 2, 3, 6, 7, 8]`; every `BKC` is `0x0000`
- `PlcfBkl` CPs: `[415, 415, 416, 425, 535, 535, 536, 1083, 1108, 1466]`

Nothing is out of range: the bound is `< ends.len() - 1 == 9` and every `ibkl`
lies in `0..=8`. The only violation is that `ibkl` 4 appears twice — on
`_Ref189041491` and `_Ref189041649`, which both start at CP 415 — while `ibkl` 5
is never used.

The fact that decides it: `PlcfBkl[4] == PlcfBkl[5] == 535`. The duplicated slot
and the unused slot hold the same CP, so the second bookmark's end is 535 either
way. With the uniqueness half removed the file yields nine well-formed
bookmarks; every `start <= end` holds and every `BKC` is zero.

Specification status: the code's own comments cite [MS-DOC] 2.8.10 only for "the
final CP of a bookmark PLC is ignored" (`bookmarks.rs:63-65`, `:196-198`).
Neither the comments nor the code cite any clause requiring `ibkl` uniqueness;
the `HashSet` is unexplained. **Unverified** whether [MS-DOC] states such a MUST,
and **unverified** whether any sentinel `ibkl` value is defined. Structurally a
repeated `ibkl` is a many-to-one map from starts to ends, which is degenerate
but determined — unlike a repeated *start* CP, which this same file has (0,0 and
415,415) and which the code already accepts.

**Provisional verdict: stricter than the format.** The range half is sound and
should stay; the uniqueness half is not derived from a cited clause and rejects
a file whose bookmarks are fully determined.

---

## B. `poi/test-data/document/test.doc` — `malformed SPRM sequence: truncated SPRM opcode at byte 3: 1 byte(s) remain`

Failing step: `PapBinTable::parse`, called with `?` from
`crates/litchi-doc/src/document/package.rs`, at
`crates/litchi-doc/src/parts/pap_bin_table.rs:259`
(`let parsed = parse_sprms(direct_sprms)?;`). Elimination: `ChpBinTable::parse`
returns an `Option` and swallows its errors; `SectionsTable` parses this file's
single `SEPX` cleanly (fc=3584, cb=36, 9 SPRMs, consumed == 36); the list and
proofing tables use `.ok()`; the stylesheet parses (`structure/stsh-poi-test.txt`
shows `strict parse: OK`).

The structure: the `PAPX` FKP on `WordDocument` page 5 (`PlcfBtePapx` FIB index
13, fc=506 lcb=20, pn = {5, 6}). Its first `PapxInFkp`, at word offset 502, has
`cb = 0`, so the length comes from `cb'` = 3 and `grpprlInPapx` is 6 bytes:
`00 00 | 31 24 00 00` — `istd` = 0, then a 4-byte grpprl. `0x2431` is
`SPRM_P_F_WIDOW_CONTROL` (`sprm_operations/model.rs:467`), size code 1, so the
SPRM is 3 bytes and a single `0x00` follows. All 24 `PAPX` entries in the file
have that shape (grpprl lengths 4, 58, 58, 62, 78, each ending in one extra
`0x00`).

The condition, `crates/litchi-doc/src/sprm.rs:194-204`:

```rust
while offset < grpprl.len() {
    let remaining = grpprl.len() - offset;
    if remaining < 2 {
        return Err(Error::Opcode { at: offset, remaining });
    }
```

What the byte most likely is: a word-alignment pad. `PapxInFkp` has two length
encodings — `cb != 0` gives `2*cb - 1` bytes (odd), `cb == 0` gives `2*cb'`
bytes (even) — so the stored run is always even. Here `istd` plus grpprl is 5
bytes, and the producer used the even form with `cb' = 3` = 6 bytes,
overshooting the real content by one. `litchi-doc` implements both forms
(`parts/fkp.rs:341-357`) and its own writer always picks the form that fits
exactly (`writer/fkp.rs:302-319`: `let extra = if (len % 2) > 0 { 1 } else { 2 };`),
so litchi never emits a pad byte and the case is untested by round-trip. Genuine
truncation is ruled out: the trailing byte is `0x00`, it is present on every
entry, and the declared length is over-long rather than short.

**Provisional verdict: stricter than the format.** `parse_sprms` is an
exact-sequence parser with no allowance for the pad byte the `cb == 0`
encoding inherently produces. A narrow fix would tolerate a single trailing
`0x00` **only at the `PapxInFkp` call site**, not loosen `parse_sprms` globally;
the `SEPX` site (`parts/sections/codec.rs:118-124`) already asserts exact
consumption and passes on this file. The claim that Apache POI's
`SprmIterator.hasNext()` uses `_offset < _grpprl.length - 1`, which would be the
same tolerance, is **unverified** — POI's source is not in this repository.

---

## Why neither was acted on in change 0640

Change 0640's brief names seven refusals and says to keep every other refusal.
Both of these move when a refusal happens, which the wave-2 briefing makes a
stop-and-report condition rather than something to fold into an adjacent change,
and both rest on specification readings this batch could not verify. They belong
to different subsystems (`parts/bookmarks.rs`; `sprm.rs` with
`parts/pap_bin_table.rs` and `parts/fkp.rs`) and each needs its own corpus
differential.
