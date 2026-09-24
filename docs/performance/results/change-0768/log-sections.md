# Log sections for change 0768

Ready-to-paste paragraphs for the coordinator, one per shared log. This change
does not edit `HOTSPOTS.md`, `REPORT.md` or `GOAL_AUDIT.md` itself; the numbers
are the record's.

---

## For `HOTSPOTS.md`

## 0768 — DOC protection is classified from its protection fields, a saved caret no longer blocks an insertion at its CP, and full FKP pages no longer corrupt their BX array

[0768](0768-doc-protection-classification.md) replaces the exact-grammar DOC
protection classifier merged in 0759 with MS-DOC's protection fields: the FIB
takes `FibBase.nFib` when `cswNew` is 0 (2.5.14, with 0x00C0/0x00C2 read as
0x00C1 per note <11>) while `cbRgFcLcb` must still match the generation; the DOP
must hold its generation's protection fields (≥ 84 bytes, ≥ 600 from Word 2003,
674/690/694 for 0x0112) instead of an exact producer length; and only the
`DopBase` locks, `lKeyProtDoc` and `Dop2003` byte 598 are read, with their MUSTs
still `Unrecognized`. The 24 of 38 fixtures the strict classifier refused
(none protected) now classify as `None`, and all 35 readable fixtures accept a
tracked insertion at CP 0 after `Selsf` keeps a recorded CP in place for an
insertion there (the one uniform choice MS-DOC 2.9.244's row and line-start MUSTs
allow). Testing it exposed a pre-existing FKP builder overflow (the first
property placed below offset 511 can take one byte more than the even-rounded
estimate, overwriting the last BX), now trimmed; three real documents'
edited output reopens again. Timed-region instructions change by −0.83% to
+0.27% on the DOC controls (`classify` −81%); the harness `large` shape swings
−20% to +31% in wall clock with glibc heap state and is 1.000 with malloc
thresholds pinned. Open: `fStyleLockEnforced` is not an editing restriction for
the classifier; `validate_fib_shape`'s `FibRgCswNew` bound is two bytes short.

---

## For `REPORT.md`

## 0768 — DOC protection classification follows MS-DOC's protection fields

[0768](0768-doc-protection-classification.md) is a correctness change with
`performance_claim: none`. After it, 49 of the repository's 57 `.doc` files
classify as unprotected (17 before), all 35 readable files in `test-data/ole/doc`
accept a tracked insertion at CP 0 (9 before), and every refusal left is
protection-independent (encryption, pre-Word 97 FIBs, invalid CFB, a malformed
SPRM sequence, an out-of-range `Selsf`). Protected states remain refused by
default and malformed protection records under every policy; the tests cover each
`DopBase` lock, the password hash, enforced `Dop2003` modes in a LibreOffice
610-byte DOP, the SHOULD-level `fProtEnabled` combinations (now `Document`) and
the protection-field MUSTs (still `Unrecognized`). Exact instructions of the
controls' timed regions (NoHeadFoot.doc and FloatingPictures.doc public
lifecycle, harness `doc_semantic_one_edit_save` tiny and large) change by −0.83%
to +0.27%. Wall clock: `docnohf` 0.990/0.992, `docfloat` 1.072 (batch 1, load
23–25, not reproduced) and 0.994, harness tiny 0.967/0.962, harness large
0.785/0.804 with both shapes in one process but 1.27–1.31 alone, tracking page
faults; 1.000 with glibc's mmap/trim thresholds pinned. The same change fixes an
FKP page overflow in the CHPX and PAPX page builders that made edited output of
three real documents unreadable.

---

## For `GOAL_AUDIT.md`

## 0768 — DOC protection classification and FKP page capacity (correctness)

[0768](0768-doc-protection-classification.md) keeps ADR 0006's "Office protection
is enforced by default" and GOAL rule 3: `ProtectionPolicy::authorize` is
unchanged, document and range protection are refused without an explicit
`AllowProtected` capability, and a malformed or incomplete protection record is
refused under every policy. What moves is the classifier's reading of MS-DOC:
fields without protection meaning (and those MS-DOC says to ignore) no longer
decide a protection verdict, which lifts non-bypassable refusals from 24 unprotected
fixtures, including every LibreOffice-written DOC. The `Selsf` remap now treats an
insertion at a recorded CP as the empty-range case of its existing start-boundary
rule; genuinely ambiguous CPs (inside replaced or removed text) stay refused. The
FKP fix changes output only where the builders previously wrote a page their own
parser refuses; every golden is unchanged. Gates pass except the known
`non_iwork_gate` `litchi-xldm` inventory failure, identical at the base.
