# 0359sa — Return legacy raw replay error classification

The coordinator returns the raw replay classification conclusions in 0359s
through `51e8e8bdbd`. The record's public DOC conflict renames and validation
receipts remain useful, but its table treats a stale raw patch base as
`InvalidFormat`, `Corrupted`, or a generic crypto `Invalid` error. Those are
not accepted exceptions to 0351 §2.4.

That section applies to every module error returned by staging or apply. A
patch or staged plan made against another base returns typed `PatchConflict`
and reaches the unified facade as `invalid_input`. Malformed serialized input
continues to use the format's corruption error, and positional source changes
remain `SourceChanged`/`io_error`.

Repair 0359sb must migrate the seven audited public raw apply paths across OLE
property sets/OLE1, crypto DataSpaces, DOC DOFR and saved-selection records,
and the DOC package patch. Keep crate-owned error boundaries and migrate all
consumers in the same change. Test actual stale replay results rather than
constructing an unrelated malformed-input error for the facade assertion.
Independently reopened matching content must continue to replay, with the
original proof scope and recipient limits intact. Malformed-input negative
controls must still classify as corruption. Any error-wrapper allocation or
ABI cost introduced by the repair needs measurement; no performance waiver is
created by this error-taxonomy correction.
