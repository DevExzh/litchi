# 0806 OLE visibility amendment review

Review status: **source-approved for the required production quality rerun**.
This is a static review of the retained visibility amendment. I did not apply
the patch, build, run Cargo, capture native data, or rerun the protected
preflight.

The earlier source-amendment review's statement that candidate public
visibility was unchanged was incorrect. I compared every helper's sealed
`candidate/before` and `candidate/after` declarations:

| Helper | Sealed candidate baseline | Candidate/current quality after |
| --- | --- | --- |
| `litchi-ole-common` | `pub` trait / `pub` struct | `pub(crate)` trait / `pub(crate)` struct |
| `litchi-opc` | `pub` / `pub` | `pub` / `pub` |
| `litchi-sign` | `pub(crate)` / `pub(crate)` | `pub(crate)` / `pub(crate)` |
| `litchi-xldm` | `pub(crate)` / `pub(crate)` | `pub(crate)` / `pub(crate)` |
| `xml-minifier` | `pub(crate)` / `pub(crate)` | `pub(crate)` / `pub(crate)` |

The current workspace and quality-amendment OLE after archive retain the
accidental narrowing. `litchi-ole-common` exports `xml_attributes` publicly,
and `litchi-crypto` imports `BytesStartExt` and calls `checked_attributes()` in
its Agile and labels readers, so the OLE regression explains the downstream
quality failure. The public trait's return type also requires
`CheckedAttributes<'a>` to be public.

The visibility amendment's before file is byte-identical to the current OLE
source and the quality-amendment OLE after archive. Its after file is exactly
the two declaration rewrites below; no other lines or files change:

```diff
-pub(crate) trait BytesStartExt {
+pub trait BytesStartExt {
...
-pub(crate) struct CheckedAttributes<'a> {
+pub struct CheckedAttributes<'a> {
```

The iterator body, helper call routing, state machine, error behavior, tests,
fixtures, and timing scope remain unchanged. The amendment restores the OLE
baseline API and leaves the other four helpers at their recorded visibility.
The full production quality gate must be rerun, including the downstream
`litchi-crypto` check. Because this is visibility-only and does not alter the
timed OPC implementation or the protected native oracle, no new micro-capture
is indicated by this source review.
