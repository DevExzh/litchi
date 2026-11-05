# 0362ul86c — final ZIP alias-closure integration proposal

- Date: 2026-10-25.
- Tested source: `e49a1062ecf59a8cf458e564af85d8880c4e9adb`.
- Parent integration at proposal time:
  `ad138ffcedb5afc8ddbdbcc69c1e0c27e064cc77`.
- Common ancestor: `4a62cdfb035cd43cbd0a28c8d1d67fbf03e19e41`.
- Status: the source lane and its focused validation are accepted; this
  record proposes one coordinator-owned atomic merge and does not perform it.

## Complete lane

The complete source lane is the following contiguous chain from the common
ancestor:

1. `7e2483fa75e3e4a69ddd76b02db492d1031add0c` removes the ZIP writer entry
   aliases and the compatibility `ErrorKind::Io` variant, and migrates every
   workspace consumer to `start_file_owned*` and `ErrorKind::IO`.
2. `a84f5761388ab5f891d4592e7feea8c30da076ab` removes the duplicate
   multiline `ErrorKind::IO(_)` matcher found during source review.
3. `e49a1062ecf59a8cf458e564af85d8880c4e9adb` repairs formatting on all
   nineteen Rust paths changed by the first two commits and records the
   scoped formatter policy.

The affected consumers are the Soapberry ZIP writer/error/office/preservation
code, tests and fuzz target; ODF package, detection, writer and source
publication consumers; OPC source-backed/error mapping; and the unified ops
ZIP classifier. The two glossary surfaces register the removed names and keep
the canonical `start_file_owned*` and `ErrorKind::IO` vocabulary. The three
lane records are:

- `0362ul86-zip-compatibility-aliases.md`;
- `0362ul86a-zip-alias-audit-correction.md`; and
- `0362ul86b-zip-formatting-repair.md`.

No deprecated alias, compatibility variant, public mutable-owner API, error
text, source chain, output layout, compression policy or limit policy is
retained by this lane.

## Exact tested-source comparison

The guarded capture tested the clean source blob at `e49a1062ec` exactly. A
read-only synthetic merge was then computed with:

```text
git merge-tree --write-tree \
  ad138ffcedb5afc8ddbdbcc69c1e0c27e064cc77 \
  e49a1062ecf59a8cf458e564af85d8880c4e9adb
```

It is conflict-free and produced tree
`0fac0c06905055ccbc2390aac1189a96ba68da7b`. The 24 paths changed by the ZIP
lane compare as follows:

| Paths | Count | Merge-tree result |
| --- | ---: | --- |
| Rust consumers, tests and fuzz target | 19 | exact e49 blobs |
| 0362ul86, 0362ul86a and 0362ul86b records | 3 | exact e49 blobs |
| `docs/API_GLOSSARY.md` and `tools/api_glossary.json` | 2 | intentional additive merge |

The path/blob comparison manifest has SHA-256
`05df3467c010106bbaf05742b87468e2f8085a06ddc11729cd937a28476b823d`.
The additive glossary result retains the parent’s 69 profiles, including the
Numbers profile, and also retains the ZIP bans for `start_entry_owned`,
`start_entry_owned_zip64` and `Io`; the tested Soapberry ZIP profile remains
unchanged in meaning.

## Validation receipts

The fresh guarded state is
`(deleted-coord-artifact)`.
All eight rows passed: scoped formatter check, ZIP all-target tests, ZIP
strict clippy, ODF streaming test, OPC source-backed test, ops/facade ODF
checks and the Soapberry glossary check. The run covered 632 ZIP, 14 ODF and
302 OPC tests (948 total). Its final target size was 835,851,100 bytes.

- `rows.json`: 30,562 bytes,
  SHA-256 `31fd779c4dcf66cc8f49a8affb464d178e737382c2be2db7766e59d6e240668d`;
- `result.json`: 32,777 bytes,
  SHA-256 `e61de4e9dc48a1a810b0c1ed05b17643a40e632fd2453f44581d0cc484da20f8`.

The earlier source-format failure is preserved at
`(deleted-coord-artifact)`:
its first-row log SHA-256 is
`6444246d15f665bdb857146323f77dc64fb61e7542e483b7c1f5cfa4ffe23b53`, with
rows/result SHA-256 `f27bbaeecbc01aee2f05d46ebf56d457fe34169c478dacf8d9ab2e9241224513`
and `66a7cf7a8d65cbe805c46ad44810d3aa422f5799dcebb5553016b64d1d26b781`.
It remains a failed workspace-wide formatting receipt; the scoped repair does
not relabel or delete it.

## One atomic coordinator proposal

After reviewing this record, the coordinator should merge the branch tip that
contains this record into the current integration at
`ad138ffcedb5afc8ddbdbcc69c1e0c27e064cc77` with one `--no-ff` merge. That
single merge must include the complete 0362ul86 chain, all migrated
consumers, the two glossary updates and all four lane records including this
proposal record. Do not cherry-pick only the formatter commit or merge the
source commits separately. No merge or additional validation was performed in
this worktree after the accepted e49 capture.
