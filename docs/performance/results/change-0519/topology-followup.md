# 0519 candidate topology follow-up

The final candidate set contains 12 owned DOCX publication profiles: six API
routes in each of `profile-after-r1` and `profile-after-r2`. Every profile has
exactly one positive `publish_document_commit_to_stream` call. The raw positive
call count for `litchi_opc::xml_splice::validate_source_xml` is zero in every
candidate profile, and the candidate annotation has no XML-validator row. The
semantic DOCX snapshot owner is unchanged: all candidate profiles retain
`litchi_docx::source_backed::Package::main_document_snapshot_with_before`, with
9,306–9,857 Ir across the matrix.

The remaining topology work is in the ZIP preservation path. In every
candidate annotation, the first direct callee of
`SourceBackedPackage::write_topology_to_stream` is
`SourceBackedPackage::write_changed_overlays_with_omissions_and_appended`.
Its inclusive Ir ranges across both profile repeats are:

| Route family | Topology Ir | Direct overlay writer Ir |
| --- | ---: | ---: |
| p128, K1 | 1,309,831–1,310,262 | 1,243,487–1,243,930 |
| p512, K1 | 1,350,046–1,351,289 | 1,187,138–1,187,266 |
| p512, K32 | 1,352,982–1,353,447 | 1,187,420–1,187,878 |

The overlay writer's dominant direct callee is
`soapberry_zip::preserve::PreservationIndex<R>::write_to_with_accounting`
(1,149,180–1,149,404 Ir for p128 and 1,187,138–1,187,878 Ir for p512).
That function then spends 1,064,838–1,065,062 Ir on p128 and
1,093,052–1,093,454 Ir on p512 in
`soapberry_zip::preserve::write_prepared_local`. Its two large direct leaves
are `soapberry_zip::accounting::write_all_counted` (about 533k Ir for p128 and
561k Ir for p512) and
`soapberry_zip::reader_at::ReaderAt::read_exact_at` (about 531k Ir in every
route).

The remaining local attribution is ZIP preservation planning and local-member
replay/source copying. Bulk-copy instruction counts do not establish removable
work or the highest program-level priority. The named fixture carries roughly
512 KiB of raw media, so memory bandwidth and required output copies must be
considered alongside native timing. Any further local investigation should
split these costs while retaining publication guards, then compare its
end-to-end opportunity with the broader OLE2/OOXML queue. Native elapsed-time
and RSS guards remain the acceptance checks for a subsequent change.

Evidence: `profile-comparison.md`, `profile-analysis.json`,
`profile-after-analysis.json`, and the inclusive annotations under
`profile-after-r1/` and `profile-after-r2/`.
