# 0362ul87an — migrate OOXML codec test consumers to transaction errors

The approved F5 run for `841d46e353e3bcb83b317079ae4d9b756a4c83d0`
(tree `f71f5a82d4893713482f05b8ab397df8c7a3ca23`) passed `fmt-scoped` and
stopped at row 1, `check-common`. The release dead-field repair succeeded;
the test build then found four stale consumers of the former package-level
publisher error shape. Custom-properties and extended-properties test
publish helpers returned `crate::Result`, although their canonical
`properties::Edit` calls return `properties::Error`.

The helpers now return `std::result::Result<bool, crate::properties::Error>`.
Signed publication assertions use the transaction's typed `Signed` error,
and the application graph assertion unwraps the transaction's `Package`
wrapper around the shared relationship error. This keeps the tests on the
shared Edit/Commit API and preserves the codec-level `crate::Error`
assertions used by the decoding tests.

The preserved receipts are `rows.json` SHA-256
`7f377cd6cd96e3ba8a22b0dba6291a180ecf281d0a54fc7a456cb9fa6f0f4619`,
`result.json` SHA-256
`6af50991460a251ea4229d0052d7e54acd5b831129b26c0eb13a74ab17d4ffa7`, and
the `check-common` log SHA-256
`0fb6bf6be89e785dd8dd78e6c93352b1deb7ad21ddc462935372124c84201c7c`.
No later row launched. The next attempt must bind a fresh source, target,
plan, driver, manifest, and fixture inventory.
