# Native iWork fixtures

These compact fixtures were created with the native macOS iWork applications
on 2026-08-07 and are intentionally checked in as package-level compatibility
fixtures. Their canonical single-file package hashes are:

| File | SHA-256 |
| --- | --- |
| `pages/basic.pages` | `21107bc9323fba6f1589152454c0b0b0cc8e239313c6a369bc4a891116601b42` |
| `numbers/basic.numbers` | `f225d5b1cd59e9da454f91a96fe8f81154bc31037c10029230e75d49b45fb693` |
| `keynote/basic.key` | `3a3d07476b45b6e543bcfba75fe38a245434176dcb3565e34570b817708b9f42` |

| File | Expected semantic content |
| --- | --- |
| `pages/basic.pages` | `Litchi native Pages fixture`, `Buffa lazy-view migration verification`, `2026-08-07` |
| `numbers/basic.numbers` | Cell `B2`: `Litchi native Numbers fixture`; cell `B3`: number `42` |
| `keynote/basic.key` | One slide containing `Litchi native Keynote fixture`, `Buffa lazy-view migration verification`, and `2026-08-07` |

Each document was saved, closed, and reopened in Pages, Numbers, or Keynote.
The application accessibility tree confirmed the expected content after the
reopen and no repair prompt was shown. The format-crate integration tests also
open the packages from disk and from borrowed archive bytes.

## Native package-directory oracles

On 2026-08-08, disposable copies of all three canonical files were opened in
their native applications and converted with **File → Advanced → Change File
Type → Package**. Each resulting directory was saved, closed, reopened from
the application's Recents view, checked for the same visible semantics, and
closed without another save. No application displayed a repair or conversion
prompt. Those app-authored directory artifacts are checked in under
`directory/`; `directory/MANIFEST.sha256` records all 46 regular members.
They contain no symbolic links or special nodes.

| Directory artifact | Regular files | Disk usage | `Index.zip` SHA-256 | Reopened tree digest |
| --- | ---: | ---: | --- | --- |
| `directory/pages/basic.pages` | 7 | 116 KiB | `c1d5ed1626ac8652ee5b2fe36f9d2224a39fb690f5ce91fc8c66a8bfe7951665` | `908f121ebeb8f53a5ca5c4b7793a3c0896d2a73d4f1f6fc5c0d96105ff154ffb` |
| `directory/numbers/basic.numbers` | 7 | 152 KiB | `5c62c8553bfcc0891866663fe474cc758bc92f7c91665ecfc6da775ce50f504a` | `5d280f54308435d099d9a3ef4a6a31b5de74c5086130d908ff80957d1fd36f98` |
| `directory/keynote/basic.key` | 32 | 552 KiB | `c0e3039a597723abd32d2f7f2f9b2fa96826bff17fc2fa1d715e7402ae575acd` | `298190c828b5be72cd06de47f5a8c95995e043f73aab915777ca6cd5e1725c77` |

The tree digest is the SHA-256 of the sorted per-file SHA-256 manifest captured
after native reopen; it is provenance evidence, not a package-format checksum.
The read-only root directory adapter deliberately freezes only `Index.zip` or
loose `Index/` semantics. It does not claim exact ZIP provenance, editing, or
preservation of `Metadata/`, `Data/`, previews, or unknown sidecars.
