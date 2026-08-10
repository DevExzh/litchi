# Standalone ODC/FODC producer evidence

Checked on 2026-08-10 for the ODC remediation review.

No genuine standalone producer-created `.odc` or `.fodc` is available in the
checked-in corpus, repository history, or the local user filesystem. The local
environment also has no LibreOffice/OpenOffice, Calligra, or OnlyOffice desktop
producer executable, no usable UNO Python bridge, and no installed local
producer package. It therefore cannot create and resave a standalone chart
through a native chart producer for this review.

The repository does contain genuine LibreOffice chart subdocuments inside
producer-created `.fods` and `.fodt` files. Those embedded XML fragments are
useful interoperability inputs for the shared chart reader, but they are not
standalone ODC/FODC packages and are deliberately not copied, repackaged, or
renamed as standalone fixture evidence.

## Primary-source search record

The 2026-08-10 follow-up checked these authoritative sources:

| Source | Pinned revision | Search and result | License/provenance |
|---|---|---|---|
| [LibreOffice core](https://github.com/LibreOffice/core) | `7d915a0d4ade0e8b9cde3fd23f0fd92c066a78e1` (`master`) | Complete, non-truncated recursive trees for `chart2`, `filter`, `odk`, `sc`, `sd`, `sw`, `test`, and `xmloff` contained no `.odc`, `.fodc`, or `.otc`. The repository-wide tree was truncated and therefore is not claimed as complete. | The upstream repository carries its own [license notice](https://github.com/LibreOffice/core/blob/master/LICENSE); no artifact was copied. |
| [Apache OpenOffice](https://github.com/apache/openoffice) | `55c04d1336cb0228ec67f5ea1b5d7cd6aa993e7d` (`trunk`) | Complete, non-truncated recursive `test` and `extras` trees contained no `.odc`, `.fodc`, or `.otc`. The repository-wide tree was truncated and therefore is not claimed as complete. | Apache OpenOffice is distributed under the repository's [Apache license notice](https://github.com/apache/openoffice/blob/trunk/LICENSE); no artifact was copied. |
| [OASIS ODF 1.4 OS Relax NG schema](https://docs.oasis-open.org/office/OpenDocument/v1.4/os/schemas/OpenDocument-v1.4-schema.rng) | ODF 1.4 OASIS Standard, SHA-256 `4034ec6be29205d5fc1ee5f42468ac6ef824287b3aba6d9289032af4fafbda7f` | Defines the optional empty `chart:coordinate-region` under `chart:plot-area` with common draw position and size attributes. A compact crate-local test case is derived from that rule. | Authoritative standards source. No schema or document artifact is vendored, and the derived XML is explicitly not producer evidence. |

General web searches and the LibreOffice OpenGrok path search likewise exposed
no standalone artifact with producer and license provenance. Search-engine
absence is not treated as proof; only the pinned, non-truncated tree results
above are used as repository evidence.

A future fixture qualifies only if its provenance records a native producer
that saved a standalone ODC/FODC artifact. The corresponding test should open
the original artifact, publish a changed file, fully reopen it, and include a
current native-application resave before claiming changed-file interoperability.
