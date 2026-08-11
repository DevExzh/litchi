# Metafile corpus provenance

The `.emf` and `.wmf` fixtures in this directory are vendored test inputs,
not an optional `3rdparty/` checkout. They preserve the source-relative paths
from the upstream repositories so the corpus tests can scan only tracked
`test-data/` files.

- LibreOffice core: commit `77ea61b9866df2203e77df2157f57e1456d68b29`
  (master on 2026-08-07), all `.emf` and `.wmf` files.
- Apache POI: commit `5be0b22971379fcc7036f4e6fb6e38c0471ddcf4`
  (trunk on 2026-08-07), all `.emf` and `.wmf` files, under
  `test-data/poi/`.

## License and notice coverage

LibreOffice-origin fixtures are covered by MPL-2.0; the retained license text
is `test-data/odf/native-resave/source/LICENSE-MPL-2.0.txt`.

Apache POI-origin fixtures are covered by Apache-2.0; the retained license
text is the repository-root `LICENSE`. The required upstream notice for the
pinned POI source is retained verbatim at `test-data/poi/NOTICE` from
`legal/NOTICE` at commit `bbf2e879c36fcd837fd1e7579f9f82cfba88883e`.
That same POI revision supplies the vendored `.xlsb` and `.vba` fixtures.

The current tracked corpus contains 182 metafiles, maintaining the
`>= 175` coverage floor in `crates/litchi-imgconv/tests/metafile_conformance.rs`.
