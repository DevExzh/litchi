# 0362ul71 — finite raw-API coverage census

## Scope

This leaf adds a callable census for the configured-limits audit. It is based
on the current source commit `ef1acb5ef0b776c9a6066029a5cac35daeb0b5eb` and
does not run Cargo or generate rustdoc. Rustdoc JSON is an input supplied by a
final gate after the source is frozen.

`tools/raw_api_coverage.py` reuses the glossary gate's `Surface` and nominal
type resolver. It walks the exact public callable paths in each supplied
rustdoc document. It does not infer a raw boundary from names such as
`parse`, `codec`, or `read`.

Each manifest entry requires:

- a profile and exact public path;
- one explicit disposition: `required_raw_limits`, `configured_owner`,
  `retained_owner`, `semantic_value`, `fixed_scalar`, `private_removed`, or
  `active_lane`;
- a source path, full blob identity, positive line, and symbol anchor; and
- non-empty evidence.

`required_raw_limits` entries additionally require the exact rustdoc input-name
list, limits argument, and nominal limits type. The checker resolves aliases by
rustdoc identity, so a same-spelled unrelated type cannot satisfy the row.
Owner and semantic rows are explicit source-bound classifications, not generic
crate exemptions. `private_removed` is stale if the callable is still public.

## Honest initial state

`tools/raw_api_coverage.json` records the verified current examples and active
returns 0362ul65 through 0362ul70, together with the existing ODF, font, crypto,
and PPTX raw lanes. Its state is deliberately `incomplete`: the final
all-profile rustdoc inputs and every public callable disposition have not been
assembled yet. It therefore cannot claim that the raw audit is complete.

Normal validation reports this state and exits zero so the manifest can be
reviewed while waves are in flight. `--require-complete` exits nonzero until
all glossary profiles and workspace packages are represented, all rustdoc
inputs are supplied, every callable has one exact entry, source identities and
limits signatures match, and no active, stale, or unclassified row remains.

The expected profile source is `tools/api_glossary.json`; the expected package
source is the workspace `Cargo.toml`. Custom source profiles are allowed for a
package before its glossary row exists, but workspace-package coverage still
requires that package to be represented.

## Checks performed

On this leaf:

```text
python3 -m unittest -v tools.test_raw_api_coverage
7 tests passed
python3 -m py_compile tools/raw_api_coverage.py
python3 tools/raw_api_coverage.py --manifest tools/raw_api_coverage.json --root . --glossary tools/api_glossary.json --json
exit 0; complete=false; loaded=0; missing glossary profiles=53; active lanes=12
python3 tools/raw_api_coverage.py --manifest tools/raw_api_coverage.json --root . --glossary tools/api_glossary.json --require-complete
exit 1 (expected: incomplete manifest and sparse checkout have no final rustdoc/Cargo inputs)
```

The synthetic rustdoc tests cover exact callable census, ordinary semantic
classification, raw limits type identity, unclassified callables, stale
private-removal rows, active-lane failure, generic-disposition rejection, and
source-blob tampering. The source-identity test uses the repository's current
tracked `tools/check_api_glossary.py` blob; no generated evidence is committed.

## Final-gate integration

The gate should register `python3 -m unittest tools.test_raw_api_coverage` as
the tooling unit row. Once the final all-features rustdoc artifacts exist, the
coordinator should bind each artifact as `--rustdoc PROFILE=PATH` and run:

```text
python3 tools/raw_api_coverage.py \
  --manifest tools/raw_api_coverage.json \
  --glossary tools/api_glossary.json \
  --require-complete \
  --rustdoc PROFILE=PATH ...
```

The final mode is intentionally dependent on those frozen rustdoc inputs. No
source-only run can prove zero unclassified public callables.
