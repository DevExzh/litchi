# Shared chart formatcode2

The implementation is committed as `625b7e121a10ee46b502a07ffe008fabc922201e`.
The retained gates passed 1,189 tests (12 ignored), including 33 focused codec
tests, plus compile, Clippy, rustdoc, formatting, and diff checks. Three generated
elements passed offline XSD validation. The source manifest includes the
preexisting OPC import-order edit present during validation; that edit is not
part of this feature commit.

The [performance evidence](../chart-formatcode2-performance/README.md) contains
1,680 samples across 28 lanes. Clones and unchanged sink writes allocated no
output memory in the measured cases. These are absolute measurements on
generated inputs, with no speedup or native Office acceptance claim.

`chart::extension::formatcode2` provides the exact MS-ODRAWXML §2.44 element and
qualified-attribute vocabulary. `Value` is the bounded semantic escaped string;
`Element` owns a standalone element, and `Attribute` retains a caller-selected
start tag. New attribute values use `write_attribute_value`; source-backed
attribute edits use `write_attribute`.

Readers decode XML and ST_Xstring separately. Literal escape-looking strings,
escaped controls, and UTF-16 surrogate pairs are handled without recursive
decoding. Writers apply the documented Office context rules for CR, tab, and LF.
Untouched parsed objects replay their exact source, including comments, CDATA,
namespace spelling, quote style, and surrounding legal XML markup. Changed
values retain unrelated source bytes. Shared-source reads retain an existing
`Arc<[u8]>`, and unchanged sink writes avoid creating an output vector.

Attribute helpers validate one explicit start tag. The `*_with_bindings`
variants accept inherited namespace bindings, with local declarations taking
precedence. They retain the caller's original tag, rather than publishing
decorated namespace text. The temporary namespace-complete view must also fit
the XML/attribute caps, so inherited declarations consume temporary headroom.
An unqualified `formatcode2` is not the global
qualified attribute. A caller must validate the parent and placement before
using this helper; the module does not search a chart for arbitrary descendants
or infer an extension owner from matching text.

The profile bounds complete XML, semantic and encoded string sizes, attributes,
namespace prefixes/URIs, and output construction. Invalid XML, ambiguous
attributes, invalid surrogate sequences, and limit violations fail before an
output is returned. These shared operations do not create package
relationships, select a chart, evaluate number formats, render charts, or
provide durable host patches.

The [implementation boundary](implementation.md) records the normative sources
and the negative native-fixture search. `corpus-scan.json` retains inspected
member hashes and skipped files. No native Office acceptance is claimed.

Reproduce the compile-first gates from the repository root:

```sh
python3 docs/report/spec-gap-validation-evidence/chart-formatcode2/run_checks.py
```

The downstream harness exercises element no-op/edit, opaque comment retention,
semantic controls and literal escapes, qualified attribute editing, and
inherited attribute context. Its three normal element outputs are checked
against vendored MS-ODRAWXML/ECMA XSDs. The attribute start-tag outputs are
diagnostic fragments, not schema-valid host-placement evidence.

Run the harness and schema checks with source-hash guards, then verify both
evidence sets:

```sh
python3 docs/report/spec-gap-validation-evidence/chart-formatcode2/run_probe.py
python3 docs/report/spec-gap-validation-evidence/chart-formatcode2/verify.py
python3 docs/report/spec-gap-validation-evidence/chart-formatcode2-performance/verify.py
```

The underlying commands are:

```sh
CARGO_TARGET_DIR=target cargo run --locked --offline \
  --manifest-path docs/report/spec-gap-validation-evidence/chart-formatcode2/harness/Cargo.toml \
  -- docs/report/spec-gap-validation-evidence/chart-formatcode2/outputs
python3 docs/report/spec-gap-validation-evidence/chart-formatcode2/validate_schema.py \
  docs/report/spec-gap-validation-evidence/chart-formatcode2/outputs/source.xml \
  docs/report/spec-gap-validation-evidence/chart-formatcode2/outputs/edited.xml \
  docs/report/spec-gap-validation-evidence/chart-formatcode2/outputs/authored.xml
```

An isolated worktree without the local ECMA archive can pass `--spec-root` to
the validator. The schema packaging adaptation is documented in its source;
XSD validation does not prove ST_Xstring decoding or chart owner placement.
