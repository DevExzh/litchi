# Current LibreOffice ODF resave provenance

This directory uses the LibreOffice 26.2.5.2 runtime pinned in
`../../office-interop/PROVENANCE.md`. Evidence always follows genuine input ->
public Litchi semantic change -> isolated-profile LibreOffice save -> Litchi
reopen/readback. Native-only edits, no-ops, embedded subdocuments, and
synthetic inputs are not promoted to evidence.

## Successful chains

| Family | Source SHA-256 | Litchi change and SHA-256 | Filter and resaved SHA-256 | Readback |
|---|---|---|---|---|
| ODT | `test-data/odf/corpus/writer-header-footer.odt`; `fda7c0be9f1135e7a30b05db6d9ddf96020ba00d87478ca4c74084c8742c5a21` | Paragraph 0 changes to `Litchi native resave 2026-08-10`; genuine formatted `meta.xml` remains exact source bytes; `2dd3b2047c89da3352adb7f3b4db027ffdcc77a2e70f4c298645e7815619f952` | `writer8`; `22e95ca413c468c8ec4de96f23f934ff6a4ee860b5089c512dafb2a0d7f74c32` | Reopened text contains the exact sentinel |
| ODS | `test-data/odf/corpus/calc-two-sheets.ods`; `67ed3f8831aa078a849badd8f2a15bdee7cf965a628ff4eb73740aa96ba0d4c0` | First sheet cell `A1` gains formula `of:=40+2`; `f942913fe6057266b9233fe539e078213a46f3903ab5ab54bac5e1076a6415b0` | `calc8`; `5d60813f36fab58a802da50fd7979b7ea503e38c8721598de093f61f60eb111b` | Reopened `content.xml` contains the exact formula |
| ODP | `test-data/odf/odp/tdf169979.odp`; `160908f993c6ba901233695b12d34c4b009142971b36dcd57c0549bf8ee5656b` | A text box named `Litchi Interop Box` containing the sentinel is added to slide 0 while the producer BOM retains exact-source provenance; `2ebbf9efb1a0b26bc60cda62c12874cdd60f0e22027a70e2c1018c44098167fb` | `impress8`; `31401a94864ee0a4b68d39287cf65159b977db798946472f4546a393f1d4a4a9` | The named text box and exact sentinel reopen |
| Formula | LibreOffice `font-styles.odf`; `2abee0da450b31c3bc87d007e85fff21714fee742d6dda9e4987415107ffb27f` | MathML token path `[0,0,0]` changes `f` -> `g` and its retained `StarMath 5.0` source changes to `g`; `8b0ce2f1415ec28d579ae0c000fc8a4570dbb3ad9b81f48366f9654db91b6508` | `math8`; `4c1ca45f31b5ac919fe82b9b09962d0d57bcffc0c801f7eae3e57882f8c5ea7c` | Both the MathML identifier and StarMath source reopen as `g` |
| ODG | LibreOffice `rhbz1870501.odg`; `46530e653ca424fd5b985813cdeeceb9f4b99589c45d8bdeb1b1256badad133f` | A positioned group descendant changes to `x=9cm`; the source ODF 1.2 manifest version is retained; `a7282b53e227fa772876e4b697bb70f7aa77e6d2b1384c5a2f6b483153d5de2a` | `draw8`; `4cf5a94733a11a9a7075284994a191cb87c41173899ee0ebbba592c57232d99c` | The descendant geometry reopens at exactly `x=9cm` |

The Formula edit deliberately changes both representations. A MathML-only
attempt was not counted: LibreOffice correctly regenerated presentation MathML
from the retained StarMath annotation and restored `f`.

At reviewed repository revision
`e925e02c426d5dbbf8b171c7a41b357382ecda79`, the Formula source was copied
byte-for-byte from
`3rdparty/libreoffice-core/starmath/qa/cppunit/data/font-styles.odf`, the ODG
source from `3rdparty/libreoffice-core/sd/qa/unit/data/odg/rhbz1870501.odg`,
and the license from `3rdparty/libreoffice-core/COPYING.MPL`. These copies
retain MPL-2.0 provenance. Their hashes are pinned here and in the checksum
test. Save logs are retained under `logs/`.

## Honest unavailable boundaries

- ODM and OTH were not attempted in this pass. Their absence is not evidence.
- ODB filter `StarOffice XML (Base)` is `IMPORT`-only. Same-package save needs
  live UNO `store()` and is not claimed by the CLI conversion harness.
- ODI has no registered LibreOffice media type or import/export filter.
- ODC has declared import/export filter `chart8`, but current LibreOffice
  `storeAsURL()` disposed the bridge and an in-place `store()` after mutation
  remained byte-identical. That behavior is not changed-file save evidence.

Reproduction:

```sh
cargo run --manifest-path tools/native-resave/Cargo.toml \
  --features ods,odf -- generate FORMAT OUTPUT
LIBREOFFICE_BIN=/temporary/runtime/program/soffice \
  python3 tools/native_odf_resave.py LITCHI_CHANGED DISTINCT_OUTPUT_DIRECTORY
cargo run --manifest-path tools/native-resave/Cargo.toml \
  --features ods,odf -- readback FORMAT LIBREOFFICE_RESAVED
python3 -m unittest tools.test_native_odf_resave
```
