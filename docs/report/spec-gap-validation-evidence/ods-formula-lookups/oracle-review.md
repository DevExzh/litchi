# Independent review of the lookup oracle and native receipt

Status: **PASS for the retained oracle and native artifacts**, with the
evidence-model limitations recorded below. This disposition does not make a
production-support, source-gate, or performance claim. No production or test
files were edited.

The review is bound to contract v2, SHA-256
`b112d66d687337912333f932c199f6e0ee5241fedc335aeaa95de572889e6aaf`, with
these current identities:

* `lookup_oracle.py`: SHA-256
  `2d9541fb319944e1d5bb24f3a79197fec2d9580bf8ba51d16ff9be9e624cbc24`.
* `lookup-goldens.json`: SHA-256
  `42083ceca906fe5842b2f1b30695310b3d23de3a9034e55fd17ff49d71bf4faa`.
* `lookup-functions-native.fods`: SHA-256
  `f084909cae337df25c37e028fa4644eac5c80a68b78522cc6c7ebbb2e592cb60`.
* `recalculated.ods`: SHA-256
  `4a6eb73fb895fffda6654a07a94a77e894210228ef7f4c10943eebbadfc05ae4`.
* Native `content.xml`: SHA-256
  `ecd94bdfc0acfff99eb3890148adaf7267243d11ae2543b987217706fc4a9175`.
* `native-results.json`: SHA-256
  `ecb3f972895a260f95fdfc0285e881d4cea872c701e3ec0752901cfcede4fbb1`.

The baseline commit remains `635fd2e1348b621426b50909cbd5765c91837306`.

## Verification

The current oracle reports 127 observations. Both
`python3 lookup_oracle.py --self-check` and `python3 lookup_oracle.py --check`
complete successfully; the latter reports `verified=true` and reproduces the
retained JSON byte-for-byte. The seven INDEX fixtures now use the evaluator
grammar at lines 746, 753, 760, 767, 774, 780, and 786:
`{1;2|3;4}` is a 2×2 array (`;` within a row and `|` between rows). Their
full-array, row-slice, column-slice, negative/bounds, and area-number
expectations now correspond to that shape.

The prior oracle findings are resolved in the current bytes. Descending
approximate MATCH retains the last qualifying candidate (lines 346–364), the
exact formula-error suffix case expects the required remaining reads, the
short LOOKUP reference case uses a genuinely in-range query, and approximate
reference rows use read envelopes. The model includes ascending and descending
Number/Text barrier cases plus mixed-type approximate midpoint rows. INDEX
list full-area, duplicate-record, and 3-D full-area descriptor rows are
present, and reference expectations carry direct/derived owner identity.

The native receipt independently reproduces with
`python3 native/reproduce.py`, returning `status=verified`, 32 formula rows,
17 documented parities, 13 documented LibreOffice divergences, and the
retained recalculated `content.xml` hash. The native input and recalculated
fixture use the same semicolon-within-row and pipe-between-row grammar. Its
divergence records retain host values such as `Err:504`, materialized range
descriptors, missing Empty-to-zero normalization, and native formula-error
precedence without treating them as normative results.

## Evidence-model limitations

* The Python oracle owns scalar, array, and descriptor types, but it has no
  resolver. Reference results, typed resolver failures, and most read counts
  are authored in `observations()`; `lookup()` itself models horizontal
  vectors only. `--self-check` validates model invariants and
  contract-bound serialization, not a full formula execution against a
  resolver.
* The oracle does not model provider/resource/cancellation/source-version
  failures or work charging. Those obligations require the Rust resource and
  integration suites. Its formula-error rows document the retained-error
  expectation but cannot independently generate typed-failure precedence.
* The native receipt contains 30 function observations from the 127-row
  oracle corpus, plus two fixture-data rows. It does not cover the full INDEX
  list/3-D descriptor matrix, all CHOOSE lazy/reference-list combinations,
  all scalar/matrix selector cases, or the mixed approximate corpus. Native
  spreadsheet cells also materialize references, so descriptor identity is
  represented by the contract expectation rather than observed from the host
  result.
* `native/reproduce.py` verifies the retained LibreOffice bytes, typed host
  outputs, formula projection, and fixture hashes. The normative fields and
  parity/divergence labels in `native-results.json` remain curated evidence;
  the reproducer does not derive them from the Python oracle. LibreOffice
  behavior is therefore an interop observation, not a normative oracle.

These limitations are explicit scope boundaries. Within the retained
contract-bound oracle and native receipt, the corrected fixture and hashes
are stable and verified.

## Independent addendum: sheet-bound extension refusal

The owner correction for `lookup.reference_extension_sheet_bound_is_na` is
supported by an independent rerun of the retained oracle. The case remains
`=LOOKUP(4;[.A1:.A4];[.B16])` with expected `#N/A`; its implementation-
independent read envelope is now **2–4**. Two key probes can reject the
out-of-extent result reference, and no result-cell read is required. The
corresponding JSON entry carries `expected_reads_min: 2` and
`expected_reads_max: 4`.

`python3 lookup_oracle.py --check` verifies the regenerated 127-observation
goldens byte-for-byte, and the contract binding is unchanged. The corrected
artifact identities are:

* `lookup_oracle.py`: SHA-256
  `a9e33442aec45cce4ee361f2bd54daaeb22d7a0778e39b7669cc3ab725f8054b`.
* `lookup-goldens.json`: SHA-256
  `ad2a9bd6c92823b1a40d012d2c64bea66a7b9a0b87c998bb16fc4fcd89887758`.
* `contract.md`: SHA-256
  `b112d66d687337912333f932c199f6e0ee5241fedc335aeaa95de572889e6aaf`.

This addendum supersedes the earlier oracle/goldens pair only for the
corrected evidence snapshot; the preceding review is retained as historical
context. It is an oracle-artifact disposition only and does not certify the
pending production CHOOSE implementation, source freeze, gates, or
performance capture.
