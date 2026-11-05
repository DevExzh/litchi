# 0362ul67 — VBA raw policy and retained payload boundary

Status: source-ready packet on `uo2/vba-raw-policy-0362ul67`, based on
`ef1acb5ef0b776c9a6066029a5cac35daeb0b5eb`. The PPT owner-policy correction
and DOC/XLS glossary closure are in follow-up `0362ul67a`. Validation remains
coordinator work; this packet deliberately contains no Cargo, native, rustdoc,
benchmark, or macro-execution result.

## Boundary classification

`litchi_vba::project::Project` is a raw parsed descriptor, not an immutable
package owner. It does not own an ingress archive or a retained policy, so its
canonical entry point is now `Project::read(bytes, &Limits)`. The former
limit-less `read` and `read_with` pair is removed; no compatibility alias is
left. Its public directory, text, decompression, parse and encode boundaries
already take the module's required `&Limits` and are registered in the new
`vba` glossary profile.

`litchi_vba::build::Payload` is different: it owns validated standalone CFB
bytes. It now retains the exact `Limits` used by `Payload::read` or
`build::Project::finish`. Its `&self` `write_into` route reopens the payload
under that retained physical policy and preflights every copied stream before
materialization. This keeps the owner immutable while preserving the existing
byte-fidelity and producer-extension copying behavior.

## Nested CFB and allocation policy

The standalone parser and retained-payload writer no longer call
`OleFileLimits::default()`. A crate-private CFB policy adapter derives input,
directory and allocation-table ceilings from `Limits::max_cfb_bytes` before
CFB parsing can allocate. Known VBA streams continue to use the existing
preflighted `read_limited_stream` path; retained unknown streams copied by
`Payload::write_into` use the same `CfbStreamBytes` admission before
`open_stream` materializes them. MS-OVBA compressed-container encode/decode
continues to reject compressed and decompressed budgets before output growth,
and project/dir/text writers retain their typed `LimitKind` errors.

No macro is compiled, interpreted, or executed. Unchanged project and stream
representations retain their existing source bytes, including compressed
`dir`, project text, module source, `PROJECTwm`, and producer extension data.

## Consumers migrated

The atomic source change migrates the shared OOXML parser and every discovered
default VBA-project parser consumer:

* `litchi-ooxml-common::vba::read_project_part` calls
  `Project::read(part.blob(), limits)`;
* DOCX `vba_project::VbaProject::project`, its relationship test, and the
  encrypted DOCX integration test take the required `&litchi_vba::Limits`;
* XLSX Active-X `Project::project` and its focused limit test take the required
  profile;
* XLSB `host::VbaProject::project` and its relationship/preservation tests take
  the required profile;
* existing DOC, XLS, PPTX and performance consumers that use `Project::open`,
  `Payload::read`, or relationship discovery retain their existing explicit
  policy paths; no no-argument `Project::read` consumer remains.

The glossary JSON adds the `litchi-vba` raw profile and required signatures for
codec, directory, project, text, builder and payload boundaries. The feature
profiles for OOXML common, DOCX, XLSX and XLSB register their relationship
parser signatures as well. The human glossary records the VBA raw row beside
the CFB and ZIP rows.

## Source checks

```text
api_glossary.json JSON parse and duplicate-key check  PASS
rustfmt --edition 2024 --check (changed Rust files) PASS
git diff --check                                  PASS
legacy Project::read_with/project_with search     PASS (no matches)
staged-size and binary guards                     pending until commit staging
```

The coordinator should run the affected feature builds, glossary/rustdoc
profiles, full consumer gate, and ABBA review after merging this packet.
