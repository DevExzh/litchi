# 0362ul39 — PPT raw record limits

This leaf completes the PPT raw/record limit migration. Public variable
record decoders now require a borrowed, selected limits profile; the former
implicit-default entry points are gone. `Record::parse` and
`RecordParser::parse_document` already require `&SourceLimits`, and the
format families use their canonical module profiles.

The external-media `Video`, `Movie`, `LinkedAudio`, `EmbeddedWav`, `CdAudio`
and `Collection` decoders now require `&external_media::Limits`. Their nested
PPT record sequences and OfficeArt owner traversal use the selected profile,
and collection object admission is bounded before growing the typed vectors.
The fixed eight-byte `Media` atom remains a fixed-width value parser and does
not gain a meaningless limits argument. Document-comparison
`SlideListTable10::parse` now requires `&document_comparison::Limits` and
checks the configured review-record ceiling before `Vec::with_capacity`.

The presentation projection maps its retained source profile into the
external-media profile. The OLE-only projection now receives the retained
`SourceLimits` profile at its public collection and container parsers. Its
cross-family media and hyperlink namespace checks use profiles derived from
that same selection; no parser path falls back to an implicit media or
hyperlink default. Variable OLE strings, metafiles, unknown records, and
container/list serialization are checked against the selected output profile
before growing the emitted record.

Hyperlink collection, base-link, and PowerPoint 9 extension parsers likewise
take the selected `SourceLimits` profile. Their nested versioned-tag and
string decoding is bounded by that profile before allocation, preserving the
fixed-width atom exception only for the fixed-width atom codecs.
InteractiveInfo parsing and text-run readback derive their module profile from
the selected source limits as well; the public `Interaction` record and byte
parsers no longer carry an implicit default.

The PPT validation names are canonicalized to `validation::{Error, Limits}`;
all crate, ops, example, test and tool consumers were migrated in this leaf.
The glossary has 60 required-signature rows in both `ppt-minimal` and
`ppt-full` profiles, including the five variable external-media decoders, the
collection parser, `SlideListTable10`, the OLE collection/container parsers,
the hyperlink collection/link/extension parsers, and both `Interaction` record
entry points.

## Validation

The approved `Cargo.lock` is SHA-256
`4100b1fbb272dd24842f33a2d6f353a3f3a97dfa7e435c4649a0e1a029bfaea1`.
The capped coordinator cargo shim was used with one build job, offline locked
resolution, incremental compilation disabled, and target
`(deleted-coord-artifact)`.

Focused commands passed:

* `cargo check -p litchi-ppt --all-features --locked --offline`;
* `cargo test -p litchi-ppt --all-features --locked --offline external_media::tests:: -- --test-threads=1` (19 passed);
* `cargo test -p litchi-ppt --all-features --locked --offline document_comparison::tests:: -- --test-threads=1` (11 passed);
* `cargo test -p litchi-ppt --all-features --locked --offline --no-run`;
* full-feature and no-default-feature rustdoc JSON generation;
* `tools/check_api_glossary.py` for `ppt-full` and `ppt-minimal`; and
* `python3 -m unittest -q tools.test_check_api_glossary` (55 passed).

No native timing or performance acceptance claim is made here. The source
changes preserve the existing wire representation and output paths; native
qualification remains a separate coordinator activity.

The commands above are the focused historical receipts from the original
`1efd4957a0fc2dc1b9efdef200cdad215360322c` migration.  They remain diagnostic
and are not substituted for the guarded rerun below.  The final source is
`6b4315487773490731f4da56b04760cbc1739f94`, clean on
`uo2/ppt-record-limits`; the semantic validation source is its parent
`884d7f7232414a6b9d4e493884aca1808286e487`.  The only delta from `884d7f`
to the final source is the clippy-equivalent return simplification in
`binary_tag_edit.rs` and the equivalent `clamp(1, MAX_EVENTS)` spelling in
`diagram/codec.rs`; no wire, limit, or output behavior changes in that delta.

The guarded rerun used Rust 1.95.0 (`rust-toolchain.toml` SHA-256
`e3a213e0d222e94d213cafbc20932eb3f76c643b4dd63756acf95192df2aa310`), the
approved `Cargo.lock` above, `.cargo/config.toml` SHA-256
`50b044c83828f2c448ee59c362afaa79de2dd47fbd390f0cd2888391f1d1119d`, and
`clippy.toml` SHA-256
`f11517bc5872d96aab42ba6519db4e78f2fdcd44512ac70a1b45adbd9929b50c`.
The f5 self-attributing runner was
`/home/zhuhe/code/litchi-worktrees/uo2-gate-runner-self-attribution/docs/unified-ops/results/0362ul29/tooling/run_safe_gate.py`
(SHA-256
`1904177040fc7c523582433028d9bba56efc30fe0d9d96beb554a2fab343691e`),
with `safe_chunked_tests.py` SHA-256
`3e03aa2227ac0002c3052cf5080d032cd2e9150f584c926b9aa0d499175ec67f` and
the adjacent cargo shim SHA-256
`495e24c6f60b238378386d86ac4158cba904992ce8870a9494f2c4a0eedd256d`.
All commands were offline and locked with one build/test job, one test
thread, incremental compilation and debug info disabled, the `20G`/no-swap
process cap, and decimal target guards of 18/19/20 GB.  The final guarded row ended
with no owned processes or builders.

Accepted receipts (all paths are outside the repository and are retained as
the coordinator's external evidence) are:

* Exact copied-payload boundary test, source
  `884d7f7232414a6b9d4e493884aca1808286e487`, state
  `(deleted-coord-artifact)`:
  log `00-ppt-record-copy-boundary-attempt-884d7-1.log`, SHA-256
  `522ecd722eb73cd2ad491ef149f31df346d1455ba50b1042319afa48722aff60`;
  target trace SHA-256
  `9727bf34fde652503ad024815098791c074b9953ceafb767c3995c2faccf60e9`.
  It proves one-tag success at 288 and refusal at 287.
* Hyperlink aggregate/profile tests, source `884d7f7232414a6b9d4e493884aca1808286e487`: log
  `01-ppt-hyperlink-focused-rerun-attempt-884d7-2.log`, SHA-256
  `1b02d4761d15af83483ce15578952eb6c2505612cf9aaa3caa3b4e4ea29b55a3`;
  trace SHA-256
  `5a6c062f08386b06bc6a431b1cff2024591a2f0e98bd5e758bcac27512328549`.
* Document-comparison profile regression, source `884d7f7232414a6b9d4e493884aca1808286e487`: log
  `02-ppt-document-comparison-profile-rerun-attempt-884d7-3.log`, SHA-256
  `522ecd722eb73cd2ad491ef149f31df346d1455ba50b1042319afa48722aff60`;
  trace SHA-256
  `47812fdf83d580b44c7e800cdf14804099eb23e5ba80ba298f71c8cf1cd3e566`.
* Full `litchi-ppt` library, source `884d7f7232414a6b9d4e493884aca1808286e487`: log
  `03-ppt-library-rerun-attempt-884d7-4.log`, SHA-256
  `687adebb3fa52cac421802c40ecaa6705d358a4bbda52ebb7c6e9c6da28e940a`;
  trace SHA-256
  `f40cab1d4289b1e355088ebd5dc89549a42ec4ea4c28daa2c3a33db53cc85b99`.
  Result: 1267 passed, 0 failed, 1 ignored.
* PPT, ops, and facade all-target check, source `884d7f7232414a6b9d4e493884aca1808286e487`: log
  `04-ppt-ops-facade-check-rerun-attempt-884d7-5.log`, SHA-256
  `e3b0c44298fc1c149afbf4c8996fb92427ae41e4649b934ca495991b7852b855`;
  trace SHA-256
  `7e64a70e95a2f5ebf9b0fd51de4d94123694c6a72b5d71c554158aabacbd1c9c`.
* Final affected all-target clippy, source
  `6b4315487773490731f4da56b04760cbc1739f94`, state
  `(deleted-coord-artifact)`:
  log `00-ppt-ops-facade-clippy-final-attempt-6b4315-1.log`, SHA-256
  `e3b0c44298fc1c149afbf4c8996fb92427ae41e4649b934ca495991b7852b855`;
  trace SHA-256
  `c041e56a8a07193a2fa791b92ee7caa81cc89bad7e2ad9d054c63c14e865b749`.
  Target peak and final size were 787,602,473 bytes; free space after was
  511,925,956,608 bytes.

The following failed attempts are intentionally retained and are not
acceptance evidence: source
`5a3d8ee9503c56da0dc9ea7be2e0bb724a85258a` failed the first compile with
nine diagnostics (log SHA-256
`af0bbb4ca8473c87f6dcb4fcfb0b5eaa191de910f3c5a2cd94d6554a4ed61e7e`);
source `768aed7e62ea7367bd2801b9a3e8fc79637075ce` reached the library but exposed the stale copied-payload
boundary at observed 208 over maximum 200 (log SHA-256
`b4eae360b5878541d111ca8b306787f900b8880895f4a5292f9362c4d610a90f`);
source `b48a2625904814cdbd1e38fd64cbab80ff9fce77` reached observed 232 over maximum 208 (log SHA-256
`b44af6bc1b842e861a3a899aa5188d2f1f392ff00c0f57029c98f304bd5ae272`);
source `6f6f3643c4f5f7b4c90d1266e46c5d46079f8192` reached observed 272 over maximum 248 in the isolated
boundary test (log SHA-256
`b35e5aaeced463f0f737387baa7d2e91e4a26c4539d44cab33d06858cf2abfcc`); and
source `884d7f7232414a6b9d4e493884aca1808286e487` clippy failed on two needless-question-mark diagnostics and
one manual-clamp diagnostic (log SHA-256
`810e835fe52c86373d72bcb91bf17579841eb2109ebf8e336e624cd01e8d6146`).
Those receipts remain preserved outside the repository.

No native timing or performance acceptance claim is made here.  The source
changes preserve the existing wire representation and output paths; native
qualification remains a separate coordinator activity.  The final source
also retains the exact shape-click admission regression: a 208-byte source
projects to a 288-byte rewritten aggregate and passes at the exact ceiling
while refusing one byte below it.

## Coordinator integration review

Final atomic `6a1317ab04e57cf1fe74b1a2637a8e0be6c46003` is based on
`939d4016092409575bb280cf3d3d1b8623cf0994`; its tree
`d302982f25348349f29dac3ef159e91b5cf6af5a` equals the preserved
`ec014b1592fcaa7b1b6eac2370eed7b28a639837` tree. The complete diff is
limited to PPT, its ops validation consumer, glossary policy and this record.
Independent review confirms the compile/error-conversion and Clippy changes
retain typed limits, profiles, wire fidelity and allocation behavior.

Root verifies every accepted and retained failed command log and Cargo trace
in the final packets, unchanged locks, clean source and empty terminal owned
process lists. The full-library/check/failed-Clippy packet is 46,949 bytes,
SHA-256 `eb9f6f393389deb10b190b3d016c5e890226bfd210d986ed99c2da24e86631f5`;
the exact-boundary packet is 25,318 bytes, SHA-256
`7bb7956e18be3bcf436db890c24b66c350640d6958839708d27911167f3870f5`.
The final Clippy packet is 4,510 bytes, SHA-256
`19988b6d35ccd59fc2223eb28b109ffa76fb4e4cefc8106a429446b241f05a15`.
These are the `custom-results-rerun.json` and `custom-results.json` files
under the external state directories above.

The exact one-tag charge is `72 + 64 + 56 + 16 + 24 + 16 + 24 + 16 = 288`.
The repeated 40-byte child parse predates this migration; existing-record
admission adds counters without copying payloads. This is not a newly
accepted performance regression. Actual native qualification and the final
integrated gate remain required.
