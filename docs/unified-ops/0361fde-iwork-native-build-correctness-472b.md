# 0361fde — iWork build and correctness preflight at 472b

This is a reduced receipt for the build/correctness preflight of the repaired
iWork native qualification packet. It contains identities and summaries only;
the complete JSON, binaries, targets, fixtures, and command logs remain in the
external packet directory below.

## Frozen inputs and execution policy

- Source endpoint: `(deleted-coord-artifact)`
- Source HEAD: `472b6705475b073e7fe08f6e11ee17fb1e29f838` (clean detached
  endpoint)
- Constructor baseline: `ddf2ac8ae2117cd78829ef94bcb610d6cbdfa0e2`
- Owner baseline: `a18593385cf3162911f00ebd3cbb1379855a8eba`
- Input root `Cargo.lock`: 139,606 bytes,
  `4100b1fbb272dd24842f33a2d6f353a3f3a97dfa7e435c4649a0e1a029bfaea1`
- Qualification runner: 109,594 bytes,
  `c6d0cd178b5c21a05ff20f9eb12fb93fcb040cd783b322f9bec8899717e325ab`
- Constructor probe: 22,799 bytes,
  `cc322c2189e8e3822472627e61c45b8e463f866adcba32858c88cc237bd17a3e`
- Owner probe: 27,230 bytes,
  `c580de1c52c2c37c1cebbc7a0e6324cd03cd827e01de09d11ca7f2d9de4b50a1`
- Cargo shim: `(deleted-coord-artifact)`,
  SHA-256 `1b22a69fd625d57565041bf6c634d8ebd1a82be8b3d0006390472c86324230d8`
- Build and preflight environment: `LITCHI_MEM_MAX=20G`,
  `CARGO_BUILD_JOBS=1`, `RUST_TEST_THREADS=1`, offline and locked Cargo,
  debug information disabled, and fresh external `TMPDIR`.

The external packet root is
`(deleted-coord-artifact)`.
The build outputs are `constructor/` and `owner/` below that root. Each target
was checked with `du -sxB1` and was below 20,000,000,000 bytes:

| packet target | bytes |
| --- | ---: |
| constructor/target-base-native | 471,359,488 |
| constructor/target-base-allocation | 471,379,968 |
| constructor/target-candidate-native | 472,481,792 |
| constructor/target-candidate-allocation | 472,502,272 |
| owner/target-base-native | 753,381,376 |
| owner/target-base-allocation | 753,582,080 |
| owner/target-candidate-native | 754,683,904 |
| owner/target-candidate-allocation | 754,819,072 |

The commands used the runner at
`$SOURCE/docs/unified-ops/results/0361fda/iwa_native_qualification.py`:

```text
prepare-fixtures $PACKET/fixtures
build constructor $PACKET/constructor --base ddf2ac8ae2117cd78829ef94bcb610d6cbdfa0e2 --candidate 472b6705475b073e7fe08f6e11ee17fb1e29f838 --lock $SOURCE/Cargo.lock
validate constructor $PACKET/constructor
build owner $PACKET/owner --base a18593385cf3162911f00ebd3cbb1379855a8eba5 --candidate 472b6705475b073e7fe08f6e11ee17fb1e29f838 --lock $SOURCE/Cargo.lock
validate owner $PACKET/owner --fixtures $PACKET/fixtures
collect constructor $PACKET/constructor --out $PACKET/constructor-collection.json
collect owner $PACKET/owner --out $PACKET/owner-collection.json
```

Here `$SOURCE` is the frozen endpoint above and `$PACKET` is the external
packet root above. `build` used the two endpoint-specific baseline refs shown
above; the owner build's requested 41-character input ending in `a8eba5` is
retained in its receipt, and Git resolved it to the canonical 40-character
commit ending in `a8eba`. `validate` ran both the native and
counting-allocation binaries. The collection commands were read-only packet
checks.

## Build receipts

Both `build` commands completed with status `complete`, with metadata,
toolchain, native, and counting-allocation commands exiting zero and timing out
false. The generated project lock identities were retained in each build JSON:

- Constructor: build JSON 1,255,772 bytes,
  `be52d49cc75697bcd391e8156cf0e6f4ff461ea99c48831942bf7085d4fbd02b`;
  base project lock 25,741 bytes,
  `c806205754ac2442382a7fc7fdb3532fa336594fa4b3742e66642780d9729542`;
  candidate project lock 25,947 bytes,
  `adf212f30aecc6e54fc46e8e54b3b0c411052e37e8ea380a872ff90a184d5160`.
- Owner: build JSON 1,354,498 bytes,
  `8e367cf61d484e8488f3a880c13262bdcfb96e45795a2e30a0bc8f17f32e435e`;
  base project lock 38,062 bytes,
  `35adbc0cdaea268db152e36f7180fb5a49c0f259ad76ec601d82687ab09d265d`;
  candidate project lock 38,075 bytes,
  `1f64efdb50642642076299f045202040b39a12e6157f249b9792cbd25c4e5daf`.

The protected binaries were copied and hashed by the runner:

| family | endpoint | native bytes / SHA-256 | allocation bytes / SHA-256 |
| --- | --- | --- | --- |
| constructor | base | 636,248 / `68c3f9e1b89b3b8176019f06e5ba59c323db37ec56145eaaf7ac20fefea983e1` | 657,800 / `d839d41d80b9bc60c621f6990e79ae260fd6a54bef181dfe49650f693913189c` |
| constructor | candidate | 665,184 / `f27589ad941e0dd517b1443ee29d13dd7a00fee81f1018007f4b99ecb436ce5d` | 686,096 / `751980e13347c6d5c5c2948c8cded01d49970ae73cc9763cba71a60c6772b772` |
| owner | base | 2,948,008 / `b6edb38db2322b92ae08e36c45423731e730d8950d2dbc5b4a18086cd87dff8b` | 3,132,920 / `808be97318f5f02be89aae202ecce245f217c2c56395d20a4d599b2c4e6bd7fa` |
| owner | candidate | 3,048,928 / `84e81f891b78e15ed326d5df5f4be8d7bd29866042f8fc50572eb1ff3ee91fee` | 3,160,488 / `8879dde402c479d20b8faf0f6516393d78fef020ed96d3d863d19bfb1bf7b8c7` |

## Fixtures and correctness

The external owner fixture manifest is 4,811 bytes,
`f936923d1a60f2670d2381507bea53d5d2e9f3bbeee85105ffdc36641502aa65`.
It contains the five required workload files: `basic.numbers`
(`f225d5b1cd59e9da454f91a96fe8f81154bc31037c10029230e75d49b45fb693`),
`basic.pages`
(`21107bc9323fba6f1589152454c0b0b0cc8e239313c6a369bc4a891116601b42`),
`media-replacement-native.key`
(`f5763984974612f078486cb2310f408cbcb7cd06ef6adca494aae19eaf62609d`),
`tone.wav`
(`27f1b6add1bd884455da7a36244d92fc26b81b622ce1671d5cb77bebffeb135a`),
and generated `footnote-fixture.pages`
(`318b06befe4d244c368f1bfa5a61dab04f00dd46530c6ce6078f316b44b43266`).
The generator target was external and 670,883,840 bytes by `du -sxB1`.

The constructor validation packet is complete: validation JSON is 1,198,680
bytes, `f3e400433dd3f407e54bdc429e16c3c44f68fecc4e6e3004d7c3309e527e54e8`;
all four endpoint variants passed all 15 routes, including source/encoded and
retained round trips. All four route reports are byte-identical with SHA-256
`c19423e71d8678f7ec15876cb403b7db28bd99f1e7cc1cc9322ae60539481720`.

The owner validation packet is complete: validation JSON is 1,284,812 bytes,
`d8c9f4bd7a0dce3b602e7466569937ee5952a421120e6761cebf68457ffb7f2c`;
all four endpoint variants passed all four routes. The reports are
byte-identical with SHA-256
`dd6e82d2556f7ec59d793bf010cf194deda57bac1c2522b1bb4792ce18dc6195` and
retain changed/replay/inverse/no-op and semantic correctness proofs.

Read-only collection independently verified both packets:

- constructor collection: 1,181,752 bytes,
  `1ca6317a44204e2d65d1c5c82820a55a4decacc9a04c4b9a6e3093a8e93380b5`;
- owner collection: 1,259,953 bytes,
  `421410175c765f678b0c4950e958c9991bf445bf5b0a7c748011957362cb98bb`.

Both collections are `status=complete` with `run=null` and
`allocation=null`. The constructor validation and owner validation commands
had zero exit status and `timed_out=false`; their retained logs have SHA-256
`e3b0c44298fc1c149afbf4c8996fb92427ae41e4649b934ca495991b7852b855` and
`fe7f14617c00a60fad7cb101053fdcf859e81225dfea7ffa6c831f354c7aec9d`,
respectively. This record is build/correctness-only evidence: no allocation
capture, native timing, ABBA, or performance acceptance claim was made.
