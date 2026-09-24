# Theme-family hardening performance evidence

This directory records a bounded current-source profile for the shared
DrawingML `themeFamily` owner after namespace, duplicate-owner, and caller
output-limit hardening. The profile is separate from production dependencies
and does not modify production code.

The targeted harness runs nine lanes in three fresh processes with twenty
measured samples per lane after two warm-ups. It uses a process-local counting
allocator for requested allocation bytes and incremental peak live bytes, and
`/usr/bin/time -v` for whole-process RSS. Fixture construction, source
validation, and family object preparation happen outside each timed closure.

Reproduce it from the repository root with:

```sh
CARGO_TARGET_DIR=/var/tmp/litchi-theme-family-hardening-profile-target \
  sh docs/report/spec-gap-validation-evidence/theme-family-hardening-performance/run_targeted.sh
```

The runner uses the native complete Theme fixture
`crates/litchi-drawingml/tests/fixtures/theme-part-native.xml`, whose SHA-256
is `50b662d8ff0e562157ff80aa19c85ff214cd21357d66457661d070f50c6fdc59`.
It builds with the pinned Rust 1.95 toolchain and an isolated target. The
source manifest, exact commands, host/build records, raw per-process JSON,
allocator receipts, and verification result are under `results/`.

The lanes are:

- `native_read`, `native_replace`, `native_remove`, and `native_add`: valid
  native Theme read and source-preserving family workflows.
- `unknown_32` and `unknown_1000`: unknown-URI direct extensions containing
  32 or 1,000 family-shaped opaque descendants, with 200 root namespace
  declarations. Each must remain readable with no typed family owner.
- `duplicate`: two supported family owners in one recognized extension; the
  scanner must refuse the ambiguous input.
- `limit_replace` and `limit_add`: caller output limit of one byte; each must
  refuse before producing a result.

The run used base HEAD
`f4f29c9306d182072aabe8eca08145470fdcfadd` with the prefix correction
applied (subsequently committed as `4814a84d7`). The source-manifest SHA-256 is
`5cb3b62847215630e723a3009e95513c5ff6a4ea919fbc8af803be9ee0849082`, and
these source hashes:

```text
codec.rs        3671ef78454309c45039741414d5385a400fd9c8a736fdfac974e704e577a910
part.rs         cb3f384a6798a343b0e92c994d982a6e196bf50311a18d5441db3fa6678e6f79
family/mod.rs   33557a195ba76b1eb793fa3bfca1b00ecec842c4016446ea969dd0da2932ace1
model.rs        e3f0e9cb33324faf29fdc94cad733cc7a95bf86ae3d583df37b29202220d7d55
transaction.rs  006472620192f0561529a2f38809381fef45c4ba07080501f40bb5a40c52b1cf
theme_part.rs  19ec1d9f75a57a505f0c8eae8338691f3eb2467bdebe8f4e16773b26785fb040
```

The measurements are absolute observations for these named inputs and source
revision. They do not establish a before/after speedup, an asymptotic proof,
or a whole-library performance claim. The earlier full XLSB host profile in
`../xlsb-theme-family/performance/` is retained as historical evidence with
its own source hashes; it is not mixed into this receipt because its codec
source predates the final prefix-validation edit.
