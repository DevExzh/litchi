# 0362ul77 — composed standalone benchmark consumer check

The actual guarded standalone consumer check passes on clean composed source
`ee7e9955d3bc6f6942a2c4434c1e1a4cfec16c63`, including the font, crypto,
Markdown and RTF migrations and both PPT benchmark repairs in 0362ul39b.
The detached capture is
`.coord/final-composed-ee7e9955-validation/source`. The command is:

```sh
cd tools/perf-baseline
cargo check --offline --locked --all-targets
```

It exits 0 after 149.2265739440918 seconds. This duration describes the
compiler command; it is not a performance measurement. The reused target
finishes at 2,851,946,744 bytes, below the 19,000,000,000-byte live and
20,000,000,000-byte hard guards. The runner uses one build job, one test
thread, zero debug/incremental artifacts and the 20G memory cap.

The receipt directory is
`.coord/final-composed-ee7e9955-validation/f5-perf-validation`. Its reduced
identities are:

| Artifact | Bytes | SHA-256 |
| --- | ---: | --- |
| `result.json` | 15,580 | `f23e236f372bcd97b01b10e2a030d1895be031d89b76f012b49050a6d66f0c27` |
| `00-perf-baseline-check-attempt-ee7e-perf.log` | 58,837 | `36e4b394581ba436566a556414ffcf07945707d93ab38d5073e67ca80c0b24df` |
| `00-perf-baseline-check-attempt-ee7e-perf-cargo-targets.tsv` | 115 | `4d6e27f0310d3109add063e6791c645f7c101793b07a8458f4a0d899bb7a48f9` |

The actual approved f5 runner has SHA-256
`1904177040fc7c523582433028d9bba56efc30fe0d9d96beb554a2fab343691e`, its
Cargo shim is
`495e24c6f60b238378386d86ac4158cba904992ce8870a9494f2c4a0eedd256d`, and
the chunk helper is
`3e03aa2227ac0002c3052cf5080d032cd2e9150f584c926b9aa0d499175ec67f`.
Rust is 1.95.0. The three approved locks remain byte-identical before and
after the command:

* workspace: `efff87af6c339a44e267987c5c79ba3fb1910cef734f0a9f212f08f865dac811`;
* performance tool: `de1f8ae53b3455a4246da0b9340896881cd72ca0295b8e52444e0d55e6c9c4ea`;
* native tool: `989da70f10deab4460a77e4afa3a1b840b4697bbea2c07e134e22d9b4c5084ae`.

The coordinator independently verifies logs, traces, source/lock identities,
target guards and all three empty terminal owned-process samples. The
earlier failed and interrupted attempts remain preserved in 0362ul39b.
This closes the queued compiler check for the composed RTF/PPT consumers.
Later source lanes still require their own composed checks; the final gate,
native execution and performance qualification remain open.
