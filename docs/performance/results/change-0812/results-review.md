# 0812 scanner instruction-localization results review

This diagnostic localizes the current scanner leaf before a candidate is
selected. Production source is unchanged, and this record makes no speedup,
causal instruction-cost, ordinary-build phase, or adoption claim. The
previous 0811 native captures remain the source of the historical offset
census; the new samples below are a separate amendment taken only because the
0811 executable could not be reconstructed byte-for-byte.

## Binary custody and evidence boundary

The rebuild used the 0811 frame-pointer probe, source manifest, locked offline
release command, two Cargo jobs, and `-C force-frame-pointers=yes`. The rebuilt
source map has the same 9,196 files as the 0811 source map, and the build
completed successfully. Its executable is 14,387,272 bytes with SHA-256
`63b55206cb606f9fa0552d2a999b7c3ba9670bd5747558aec4ba29cefdc0c231`.

The sealed 0811 frame-pointer executable has the same length but SHA-256
`b8b89fc5c4c062ab79df46c9105617327d8554b40baf8125d20351149d032062`.
Because the hashes differ, the old 0811 sampled instruction offsets are not
joined to the rebuilt disassembly. The retained old census still records
`+0x24d` in 151/217 and 223/254 scanner self-leaf samples, but those counts are
historical, unmapped observations. They are not evidence about the new
instruction at the same numerical offset.

The fresh amendment ran two serial large-capture profiles against the actual
rebuilt executable. Each used 100 measured samples, zero warmup, CPU 12,
`cycles:u` at 499 Hz, frame-pointer call graphs, and `--no-buildid-cache`.
Both workload commands and both decodes exited successfully. The two reports
contain 200 measured outputs and retain the sealed current-source output,
source, and semantic oracle values;
the raw perf data and decoded frames are retained as deterministic gzip
members. No timing comparison with the old executable was made.

## Exact instruction localization

The fresh executable's `scan_processed_xml` symbol begins at `0x12a2d0` and is
2,588 bytes. Immediately after the `Reader::read_event_impl` call at function
offset `+0x21c`, the success path contains three visible 40-byte payload-copy
chains before the event discriminant dispatch at `+0x296`:

```text
+0x242..+0x256   Result payload -> temporary
+0x25a..+0x270   temporary -> temporary
+0x274..+0x28f   temporary -> event-dispatch storage
```

The concentrated sampled instruction is at `+0x24d`:

```text
movups 0x10(%rcx),%xmm1
```

The exact line map attributes sampled address `+0x24d` to
`core::result::Result<T,E>::map_err` (`result.rs:967`), in the
`reader.read_event().map_err(xml_error)?` expression at `codec.rs:389`;
separately queried later address `+0x27d` maps only to `codec.rs:389`. The
assembly therefore confirms the source hypothesis: the
success-path error-type conversion is materialized as aggregate `Event` payload
transport in this exact frame-pointer binary. It does not establish how much
ordinary native time those copies consume, and the three chains must not be
treated as three independent costs.

The fresh exact-binary frame census is:

| Repeat | Whole samples | Exact-owner samples | Scanner self-leaf samples | `+0x24d` samples | Unknown interiors | Lost lines |
| ---: | ---: | ---: | ---: | ---: | ---: | ---: |
| 0 | 3,066 | 960 | 231 | 193 | 2 | 0 |
| 1 | 3,095 | 970 | 227 | 178 | 0 | 0 |

The `+0x24d` counts are descriptive sampled IP counts from the fresh
executable. Sampling skid and the material frame-pointer perturbation measured
in 0811 prevent converting them into a causal cycle share or an ordinary-build
phase fraction.

## Bounded next candidate

The smallest candidate justified by this localization is to replace only the
`map_err(...)?` transport with a direct `Result` match:

```rust
match reader.read_event() {
    Ok(Event::Start(element)) => { /* existing arm */ }
    Ok(Event::Empty(element)) => { /* existing arm */ }
    Ok(Event::End(_)) => { /* existing arm */ }
    Ok(Event::DocType(_) | Event::PI(_)) => { /* existing refusal */ }
    Ok(Event::CData(_)) => { /* existing refusal */ }
    Ok(Event::Eof) => break,
    Ok(_) => {},
    Err(error) => return Err(xml_error(error)),
}
```

This tests whether the intermediate error-type conversion and `?` boundary
can be removed from the successful event path. `Result::map_err` is inline, so
the compiler may already have emitted the best equivalent code; a candidate
that produces identical release assembly has no reason to proceed. The
candidate must leave `Reader<&[u8]>`, parser configuration, every event arm,
resolver push/pop timing, checked attributes, duplicate checks, UTF-8 and
unescape behavior, limits, value ownership, and error identity/order exactly
unchanged. It must first pass the buffered differential oracle and the full
malformed namespace/refusal corpus, then receive fresh release quality,
output/semantic, resource, and owner-scoped native workflow evidence before
any retention decision.

The namespace and checked-attribute fusion idea remains rejected as a local
follow-up: the pinned public APIs perform distinct walks with distinct
ordering and refusal behavior. This result also does not revive the rejected
0806 iterator candidate, introduce a second parser, or justify unsafe code.
The current scanner/event hypothesis remains open for a bounded qualification;
the broad OLE2/OOXML goal remains active and iWork is excluded.
