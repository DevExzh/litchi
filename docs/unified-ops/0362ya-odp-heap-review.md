# 0362ya — ODP source-backed heap review

This receipt records the second ODP heap candidate after the rejected lazy
projection in 0362y. It is review evidence for the residual from 0358a; it
does not close the residual until the native paired run is repeated on a quiet
host.

The accepted 0358a probe used a generated ODP with 2,000 slides, a 591,185
byte `content.xml`, and a 4 MiB incompressible media member. Its source archive
was 4,215,428 bytes. The former route streamed `content.xml` while planning a
tail append. The 0358a source-backed route opened a
`SourceBackedPresentation`, retained the validated XML, committed the append,
and wrote the result. Across 15 samples, the lifecycle peak grew from
5,513,403 to 6,107,267 bytes (+593,864 bytes, 1.108x); output length and the
recorded SipHash stayed identical. These are the measurements in
`results/0358a/sb-before-1.json`, `sb-after-2.json`, `sb-after-3.json`, and
`summary.json`, described in 0358a §6–7.

The 0362y implementation dropped the open-time string but reread and
re-decoded `content.xml` on the first semantic operation. A matched counting
allocator probe showed why that is not sufficient: the semantic pass added
691,175 allocated bytes relative to an owner that already held the XML, and
the first semantic/no-op lifecycles still had a net allocation increase after
the lower open cost was included. The probe used the same generated archive
and started counting after the source byte clone, so source ownership is not
part of these numbers.

The candidate in this branch makes the cold path bounded without changing raw
getter behavior:

* open validates unencrypted `content.xml` and `styles.xml` through the
  verified fixed-buffer member reader, checking UTF-8, the materialized-size
  ceiling, the office body marker and prologue doctype without retaining the
  member payload;
* encrypted members use the existing materialized path, preserving password
  handling and its error ordering;
* the raw `content_xml()` and `styles_xml()` getters retain an `Arc<String>`
  only when requested, removing the temporary `String` to `Arc<str>` copy;
* slide, text, tolerance and sequential-text routes decode temporary source
  projections and drop them after the parser or sink pass, so a cold append
  and replay do not retain `content.xml` merely because the owner was opened;
* source-version checks remain before, after and around each source read, and
  raw publication still uses the retained cache only when publication needs
  the exact XML projection.

The matched allocation-only comparison below uses the same six owner
lifecycle shapes in isolated runs. `peak` is the allocator's peak live bytes
after the source clone; `allocated` is cumulative bytes requested, including
realloc growth. The eager reference is the 0358a owner with the `Arc<String>`
construction fix but without the streaming/direct semantic changes.

| lifecycle | candidate peak | eager peak | peak delta | candidate allocated | eager allocated | allocated delta |
|---|---:|---:|---:|---:|---:|---:|
| open | 116,001 | 693,231 | -577,230 | 200,325 | 777,932 | -577,607 |
| open + first slide | 694,351 | 693,231 | +1,120 | 3,325,010 | 3,211,442 | +113,568 |
| open + first text | 1,501,757 | 1,501,373 | +384 | 4,786,692 | 4,673,124 | +113,568 |
| open + unchanged write | 1,285,577 | 1,285,153 | +424 | 1,582,282 | 1,468,674 | +113,608 |
| open + append commit + write | 612,845 | 1,204,071 | -591,226 | 5,853,743 | 6,431,822 | -578,079 |
| open + replay + write | 612,845 | 1,204,071 | -591,226 | 5,853,775 | 6,431,854 | -578,079 |

The semantic pass itself still reads and allocates a fresh XML projection:
subtracting the open rows gives 691,175 more cumulative allocated bytes for
the first slide and first text routes (and 691,215 for unchanged write) than
the eager reference's corresponding pass. The retained cold payload is gone,
and the measured peak is effectively level for semantic routes, but the extra
read/decode work remains an open performance question. The candidate therefore
does not claim that 0358a's native time or allocation residual is closed.

The probe also recorded candidate allocator calls: open 72 allocations/7
reallocations, first slide 20,093/23, first text 34,090/7,947, unchanged write
86/9, append 36,185/99, and replay 36,189/99. Native call and instruction
counts were not collected while the shared host was loaded; the required
quiet-host ABBA timing and callgrind comparison remains follow-up work. The
0358a source-backed route's 15-sample timing was 13,131.943 µs before and
13,296.183/13,314.163 µs on its two after legs.

The old 0362y source, tests and allocation receipt remain in history as the
rejected candidate. This slice keeps its failure evidence rather than
rewriting the record. The new source and tests add coverage that semantic
reads do not populate the raw cache, ordinary text projection remains source
fresh, and the existing malformed-input, source-change, repeated-reader,
sequential-sink and exact-output tests continue to exercise the same parser
and freshness contracts.

## Validation

The following focused checks passed on the candidate parent `681e5ec809`:

| Check | Result |
|---|---|
| `cargo fmt --all -- --check` | passed after formatting |
| `cargo test --offline -p litchi-odp --test source_backed --test sequential_text` | 28 passed |
| `cargo test --offline -p litchi-odp --lib package::source::tests` | 2 passed |

The allocation probe was temporary and remains outside the committed source
tree. Its SHA-256 was
`d47d8bf1960e3e305bf42e6de3f2dbe40cc2185fd921271c55ae60f1159cfd7c`; the
regenerating command was:

```text
cargo run --offline --release -p litchi-ops --features odp --example odp_alloc_probe -- all
```

No native timing or callgrind run is represented by this receipt. A quiet-host
ABBA run must compare open, first semantic read, unchanged publication,
append, replay, output bytes and source-read counts before this candidate can
be treated as an accepted performance repair.
