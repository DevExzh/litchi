# Supplementary CPU profile scope

Two media-rich profiles retain 100 samples and three warmups each, using
cycles:u at 499 Hz and a 16,384-byte DWARF stack dump. Bytes has 17,217 decoded
CPU samples; files has 16,635. These 200 benchmark rows are supplementary and
are not part of the formal 960-sample baseline.

The frozen `profile-audit.py` searches for a qualified iteration name. Perf's
DWARF display uses the unqualified name, so that initial analysis reports no
iteration attribution. Its unchanged result remains in `profile-summary.json`.
The separate `profile-symbol-audit.py` also recognizes the exact unqualified
name. `checks/profile-symbol-resolution.json` records that the hash-bound
binary has exactly one such symbol, in `pptx_provider_lifecycle`.

| Profile | Resolved iteration | Corpus builder | Other/unresolved | Total |
| --- | ---: | ---: | ---: | ---: |
| media-rich bytes | 2,396 | 830 | 13,991 | 17,217 |
| media-rich file | 2,074 | 785 | 13,776 | 16,635 |

The resolved iteration stacks include plan ancestry (698/614 samples),
publication ancestry (515/518), hashing helper ancestry (975/830), and small
open/RSS-observer counts. These ancestry counts overlap; they are not exclusive
API CPU totals. Iterations also include setup, observers, checks and drops,
which are outside the formal API clocks. Decoded iteration leaves include SHA
intrinsics, while frequent unassigned leaves include Deflate implementation
functions. The large unassigned population prevents complete API attribution
or a defensible end-to-end hotspot ranking from these profiles alone.

Further profiling should improve stack completeness and preserve setup/check
attribution before choosing an optimization from this CPU evidence. Neither
sample percentages nor these incomplete ancestry groups are API wall-time
fractions, speedup evidence, or proof that all unassigned work is corpus setup.
The raw perf data and decoded stacks remain available for alternative analysis.

The first recording completed before its wrapper rejected the profile sample
count. The explicit validation amendment preserves that recording and the
original validator/build, revalidates every unchanged formal report, and allows
100 samples only for the two declared profile roles. The continuation records
the first command as reconstructed from the frozen driver/protocol, with its
unretained original start timestamp left unavailable. It does not rerecord it.
