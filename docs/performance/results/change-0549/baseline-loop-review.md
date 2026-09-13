# 0549 baseline loop review: `CheckedBitSet` membership and mark

This is a read-only follow-up to the sealed 0548 OLE2 experiment. It checks
whether the restored baseline `SectorChainScratch::collect_exact` still emits
the source-level duplication between `CheckedBitSet::contains` and
`CheckedBitSet::insert`, and whether a checked fused operation has a measurable
machine-level opportunity. No source, build, benchmark, profile, or capture was
changed or run for this review.

## Evidence binding

The measured baseline source is the current production source at
`crates/litchi-cfb/src/file.rs`, SHA-256
`72fb5295dfbba5f3ed2192a9c24ecbf7d8aa2759f418b77563a4e2ff1042d03b`. The
baseline source manifest records the same file hash and has SHA-256
`fa85f76972a37f52644c60b955c9125ed80e4d78e2434ec54b77c3cb19e64163`. The
release binary used for the disassembly has SHA-256
`960ef0f2557ae1d81c958c06d17496a380ef51f75902c85ed5e76aa2ef8ebec5`, and the
0548 plan binding is
`575316305da224ab1a4c673a9d14c0912e600e473e18d8f61bf4c32dc4339fa9`.

| Artifact | SHA-256 |
| --- | --- |
| `baseline/assembly-index.json` | `92a83269111b7e271b4bf83b2633fc116bcabdceab1711cb454e3199cfc740fb` |
| `baseline/assembly-1.stdout` (`CheckedBitSet::insert`) | `63e74e0a8fd13188fe90d081843d5f58eb6e6d7f61345f17c1c2cb3b494a5dff` |
| `baseline/assembly-1.receipt.json` | `82ebad8a4296baf2943ff3f5670aef0691275766564564af45d4138412422cb6` |
| `baseline/assembly-2.stdout` (`collect_exact`) | `c8f23242ee89bfbadde2c96e4f1750be3df3dfd668b487678a32db9f33c520ab` |
| `baseline/assembly-2.receipt.json` | `03e2084d2b291448fce3298fe8e346a117c6551cd54654e3a6f008d3b587998d` |
| `baseline/instruction-analysis.json` | `0440659c68bce9d0e053485bce3bb3765be977f790a7d451df3cbacbdee69a0e` |
| `baseline/profile-analysis.json` | `32111b693e42f85d2f8f754fb8c5bc45eac0f447bce196ba7facdc144823a011` |
| `0548/adverse-review.json` (candidate comparison) | `016c05eeb883c2ae3e1d246d1d5d056e1f79937f4fcdcef9df783f5ab8356e52` |

`assembly-2` is the exact symbol
`_ZN10litchi_cfb4file18SectorChainScratch13collect_exact17h6a8f4566950984e2E`,
at `0x2f223c0`, with 1,436 bytes. The instruction analyzer found one exact
collector row, 313 static instructions, six setup dumps, and 40 timed dumps.
It found one exact `insert` row and no `contains` symbol or optional helper
fragment. The profile analyzer independently reports eight profiles, six setup
dumps, and 40 timed constructor dumps. These are Callgrind Ir and instruction
attribution records; they do not establish operation-local wall time or
hardware uops.

## Source and emitted loop

The restored source calls `contains` at lines 2848–2851 and `insert` at lines
2853–2857 of `file.rs`. The helpers themselves are lines 39–46 and 48–62.
Both helpers contain a logical bit-length check, compute `word = bit / 64` and
`mask = 1 << (bit % 64)`, then perform a checked word access. The collector also
checks the allocation-table slot before those calls.

The hot path in `collect_exact` is the interval `+0x1c0..+0x1fd` relative to
the symbol start (`0x2f22580..0x2f225bd`). The relevant instructions are:

| Relative offset | Absolute address | Emitted operation | What it represents |
| ---: | ---: | --- | --- |
| `+0x1b6` | `0x2f22576` | `lea 0x30(%r14),%r10` | Address of `visited.bit_len` |
| `+0x1c0` | `0x2f22580` | `cmp %r12,(%r10)` | One unsigned current-slot/bit-length bound check |
| `+0x1c3` | `0x2f22583` | `jbe ...+0x2da` | Bound-error path |
| `+0x1c9` | `0x2f22589` | `mov %r12,%rdi` | Copy current slot |
| `+0x1cc` | `0x2f2258c` | `shr $0x6,%rdi` | One word-index computation |
| `+0x1d0` | `0x2f22590` | `cmp 0x28(%r14),%rdi` | One backing-word-length check |
| `+0x1d4` | `0x2f22594` | `jae ...+0x324` | Backing-word error path |
| `+0x1da..+0x1e3` | `0x2f2259a..0x2f225a3` | `mov $1,%r8d`; `mov %r12d,%ecx`; `shl %cl,%r8` | One mask materialization, used for marking |
| `+0x1e6` | `0x2f225a6` | `mov 0x20(%r14),%rcx` | One words-base load |
| `+0x1ea` | `0x2f225aa` | `mov (%rcx,%rdi,8),%r9` | One explicit visited-word load |
| `+0x1ee` | `0x2f225ae` | `bt %r12,%r9` | One membership test; the register form consumes the low six bit positions |
| `+0x1f2` | `0x2f225b2` | `jb ...+0x3d0` | Duplicate-sector error path |
| `+0x1fd` | `0x2f225bd` | `or %r8,(%rcx,%rdi,8)` | One memory read-modify-write mark |

The first check is the single check that survives for the caller's slot
bound and the two helpers' logical `bit_len` conditions: `prepare_visited`
sets `visited.bit_len` to `allocation_table.len()`. Disassembly cannot assign
which source condition LLVM selected as its origin, but there is no second
logical bit-length compare in this loop. The backing-word check also occurs
once, so the `get` and `get_mut` checks are coalesced to one check. The
allocation-table lookup after the slot check is emitted as a direct load at
`+0x23c` (`0x2f225fc`), with no second slice-bound branch.

The source-level word and mask calculations are also already coalesced: the
loop has one `shr` and one variable `shl`, not one copy for each helper. The
membership operation uses `bt` against the loaded register and therefore does
not materialize a separate `mask` for `contains`. There is only one explicit
visited-word load. The `or` has a memory operand, so it performs another
architectural read as part of its read-modify-write; it is not a second
explicit load instruction, and this disassembly alone does not make a cache or
uop claim.

For comparison, the out-of-line `CheckedBitSet::insert` symbol starts at
`0x2f142c0`. Its relevant offsets are `+0x0c` (`cmp bit_len`), `+0x15`
(`shr`), `+0x19` (`cmp words.len`), `+0x1f` (load words pointer), `+0x23..+0x2a`
(`1 << bit`), and `+0x2d` (memory `or`). It has no explicit word load because
standalone `insert` does not need to return the old bit. `contains` has no
out-of-line assembly row. This confirms that the collector row, rather than
the standalone helper row, is the right place to judge the combined call site.

## Profile attribution

The exact `collect_exact` function was attributed separately from its direct
callees. Each row below is per timed Callgrind dump; both repeats have the same
collector counts for the corresponding group.

| Profile group | Calls per dump | Collector self Ir | Direct-callee Ir | Collector inclusive Ir | Constructor inclusive Ir summed over five dumps (R1 / R2) |
| --- | ---: | ---: | ---: | ---: | ---: |
| XLS-owned | 10 | 1,120,228 | 43,120 (R1), 43,080 (R2) | 1,163,348 (R1), 1,163,308 (R2) | 11,318,644 / 11,317,729 |
| CFB few-large | 4 | 1,114,389 | 21,638 | 1,136,027 | 10,263,764 / 10,263,764 |
| CFB many-small | 256 | 152,897 | 11,347 | 164,244 | 13,975,077 / 13,975,077 |
| CFB tiny | 3 | 1,040 | 208 | 1,248 | 244,672 / 244,672 |

The dominant collector observations are therefore the XLS-owned and CFB
few-large rows. The retained 0548 comparison is consistent with this
attribution: its checkpoint candidate removed the baseline visited-bit
sequence but raised collector self Ir by 327,970 (5.855415%) for XLS-owned and
327,145 (5.871289%) for CFB few-large. This is evidence against adding scalar
state before proving that it removes emitted work; it is not evidence that a
different fused primitive will regress.

## Fused test-and-mark feasibility

A private checked operation with a result such as
`Result<bool, OleError>` could express the intended semantics: validate the
bit, locate one existing word, return whether the bit was already set, and set
it when absent. The caller would then report the same first duplicate and keep
the same append and successor-check order. The operation must retain the
current invalid-index precedence, error text, allocation labels and order,
zero-fill, scratch reset/reuse, separate FAT/MiniFAT maps, and collect-before-
claim behavior. Existing standalone `contains` and `insert` callers should
remain unchanged.

However, the safe source rewrite alone has no demonstrated machine-level
opportunity here. The current optimized loop already has one logical check,
one backing check, one word-index shift, one mask, one explicit word load, one
`bt`, and one memory `or`. A safe fused helper that performs
`let old = *value & mask != 0; *value |= mask` is likely to lower to the same
`mov`/`bt`/`or` shape, so it must be admitted only after exact release
disassembly proves a difference.

The concrete machine-level opportunity would be a genuine memory
test-and-set, for example an architecture-specific `bts`-equivalent that
returns the previous carry flag. Such an instruction could remove the explicit
word load, the mask materialization, and the separate `or` by doing one memory
read-modify-write. It would still need the checked logical and backing bounds
beforehand. If a prototype uses x86 `bts`, the full bit index must be applied to
the words base address; combining a full index with the already word-offset
address would address the wrong word. A portable safe implementation, an
architecture-specific intrinsic, or any inline assembly must each be inspected
for its actual lowering and safety contract. No unsafe indexing, unchecked
pointer arithmetic, atomics, or new dependency is justified by this baseline
alone.

## Recommendation and admission constraints

Do not retain a source-only `contains`/`insert` fusion on appearance. It should
be a narrowly scoped follow-up candidate only if a release build shows an
actual reduction in the `+0x1c0..+0x1fd` work, preferably eliminating the
explicit load/mask/OR sequence through a valid test-and-set lowering while
keeping the two bound checks. The candidate must then earn fresh matched
native XLS/CFB, allocation, exact profile attribution, malformed-input guard,
correctness, and final quality evidence. The sealed 0548 evidence cannot be
reused for admission.

If the fused candidate still emits the same `mov`/`bt`/`or` sequence, reject it
without further runtime capture and investigate a different representation or
compiler-supported primitive. The exact visited-bit walk remains the
authoritative cycle detector; no checkpoint or speculative replay is needed
for this hypothesis. OLE2 and OOXML remain the active optimization priority,
and ODF stays deferred.
