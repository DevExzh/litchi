# 0359rc — RTF underline call-inclusive attribution

- Date: 2026-10-22.
- Status: attribution complete; the 0359rb candidate remains unaccepted and
  native qualification remains open.
- Candidate source commit: `68f976eb00`.

This record corrects the successful-route control in 0359rb. The earlier
no-inline `success` wrapper was measured only as a direct function, so its
reported 200/194 Ir per call excluded the edit and commit callees. This
record uses Callgrind call-inclusive function and edge costs for the same
non-empty operation, and repeats the cold comparison with the already-built
current and candidate binaries.

## Workload and method

The retained private probe is
`/home/zhuhe/code/litchi-worktrees/.coord/rtf-refusals/private-probes/underline_route_probe.rs`
(1,177 bytes, SHA-256
`c2f7ec5c3f121cabf3b143bd459cbed9dbfebe08a17e0ef7eca1520331a3a618`). Its
cold mode parses 34 fresh `Alpha` packages and successfully applies the
non-empty span `0..5` and commits it once per package. Its warm mode
materializes one package's text and repeats the same successful edit and
commit 34,000 times. The 0359b refusal probe keeps the eight original
scenarios and makes 34,000 unsupported empty-span calls.

Callgrind 3.26.0 collected `Ir` on CPU 7 with:

```sh
TMPDIR=<external-tmp> taskset -c 7 valgrind --tool=callgrind \
  --dump-instr=yes --collect-jumps=yes \
  --callgrind-out-file=<external-profile> <binary> [mode]
```

The four new cold profiles were run in alternating order current, candidate,
current, candidate. `callgrind_annotate --inclusive=yes --tree=none` supplied
function totals, and the raw `cfn` records supplied caller-to-callee costs.
The two runs of each binary have identical function and edge values. This
rules out run-to-run variation in the measured cold delta; it does not by
itself identify a source-level cause for the changed allocator descendants.

## Cold successful route

The repeated cold profiles report these call-inclusive values. The edge rows
are the cost charged to that exact caller-to-callee edge, not a direct body
cost.

| Function or edge | Current | Candidate | Delta |
| --- | ---: | ---: | ---: |
| Whole profile | 2,639,304 | 2,658,402 | +19,098 |
| `underline_route_probe::main` | 2,343,392 | 2,362,468 | +19,076 |
| `main → Tree::parse_internal` | 674,232 | 699,272 | +25,040 |
| `main → success` | 1,599,750 | 1,593,580 | −6,170 |
| `success → Edit::commit_staged` | 1,538,068 | 1,532,624 | −5,444 |
| `success → underline_for_span` | 8,840 | 8,840 | 0 |
| `Edit::commit_staged → Package::from_bytes_with_options` | 1,044,644 | 1,039,026 | −5,618 |
| `Package::text` | 7,519 | 7,519 | 0 |
| `Tree::parse_internal` | 1,678,246 | 1,695,953 | +17,707 |
| `Tree::parse_string` | 1,525,554 | 1,542,435 | +16,881 |
| `Parser::parse` | 666,820 | 681,279 | +14,459 |
| `Parser::parse_group` | 483,352 | 497,980 | +14,628 |
| `Lexer::tokenize_with_spans` | 413,362 | 414,794 | +1,432 |
| `RawVecInner::finish_grow` | 225,466 | 240,247 | +14,781 |

The direct bodies of the probe's success function, `Edit::commit_staged`, the
parser functions and `underline_for_span` are unchanged in the flat profiles.
For example, the direct success body is 6,800 Ir per call in both binaries,
and the direct `Edit::commit_staged` body is 61,404 Ir per call in both. The
inclusive successful route therefore does not show an added edit or commit
cost. Under the original distinct executable paths, the cold whole-profile
increase is concentrated in the parser's nested allocation descendants shown
below. The same-path control later in this record removes that increase.

## Nested cold parser edges

These exact edges explain the observed parser delta in the profiles. Their
logical call counts match between binaries and are identical in both repeats.

| Caller → callee | Calls | Current | Candidate | Delta |
| --- | ---: | ---: | ---: | ---: |
| `parse_group → RawVec::grow_one` | 68 | 20,158 | 36,169 | +16,011 |
| `RawVec::grow_one → RawVecInner::finish_grow` | 272 | 47,024 | 61,431 | +14,407 |
| `RawVecInner::finish_grow → realloc` | 374 | 145,088 | 153,848 | +8,760 |
| `realloc → _int_realloc` | 375 | 125,615 | 133,913 | +8,298 |
| `0x4c3d680 → __memcpy_avx_unaligned_erms` | 299 | 10,221 | 18,905 | +8,684 |
| `RawVecInner::finish_grow → malloc` | 612 | 36,382 | 42,403 | +6,021 |
| `_int_realloc → _int_malloc` | 250/262 | 53,168 | 46,692 | −6,476 |

The `hashbrown::RawTable` drop cost is 5,508 Ir in both profiles, so the
observed change is not a changed hash-table drop path. The repeated profiles
also preserve the same parser and allocation edge call counts. These are
measured allocator/libc descendants for the distinct-path runs. The fixed-path
identity control below tests whether they persist when executable path,
`argv[0]` and process startup state are held constant.

## Same-path identity control

To control for process identity and allocator startup state, each already-built
binary was copied in turn to the same external path:

`/home/zhuhe/code/litchi-worktrees/.coord/rtf-refusals/same-path-20261022/underline_route_probe`

The current and candidate copies retained their original SHA-256 values, and
each was run with the same `argv[0]`, environment, CPU 7 pinning and cold
successful probe. The two current and two candidate profiles were identical
for every listed value:

| Function or edge | Current r1/r2 | Candidate r1/r2 |
| --- | ---: | ---: |
| Whole profile | 2,657,559 | 2,657,559 |
| `underline_route_probe::main` | 2,361,674 | 2,361,674 |
| `Tree::parse_internal` | 1,695,137 | 1,695,137 |
| `Tree::parse_string` | 1,542,435 | 1,542,435 |
| `Parser::parse` | 681,279 | 681,279 |
| `Parser::parse_group` | 497,980 | 497,980 |
| `RawVecInner::finish_grow` | 240,247 | 240,247 |
| `success` | 1,593,170 | 1,593,170 |
| `Edit::commit_staged` | 1,532,214 | 1,532,214 |
| `hashbrown::RawTable` drop | 5,508 | 5,508 |

The allocator, realloc and memcpy descendants were also identical. Thus the
`+25,040 Ir` `main → Tree::parse_internal` edge and the associated cold parser
increase from the distinct-path runs do not persist under the matched process
identity. This is an identity control, not a native timing qualification, and
no allocated-size or growth-size comparison is needed because the instruction
delta itself disappears under the control. The cold route remains open for the
owner's isolated native run.

## Warm successful route

The retained warm profiles now have call-inclusive route costs. They separate
the repeated non-empty edit from first-package parsing:

| Function or edge | Current | Candidate | Delta |
| --- | ---: | ---: | ---: |
| Whole profile | 1,581,012,112 | 1,576,402,637 | −4,609,475 |
| `underline_route_probe::main` | 1,580,716,501 | 1,576,107,027 | −4,609,474 |
| `main → success` (34,000 calls) | 1,580,556,062 | 1,575,946,609 | −4,609,453 |
| `success → Edit::commit_staged` | 1,524,866,086 | 1,520,256,516 | −4,609,570 |
| `success → underline_for_span` | 8,840,000 | 8,840,000 | 0 |
| `Edit::commit_staged → Package::from_bytes_with_options` | 1,035,875,428 | 1,034,342,970 | −1,532,458 |
| `Package::text` | 7,310,377 | 7,310,209 | −168 |

The direct `success` and `Edit::commit_staged` bodies remain 6,596,000 and
61,404,000 Ir respectively in both warm profiles. The inclusive warm route
improves; the first cold parser increase does not appear in this warmed
non-empty control.

## Refusal route and matched normal controls

The repeated refusal profiles use the existing 0359b eight-scenario probe.
The unsupported-content edge is a named, non-inlined method in this binary:

| Function or edge | Current | Candidate | Delta |
| --- | ---: | ---: | ---: |
| Whole profile | 627,438,721 | 626,418,455 | −1,020,266 |
| `refusal_contract_bench::main` | 627,140,674 | 626,120,384 | −1,020,290 |
| `main → Edit::set_text_underline` (34,000 calls) | 1,496,000 | 476,000 | −1,020,000 |
| `Edit::set_text_underline` Ir per call | 44 | 14 | −30 |
| `Edit::commit_staged` inclusive | 347,859,241 | 347,858,413 | −828 |
| `Edit::commit_staged` direct | 3,553,967 | 3,553,967 | 0 |
| `Package::text` inclusive | 2,206,446 | 2,206,446 | 0 |
| `Package::apply` inclusive | 1,679,600 | 1,679,600 | 0 |
| `Edit::replace_text` inclusive | 1,489,647 | 1,489,647 | 0 |
| `tail_append::Edit::commit` inclusive | 127,067,737 | 127,068,219 | +482 |

The normal operation and publication blocks retain the same direct instruction
costs in the matched probe: `replace_text` 569,279, `commit` 268,167,
`commit_staged` 3,553,967, tail-append publication 14,892, `to_bytes` 2,652,
and `apply` 1,159,400. The input, output checks and dispositions are the same
as 0359b. Existing 0359b heaptrack captures also found identical allocation
summaries for the refusal comparison. No new allocation claim is made for the
successful cold route from Callgrind alone.

## Acceptance boundary

The inclusive evidence establishes that the fast path removes the refusal
producer cost and does not add a measured successful edit or commit body cost.
The distinct-path cold parser increase is eliminated by the fixed executable
path control, so it is not a persistent candidate cost in this matched probe.
No source-level repair was tested in this follow-up, and the candidate remains
unaccepted. Owner acceptance still requires isolated native A-B-B-A timing and
paired heaptrack on the exact current/candidate binaries, followed by the
complete RTF correctness, corpus/fidelity and glossary checks.

## External artifacts

The profiles remain outside the repository. The eight new Callgrind files are
the alternating refusal and cold-success runs:

| Artifact | Size | SHA-256 |
| --- | ---: | --- |
| `current-r1.callgrind` | 516,785 B | `b607d9e7217315330fc6abfd17a28c47f887dff67bfc5114734c6f29543e7952` |
| `fast-r1.callgrind` | 516,835 B | `5435e519b19664e6d6dbb1d5dbb1d2eeea6675b71804231fe0e99fa6011eba8a` |
| `current-r2.callgrind` | 516,355 B | `776d30a140562d5666bec18198606e4f93992a54ef59962be10998182892006f` |
| `fast-r2.callgrind` | 516,952 B | `8beb35066d0bee7c382f014ea6dd98c136816555618280a7ac6714b4f62754e0` |
| `success-current-r1.callgrind` | 442,304 B | `a763166d21026ded6cff9ae35fb7cb6b17482146522ac23e939731a4be8d04b1` |
| `success-fast-r1.callgrind` | 442,953 B | `2132e89af5f7d8de0ac80b080949b6f7fae9ef7e68825f4231016e56208946a5` |
| `success-current-r2.callgrind` | 442,304 B | `9e050a9bfb1b1d27ca8fbeb5812d968ce459e78965f53e6fd2a16578af3fca0d` |
| `success-fast-r2.callgrind` | 442,953 B | `5d47d3df8cb2315bab9aae9d14de45778fcf1c741c9467a1933e6177897d730a` |

The same-path control profiles are:

| Artifact | Size | SHA-256 |
| --- | ---: | --- |
| `same-current-r1.callgrind` | 442,474 B | `501726679a0aa4f17aaa96ff94e939eae1987c92236c3926b083015f6337be22` |
| `same-fast-r1.callgrind` | 442,322 B | `2879dc0c865203aa5aec52d8a31ef6f562399921f08d3b455e64d9454af69391` |
| `same-current-r2.callgrind` | 442,474 B | `4496c6eca88257470b1d32d9bc30b7e22624361e641a16fde374224291e915a2` |
| `same-fast-r2.callgrind` | 442,322 B | `a99d6731723e183383a53e5d95b19a311592c69564dfa6b72a30fe8b81070895` |
