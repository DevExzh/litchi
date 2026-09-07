# Change 0459 lookup review

## Verdict

The one-line `ElementAttrs::lookup` experiment was semantically sound but was
rejected on performance evidence and reverted. No production change remains.
The restored source is revision `36df468bd` with the restored
`validation.rs` hash recorded in `decision.json`.

The measured candidate changed

```rust
cached.namespace.matches(namespace_uri) && cached.local_name.as_ref() == local_name
```

to

```rust
cached.local_name.as_ref() == local_name && cached.namespace.matches(namespace_uri)
```

in `crates/litchi-odp/src/codec/parser/codec/xml/validation.rs`.

## Semantic review

The reorder is exact. `ResolvedAttributeNamespace::matches` is a read-only
boolean check: `Bound` compares the borrowed URI bytes and `Unbound` or
`Unknown` returns false. `LocalName::as_ref()` is also a read-only byte-slice
comparison. Neither operand can return an error, mutate the resolver, consume
the attribute iterator, or alter lifetimes.

`ElementAttrs::get` still owns iterator advancement, malformed-attribute and
duplicate detection, and document-order first-match behavior. Value decoding
is still reached only after both predicates succeed. Therefore the reorder
cannot change error precedence, normalization, or returned values. Existing
validation tests already compare the cache with the unchanged one-shot reader
and cover same-local/different-namespace attributes, style fallback,
namespace shadowing, unknown prefixes, unqualified attributes, long URIs,
malformed attributes, duplicates, and harvest order.

The adversarial performance case is worth recording: the old order rejects
`Unbound` and `Unknown` entries after only the namespace discriminant, whereas
the candidate first compares their local names. The candidate can also do a
short local comparison for a same-local, wrong-namespace `Bound` entry. The
candidate therefore has a workload-dependent tradeoff; it is not a universal
optimization even though it is semantically equivalent.

## Assembly evidence

The compiler did not already reorder the baseline condition. These commands
were used against the retained release binaries:

```sh
nm -C /tmp/litchi-goal-0459/baseline/litchi-perf-baseline \
  | rg 'ElementAttrs.*lookup'
objdump -d --demangle --no-show-raw-insn \
  --disassemble='<litchi_odp::codec::parser::codec::xml::validation::ElementAttrs>::lookup' \
  /tmp/litchi-goal-0459/baseline/litchi-perf-baseline
objdump -d --demangle --no-show-raw-insn \
  --disassemble='<litchi_odp::codec::parser::codec::xml::validation::ElementAttrs>::lookup' \
  /tmp/litchi-goal-0459/candidate/litchi-perf-baseline
```

The relevant baseline sequence was:

```text
cmpq $0x0,(%rsi)       # Bound discriminant
cmp    %rcx,0x10(%rsi) # namespace URI length
call   bcmp            # namespace URI bytes
cmp    %r15,0x48(%r14) # local-name length
call   bcmp            # local-name bytes
```

The candidate sequence was:

```text
cmp    %r9,0x48(%rsi)  # local-name length
call   bcmp            # local-name bytes
cmpq   $0x0,(%r14)     # Bound discriminant
cmp    %r15,0x10(%r14) # namespace URI length
call   bcmp            # namespace URI bytes
```

Thus the source order reached the generated comparison order under the
release LTO build. The candidate could skip long URI comparisons when a
bound attribute's local name mismatched, but it could not remove namespace
work for same-local candidates or non-bound entries.

## Profile and measurement bounds

The unchanged-source diagnostic is a whole-process sampled profile, including
setup and warmups; inclusive shares overlap. In the DWARF recording,
`ElementAttrs::get` was 18.50% inclusive of the commit group and 30.48%
inclusive of snapshot opening; `ElementAttrs::lookup` was 6.12% and 9.99%
respectively. `memcmp` was 4.18% of commit and 9.27% of snapshot opening.
These figures establish a real hotspot but do not imply that the one-line
reorder can save the whole `get` share. It does not affect scan-time namespace
resolution, drawing-attribute harvesting, or transaction staging work.

The frozen matrix retained 720 operations across two repeats, normal and
allocator binaries, and tiny/medium/large corpora. The normal p50 deltas were
`-0.994%` (R1 medium), `-0.227%` (R1 large), `+13.327%` (R2 medium), and
`+0.131%` (R2 large). The predeclared gate required at least 5% improvement on
both medium and large inputs in both repeats; it failed. Six adverse timing
flags were retained, including `+43.039%` allocator-large p50. Allocation
volume, allocation/deallocation calls, peak-above-entry, retained-live
delta, and process maximum RSS showed no practically useful improvement.

The candidate epoch passed the serialized ODP release suite: **368 tests
passed, 0 failed**. No additional test is needed to prove commutativity of
these two pure predicates. The production experiment was nevertheless
reverted because the measured gate failed; the timing cause is not
established by the profile.

## Alternatives and follow-up

If this narrow idea is revisited, an enum-first form could preserve the old
cheap rejection for `Unbound` and `Unknown`, then check local name before URI
bytes for `Bound`. A local-name index could reduce repeated replay scans, but
it would add per-element memory and must preserve earliest document order,
duplicate behavior, and malformed-attribute reachability.

The accepted next direction is sharing staging metadata and `ContentSource`
event traversal and namespace maintenance with differential preservation and
error-priority tests. Retaining `Package` or `Presentation` is not a useful
alternative here: the family reopen is at most about 3% of transaction work,
while retaining the large XML owner increases memory retention.
