# 0534 final CFB physical role/FAT source review

`scope: read-only review of the frozen 0534 runtime and test diff`

`revision: 7c1d3911da286b8a9dd0fde15475bbbeed2f84c9`

`baseline source: crates/litchi-cfb/src/file.rs SHA-256 bb1928970dc3d652c5091d521cad86c8873ab8352eee5d2eed18a6aa4dd4c20b`

`candidate source: crates/litchi-cfb/src/file.rs SHA-256 a7dfb85e47de862f80350c335d4badc97df0b9cc68aa9ef1d2b8bd2309588e6f`

`candidate.patch SHA-256 8d6d393f6dc96aca8fbf89884790ab379047b24f57a8f0a563fd11bea3a96432`

`candidate/source.patch SHA-256 ce5a70b56fe9ad07ce24e6c512a63ef94efa754f0ad8fc85e3f647aaf1c0eeec`

`tests.patch SHA-256 d9af40480d8fa20c631d4cc2a56ee7f7c391c19d63af8517607887ecc3f92dcb`

`performance_claim: none`

This review covers the actual runtime hunks in [`candidate.patch`](candidate.patch)
and [`candidate/source.patch`](candidate/source.patch), plus the formatted
test hunk in [`tests.patch`](tests.patch), against the frozen [0534 plan](plan.json).
No Rust source, build, test,
benchmark, profiler, or analyzer was run for this review. The accepted ADR
manifest remains byte-identical to 0533 (`adr-manifest.json` SHA-256
`8aed97d378425650f3b355a3bd0ec5d2eb8000b9d28ba2a900cd8f35585e541f`).

## Actual source delta

The only runtime change is the body of the private
[`validate_physical_sector_layout`](../../../../crates/litchi-cfb/src/file.rs#L1178)
method. Its per-sector `fat.get(sector).copied()` lookup is replaced with
`sector_roles.iter().zip(&fat).enumerate()`, followed by one length check when
the FAT is shorter than the role map. The unclaimed-marker branch and both
error format strings are unchanged. There are no changes to `claim_sector`,
`claim_chain`, role publication, loaders, constants, call order, public API,
dependencies, unsafe code, or resource policy.

The paired iterator visits exactly the common prefix in ascending physical
sector order. For every pair it continues to check only
`PhysicalSectorRole::Unclaimed` and accepts only `FREESECT`; any other marker
returns the existing exact message:

```text
unclaimed physical sector {sector} has FAT marker 0x{entry:08X}
```

After that prefix, `fat.len() < sector_roles.len()` returns the existing exact
missing-entry message at `sector = fat.len()`:

```text
FAT does not contain an entry for physical sector {fat.len()}
```

This preserves the required precedence: an earlier bad unclaimed marker is
reported before a later missing-FAT entry. A valid prefix reports the first
missing index. An equal-length pair checks every role, a longer FAT leaves its
tail untouched, and empty roles remain successful even with a nonempty FAT.
The implementation borrows both vectors through `&self`, performs no success-
path allocation or mutation, and uses no physical offset, pointer, unchecked
index, or wrapping arithmetic.

## Test diff and baseline retention

The candidate adds two focused private tests:

* `physical_layout_checks_every_role_and_fat_marker` covers all seven role
  variants against `FREESECT`, `ENDOFCHAIN`, `FATSECT`, `DIFSECT`,
  `MAXREGSECT`, and an ordinary marker. It asserts the exact unclaimed error,
  successful treatment of claimed roles, and unchanged role/FAT vectors.
* `physical_layout_preserves_prefix_order_and_padding_contract` covers empty
  and equal arrays, short FATs at zero and after a valid prefix, marker-before-
  missing precedence, long arbitrary FAT padding, and empty roles with long
  padding. It also asserts no mutation for every row.

Both tests call the unchanged private method on a fresh synthetic
`OleFile`; neither names or assumes the candidate's paired iterator or
post-loop length branch. They therefore remain valid if the runtime hunk is
rejected and the baseline `get`-based loop is restored. The tests can stay in
the baseline independently: their helper and role table are existing test
support, and their expected results are the baseline method's behavior.

The rows provide the important length and precedence coverage. They do not
duplicate a short-FAT row with a valid `Unclaimed` prefix; the short claimed
prefix and bad-marker precedence rows exercise the same tail and ordering
branches. This is a minor matrix extension opportunity, not a source
correctness blocker, because the exact short-entry result and unclaimed
precedence are both asserted and the source preserves the baseline logic.

The existing real-file
`tolerates_nonfree_fat_padding_beyond_the_physical_file` test remains the
compatibility guard for producer padding, and the surrounding malformed-input,
chain, ownership, and role-claim tests remain applicable. No test asserts
iterator shape, instruction layout, allocation counts, or benchmark timing.

## Source disposition

The runtime hunk is narrow and contract-preserving. Static review found no
material source blocker, and both focused tests are safe to retain on the
baseline if the candidate is declined. Final retention still depends on the
parent-controlled exact-source custody, native/allocation/profile/assembly,
quality, and performance gates in the frozen plan; this review provides no
performance approval and does not claim those gates passed.
