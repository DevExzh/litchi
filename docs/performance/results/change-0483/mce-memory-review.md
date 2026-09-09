# MCE memory audit for the DOCX settings admission path

Status: **source audit complete; partial runtime evidence available; final change acceptance pending the full suite**.
This is a source-derived owner audit of
`litchi_ooxml_common::mce::process_markup_compatibility` and the current MCE
call made by the source-backed DOCX settings check. It is not a process RSS
measurement, and it does not claim that allocator physical capacity or
allocator metadata is identical to the requested logical bytes. No concrete
production blocker remains in the source mapping reviewed here.

The relevant requirements are ADR 0005's finite, hierarchical resource
budgets and fallible reservations and ADR 0006's fail-closed validation. The
source inspected here is `crates/litchi-ooxml-common/src/mce/codec.rs`; the
settings call sites are `crates/litchi-docx/src/source_backed/tail_append.rs`,
`crates/litchi-docx/src/source_backed/tail_append/mce_workspace.rs`,
`crates/litchi-docx/src/settings/extensions/package.rs`, and
`crates/litchi-docx/src/settings/document/{package,codec}.rs`.

## Fast path and the actual input to the bound

The first branch of `process_markup_compatibility` checks the input ceiling
and then searches for the literal MCE namespace URI (`codec.rs:535-549`). If
the literal is absent, it returns `Cow::Borrowed(xml)` and a zero report. That
branch does not construct the reader, event buffer, stack, output vector, or
any MCE-owned string. The output ceiling is still checked before returning.

The test is a byte-substring test, so a URI in ordinary text or an unrelated
attribute still takes the allocating path. A source-backed caller may use the
borrowed result only after this branch has returned; it cannot assume that a
document with no actual `mc:` element will take the fast path.

When the URI is present, let `S = xml.len()` for this particular pass and
`O = lim.max_output_bytes`. The input check proves `S <=
lim.max_input_bytes`, but it does not bound any individual XML event more
tightly than `S`. `max_depth`, `max_namespace_bindings`,
`max_directive_tokens`, and `max_choices_per_alternate` are finite, but the
legacy `Limits` type has no per-event attribute, name, decoded-value, or
directive-context byte ceiling.

## Owners created by one MCE pass

`Reader::from_reader(xml)` borrows the input, but the `read_event_into` path
uses a caller-owned event `Vec<u8>` and quick-xml's `ReaderState`. With its
default end-name checking, the reader retains the names of open elements in
`opened_buffer` and their start indexes in `opened_starts`. quick-xml records a
new start name before `start` checks `st.len()`, so a rejected over-depth start
can temporarily occupy `max_depth + 1` reader levels. The MCE stack itself is
checked before it builds a `Frame`, and can retain `max_depth` frames. These
are separate stacks.

For a start or empty event, the MCE code then creates the following owners.

* `raw: Vec<(String, String)>` stores a new `String` for every qualified
  attribute name and every decoded value (`codec.rs:653-667`). The
  quick-xml attribute iterator simultaneously owns its `Vec<Range<usize>>`
  duplicate-check state and, after its small-attribute threshold, a
  `HashSet<u64>` of name hashes.
* `local_namespaces` clones namespace prefixes and values before moving them
  into one `Arc<NamespaceLayer>` (`codec.rs:668-682`). A live `Namespaces`
  value keeps one parent chain of these layers. `Ctx::clone` and each frame
  clone only the `Arc` handle, so inherited namespace declarations are not
  copied once per frame. Re-declaring an existing prefix does not increase
  the binding counter, so `max_namespace_bindings` does not bound the number
  of declaration strings retained by the layers; the source-byte ceiling and
  the number of live layers do.
* `Ctx::clone` is shallow for both `Arc` chains, but each `Frame` retains a
  copy of the `Ctx` and a `Mode`. An emitting frame owns a `String` containing
  the qualified element name. The local `q` and the cloned name in
  `Mode::Emit` coexist while the frame is being built. The `Frame` copies add
  pointer-sized handles, not copies of any namespace or directive set.
* The MCE directive vector owns expanded local-name strings and borrows values
  from `raw`. `local_ign`, `new_ignorable`, and the three target sets are
  `HashSet`s. `local_ign` is a current-start temporary. `new_ignorable` is
  moved into a new `DirectiveLayer` only for URI values that are not already
  effective in the parent; inherited ignorable URIs therefore are not cloned
  into every layer. The three target sets retain one owned `NamePattern` per
  accepted `ProcessContent`, `PreserveElements`, or `PreserveAttributes`
  occurrence; their parent layers remain reachable through one `Arc` chain,
  so those target patterns are the genuine cross-depth amplification.
  Temporary `HashSet<&str>` values check duplicate `Ignorable` and
  `MustUnderstand` prefixes. These temporary sets overlap the current
  directive sets only during the current start event.
* `BoundedOutput` reserves `min(S, O)` immediately and can grow to `O` as
  `write_start`, text, CDATA, comments, references, and end tags are emitted.
  The output is retained in the returned `Cow::Owned` on success.
* `expand`, `parse_qname_target`, decoded-value conversion, and
  `MustUnderstand`/malformed-input diagnostics create transient owned names
  or error strings. They are not covered by the vector capacities above.

The code uses fallible reservation for several vectors and sets, but it uses
ordinary `to_string`, `clone`, `to_owned`, and `into_owned` for many payload
strings. An outer execution reservation therefore does not by itself turn all
MCE allocation failures into typed errors.

## Checked envelope

The following is a source-derived requested-storage envelope for the current
representation. It separates owners that are cloned from source text from
owners that are shared through `Arc`. A helper can use measurements from a
bounded preflight, or substitute the finite profile ceilings shown below.

The calculation uses the same pinned geometric convention as the XML
minifier audit. For a non-empty allocation, define:

```text
Str(bytes, count) = 2 * bytes + 8 * count
V<T>(count)       = max(8 * size_of::<T>(),
                        2 * count * size_of::<T>())
H<T>(count)       = max(8 * (size_of::<T>() + 1),
                        4 * (count + 1) * (size_of::<T>() + 1))
```

`Str` is a twofold byte-capacity allowance with an eight-byte per-`String`
floor. `V` is the corresponding geometric `Vec` allowance. `H` charges one
key slot plus one control byte with the fourfold table allowance and the
eight-entry floor used by `xml-minifier::audit`. A zero-count allocation
contributes zero. These are pinned requested-capacity terms, including table
growth headroom; they are not a portable claim about allocator metadata or
process RSS. The implementation must keep the pin next to the helper and
revisit it when its Rust/hash-table implementation changes.

Let:

```text
S  = actual input bytes for this MCE pass
O  = lim.max_output_bytes
D  = lim.max_depth
L  = D + 1                         // quick-xml depth-before-check floor
A  = maximum attributes in one start or empty event (A <= S)
T  = lim.max_directive_tokens per start element
B  = lim.max_namespace_bindings (fresh prefixes only)
C  = lim.max_choices_per_alternate (a scalar counter)

Kns = number of live NamespaceLayer allocations (Kns <= D)
n_l = local namespace declaration count in namespace layer l
N   = sum(n_l)                     // repeated redeclarations count
NS  = sum(prefix_bytes + uri_bytes for every local declaration)

U  = maximum bytes in one expanded namespace URI
Q  = maximum bytes in one qualified or local name

Kd  = number of live DirectiveLayer allocations (Kd <= D)
j_l = ignorable URI entry count in directive layer l
p_l, e_l, a_l = target counts in that layer's process,
                preserve_elements, and preserve_attributes sets
J   = sum(j_l)
Jb  = sum(bytes of the distinct URI strings in those j_l entries)
I   = current local_ignorable entry count (distinct URI strings)
Ib  = bytes in the current local_ignorable URI strings
Rp, Re, Ra = sums of p_l, e_l, a_l; R = Rp + Re + Ra
Pb  = bytes in all retained NamePattern URI and local-name strings
Qsum = bytes in the local-name fields of those NamePatterns
M   = current MCE directive-attribute count (M <= A)
E0  = fixed diagnostic/error-string overhead for this implementation
```

`N <= S`, `U <= S`, and `Q <= S` are safe source-only ceilings. The root
context starts with one binding, so at most `B - 1` declarations introduce a
fresh prefix; repeated prefix declarations still add a layer string and are
bounded only by `N` and `S`. `Kns` is at most `min(D, N)` and `Kd` is at most
`D`. Every retained target is one accepted directive token, so `R <= D*T`
and `Pb <= R*U + Qsum`, with `Qsum <= R*Q` when only a profile ceiling is
available.

The important distinction is in the ignorable counts. `local_ign` is the
temporary set for the current start event, so `I <= T`; it is not an
aggregate over the stack. `new_ignorable` is moved into the current
`DirectiveLayer` only when `c.is_ignorable(namespace)` is false. Because the
parent chain is retained and that lookup walks every ancestor, URI values in
different `DirectiveLayer::ignorable` sets are distinct effective values; an
inherited URI is not copied into each descendant layer. Thus `Jb` is a sum of
distinct retained URI expansions, not `D*T` copies of every inherited URI.
Without a preflight, `J <= min(D*T, S)` and `Jb <= J*U` are conservative.

For the current source, a checked owner sum is:

```text
event_buffer       = V<u8>(S)
reader_open_names  = V<u8>(S) + V<usize>(L)
mce_frames         = V<Frame>(D) + Str(S + 2*Q, D + 2)
raw_attributes     = Str(S, 2*A) + V<(String,String)>(A)
quickxml_attrs     = V<Range<usize>>(A) + H<u64>(A)
namespace_layers   = Str(NS, 2*N)
                    + sum_l(V<(String,String)>(n_l))
                    + Kns * (arc_ns + size_of::<NamespaceLayer>())
directive_vector   = Str(S, M) + V<(String,&str)>(A)
directive_payload  = Str(Jb, J) + Str(Ib, I) + Str(Pb, 2*R)
directive_tables   = H<String>(I) + H<&str>(T)
                    + sum_l(
                        H<String>(j_l)
                        + H<NamePattern>(p_l)
                        + H<NamePattern>(e_l)
                        + H<NamePattern>(a_l)
                      )
                    + Kd * (arc_dir + size_of::<DirectiveLayer>())
transient_names    = Str(U + Q, 2) + Str(U + Q, 2) + Str(S + E0, 2)
output             = V<u8>(O)
fixed_state        = size_of::<Reader<&[u8]>>()
                    + size_of::<BoundedOutput>()
                    + size_of::<Report>()
                    + capabilities_owner

mce_workspace = checked_sum(
    event_buffer,
    reader_open_names,
    mce_frames,
    raw_attributes,
    quickxml_attrs,
    namespace_layers,
    directive_vector,
    directive_payload,
    directive_tables,
    transient_names,
    output,
    fixed_state,
)
```

The `Str(S + 2*Q, D + 2)` term covers the source-sized qualified names
already retained by emitting frames, plus the current `q` and its cloned
`Mode::Emit` name, which coexist before `close` moves the frame. The
`transient_names` terms cover the expanded element name, one in-flight
expanded attribute/target name, and a decoded-value or diagnostic string. `E0`
is a finite implementation constant (or a profile field) for the fixed
diagnostic prefix/error formatting overhead; the dynamic offending source
bytes are charged by `S`.
`arc_ns` and `arc_dir` are the pinned allocation-header terms for one
`Arc<NamespaceLayer>` or `Arc<DirectiveLayer>` (the layer object sizes are
charged separately); `sum_l` is over the live layers, including the current
layer while its local sets are being built. The raw value term is source-sized
because XML character/reference decoding does not increase the byte total
beyond the source event; a caller that uses a different decoder can substitute
its measured decoded-event ceiling.

The `H` terms are per actual `HashSet`, not one table for an aggregate count:
each `DirectiveLayer` owns four separate sets, and `local_ign` owns another
set while the current layer is being built. The one `H<&str>(T)` term covers
whichever of the scoped `Ignorable` or `MustUnderstand` duplicate-check sets
is live; those scopes are sequential in `start`. The fourfold table/control
factor and growth floor are the pinned geometric allowance, so this formula
does not pretend that a `HashSet<T>` costs only `count * size_of::<T>()`.

For a no-preflight profile, use `A=N=U=Q=S`, `NS<=S`, `n_l<=S`, `I<=T`,
`Ib<=T*U`, `R<=D*T`, `Pb<=D*T*(U+Q)`, `J<=min(D*T,S)`, and `Jb<=J*U`.
The per-layer
table sums can be charged as `D * (H<String>(T) + 3*H<NamePattern>(T))`,
with `H<String>(T)` for `local_ign` and `H<&str>(T)` for the transient prefix
set. The namespace vector sum can similarly be charged as
`D*V<(String,String)>(S)` when no event attribute ceiling exists. A streaming
profile's finite event, attribute, name, and context ceilings make those terms
much smaller and allow the helper to reject before `codec::start` materializes
the raw event.

For this profile, `B` and `C` are finite, non-zero policy values. `B` limits
fresh visible prefix bindings but does not limit repeated declaration strings,
and `C` limits only the scalar `choices` field in `Mode::Alt`; neither is a
substitute for the source, layer, or target terms above. The `capabilities_owner` term is
the complete heap footprint of a caller-created `Capabilities` value; it is a
fixed profile term for the DOCX baseline and extension namespaces, but it is
live inside `extensions::process_bytes_with_limits` and must be included
there.

Using an unexplained `2 * input`, `4 * input`, or other fixed multiplier would
hide both the real `D*T*(U+Q)` target-pattern amplification and the separate
per-table owners. The pinned `V`, `Str`, and `H` terms keep those owners
visible while giving the admission guard an implementable checked sum.

### Current mapping to the settings preflight

The current `SettingsGuardFacts` can feed a conservative owner-local helper
without exposing the private MCE structs. Use `S = source_bytes`, `O =
mce_limits.max_output_bytes`, `D = mce_limits.max_depth`, `T =
mce_limits.max_directive_tokens`, `NS = max_namespace_buffer`, and `N =
namespace_declaration_count`. The preflight's semantic attribute count excludes
`xmlns` attributes, so the current mapping sets `A = max(1, max_token)`;
`codec::start` puts those declarations into `raw` and the quick-xml duplicate
checker sees them. Set `Q = max(1, max_token)` and `U = NS`; these are safe
event/name and retained-context ceilings from the existing facts. Use
`Kns = min(D, N)` and charge each namespace layer as
`2*size_of::<usize>() + size_of::<Option<Arc<()>>>() +
size_of::<Vec<(String,String)>>()`, followed by the `V` and `Str` terms.

`mce_directive_owned_bytes` is an upper bound for `Pb` because the preflight
counts every directive token plus each resolved target URI/local pair; its
extra token bytes are harmless. `mce_directive_tokens` is a document total,
so use `Ttot = min(mce_directive_tokens, S)` for `R`, `J`, and the retained
pattern string count. For the per-event temporary sets, use
`I = min(T, Ttot)` and `H<&str>(I)` when a separate per-event token maximum is
not recorded. Let `Kd = min(D, Ttot)`. The sum of the per-layer hash tables can
then be charged without knowing their distribution as
`4*(Ttot + Kd)*(size_of::<String>() + 1)` for ignorable tables and three times
`4*(Ttot + Kd)*(pattern_slot + 1)` for the target families, with `H<String>(I)`
and `H<&str>(I)` for the current temporaries. `pattern_slot` can be pinned as
`size_of::<Name>() + size_of::<String>() + size_of::<usize>()`, which
overbounds the private `NamePattern` enum on the supported layout.

### Historical pre-integration omissions (resolved)

The pre-integration design required `12*size_of::<usize>()` per
`Frame` as the declared-layout upper bound: `Ctx` has two `Arc` handles and a
binding word, `Mode::Emit` owns a three-word `String`, and the enum/frame
alignment can reach nine words. Charge `V<Frame>(D)` with that 12-word item
size, then `Str(S + 2*Q, D + 2)` for frame names. Each namespace or directive
layer also needs a two-word `Arc` allocation header; for the private directive
object, a safe declared object size is
`size_of::<Option<Arc<()>>>() + 4*size_of::<HashSet<String>>()` because all
four `HashSet` wrappers have the same layout. Add these terms to the layer
size and do not rely on cloned context handles to pay for the layer.
Before the owner-local helper existed, the settings workspace terms omitted
the reader open name/index capacities, output `V<u8>(O)`, quick-xml
range/hash tables, per-layer directive hash tables, both layer allocation
headers, and temporary decoded/expanded names; the then-current
`depth * 8 * usize` frame estimate could also understate the declared `Frame`
layout. The current `mce_workspace` helper charges each of those owners and
uses the declared 12-word frame upper bound described above.

## The directive amplification is real

The product term is reachable with valid MCE-shaped input. Bind one long
ignorable namespace URI once, then put up to `T` distinct
`ProcessContent`/`Preserve*` targets on each of `D` nested elements. Every
target pattern clones the same URI into its owned `Name`, even though the URI
appears only once in the source. With a large URI of length `U`, the retained
target strings alone are `D*T*(U+Q)` bytes. The
`max_namespace_bindings` counter alone does not prevent this: the nested
elements reuse one already-bound prefix.

Thus a bound based only on `S`, `O`, and a small constant number of source
copies is unsound. A finite profile is possible only when `D`, `T`, and name
bytes are all finite and the product is charged. The current legacy
`mce::Limits` provides no name-byte or context-byte ceiling, so substituting
`U=Q=S` is the only source-only bound and can be prohibitively large.
Even with a 4 MiB settings input, a profile permitting a roughly 64 KiB
namespace URI and hundreds of thousands of short target tokens can retain
many tens of GiB of cloned target-name payload before the input ceiling is
reached. The exact worst case depends on the source token packing; the point
is that it is a product of real source-reachable owners, not a constant input
copy.

## Current DOCX settings integration checkpoint

The source-backed settings path now performs one explicitly bounded byte-buffer
MCE pass and keeps the processed `Cow` live through the model phase. The
`PartData` returned by `settings_part.data()` remains live for the whole block;
when the package cache is managed, its decoded payload reservation therefore
continues to account for the original source owner. The MCE leases account for
the additional owners created by preprocessing.

`bounded_settings_mce_limits` uses the authenticated source length as
`max_input_bytes`, while `max_output_bytes` is the checked
`max_settings_xml_bytes` ceiling. The latter must be allowed to exceed the
source length because `write_start` emits every effective namespace declaration
on an emitted element; a child can therefore repeat inherited declarations.
`mce_workspace_requirement` passes that same output ceiling to
`mce_workspace::Profile`, and `output_memory_requirement` charges exactly the
same `V<u8>(O)` term as the `output_owner` term in the aggregate helper.

The reservation sequence is now source-derived and composable:

1. The source and root guards take short-lived scanner reservations while
   `settings_data` remains live.
2. The operation reserves `output_bound = V<u8>(O)` and then
   `scratch_bound = mce_workspace - output_bound` before calling the MCE
   processor. A scratch-reservation failure drops the already acquired output
   lease through normal scope unwinding.
3. After the call, the scratch lease is dropped. A borrowed fast-path result
   drops the output lease as well; an owned result retains the full output
   lease while the post-MCE guard and model run.
4. The post-MCE guard recounts the actual processed bytes and namespace facts
   with the remaining workspace allowance. Its scanner lease ends when the
   guard returns. The model envelope is then checked as
   `retained_output + model_workspace` and reserved while the processed output
   remains live.

This sequence does not add the post-MCE scanner lease to the model peak because
that lease is released before `model_workspace` is reserved. It also does not
charge the source bytes a second time in the MCE helper: the managed
`PartData`/cache owner already holds the source payload reservation, while the
MCE helper charges only its borrowed parser state, cloned context, and output.

The source-derived profile has these concrete properties:

* `facts.max_namespace_bindings` is passed to the MCE policy. It includes the
  guard's implicit binding floor and source declarations. This matters because
  `Ctx::root` starts with `bindings: 1`; passing the declaration-only total was
  one binding too small and rejected a small valid MCE fixture.
* `namespace_bytes` is the maximum live namespace buffer and is no longer
  clipped to `source_bytes`. `namespace_declaration_count` is only the
  aggregate declaration count used to bound live layer vectors and their
  per-layer eight-entry floors. Since inherited `Arc<NamespaceLayer>` chains
  share their parents, this does not multiply one URI by the depth.
* `directive_tokens` and `directive_owned_bytes` are document totals. The
  helper distributes the total across at most `min(depth, directive_tokens)`
  layers, charges all four per-layer hash tables, and separately charges the
  current temporary sets. This preserves the real repeated-target URI/local
  amplification while avoiding a clone of every inherited directive set.
* `max_directive_tokens_per_event` currently receives the document total because
  the preflight has not recorded a per-event maximum. That is safe but can
  overreserve the current-event hash tables. A future tightening can add a
  per-event fact without changing the aggregate bound or MCE policy.
* `namespace_copy_bytes` is an aggregate scanner fact. It is not used as a
  live MCE context term; the MCE layers retain only the current parent chain,
  so `max_namespace_buffer` is the correct payload input for this helper.

The post-MCE guard sets quick-xml's per-element declaration ceiling to
`bytes.len().clamp(1, 256)`, matching the scanner workspace envelope. This is
a deliberate finite parser admission ceiling and does not understate the MCE
lease. A valid MCE output that places more than 256 effective inherited
declarations on one emitted start tag will be refused during recount; if that
profile is intended to accept such documents, the declaration ceiling and its
scanner reservation must be widened together.

The remaining arithmetic is therefore a conservative requested-capacity
envelope, rather than an unexplained source multiplier. The only correctness
fix found in this checkpoint was the implicit namespace-binding baseline; the
current source call site supplies it.

## Acceptance evidence and remaining review status

The integration owner reports passing runtime coverage for borrowed
empty-settings processing and MCE output that expands through inherited
namespace declarations. Existing memory-limit cases cover the surrounding
workspace reservations. The narrow-output-cap and MustUnderstand cases are
present in test source but have not yet run in the full gate. This document
records the source audit and did not run Cargo, tests, formatting, or Git
commands independently.

The remaining acceptance evidence is the full-suite result, including the
narrow-output-cap and MustUnderstand cases, followed by the repository's normal
change manifest/update review. Those are evidence tasks, not unresolved
arithmetic findings in this audit.

There is no concrete production blocker identified here. The 256 declarations
per output element setting is an intentional finite post-MCE parser ceiling;
it only becomes a production issue if the accepted DOCX profile is required to
admit an output element with more than 256 effective inherited declarations.
The aggregate document total used for the current-event directive table is a
safe over-reservation and can be tightened later with a per-event fact.
