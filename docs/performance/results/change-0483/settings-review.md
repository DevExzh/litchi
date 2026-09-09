# DOCX tail append settings topology review

Status: **accepted**. This is a read only review of the current settings
admission path used by `source_backed::tail_append`. The findings are about
the code presently in the worktree; root reports 1,378 all-feature tests
passing on the frozen source, including 40 focused tail tests. No build or
test was run by this review. The relevant contract is the bounded, fail closed
one paragraph closure. A settings helper is acceptable only when its complete
live owner set is covered by the operation's explicit finite policy.

## What the current path preserves

The current path now performs a guarded source preflight, one settings-capable
MCE pass, and the complete borrowed `DocumentSettings` grammar. The full
codec checks the three values that can make a changed document unsafe to
admit:

* `w:documentProtection` defaults to enabled when `w:enforcement` is absent;
* `w:writeProtection` defaults to enabled when `w:val` is absent; and
* `w:trackRevisions` defaults to enabled when `w:val` is absent.

`validate_topology` refuses when any of those flags is true. The full codec
also rejects duplicate protection/tracking settings, keeps the
package/main/settings dialect comparison, validates attached-template
relationships without collecting a matching `Vec`, and checks mail-merge IDs
against the original borrowed `Relationships` map. The root probe and model
reservations are sequential on the effective package context. This review
audits the checked requested-storage terms and records their boundary
evidence.

## Historical findings now resolved

The earlier reduced-projection review identified five blockers. The current
worktree resolves their semantic shape: MCE is processed once with the DOCX
settings capability set; the root probe owns a separate sequential reservation;
the full settings and mail-merge grammar is used; duplicate protection and
tracking nodes fail closed; and attached-template matching is streaming rather
than a collected relationship `Vec`. Those earlier details are retained here
only to explain why the current review requires the full model path. The
current helper covers the complete additive model/storage owner set; the
integration findings below confirm that coverage and distinguish proof
hardening from semantic blockers.

The external source/header relationship mode remains safe under the topology
caller's package-wide external-relationship rejection. Keep that invariant
adjacent to the relationship helper if it is reused elsewhere.

### 5. Current integration audit

The current `validate_topology` ownership sequence is now coherent:

* `mce_workspace::memory_requirement` includes the bounded output owner. The
  caller splits that result into an output lease and an MCE scratch lease,
  keeps both while `process_bytes_with_mce_limits` runs, and drops scratch as
  soon as the MCE pass returns. A borrowed fast-path result releases the
  output lease; an owned result keeps the output lease while later phases use
  the processed bytes. The MCE input cap remains the actual source length,
  while the output cap is the explicit `max_settings_xml_bytes` limit, so
  namespace reinjection may produce output larger than the source without
  escaping the reserved output envelope.
* The processed bytes are recounted with `guard_settings_xml` after MCE. Its
  scanner reservation is checked against the remaining workspace after the
  retained output lease, then released before model admission. The model
  helper receives the post-MCE byte length and facts, and the model lease is
  checked and reserved together with the retained output. This closes the
  source-facts-versus-MCE-output gap and keeps the planned root-probe lease
  outside the model total.
* The settings part's decoded bytes remain held by the OPC cache reservation
  in managed contexts while these leases are active. The topology path does
  not silently omit the effective package execution context when the caller
  does not pass one.

The root and scanner resolver caps are now explicit. `inspect_root_dialect`
sets the 256-declaration cap, and both source and post-MCE `guard_settings_xml`
passes set the same cap (or the processed byte length when smaller). The full
mail-merge, extension, and direct-settings readers receive an immutable slice
only after that post-MCE guard has accepted it. Their quick-xml default of 256
therefore cannot be used to bypass the guard's per-element declaration bound;
keeping the dependency default pinned or setting the same cap in those readers
would still make the invariant easier to audit, but it is not a current
unbounded-allocation finding.

Cancellation is also a documented boundary rather than a missing owner. The
caller checks the token before and after the synchronous MCE call and before
and after the synchronous full model/relationship call, and prepays the
bounded model work for that phase. A token cancelled during one of those
calls is observed at the next boundary. The public `Options` documentation
describes settings validation boundaries, so this is not a blocker unless a
future API promise changes to require interruption inside a codec call.

The helper now covers the three concrete terms that were previously open:

* `resolver_storage` charges four geometric binding slots per guarded binding,
  which covers the reader's retained capacity and the exact binding-vector
  clone made by each model-parser event. Its four namespace-byte layers cover
  the corresponding reader/clone buffer capacities.
* `semantic_bytes` charges twice the counted semantic payload plus an eight-byte
  floor for each concrete name/namespace/value field. This covers the lossy
  name construction, decoded settings strings, and their geometric capacity
  without multiplying the whole input by an arbitrary factor.
* The direct settings model charges `2 * token + 8` for the attached-template
  relationship ID. That matches the parser's actual caller-token bound while
  preserving its existing full grammar; the attached target URI remains
  bounded by the package validator's 32 KiB check.

I found no additional uncharged heap owner in the current full-parser order.
Root's final focused-tail validation supplies the runtime boundary evidence
for this checked envelope, including an attached-template ID near the token
ceiling, resolver clones, lossy namespace strings, and unknown-extension
rewrite overlap. The source-bound arithmetic is conservative and accepted,
rather than blocked by a known missing owner.

The remaining accounting is conservative rather than silently missing an
owner: MCE raw attributes use the authenticated token ceiling because the
preflight's per-event attribute count excludes `xmlns`, and the post-MCE guard
recounts namespace bindings, semantic strings, events, and depth before the
model lease. The root-probe reservation, MCE scratch lease, post-MCE scanner
lease, model lease, and retained output lease are sequential/overlapping in
the stated order, with each overlap checked against the same managed budget.

### 6. Current full grammar path and its owner obligations

The worktree now uses a borrowed package-boundary helper that accepts the
settings bytes and original `Relationships` map, runs the existing full
settings/mail-merge/extension grammar, and returns the security values after
relationship validation. This preserves protection defaults, duplicate/order
rules, namespace checks, ODSO validation, and relationship semantics. It
consumes one already-selected MCE output with the settings capability set;
`extract_from_part` remains the ordinary package API and is not the bounded
tail path.

That change is not a license to charge the helper as `input * N`. Its live
owners are concrete and need a checked finite envelope. Let `S` be the actual
decoded settings source bytes (and account for the package cache reservation,
or transfer that ownership into this operation), `T` the caller token bound,
and `D` the caller depth bound. Before parsing, the admission calculation must
cover the peak of:

* one MCE output plus its start-tag attributes, namespace bindings, directive
  sets, alternate-choice state, and open frames, all derived from `S`, `T`,
  `D`, and the explicit MCE caps;
* the full mail-merge `Node` graph while schema validation runs. Each node owns
  an `OwnedNamespace`, local name, attribute vector and strings, child vector,
  and the parser retains the root plus the open stack. Its finite cap is
  governed by `MAX_NODES`, `MAX_ATTRIBUTES_PER_NODE`,
  `MAX_STRING_BYTES`, `MAX_FIELD_MAPS`, `MAX_RECIPIENTS`, and the
  relationship-ID limit; a source-size bound still has to include per-node and
  per-vector overhead rather than assuming the graph is no larger than `S`;
* the complete settings model and its temporary parser state, including
  compatibility vectors, smart tags, the extension vector, and opaque
  extension copies. `Extensions::parse` can copy unknown ranges and
  `make_self_contained` can append active namespace declarations, so the
  opaque total needs a checked expansion bound using `MAX_EXTENSIONS`,
  `MAX_OPAQUE_BYTES`, namespace-binding limits, and `S`;
* transient `NsReader` events/resolvers, `active_bindings` and
  `make_self_contained` scratch vectors, and the bounded relationship-ID
  summary; and
* relationship-validation scratch. The current attached-template helper
  now counts matching relationships in a scalar; retain that streaming shape
  so relationship-map cardinality does not become a hidden allocation.

The current extraction order drops the mail-merge tree before building the
extensions model, so the peak is the maximum of those phases plus the final
direct settings model. A later security scan should not be added if the full
model already supplies the flags. If the derived sum does not fit
`max_workspace_bytes`, refuse admission with the typed resource error. Do not
narrow the accepted document grammar to inputs whose flags happen to be
disabled.

The root probe is now a separate sequential owner. Keep its reservation out of
the model phase total, and keep one model reservation whose ownership is clear
for managed and unmanaged contexts; a borrowed signature alone does not make
model or cache allocations free.

### 7. Concrete owner expression for the current full-model order

The current borrowed helper parses the full mail-merge tree first, then
`Extensions::parse`, then the direct `DocumentSettings` model. That order is
useful because the tree is dropped before the extension model is built. The
post-MCE `guard_settings_xml` pass now establishes the input-size,
token/event/depth, namespace, attribute, and semantic-byte facts before that
owning call. The parser's local `MAX_NODES = 1_000_000` and `MAX_DEPTH = 128`
checks remain semantic refusal limits; the workspace helper may reserve from
the broader guarded depth/node facts and therefore stays conservative. The
guard's token window also bounds the otherwise early
`read_event().into_owned()` allocation in the full parser.

For a processed XML slice, define these measured values before admitting the
full parser:

```text
X       = processed XML byte length
T       = caller token ceiling (also enforced before every read_event().into_owned())
D       = min(caller depth ceiling, 128)       # mail-merge Node::stack limit
E       = element events counted by the guard
A       = non-xmlns attributes counted by the guard
F       = w:fieldMapData children in the ODSO subtree
U       = bytes of resolved namespace/prefix strings that own_namespace() will copy
V       = bytes of decoded Node attribute values retained by the tree
Vsel    = bytes of Node attribute values cloned into MailMerge Settings
R       = sum of raw unknown-extension ranges copied by copy_range()
Q       = sum of namespace declaration bytes added by make_self_contained()
```

The post-MCE guard first checks `X <= min(caller max_settings_xml_bytes,
16 MiB)` and applies the caller's token/event/depth policy before the model
lease is reserved. The existing constants then give `E <= 1_000_000` for the
mail-merge tree, `D <= 128`, `F <= 16_384`, and at most 256 non-namespace
attributes on each mail-merge node. The guard's measured facts are still
needed for workspace admission because these local constants do not describe
the actual string and vector capacities.

For the mail-merge tree, charge the following actual capacities (or make a
counting pass and call `try_reserve_exact` so these capacities are known before
the owning pass):

```text
G_tree = size_of::<Node>()                    # root Option storage
        + capacity(stack) * size_of::<Node>()
        + sum_i(capacity(node_i.children) * size_of::<Node>())
        + sum_i(capacity(node_i.attributes) * size_of::<Attribute>())
        + C_tree_strings                       # names, namespaces, values
        + C_resolver_clone                     # reader + per-event clone
        + reallocation_overlap
```

The node and attribute values themselves live in the `stack`, child, and
attribute `Vec` buffers, so adding `N * size_of::<Node>()` or
`A * size_of::<Attribute>()` on top of those capacities would double-count
them. `N` is still the counted element total (`N <= min(E, 1_000_000)`) used
to derive the per-node capacities. `C_tree_strings` is the implementation's
geometric capacity for the counted lossy names/namespaces plus decoded
attribute values. The output length of
`String::from_utf8_lossy(...).into_owned()` is at most three times its input
bytes, but replacement-heavy values can leave capacity above that final
length; charge that capacity or reserve exact lengths. `U` must be counted
separately: every element and attribute namespace is converted into a new
owned `String` by `own_namespace()`, even when the same URI was declared once
on an ancestor. Thus a long root namespace URI can be cloned once per node and
is not bounded by `X` alone. The decoded attribute-value bytes are included in
the counted value term; retain their measured sum as `V` if the parser uses a
decoder whose expansion is not covered by the raw-byte bound.

`C_resolver_clone` is not just `size_of::<NsReader>()`: quick-xml's
`NamespaceResolver::clone()` clones both its namespace byte buffer and its
binding vector on every model-parser event. The current helper's four-layer
namespace-byte and binding terms cover the reader's geometric capacity and one
per-event clone capacity, with the binding-vector growth factor included.

The current `Vec::push` calls leave capacities allocator-dependent and can
hold old and new buffers during growth. `capacity(...)` above therefore means
the actual reserved capacities plus the largest simultaneous old/new buffer
at a reallocation. A prepass that counts each node's children and attributes,
followed by `try_reserve_exact`, removes that unknown term. Without that
change, no proof from only `N`, `A`, and `D` exists; a fixed multiplier on
`X` is not a substitute.

While `parse_mail_merge` walks the retained tree, the selected model owns a
second copy of selected attribute values. Charge:

```text
G_mm_model = size_of::<MailMergeSettings>()
            + capacity(field_maps) * size_of::<FieldMap>()
            + M_text + M_relationship + M_temporary

M_text         = Vsel for connectString/query/addressFieldName/mailSubject,
                 ODSO udl/table/type, and fieldMapData name/mappedName/lid
M_relationship = at most 4 * min(T, 1024) bytes
                 (dataSource, headerSource, odso/src, odso/recipientData)
M_temporary    = max selected attribute value length <= min(T, 16 MiB)
```

`F <= 16_384` is a count limit, not a justification for multiplying the
entire input by 16,384. `M_text` must be the counted sum of the selected
attribute values (`Vsel <= V`); the per-field `MAX_STRING_BYTES = 16 MiB`
check remains a necessary local guard. `schema_attribute()` clones one
attribute value for each lookup, so the temporary term is one such field at a
time. The settings mail-merge path does not parse the recipient-data part; if
that parser is added, its `MAX_RECIPIENTS = 1_000_000`, decoded unique-tag
limit, and base64 compact/decoded/re-encoded temporaries need a separate
term.

The extension model has a different, measurable amplification. For each
unknown direct extension `j`, let `R_j` be the copied source range and `Q_j`
the declarations inserted to make it self-contained. Then:

```text
G_ext_model = capacity(extension_values) * size_of::<Extension>()
            + R + Q
G_ext_peak  = G_ext_model
            + max_j(R_j + bind_scratch_j + declared_scratch_j)
            + event_capacity
```

`R = sum(R_j) <= X`, but `Q = sum(Q_j)` can exceed `X` because the same
ancestor bindings are repeated on each unknown child. The exact value is
already available from `make_self_contained()`:
`Q_j = sum(10 + prefix_len + namespace_len)` for each active binding not
declared on that child. Charge this sum before retaining the output. During
one rewrite, `copy_range()`'s `R_j` and the new `R_j + Q_j` output overlap, so
the extra `max_j(R_j)` term is required. `active_bindings()` also owns one
prefix/namespace tuple vector and `make_self_contained()` owns a declared
prefix `HashSet`; charge their measured bytes as `bind_scratch_j` and
`declared_scratch_j`. The model's typed extensions are scalar; unknown
`OpaqueExtension` buffers are the heap owner. `Extensions::push()` enforces
`R_j + Q_j <= MAX_OPAQUE_BYTES` for each child and at most
`MAX_EXTENSIONS` children; use those as rejection limits while still charging
the measured `R` and `Q` sums.

The direct settings model is mostly scalar. Its additional heap owners are

```text
G_settings_model = capacity(compatibility_options)
                       * size_of::<CompatibilityOption>()
                 + capacity(compatibility_settings)
                       * size_of::<CompatibilitySetting>()
                 + capacity(smart_tag_types) * size_of::<SmartTagType>()
                 + C_compat_strings + C_smart_tag_strings
                 + C_attached_relationship_id
                 + C_attached_target_uri       # only if the full helper copies it
```

`C_compat_strings` is the counted bytes of each `compatSetting` name, URI,
and value. There is no individual length cap in `CompatibilitySetting`, so
the aggregate must be charged from the processed source or given an explicit
owner-local limit. Each smart-tag declaration is bounded to 2,083 Unicode
characters for its namespace URI and URL and 255 for its name; using four
bytes per scalar gives a per-item UTF-8 ceiling of `8,332 + 8,332 + 1,020`
bytes, with `C_smart_tag_strings` still preferably counted from the parsed
attributes. Smart-tag count has no dedicated model cap, so
`smart_tag_types.len()` is bounded only by the counted settings elements and
its vector capacity must be included. The attached-template target is capped
at 32,768 bytes. The current helper charges
`C_attached_relationship_id` as `2 * T + 8`, covering the caller token bound,
the `String` capacity, and the parser's temporary allowance without changing
the full grammar.

The full parser's peak is therefore the maximum of its non-overlapping phases,
plus the borrowed source and the separately accounted MCE output:

```text
P_full = source/cache
       + MCE_output_and_MCE_state
       + max(
           G_tree + G_mm_model + tree_parser_scratch,
           G_ext_model + extension_parser_scratch,
           G_ext_model + G_mm_model + G_settings_model
               + direct_parser_scratch
         )
```

`tree_parser_scratch` and `direct_parser_scratch` each include one guarded
event of at most `T`, the resolver/namespace state for `D`, and the per-event
resolver clone; extension scratch includes the `R_j/Q_j` overlap above and
its resolver clone. The root/source relationship map stays borrowed. This
expression is source-sized where the code permits it, and it exposes the
non-source-sized owners (`U` namespace clones, `Q` self-contained
declarations, and resolver clone capacity) instead of hiding them in an
unexplained input multiplier. If those measured sums cannot fit the residual
workspace, reject before the first owning parse.

### 8. Owner terms now encoded by the workspace checkpoint

The private
[`settings_workspace.rs`](../../../../crates/litchi-docx/src/source_backed/tail_append/settings_workspace.rs)
checkpoint now has a checked `model_memory_requirement` expression for the
full borrowed model phases. It charges the following owners from the guarded
facts rather than using the former `3 * source` settings shortcut:

* the complete mail-merge `Node`/`Attribute` tree, with geometric child,
  parser-stack, and attribute capacities; `Attribute` includes an explicit
  enum-discriminant word in addition to its three `String` payloads;
* cloned decoded semantic strings with geometric capacity and small-string
  floors, fixed `MailMergeSettings`/`DataSourceObject` storage,
  relationship-ID ceilings, and the two-slot geometric `field_maps` capacity;
* opaque extension source ranges plus self-contained rewrite bytes, their
  overlap and binding/declaration scratch, the `Extension` vector's geometric
  floor, parser token/opened-name/index and resolver storage, and the true
  `MAX_EXTENSIONS * MAX_OPAQUE_BYTES` aggregate output ceiling; and
* direct `DocumentSettings` storage, compatibility/smart-tag vector floors,
  decoded strings, the token-sized attached-template relationship ID, the
  32 KiB attached-template target, and another parser token/resolver term.

This is a source-derived requested-storage envelope with checked arithmetic,
and the current `settings_workspace_requirement` call already includes it.
The integration retains a per-event namespace-declaration maximum for MCE raw
attributes and recounts the actual MCE output before the model lease. The
post-MCE guard's immutable slice and explicit 256-declaration cap are the
input invariant for the owning model readers; their per-event resolver clones
are included in the four-layer binding-capacity term. The attached-template
relationship ID uses the caller token bound with a geometric allowance, so it
does not need a new parser-specific semantic limit.

The following list records the pre-checkpoint gaps. It is historical context,
not a statement that the current path still uses the old projection. The
current helper now closes those owner gaps; the boundary-test evidence is
recorded above and below.

Before the checkpoint, the implementation's mail-merge term was a
conservative geometric reservation for the `Node` and `Attribute` buffers.
The missing model and extension owners were:

* `Extensions::parse` retains each unknown range and may allocate a second
  `R_j + Q_j` buffer while the copied `R_j` is still live. The current
  `settings_model = 3 * source + ...` has no measured `Q` or per-child overlap
  term. Add `unknown_raw_bytes`, `unknown_added_namespace_bytes`, and the
  largest active-binding/declared-prefix scratch, or reject when those values
  cannot be counted. The per-child `MAX_OPAQUE_BYTES` and
  `MAX_EXTENSIONS` checks remain useful local ceilings.
* The extension `values` vector grows through `reserve_one`; its geometric
  capacity and `Extension` slots are not represented. Add the counted typed
  and unknown extension count with the actual `Vec` capacity, or reserve the
  fixed `MAX_EXTENSIONS * size_of::<Extension>()` ceiling.
* The full mail-merge model clones selected `Node` attribute values while the
  tree is still live. Add `Vsel`, the model's fixed `MailMergeSettings` and
  `DataSourceObject` storage, and `2 * field_map_count * size_of::<FieldMap>()`
  for the current geometric `field_maps` growth. The existing `semantic` term
  accounts for the tree strings, not this second model copy or its vector
  slots.
* The direct settings model grows compatibility-option,
  compatibility-setting, and smart-tag vectors independently. Add their
  measured counts/capacities and fixed item sizes; `nodes * 3 *
  size_of::<String>()` is not a proof for a `CompatibilitySetting` or
  `SmartTagType` vector capacity. Compatibility-setting strings have no
  individual model limit, so charge their aggregate decoded bytes from the
  processed XML. Smart-tag count has no dedicated cap beyond settings nodes.
* Attached-template validation now avoids the relationship-match `Vec`, but
  the full model still copies the external target URI into
  `AttachedTemplate::target_uri`; add the 32 KiB target ceiling and a finite
  relationship-ID ceiling. The current document parser only rejects an empty
  ID, so a token-sized ID can otherwise escape the model term.
* The model readers create fresh `NsReader` resolvers after the guard. Their
  requested storage is covered only if the guard's expanded namespace-byte and
  binding facts are applied to each reader, including the per-event resolver
  clone. The root probe and guard now set the declaration cap; the owning
  readers consume only the immutable, already-guarded slice.
* The MCE raw-attribute vector retains namespace declarations too, while
  `max_attributes_per_event` currently counts only non-`xmlns` attributes.
  Track the maximum namespace declarations on one event and add those tuple
  slots to the MCE raw-attribute term; the aggregate declaration count alone
  does not bound one event's allocation.

The checkpoint encodes the first four items as finite requested-storage terms
tied to counted post-MCE facts and standard geometric vector capacity. The
MCE output is now rescanned before model admission, so source facts do not
silently stand in for a different processed input. The current source-bound
owner proof has no identified missing term. Root's final all-feature and
focused-tail validation supplies the boundary evidence, without an RSS
measurement or an unexplained `N * input` multiplier.

## Acceptance evidence

The implementation and tests show:

1. the implemented settings-capability MCE pass keeps caller-derived finite
   depth, namespace/directive/choice policy and an allocation/token guard;
2. the post-MCE guard's facts are applied to the full model phases, including
   resolver/event storage, per-event resolver clones, geometric string
   capacity, and the additive extension/model terms listed above;
3. the root probe keeps one sequential lease with no nested double reservation,
   while retaining the effective package context for default options;
4. duplicate security flags, malformed/duplicate/out-of-order mail-merge
   children, extension namespace spoofing, selected MCE branches, and missing
   or wrong-mode relationships all refuse before publication; and
5. exact-boundary managed budget tests plus unmanaged finite-limit tests that
   exercise a long settings token, deep MCE input, many namespace declarations,
   and a large relationship map. The tests assert typed refusal or allocation
   failure and no output, rather than only checking a successful small
   settings fixture; and
6. the borrowed full-model helper receives the processed `X`, `T`, `D`, and
   event/attribute counts, checks the `P_full` owner expression before any
   owning parse, and either exact-reserves those capacities or charges each
   fallible growth against the same lease. A static extension/node cap by
   itself is not evidence that the caller's workspace is sufficient; the
   attached-template relationship ID must also remain within its charged
   token-bound term.

The focused tail coverage is recorded in
[`source_backed_tail_append.rs`](../../../../crates/litchi-docx/tests/source_backed_tail_append.rs);
the reviewed ownership arithmetic is in
[`settings_workspace.rs`](../../../../crates/litchi-docx/src/source_backed/tail_append/settings_workspace.rs).
Root resolver/output ownership, full-parser owner accounting, post-MCE fact
dominance, cancellation boundary documentation, and relationship validation
are complete for this review; no concrete missing owner remains identified.
