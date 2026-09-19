# 0697 instruction review — MCE ownership attribution

This is a measurement-only reviewer conclusion for the 0697 packet. It does
not propose a production change or establish a performance claim. The source
revision is HEAD `1bf58ace2c2d69db4ab88e21f62aff5f1ba7cf1e`. The inspected
binary is
`/home/zhuhe/code/litchi-target-0697/release/mce-attribution-0697`, SHA-256
`02312667a7b0931fa37501702315241c1b0c4e6954d87d49ab625098a6dddd75`; the
profile input is `/home/zhuhe/code/litchi-0697-profile/all.data`, SHA-256
`831d7428e147b6c3ae4308947f66e5abcf66976604f0230b00b6fe319302a073`.

The final instruction receipt records successful address-bounded disassembly
and annotation commands for the same binary/profile pair. The bounded symbol
ranges are `start` `0x42a40..0x470e3`, `drop_in_place<Ctx>`
`0x47460..0x474d6`, and `drop_in_place<Inherited>` `0x47670..0x476e6`.
The earlier `instructions-initial/` object-selection output was empty and is
excluded from this conclusion; the final symbol-bound output is the valid
instruction evidence.

## Profile and sample interpretation

The isolated MCE profile reports zero lost samples and these self shares:

| Symbol | Self share |
| --- | ---: |
| `litchi_ooxml_common::mce::codec::start` | 26.52% |
| `drop_in_place<litchi_ooxml_common::mce::codec::Inherited>` | 7.16% |
| `drop_in_place<litchi_ooxml_common::mce::codec::Ctx>` | 4.86% |

The `start` annotation contains 649 samples, the `Inherited` destructor
annotation 175, and the `Ctx` destructor annotation 119. These values use the
isolated packet denominator and are not whole-workflow improvement numbers.
The inclusive `start` callgraph attributes 22.02% to inlined
`NamespaceLayer` `Arc` clone regions: `after` 8.71%, parent `Ctx` cloning
8.37%, temporary inherited namespace observation 2.53%, and temporary
inherited emitted-boundary observation 2.41%. Inclusive regions overlap the
`start` self share and are diagnostic only.

Perf annotation places samples at the branch immediately after several
`lock incq`/`lock decq` instructions. A following-branch sample supports the
nearby source region, but skid means it does not measure the exact preceding
atomic instruction. The counts below are sample counts from this run, not
dynamic execution counts, refcount contention measurements, or per-operation
costs. A zero count does not prove that a path is never executed.

| Operation | Atomic address | Following branch | Samples at branch |
| --- | ---: | ---: | ---: |
| Parent `Ctx` namespace clone | `0x4311a` | `0x4311e` | 205 |
| Parent `Ctx` directive clone | `0x43131` | `0x43135` | 0 |
| Temporary `Inherited.ns` clone | `0x43168` | `0x4316c` | 62 |
| Temporary `Inherited.emitted` clone | `0x432fb` | `0x432ff` | 59 |
| `after` carried emitted-boundary owner, branch path | `0x43814` | `0x43818` | 0 |
| `after` carried/emitted owner, branch path | `0x439a5` | `0x439a9` | 0 |
| `after` ordinary child emitted-boundary owner | `0x45b2e` | `0x45b32` | 213 |
| `after` carried owner, branch path | `0x4689a` | `0x4689e` | 0 |
| Temporary `Inherited.ns` release | `0x4767f` | `0x47683` | 132 |
| Temporary `Inherited.emitted` release | `0x47697` | `0x4769b` | 43 |
| `Ctx` namespace release | `0x4746f` | `0x47473` | 119 |
| `Ctx` directive release | `0x47487` | `0x4748b` | 0 |

The two temporary clone regions account for 121 of the 649 `start` samples;
the two `Inherited` release branches account for all 175 destructor samples;
the namespace release branch accounts for all 119 `Ctx` destructor samples.
This concentrates the narrow ownership hypothesis, while the 213 samples at
`0x45b32` identify a separate child-frame owner that must remain semantically
owned.

## Source-to-instruction mapping

In `crates/litchi-ooxml-common/src/mce/codec.rs`, `Inherited` owns its two
temporary options at lines 441–444. The `start` setup at lines 763–767 first
clones the parent `Ctx` with
`st.last().map_or_else(Ctx::root, |f| f.ctx.clone())`, then clones
`c.ns.head` and `st.last().and_then(|f| f.emitted_ns.clone())` into
`Inherited`. The final assembly maps the parent `Ctx` namespace and directive
owners to `0x4311a` and `0x43131`, and the two temporary `Inherited` owners to
`0x43168` and `0x432fb`.

`Inherited::after` at lines 451–458 is the ownership boundary for the child
`Frame`: it clones `ctx.ns.head` for an emitted element or carries the prior
`self.emitted` owner for a dropped/unemitted element. The inlined owner clones
at `0x43814`, `0x439a5`, `0x45b2e`, and `0x4689a` are therefore not all
removable. In particular, `0x45b2e` is the ordinary-path child emitted-boundary
owner with 213 following-branch samples. `Frame.emitted_ns` at lines 602–610
must continue to own this nearest emitted scope across later children and end
events.

`write_start` consumes the inherited values read-only through `hoists` and
`Namespaces::for_each_hoisted` at lines 1349–1381. The `Ctx` fields at lines
470–475 remain independently owned because the child can install local
namespace state, take or replace directives, and change opacity. The temporary
destructor ranges show the corresponding two `Inherited` releases at
`0x4767f` and `0x47697`; `Ctx` releases its namespace and directive fields at
`0x4746f` and `0x47487`.

## Reviewer conclusion and proof seam

The highest-value semantically narrow next experiment remains a borrowed
`Inherited<'a>` view whose `ns` and `emitted` fields are references to the
parent `Frame` owners. It targets only the two temporary clone/drop pairs at
`0x43168`/`0x432fb` and `0x4767f`/`0x47697`. It should preserve the parent
`Ctx` clone and every `Inherited::after` child owner. The sample evidence is
consistent with this target, but does not predict its speed or allocation
effect.

The borrow source must be the parent frame's fields, not `c.ns.head`:
`c` can be replaced by `with_local` at lines 781–783, while the inherited
scope is the pre-local parent scope. The lifetime seam must be structural:

1. Complete the existing validation and namespace/directive work.
2. Finish `st.last_mut()` AlternateContent bookkeeping at lines 981–1061,
   retaining only the owned `active`/`Mode` results.
3. In a short inner block, obtain `st.last()` and construct the borrowed view
   from `parent.ctx.ns.head.as_ref()` and `parent.emitted_ns.as_ref()`.
4. Use the view for `write_start`, call `after` to produce the owned
   `Frame.emitted_ns`, and construct the owned `Frame`.
5. End that block before `close` at lines 1134–1148 can reserve or push into
   `st`; no `Inherited<'_>` reference, alias, closure capture, or helper return
   may cross `close`.

The opaque early return and both direct AlternateContent child returns need the
same owned-frame-before-`close` boundary. This sequencing preserves Alt
selection/error order while allowing the parent mutation to finish before an
immutable view is borrowed. A larger `Ctx`/`Frame` redesign remains a separate
candidate because it would change owners that survive to siblings, dropped
wrappers, and end events.

Retention requires the existing shared oracle and focused MCE/refusal matrix
to prove exact output bytes, `Cow` ownership, `Report`, typed refusal identity,
QName/directive/error order, and all namespace/depth/directive/choice limits.
It also requires dropped-wrapper and `AlternateContent` cases where the
nearest emitted scope differs from the immediate context, plus declaration
heavy/native controls and frozen post-change instruction evidence. Allocation,
RSS, stack/code size, and end-to-end timing must be measured separately; this
packet establishes none of them.
