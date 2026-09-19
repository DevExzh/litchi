# Change 0692 code review

Review scope: the capture-local PPTX slide root/name projection in
`crates/litchi-pptx/src/parts/slide.rs`,
`crates/litchi-pptx/src/presentation/package.rs`,
`crates/litchi-pptx/src/presentation/model.rs`, and
`crates/litchi-pptx/src/opened/model.rs`, compared with baseline commit
`e9bd3360c`. This review follows the accepted preservation, validation, and
bounded-resource constraints in ADRs 0003, 0005, and 0006.

## Findings

The production ordering is consistent with the pre-change capture. The main
part's slide-reference parser still runs before the borrowed presentation
view. `capture_slides` uses the same unvalidated catalog and per-reference
relationship/content-type checks as the old `Presentation::slides` path; the
opened identity loop still checks relationship type, target equality, and
duplicate identities before consuming a slide name. Root failures remain
immediate. Only the first name failure is deferred, so a later root or
relationship failure discovered while the contextual vector is built still
wins, while an earlier name failure wins over later identity or name work.

`SlidePart::from_part_with_name` keeps the existing public `from_part` and
`name` behavior as wrappers over the factored byte-slice readers. It validates
the content type, processes MCE once, validates the root, projects the name,
drops the processed `Cow`, and only then allocates the part-name fallback.
Processed XML is not retained in a public or snapshot field.

The scratch vector reserves exactly the reference count with
`try_reserve_exact`; the caller checks `limits.max_parts` before entering the
helper. The first name error is the only retained error, and later slides skip
name projection after that point. These choices preserve typed failures and
avoid an unbounded error/result vector.

The accepted transient-memory tradeoff is real: each successful name and its
`CaptureSlide` entry remains alive while later slide roots are validated. The
name allocation is moved into the final snapshot vector, so it is not copied,
but it can overlap later processed XML and adds one scratch entry per slide.
This is bounded by the existing slide/part limits and should remain called out
when reporting memory results.

## ABI size check

The source layout is unambiguous: `SlidePart<'a>` contains one `&'a dyn Part`
(a two-word fat pointer), `Slide<'a>` contains an `&OpcPackage` plus that
`SlidePart`, and `Option<String>` has the same size as `String` on the target
ABI. A standalone `rustc 1.95.0 (59807616e 2026-04-14)` probe mirrored those
fields without Cargo or workspace dependencies. The exact standalone source
was:

```text
use std::mem::size_of;
trait Part {}
struct OpcPackage;
#[derive(Clone, Copy)]
struct SlidePart<'a> { part: &'a dyn Part }
struct Slide<'a> { package: &'a OpcPackage, part: SlidePart<'a> }
struct CaptureSlide<'a> { slide: Slide<'a>, name: Option<String> }
fn main() {
    println!("usize={} SlidePart={} Slide={} OptionString={} CaptureSlide={}",
        size_of::<usize>(), size_of::<SlidePart<'static>>(),
        size_of::<Slide<'static>>(), size_of::<Option<String>>(),
        size_of::<CaptureSlide<'static>>());
}
```

It was compiled from standard input with `rustc --edition=2021 -O` and ran as
the temporary `/tmp/litchi-0692-abi` executable. Its output was:

```text
usize=8 SlidePart=16 Slide=24 OptionString=24 CaptureSlide=48
```

Thus the capture scratch entry is 48 bytes on the measured 64-bit target,
versus 24 bytes for the old `Slide` entry, before the separately allocated
name bytes and `Vec` capacity are counted. This is a layout observation, not a
whole-operation peak-memory measurement.

## Required focused test coverage

The frozen test diff now directly exercises the private projection and compares
successful output and failures with the old two-phase sequence. It covers
Transitional and Strict roots, explicit/empty/fallback names, MCE input, a
malformed tail, deferred first-name failure, a later bad root, catalog-level
duplicate IDs, missing relationship after a prior name error, duplicate target
catalog errors before names, and notes validation after slide processing. The
helper functions are all used by concrete tests. Error comparisons now check
both the `Error` discriminant and `Debug` value against the legacy oracle,
alongside focused display fragments.

The second identity loop's duplicate-part/target guard cannot be reached by an
ordinary immutable OPC graph: `PresentationPart::slide_references` already
rejects duplicate IDs, relationship IDs, and case-folded target names before
contextual capture. The new duplicate-target cases document and test that
earlier precedence with both an earlier and a later malformed name. The
identity-loop guard remains appropriate defensive protection for a foreign or
otherwise changing `Part` implementation, but no such implementation is
constructed by these package fixtures.

Review result: no production correctness blocker found in the frozen diff.
