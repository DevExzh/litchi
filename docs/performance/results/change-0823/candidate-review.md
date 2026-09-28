# 0823 candidate review

Reviewed at the frozen base `b76786208d04310b4a033b39d70de439a64fcf69`.
This is a read-only source review; no Cargo, compiler, formatter, test, or
workload command was run.

## Source identity and scope

The archived `candidate/before/reader.rs` is byte-identical to the checked-out
HEAD file and has SHA-256
`58e88e86ffc9d4053b694d7928377a05beacbfcb47368cefc9d2fc229b1af26e`.
The archived after file has SHA-256
`a08a48f25c9e2e5f367bd67c579614f1c0c1311784ec81dd30d201b69c3db7f2` and
the patch has SHA-256
`e4a97bd62ec14043922ba0babeb08bcf00a0e6ce110205d4ff67d7d14cd347c3`.
The production diff is confined to `shape/reader.rs`: `Scan` uses a plain
`Reader` and one `NamespaceResolver`; the old `NsReader` loop is test-only.
No public API, dependency, MCE path, record model, element helper, limit, or
source-span contract changes.

## Runtime loop review

The candidate reproduces quick-xml 0.41 `NsReader::read_event` ordering:

1. capture `start` and `decoder`;
2. apply a deferred Empty/End `resolver.pop()` immediately before the next
   underlying `Reader::read_event()`;
3. push Start/Empty declarations with
   `NamespaceError -> quick_xml::Error` conversion, and mark Empty/End for the
   next pop;
4. compute `end` only after the push;
5. resolve Start/Empty/End with `resolver.resolve_element(...)` while the
   event's scope is active;
6. run the existing scanner arm unchanged.

This preserves namespace-error precedence over `end` position and node/depth
checks, keeps an End event in its parent scope, and restores an Empty/End
scope before the following read. Reader defaults and parser errors remain the
same. `position` was correctly specialized for `Reader<&[u8]>`; the separate
test-only `ns_position` preserves the old oracle's offset path. Other PPTX
modules retain their own `NsReader` imports.

The retained `scan_with_nsreader_oracle` is the pre-change scanner body,
including the same handlers, error strings, limit checks, and span accounting.
The parity helper compares complete private records and retained text on
success, and typed error debug output on refusal. This is a suitable oracle
structure because the only production seam changed is event namespace
transport.

## Oracle coverage required before adoption

The refreshed archived tests now positively assert the rich fixture's DML text,
shape names, IDs, and bounds; cover default, Strict, and unknown roots; compare
CDATA, entities, duplicate attributes, malformed XML, and explicit DTD/PI
refusals; exercise reserved `xml`/`xmlns` declarations on both Start and Empty
events; and assert both the 256 accepted and 257 refused declaration cases for
Start and Empty forms. The BOM and node/depth/shape/text boundary parity cases
remain present. The final `borrowed_resolver_restores_shadowed_default_drawingml_scope`
case adds a positive `ABCD` text assertion: an Empty `t` with a shadowing default
namespace is followed by restored default-DML text, and a nested Start rebind is
also restored.

The existing public Scene/MCE tests remain necessary because the private
scanner oracle starts after MCE preprocessing. These additions are test-only
and stay within the single production-file allowlist.

## Disposition

Runtime ordering, old-oracle identity, source identity, and the refreshed
positive/error coverage are **accepted**. No semantic blocker was found. No
performance conclusion follows from this review; the fresh 0823 workflow,
real-file, allocation, RSS, and repaired probe-quality gates remain required.
