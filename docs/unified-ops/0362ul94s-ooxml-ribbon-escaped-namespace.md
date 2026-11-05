# 0362ul94s — Ribbon escaped-namespace source publication repair

This source-only repair starts from `3774fe4374dfb66e74b98b1ae742672ed6cad37f`
(tree `3832ac43f61d2ddb1e09ba57f6faca142b2f263e`) on the fresh branch
`uo2/ooxml-ribbon-94s`. It addresses the one Ribbon failure in the actual
runtime capture; it makes no acceptance claim and does not rerun Cargo. Only
redaction audit 22 remains deferred.

The captured `ribbon::tests::xml_declarations_characters_namespaces_and_inert_instructions_are_strict`
test supplied a valid default namespace URI containing the XML character
reference `&#x69;`. Ribbon's semantic reader correctly resolves that URI and
the test requires the original bytes to remain unchanged. Replay then sent the
new XML part through `OpcPackage::try_add_owned_xml_part_bytes`, whose stricter
source-splice proof intentionally requires literal namespace bindings. The
result was the capability error:

```text
Opc(SourceBackedOverlayUnavailable { reason: "source XML requires literal namespace bindings" })
```

The fix keeps that source-splice rule unchanged. `OpcPackage` now has explicit,
limit-taking authored whole-XML insertion and replacement primitives. They
check the recipient `SourceLimits`, run the bounded `xml_minifier` source
publication audit and bounded OPC XML structural check before retaining the
payload, and keep the exact shared allocation in `source_xml_parts` for
source-preserving publication. Ribbon's graph replay uses these primitives for
new and changed opaque XML members;
Ribbon's own semantic validator remains responsible for the Custom UI grammar.
No bytes are normalized or copied, and the existing owned/source-splice APIs
are not weakened or aliased.

The regression extends the failing test to replace the escaped-namespace part
as well as add it, asserting exact byte fidelity after each load. Invalid XML,
namespace, character, declaration, and inert-instruction cases retain their
existing refusal assertions.

Runtime evidence is retained outside Git:

| Artifact | Bytes | SHA-256 |
| --- | ---: | --- |
| `common-tests/0000.log` | 222,716 | `668db8f2afc172f4abacf8ce06d36934cc1ca6a89d3ee7c404cc805abe5a8153` |

The capture root is
`(deleted-coord-artifact)`.
The capture reported 511 passing, 17 failing, and the named Ribbon failure;
the other failures remain separate work. No Cargo, compiler, build, rustdoc,
native, or performance command was run for this repair.
