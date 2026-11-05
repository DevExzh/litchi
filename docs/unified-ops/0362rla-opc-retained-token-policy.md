# 0362rla — retain OPC source policy in XML tokens

Status: source repair complete; focused validation and native performance
qualification remain pending.

`OwnedXmlPart` now retains the `SourceLimits` profile that admitted its bytes.
XML attribute, element and child-sequence edits validate authored fragments and
the final publication against that retained profile before allocating output.
The policy is metadata, so equal-content tokens captured under different
profiles remain content-equal; a recipient with a tighter profile re-admits the
bytes before installing provenance.

Content-types and relationships owners, structural edit plans, fragment plans,
and metadata fragment materialization intersect requested limits with the
token's retained profile and the package's retained ingress profile. Reparsed
bindings use that effective profile, and source-preserving no-ops carry the
narrowed token policy forward. Authored content-type overrides and relationship
additions now perform output admission against the retained part, entry and
family ceilings before their output buffers are allocated. Every changed XML
token is captured with the effective limits rather than the default profile.
Direct metadata replacement APIs apply the same owner intersection, including
source-byte relationship replacement. Batch source-token publication re-admits
both replacement XML members and the content-type binding under the recipient
before validating graph changes or reserving package state; a producer token's
retained binding is not used as a recipient policy proof.

The source relationship edit plan now also reparses an unchanged/no-op member
when `materialize` receives a narrower profile. Generic XML preflight alone
does not check typed relationship targets or per-member relationship count, so
this second admission closes the broad-plan/tight-materialize path. The
parallel content-types no-op path already reparses `ContentTypeMap` under the
narrowed profile; the regression keeps that behavior covered by review.

The focused regressions cover:

* custom-limited XML tokens rejecting an oversized element edit without
  changing the source token;
* equal-content replay from a default-profile producer into a tighter recipient,
  including a rejection for an actual oversized replacement;
* custom-limited content-type and relationship tokens refusing authored output
  that exceeds their admitted XML budget.
* default capture and plan wrappers inheriting an explicitly admitted owner
  policy; and
* a broad content-types/relationships batch token being refused atomically by
  a tighter recipient before its publication changes; and
* a broad no-op relationship plan refusing tight materialization when its
  target-byte or relationship-count ceiling is smaller than the retained
  source.

No public compatibility alias or new default-limit entry point was added. The
existing default wrappers are fresh-package/test conveniences; opened source
tokens retain their actual admission policy through every edit and replay path.

Validation is intentionally recorded after the bounded focused test/check run;
no native performance result is claimed by this source repair.

The first combined runtime check exposed a fixture issue rather than a
recipient-policy bypass: the broad-token test built a target of 4,117 bytes
from a producer opened under the default 4,096-byte target ceiling. Its
producer now opens under an explicitly broader 8,192-byte target policy, while
the recipient remains at its tighter admitted XML budget and the atomic
refusal assertion is unchanged. The fixed-owned-XML regression likewise now
raises the authored package's retained owner policy before checking the
independent 32 MiB `OwnedXmlPart` ceiling; the caller argument alone cannot
widen that owner policy. These corrections preserve both production
intersection rules and the intended refusal classifications.
