#![allow(
    clippy::unwrap_used,
    clippy::expect_used,
    reason = "test assertions panic on failure by design"
)]
//! The signature staleness token must be a function of the signature graph.
//!
//! `content_control::package::signature_token` serializes the reachable
//! signature graph into an `Arc<[u8]>` that `require_current` compares with `!=`
//! whenever an exact no-op package patch is published against a signed package.
//! The part names it walks are sorted, but the relationship records inside the
//! token were emitted straight from `Relationships::iter()` — a `HashMap` walk —
//! both for the root signature relationships and for each signature part's own
//! relationships. A signature origin that owns two signature relationships is
//! the ordinary shape for a twice-signed document, and it made the token differ
//! byte-for-byte between two loads of the same package, so an exact no-op was
//! refused as "package signature topology is stale" about half the time.
//!
//! ADR 0006 requires serialization to be deterministic unless a `Clock`, actor
//! identity, or cryptographic RNG is explicitly supplied; nothing supplies one
//! here. ADR 0003 requires a source-checked patch to conflict on a real overlap.
//! An exact no-op on an untouched package is neither.
//!
//! Each `Relationships::new()` builds a fresh `HashMap`, whose `RandomState` is
//! seeded from a per-thread counter that advances on every construction, so
//! rebuilding the same package in one process varies the visit order exactly as
//! separate processes do.

use litchi_docx::{Error, Package};
use litchi_opc::constants::{content_type as ct, relationship_type as rt};
use litchi_opc::part::BlobPart;
use litchi_opc::{OpcPackage, PackURI};

/// Repetitions per determinism assertion; 2^-128 accidental agreement.
const REPEATS: usize = 128;

const ORIGIN: &str = "/_xmlsignatures/origin.sigs";
const SIGNATURE_RELATIONSHIP: &str =
    "http://schemas.openxmlformats.org/package/2006/relationships/digital-signature/signature";

/// Which signature graph to build. `Package::edit_opc` unsigns its candidate
/// before running the closure, so the whole graph has to be installed in one
/// call; a variant is a different build, never an edit of a built one.
#[derive(Clone, Copy, PartialEq, Eq)]
enum Shape {
    /// One signature: the origin owns one relationship and the package root
    /// owns one signature relationship, so no order exists to get wrong.
    Once,
    /// Two signatures: the origin owns two relationships and the package root
    /// owns three signature relationships. This is the twice-signed shape.
    Twice,
    /// `Twice`, with one signature payload genuinely different.
    TwiceTampered,
}

fn signed_package(shape: Shape) -> Package {
    let mut package = Package::new().unwrap();
    package
        .edit_opc(|opc: &mut OpcPackage| {
            opc.try_add_part(Box::new(BlobPart::new(
                PackURI::new(ORIGIN).unwrap(),
                ct::OPC_DIGITAL_SIGNATURE_ORIGIN.to_owned(),
                b"origin".to_vec(),
            )))?;
            opc.rels_mut().add_relationship(
                rt::DIGITAL_SIGNATURE_ORIGIN.to_owned(),
                "_xmlsignatures/origin.sigs".to_owned(),
                "rIdOrigin".to_owned(),
                false,
            );

            let signatures: &[(&str, &str)] = match shape {
                Shape::Once => &[("sig1.xml", "rIdSignatureAlpha")],
                Shape::Twice | Shape::TwiceTampered => &[
                    ("sig1.xml", "rIdSignatureAlpha"),
                    ("sig2.xml", "rIdSignatureBeta"),
                ],
            };
            for (target, r_id) in signatures {
                let payload = if shape == Shape::TwiceTampered && *target == "sig2.xml" {
                    "<Signature>tampered</Signature>".to_owned()
                } else {
                    format!("<Signature>{target}</Signature>")
                };
                opc.try_add_part(Box::new(BlobPart::new(
                    PackURI::new(format!("/_xmlsignatures/{target}")).unwrap(),
                    ct::OPC_DIGITAL_SIGNATURE_XMLSIGNATURE.to_owned(),
                    payload.into_bytes(),
                )))?;
                opc.get_part_mut(&PackURI::new(ORIGIN).unwrap())?
                    .rels_mut()
                    .add_relationship(
                        SIGNATURE_RELATIONSHIP.to_owned(),
                        (*target).to_owned(),
                        (*r_id).to_owned(),
                        false,
                    );
                if shape != Shape::Once {
                    // Extra root-level signature relationships exercise the
                    // token's root loop as well as its per-part loop.
                    opc.rels_mut().add_relationship(
                        SIGNATURE_RELATIONSHIP.to_owned(),
                        format!("_xmlsignatures/{target}"),
                        format!("{r_id}Root"),
                        false,
                    );
                }
            }
            Ok::<_, Error>(())
        })
        .unwrap();
    package
}

/// Capture an exact no-op against one build of the package and publish it
/// against an independent build of the same package.
fn noop_across_two_builds(shape: Shape) -> litchi_docx::Result<()> {
    let source = signed_package(shape);
    let mut target = signed_package(shape);
    let snapshot = source.content_control_snapshot()?;
    let commit = snapshot.edit()?.commit()?;
    assert!(!commit.changed(), "the transaction stages no edit");
    assert!(commit.patch().is_noop(), "the patch is an exact no-op");
    target.apply_content_controls(&commit)
}

#[test]
fn exact_noop_is_accepted_against_every_build_of_the_same_signed_package() {
    for attempt in 0..REPEATS {
        noop_across_two_builds(Shape::Twice).unwrap_or_else(|error| {
            panic!(
                "attempt {attempt}: an exact no-op was refused against an identical build of the \
                 same signed package: {error}"
            )
        });
    }
}

#[test]
fn single_signature_packages_keep_accepting_their_noop() {
    for attempt in 0..REPEATS {
        noop_across_two_builds(Shape::Once).unwrap_or_else(|error| {
            panic!("attempt {attempt}: single-signature no-op was refused: {error}")
        });
    }
}

#[test]
fn a_genuinely_different_signature_graph_is_still_refused() {
    let source = signed_package(Shape::Twice);
    let mut target = signed_package(Shape::TwiceTampered);
    let snapshot = source.content_control_snapshot().unwrap();
    let commit = snapshot.edit().unwrap().commit().unwrap();
    assert!(
        target.apply_content_controls(&commit).is_err(),
        "a changed signature payload must still be refused"
    );
}
