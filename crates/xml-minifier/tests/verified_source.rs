//! `VerifiedSource` exists only for bytes that passed `verify_source`, and
//! covers exactly the audited allocation: the same address and length, under
//! the same limits (change 0754).
#![allow(
    clippy::unwrap_used,
    clippy::expect_used,
    reason = "each test states one fixed expected verdict"
)]

use std::sync::Arc;

use xml_minifier::audit::{
    Limits, ReplacementError, ReplacementProof, Resource, VerifiedSource, verify_source,
    verify_source_replacement,
};

const VALID: &str = "<?xml version=\"1.0\"?>\n<w:document xmlns:w=\"urn:w\">\n  <w:body><w:p><w:r><w:t>one</w:t></w:r></w:p><w:p><w:r><w:t>two</w:t></w:r></w:p></w:body>\n</w:document>";

fn shared(text: &str) -> Arc<Vec<u8>> {
    Arc::new(text.as_bytes().to_vec())
}

#[test]
fn a_proof_is_issued_only_for_bytes_the_source_audit_accepts() {
    let limits = Limits::default();
    let proof = VerifiedSource::verify(shared(VALID), limits).unwrap();
    assert_eq!(proof.bytes().as_slice(), VALID.as_bytes());
    assert_eq!(proof.limits(), limits);

    // Every refusal is the `verify_source` error itself.
    for invalid in [
        "<w:p>",
        "<x:p/>",
        "<!DOCTYPE d><d/>",
        "<a/><b/>",
        "<a b=\"<\"/>",
    ] {
        let error = VerifiedSource::verify(shared(invalid), limits).unwrap_err();
        assert_eq!(
            error,
            verify_source(invalid.as_bytes(), limits).unwrap_err(),
            "{invalid}"
        );
    }
}

#[test]
fn a_proof_covers_only_the_audited_allocation_under_the_audited_limits() {
    let limits = Limits::default();
    let bytes = shared(VALID);
    let proof = VerifiedSource::verify(Arc::clone(&bytes), limits).unwrap();

    // The audited slice, through the proof, a clone of it, or the caller's
    // own reference to the same allocation.
    assert!(proof.covers(proof.bytes().as_slice(), limits));
    assert!(proof.clone().covers(bytes.as_slice(), limits));

    // Equal bytes in another allocation are a copy, not the audited bytes.
    let copy = bytes.as_slice().to_vec();
    assert_eq!(copy.as_slice(), bytes.as_slice());
    assert!(!proof.covers(copy.as_slice(), limits));

    // A prefix or suffix of the audited bytes is not the audited document.
    assert!(!proof.covers(&bytes[..bytes.len() - 1], limits));
    assert!(!proof.covers(&bytes[1..], limits));
    assert!(!proof.covers(&[], limits));

    // Other limits are another audit.
    let other = limits.narrow(Resource::Depth, 64);
    assert_ne!(other, limits);
    assert!(!proof.covers(bytes.as_slice(), other));
}

#[test]
fn a_modified_payload_is_a_new_allocation_that_no_proof_covers() {
    let limits = Limits::default();
    let bytes = shared(VALID);
    let proof = VerifiedSource::verify(Arc::clone(&bytes), limits).unwrap();

    // The proof keeps the allocation shared, so the only way to change the
    // bytes in safe code is to copy them out first.
    let mut modified = bytes;
    Arc::make_mut(&mut modified)[1] = b'!';
    assert_ne!(modified.as_slice(), proof.bytes().as_slice());
    assert!(!proof.covers(modified.as_slice(), limits));
    assert!(proof.covers(proof.bytes().as_slice(), limits));
    assert_eq!(proof.bytes().as_slice(), VALID.as_bytes());
}

#[test]
fn a_replacement_proof_has_the_pair_audits_verdict_and_order() {
    let limits = Limits::default();
    let replacement_text = VALID.replace("two", "zwei");

    // Both sides pass: the proof covers the replacement allocation.
    let replacement = shared(&replacement_text);
    let (proof, how) =
        VerifiedSource::verify_replacement(VALID.as_bytes(), Arc::clone(&replacement), limits)
            .unwrap();
    assert!(proof.covers(replacement.as_slice(), limits));
    assert!(matches!(how, ReplacementProof::Window { .. }), "{how:?}");
    assert_eq!(
        how,
        verify_source_replacement(VALID.as_bytes(), &replacement, limits).unwrap()
    );

    // A failing original is reported as the original's error, before the
    // replacement is examined.
    let error =
        VerifiedSource::verify_replacement(b"<w:document>", shared(&replacement_text), limits)
            .unwrap_err();
    assert_eq!(
        error,
        ReplacementError::Original(verify_source(b"<w:document>", limits).unwrap_err())
    );

    // A failing replacement is reported as the replacement's error, and no
    // proof exists.
    let broken = VALID.replace("<w:t>two</w:t>", "<w:t>two</w:q>");
    let error =
        VerifiedSource::verify_replacement(VALID.as_bytes(), shared(&broken), limits).unwrap_err();
    assert_eq!(
        error,
        ReplacementError::Replacement(verify_source(broken.as_bytes(), limits).unwrap_err())
    );
    let undeclared = VALID.replace("<w:t>two</w:t>", "<w:t>two</w:t><x:y/>");
    assert!(matches!(
        VerifiedSource::verify_replacement(VALID.as_bytes(), shared(&undeclared), limits),
        Err(ReplacementError::Replacement(_))
    ));
}

#[test]
fn a_proof_never_prints_its_payload() {
    let proof = VerifiedSource::verify(shared(VALID), Limits::default()).unwrap();
    let printed = format!("{proof:?}");
    assert!(!printed.contains("document"), "{printed}");
    assert!(printed.contains(&VALID.len().to_string()), "{printed}");
}
