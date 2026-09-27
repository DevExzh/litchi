//! Bounded, inert `PresentationML` notes-slide and notes-master package graphs.

mod codec;
mod model;
mod package;
mod transaction;
mod validation;

pub use codec::{master_xml, write_text, write_text_with};
pub(crate) use codec::{rewrite_text, root_conformance_from_processed};
pub use model::{Conformance, Graph, Link, Master, Slide, Theme};
pub(crate) use package::{
    SlideRootMemo, SlideRootProof, apply_commit, apply_patch, clear_checked, load, load_snapshot,
    load_snapshot_with_slide_root_proofs, remove_checked,
};
#[cfg(test)]
pub(crate) use package::{clear, put, remove, with_refused_slide_root_reservation};
pub use transaction::{Commit, Patch, Revision, Snapshot, Transaction};

pub(crate) const P: &str = "http://schemas.openxmlformats.org/presentationml/2006/main";
pub(crate) const PS: &str = "http://purl.oclc.org/ooxml/presentationml/main";
pub(crate) const A: &str = "http://schemas.openxmlformats.org/drawingml/2006/main";
pub(crate) const AS: &str = "http://purl.oclc.org/ooxml/drawingml/main";
pub(crate) const R: &str = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
pub(crate) const RS: &str = "http://purl.oclc.org/ooxml/officeDocument/relationships";
pub(crate) const SLIDE_CT: &str =
    "application/vnd.openxmlformats-officedocument.presentationml.slide+xml";
pub(crate) const THEME_CT: &str = "application/vnd.openxmlformats-officedocument.theme+xml";
pub(crate) const MAX_PRESENTATION_XML: usize = 32 * 1024 * 1024;
pub(crate) const MAX_SLIDE_XML: usize = 16 * 1024 * 1024;
pub(crate) const MAX_NOTES_XML: usize = 8 * 1024 * 1024;
pub(crate) const MAX_MASTER_XML: usize = 16 * 1024 * 1024;
pub(crate) const MAX_THEME_XML: usize = 16 * 1024 * 1024;
pub(crate) const MAX_TOTAL_BYTES: usize = 64 * 1024 * 1024;
pub(crate) const MAX_NOTES_SLIDES: usize = 4096;
pub(crate) const MAX_OWNED_PARTS: usize = 4_096;
pub(crate) const MAX_NODES: usize = 100_000;
pub(crate) const MAX_DEPTH: usize = 128;
pub(crate) const MAX_ATTRIBUTES: usize = 500_000;
pub(crate) const MAX_ATTRIBUTE_BYTES: usize = 8 * 1024 * 1024;

pub(crate) fn checked_add(left: usize, right: usize, label: &str) -> crate::Result<usize> {
    left.checked_add(right)
        .ok_or_else(|| invalid(format!("PPTX notes {label} overflow")))
}

pub(crate) fn own_blob(blob: &[u8], resource: &'static str) -> crate::Result<Vec<u8>> {
    let mut owned = Vec::new();
    owned
        .try_reserve_exact(blob.len())
        .map_err(|source| allocation(resource, source))?;
    owned.extend_from_slice(blob);
    Ok(owned)
}

pub(crate) fn xml_error(error: impl std::fmt::Display) -> crate::Error {
    crate::Error::Xml(error.to_string())
}

pub(crate) fn invalid(message: impl Into<String>) -> crate::Error {
    crate::Error::Invalid(message.into())
}

pub(crate) fn allocation(
    resource: &'static str,
    source: std::collections::TryReserveError,
) -> crate::Error {
    crate::Error::Allocation { resource, source }
}

pub(crate) fn limit(resource: &'static str, limit: usize) -> crate::Error {
    crate::Error::Limit { resource, limit }
}

/// The namespace a resolved XML name is bound to as a string.
///
/// Exact known OOXML URIs reuse static constants; other bound values borrow
/// from the resolver. An unbound name is the empty namespace; an undeclared
/// prefix is refused.
pub(crate) fn resolved<'a>(value: quick_xml::name::ResolveResult<'a>) -> crate::Result<&'a str> {
    match value {
        quick_xml::name::ResolveResult::Bound(quick_xml::name::Namespace(value)) => {
            if let Some(known) = known_namespace(value) {
                return Ok(known);
            }
            std::str::from_utf8(value).map_err(xml_error)
        },
        quick_xml::name::ResolveResult::Unbound => Ok(""),
        quick_xml::name::ResolveResult::Unknown(prefix) => Err(invalid(format!(
            "unbound XML prefix '{}'",
            String::from_utf8_lossy(prefix.as_ref())
        ))),
    }
}

/// Returns a static known PresentationML, DrawingML, or relationship namespace
/// only for an exact byte match. Distinct constant lengths avoid scanning
/// unrelated namespace values.
#[inline]
fn known_namespace(value: &[u8]) -> Option<&'static str> {
    match value.len() {
        length if length == P.len() && value == P.as_bytes() => Some(P),
        length if length == PS.len() && value == PS.as_bytes() => Some(PS),
        length if length == A.len() && value == A.as_bytes() => Some(A),
        length if length == AS.len() && value == AS.as_bytes() => Some(AS),
        length if length == R.len() && value == R.as_bytes() => Some(R),
        length if length == RS.len() && value == RS.as_bytes() => Some(RS),
        _ => None,
    }
}

#[cfg(test)]
mod tests;

#[cfg(test)]
mod known_uri_tests {
    use super::{A, AS, P, PS, R, RS, resolved};
    use crate::{Error, Result};
    use quick_xml::name::{Namespace, ResolveResult};

    const KNOWN: [(&str, &str); 6] = [(P, P), (PS, PS), (A, A), (AS, AS), (R, R), (RS, RS)];

    #[derive(Debug, PartialEq, Eq)]
    enum Observed {
        Value(String),
        Xml(String),
        Invalid(String),
    }

    fn observe(result: Result<&str>) -> Observed {
        match result {
            Ok(value) => Observed::Value(value.to_owned()),
            Err(Error::Xml(message)) => Observed::Xml(message),
            Err(Error::Invalid(message)) => Observed::Invalid(message),
            Err(error) => panic!("unexpected error variant from namespace resolver: {error:?}"),
        }
    }

    /// The pre-candidate implementation, kept independent of `resolved` so the
    /// focused tests compare both values and refusal messages.
    fn original<'a>(value: ResolveResult<'a>) -> Result<&'a str> {
        match value {
            ResolveResult::Bound(Namespace(value)) => {
                std::str::from_utf8(value).map_err(|error| Error::Xml(error.to_string()))
            },
            ResolveResult::Unbound => Ok(""),
            ResolveResult::Unknown(prefix) => Err(Error::Invalid(format!(
                "unbound XML prefix '{}'",
                String::from_utf8_lossy(prefix.as_ref())
            ))),
        }
    }

    fn assert_bound_parity(value: &[u8]) {
        let actual = observe(resolved(ResolveResult::Bound(Namespace(value))));
        let expected = observe(original(ResolveResult::Bound(Namespace(value))));
        assert_eq!(actual, expected, "namespace bytes: {value:?}");
    }

    #[test]
    fn exact_known_namespaces_do_not_borrow_the_input() {
        for &(namespace, expected) in &KNOWN {
            let input = namespace.as_bytes().to_vec();
            let actual = resolved(ResolveResult::Bound(Namespace(&input)))
                .expect("known namespace should be valid UTF-8");
            assert_eq!(actual, expected);
            // Equal string constants need not have the same address. The known
            // path must return static bytes instead of borrowing this live Vec.
            assert_ne!(actual.as_ptr(), input.as_ptr());
            assert_bound_parity(&input);
        }
    }

    #[test]
    fn every_single_byte_substitution_keeps_fallback_and_xml_errors_exact() {
        for &(namespace, _) in &KNOWN {
            let source = namespace.as_bytes();
            for index in 0..source.len() {
                for replacement in u8::MIN..=u8::MAX {
                    if replacement == source[index] {
                        continue;
                    }
                    let mut substituted = source.to_vec();
                    substituted[index] = replacement;
                    if std::str::from_utf8(&substituted).is_ok() {
                        let actual = resolved(ResolveResult::Bound(Namespace(&substituted)))
                            .expect("a valid substitution should remain accepted");
                        let expected = original(ResolveResult::Bound(Namespace(&substituted)))
                            .expect("the original valid substitution should be accepted");
                        assert_eq!(actual, expected);
                        assert_eq!(actual.as_ptr(), substituted.as_ptr());
                    } else {
                        assert_bound_parity(&substituted);
                    }
                }
            }
        }
    }

    #[test]
    fn prefix_suffix_unicode_and_vendor_namespaces_are_borrowed_fallbacks() {
        for &(namespace, _) in &KNOWN {
            let cases = [
                format!("vendor:{namespace}"),
                format!("{namespace}:suffix"),
                format!("urn:vendor:{namespace}"),
                format!("{namespace}/名前/🚀"),
            ];
            for case in cases {
                let input = case.into_bytes();
                let actual = resolved(ResolveResult::Bound(Namespace(&input)))
                    .expect("vendor namespace should be valid UTF-8");
                assert_eq!(actual, std::str::from_utf8(&input).unwrap());
                assert_eq!(actual.as_ptr(), input.as_ptr());
                assert_bound_parity(&input);
            }
        }
    }

    fn same_length_unicode_namespace(length: usize) -> String {
        let mut value = String::from("urn:vendor:");
        while value.len() + "名".len() <= length {
            value.push('名');
        }
        while value.len() < length {
            value.push('x');
        }
        assert_eq!(value.len(), length);
        value
    }

    #[test]
    fn unicode_vendor_namespaces_can_share_each_known_byte_length() {
        for &(namespace, _) in &KNOWN {
            let text = same_length_unicode_namespace(namespace.len());
            let input = text.as_bytes();
            assert_ne!(input, namespace.as_bytes());
            let actual = resolved(ResolveResult::Bound(Namespace(input)))
                .expect("same-length Unicode vendor namespace should be valid");
            let expected = original(ResolveResult::Bound(Namespace(input)))
                .expect("original same-length Unicode namespace should be valid");
            assert_eq!(actual, expected);
            assert_eq!(actual.as_ptr(), input.as_ptr());
            assert_bound_parity(input);
        }
    }

    #[test]
    fn empty_and_unbound_namespaces_keep_the_existing_success_value() {
        assert_bound_parity(b"");
        assert_eq!(
            observe(resolved(ResolveResult::Unbound)),
            Observed::Value(String::new())
        );
        assert_eq!(
            observe(original(ResolveResult::Unbound)),
            Observed::Value(String::new())
        );
    }

    #[test]
    fn unknown_prefix_keeps_the_exact_invalid_message() {
        let actual = observe(resolved(ResolveResult::Unknown(b"missing".to_vec())));
        let expected = observe(original(ResolveResult::Unknown(b"missing".to_vec())));
        assert_eq!(actual, expected);
        assert_eq!(
            actual,
            Observed::Invalid("unbound XML prefix 'missing'".into())
        );

        let actual = observe(resolved(ResolveResult::Unknown(vec![0xff, b'x'])));
        let expected = observe(original(ResolveResult::Unknown(vec![0xff, b'x'])));
        assert_eq!(actual, expected);
        assert_eq!(actual, Observed::Invalid("unbound XML prefix '�x'".into()));
    }
}
