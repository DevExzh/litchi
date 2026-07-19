//! Inert OpenDocument digital-signature metadata.
//!
//! This module parses the ODF signature containers and XMLDSIG metadata without
//! verifying signatures, resolving external references, or executing macros.

use litchi_core::{Error, Result};
use quick_xml::{
    XmlVersion,
    encoding::Decoder,
    events::{BytesStart, Event},
    name::{Namespace, ResolveResult},
    reader::NsReader,
};

const ODF_SIGNATURE_NAMESPACE: &[u8] =
    b"urn:oasis:names:tc:opendocument:xmlns:digitalsignature:1.0";
const XML_SIGNATURE_NAMESPACE: &[u8] = b"http://www.w3.org/2000/09/xmldsig#";
const DUBLIN_CORE_NAMESPACE: &[u8] = b"http://purl.org/dc/elements/1.1/";
const XADES_NAMESPACE_PREFIX: &[u8] = b"http://uri.etsi.org/01903/";
const MAX_SIGNATURE_XML_BYTES: usize = 16 * 1024 * 1024;
const MAX_XML_DEPTH: usize = 128;
const MAX_XML_EVENTS: usize = 1_000_000;
const MAX_SIGNATURES: usize = 128;
const MAX_REFERENCES: usize = 4096;
const MAX_TRANSFORMS: usize = 32;
const MAX_CERTIFICATES: usize = 64;
const MAX_ATTRIBUTE_BYTES: usize = 64 * 1024;
const MAX_BASE64_BYTES: usize = 8 * 1024 * 1024;

pub(crate) const DOCUMENT_SIGNATURE_PATH: &str = "META-INF/documentsignatures.xml";
pub(crate) const MACRO_SIGNATURE_PATH: &str = "META-INF/macrosignatures.xml";

/// Signature metadata stored in an OpenDocument package.
#[derive(Debug, Clone, Default, PartialEq, Eq)]
pub struct OdfDigitalSignatures {
    pub document_signatures: Vec<OdfDigitalSignature>,
    pub macro_signatures: Vec<OdfDigitalSignature>,
}

impl OdfDigitalSignatures {
    pub fn is_empty(&self) -> bool {
        self.document_signatures.is_empty() && self.macro_signatures.is_empty()
    }

    pub fn len(&self) -> usize {
        self.document_signatures.len() + self.macro_signatures.len()
    }
}

/// One inert XMLDSIG signature record.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct OdfDigitalSignature {
    pub id: Option<String>,
    pub canonicalization_method: String,
    pub signature_method: String,
    pub references: Vec<OdfSignatureReference>,
    /// Whitespace-normalized base64 signature bytes.
    pub signature_value: String,
    /// Whitespace-normalized base64 DER certificates, in document order.
    pub x509_certificates: Vec<String>,
    /// XAdES `SigningTime`, or the legacy Dublin Core signature date.
    pub signing_time: Option<String>,
}

/// One `ds:Reference` from a signature's signed-info block.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct OdfSignatureReference {
    pub uri: String,
    pub type_uri: Option<String>,
    pub transforms: Vec<String>,
    pub digest_method: String,
    /// Whitespace-normalized base64 digest bytes.
    pub digest_value: String,
}

#[derive(Default)]
struct SignatureBuilder {
    depth: usize,
    id: Option<String>,
    signed_info_depth: Option<usize>,
    canonicalization_method: Option<String>,
    signature_method: Option<String>,
    references: Vec<OdfSignatureReference>,
    reference: Option<ReferenceBuilder>,
    signature_value: Option<String>,
    certificates: Vec<String>,
    signing_time: Option<String>,
    legacy_signing_time: Option<String>,
}

struct ReferenceBuilder {
    depth: usize,
    uri: String,
    type_uri: Option<String>,
    transforms: Vec<String>,
    digest_method: Option<String>,
    digest_value: Option<String>,
}

#[derive(Clone, Copy)]
enum TextTargetKind {
    DigestValue,
    SignatureValue,
    Certificate,
    SigningTime,
    LegacySigningTime,
}

struct TextTarget {
    depth: usize,
    kind: TextTargetKind,
    value: String,
}

pub(crate) fn parse_signature_container(xml: &[u8]) -> Result<Vec<OdfDigitalSignature>> {
    if xml.len() > MAX_SIGNATURE_XML_BYTES {
        return invalid("ODF signature XML exceeds 16 MiB");
    }

    let mut reader = NsReader::from_reader(xml);
    reader.config_mut().check_end_names = true;
    reader.config_mut().expand_empty_elements = false;
    let mut buffer = Vec::new();
    let mut depth = 0usize;
    let mut events = 0usize;
    let mut root_seen = false;
    let mut root_closed = false;
    let mut signatures = Vec::new();
    let mut signature: Option<SignatureBuilder> = None;
    let mut text_target: Option<TextTarget> = None;

    loop {
        events = events
            .checked_add(1)
            .ok_or_else(|| format_error("ODF signature XML event count overflow"))?;
        if events > MAX_XML_EVENTS {
            return invalid("ODF signature XML exceeds the event limit");
        }

        let decoder = reader.decoder();
        let (namespace, event) = reader
            .read_resolved_event_into(&mut buffer)
            .map_err(|error| format_error(format!("invalid ODF signature XML: {error}")))?;
        match event {
            Event::Start(ref element) => {
                if root_closed {
                    return invalid("ODF signature XML contains content after its root element");
                }
                if depth >= MAX_XML_DEPTH {
                    return invalid("ODF signature XML exceeds the nesting limit");
                }
                let local = element.local_name();
                let local = local.as_ref();
                if depth == 0 {
                    if root_seen
                        || !namespace_is(&namespace, ODF_SIGNATURE_NAMESPACE)
                        || local != b"document-signatures"
                    {
                        return invalid("invalid ODF digital-signature root element");
                    }
                    root_seen = true;
                } else if depth == 1
                    && namespace_is(&namespace, XML_SIGNATURE_NAMESPACE)
                    && local == b"Signature"
                {
                    if signature.is_some() || signatures.len() >= MAX_SIGNATURES {
                        return invalid("invalid or excessive ODF digital signatures");
                    }
                    signature = Some(SignatureBuilder {
                        depth,
                        id: attribute(decoder, element, b"Id")?,
                        ..SignatureBuilder::default()
                    });
                } else if let Some(current) = signature.as_mut() {
                    process_start(
                        decoder,
                        &namespace,
                        element,
                        depth,
                        current,
                        &mut text_target,
                    )?;
                } else if local == b"Signature" {
                    return invalid("signature vocabulary uses the wrong namespace or location");
                }
                depth += 1;
            },
            Event::Empty(ref element) => {
                if depth == 0 || root_closed {
                    return invalid("empty or trailing ODF signature root content");
                }
                let local_name = element.local_name();
                let local = local_name.as_ref();
                if let Some(current) = signature.as_mut() {
                    process_empty(decoder, &namespace, element, depth, current)?;
                } else if local == b"Signature" {
                    return invalid("signature vocabulary uses the wrong namespace or location");
                }
            },
            Event::Text(ref value) => {
                let decoded = value
                    .decode()
                    .map_err(|error| format_error(format!("invalid signature text: {error}")))?;
                if let Some(target) = text_target.as_mut() {
                    append_bounded(&mut target.value, &decoded, MAX_BASE64_BYTES)?;
                } else if signature.is_none() && !decoded.trim().is_empty() {
                    return invalid("text is not allowed outside an ODF signature");
                }
            },
            Event::CData(ref value) => {
                if let Some(target) = text_target.as_mut() {
                    let decoded = value.decode().map_err(|error| {
                        format_error(format!("invalid signature CDATA: {error}"))
                    })?;
                    append_bounded(&mut target.value, &decoded, MAX_BASE64_BYTES)?;
                } else if !value.is_empty() {
                    return invalid("unexpected CDATA in ODF signature XML");
                }
            },
            Event::End(ref element) => {
                depth = depth
                    .checked_sub(1)
                    .ok_or_else(|| format_error("ODF signature XML depth underflow"))?;
                if text_target
                    .as_ref()
                    .is_some_and(|target| target.depth == depth)
                {
                    let target = text_target.take().expect("target depth checked");
                    finish_text_target(target, signature.as_mut())?;
                }
                if let Some(current) = signature.as_mut() {
                    if current
                        .reference
                        .as_ref()
                        .is_some_and(|value| value.depth == depth)
                    {
                        finish_reference(current)?;
                    }
                    if current.signed_info_depth == Some(depth) {
                        current.signed_info_depth = None;
                    }
                }
                if signature.as_ref().is_some_and(|value| value.depth == depth) {
                    signatures.push(finish_signature(
                        signature.take().expect("signature depth checked"),
                    )?);
                }
                if depth == 0 {
                    let local_name = element.local_name();
                    if local_name.as_ref() != b"document-signatures" {
                        return invalid("ODF signature XML root end tag is invalid");
                    }
                    root_closed = true;
                }
            },
            Event::GeneralRef(_) | Event::DocType(_) | Event::PI(_) => {
                return invalid("entities, DTDs, and processing instructions are prohibited");
            },
            Event::Decl(_) | Event::Comment(_) => {},
            Event::Eof => break,
        }
        buffer.clear();
    }

    if !root_seen || !root_closed || depth != 0 || signature.is_some() || text_target.is_some() {
        return invalid("incomplete ODF signature XML");
    }
    if signatures.is_empty() {
        return invalid("ODF signature container must contain at least one signature");
    }
    Ok(signatures)
}

fn process_start(
    decoder: Decoder,
    namespace: &ResolveResult<'_>,
    element: &BytesStart<'_>,
    depth: usize,
    signature: &mut SignatureBuilder,
    text_target: &mut Option<TextTarget>,
) -> Result<()> {
    let local_name = element.local_name();
    let local = local_name.as_ref();
    if namespace_is(namespace, XML_SIGNATURE_NAMESPACE) {
        match local {
            b"SignedInfo" if signature.signed_info_depth.is_none() => {
                signature.signed_info_depth = Some(depth);
            },
            b"Reference" if signature.signed_info_depth.is_some() => {
                start_reference(decoder, element, depth, signature)?;
            },
            b"DigestValue" if signature.reference.is_some() => {
                start_text(text_target, depth, TextTargetKind::DigestValue)?;
            },
            b"SignatureValue" => {
                start_text(text_target, depth, TextTargetKind::SignatureValue)?;
            },
            b"X509Certificate" => {
                start_text(text_target, depth, TextTargetKind::Certificate)?;
            },
            b"CanonicalizationMethod" | b"SignatureMethod" => {
                process_algorithm_element(decoder, element, local, signature)?;
            },
            b"Transform" | b"DigestMethod" if signature.reference.is_some() => {
                process_algorithm_element(decoder, element, local, signature)?;
            },
            _ => {},
        }
    } else if local == b"Signature" || local == b"SignedInfo" || local == b"Reference" {
        return invalid("XMLDSIG vocabulary uses the wrong namespace");
    } else if namespace_prefix_is(namespace, XADES_NAMESPACE_PREFIX) && local == b"SigningTime" {
        start_text(text_target, depth, TextTargetKind::SigningTime)?;
    } else if namespace_is(namespace, DUBLIN_CORE_NAMESPACE) && local == b"date" {
        start_text(text_target, depth, TextTargetKind::LegacySigningTime)?;
    }
    Ok(())
}

fn process_empty(
    decoder: Decoder,
    namespace: &ResolveResult<'_>,
    element: &BytesStart<'_>,
    depth: usize,
    signature: &mut SignatureBuilder,
) -> Result<()> {
    let local_name = element.local_name();
    let local = local_name.as_ref();
    if namespace_is(namespace, XML_SIGNATURE_NAMESPACE) {
        match local {
            b"CanonicalizationMethod" | b"SignatureMethod" => {
                process_algorithm_element(decoder, element, local, signature)?;
            },
            b"Transform" | b"DigestMethod" if signature.reference.is_some() => {
                process_algorithm_element(decoder, element, local, signature)?;
            },
            b"Reference" if signature.signed_info_depth.is_some() => {
                start_reference(decoder, element, depth, signature)?;
                finish_reference(signature)?;
            },
            b"Signature" | b"SignedInfo" | b"DigestValue" | b"SignatureValue"
            | b"X509Certificate" => {
                return invalid("required XMLDSIG element cannot be empty");
            },
            _ => {},
        }
    } else if matches!(local, b"Signature" | b"SignedInfo" | b"Reference") {
        return invalid("XMLDSIG vocabulary uses the wrong namespace");
    }
    Ok(())
}

fn process_algorithm_element(
    decoder: Decoder,
    element: &BytesStart<'_>,
    local: &[u8],
    signature: &mut SignatureBuilder,
) -> Result<()> {
    let algorithm = required_attribute(decoder, element, b"Algorithm")?;
    match local {
        b"CanonicalizationMethod" if signature.signed_info_depth.is_some() => set_once(
            &mut signature.canonicalization_method,
            algorithm,
            "canonicalization method",
        ),
        b"SignatureMethod" if signature.signed_info_depth.is_some() => set_once(
            &mut signature.signature_method,
            algorithm,
            "signature method",
        ),
        b"Transform" => {
            let reference = signature
                .reference
                .as_mut()
                .ok_or_else(|| format_error("ds:Transform appears outside ds:Reference"))?;
            if reference.transforms.len() >= MAX_TRANSFORMS {
                return invalid("signature reference has too many transforms");
            }
            reference.transforms.push(algorithm);
            Ok(())
        },
        b"DigestMethod" => {
            let reference = signature
                .reference
                .as_mut()
                .ok_or_else(|| format_error("ds:DigestMethod appears outside ds:Reference"))?;
            set_once(&mut reference.digest_method, algorithm, "digest method")
        },
        _ => Ok(()),
    }
}

fn start_reference(
    decoder: Decoder,
    element: &BytesStart<'_>,
    depth: usize,
    signature: &mut SignatureBuilder,
) -> Result<()> {
    if signature.reference.is_some() || signature.references.len() >= MAX_REFERENCES {
        return invalid("nested or excessive signature references");
    }
    signature.reference = Some(ReferenceBuilder {
        depth,
        uri: required_attribute(decoder, element, b"URI")?,
        type_uri: attribute(decoder, element, b"Type")?,
        transforms: Vec::new(),
        digest_method: None,
        digest_value: None,
    });
    Ok(())
}

fn finish_reference(signature: &mut SignatureBuilder) -> Result<()> {
    let reference = signature
        .reference
        .take()
        .ok_or_else(|| format_error("missing signature reference"))?;
    signature.references.push(OdfSignatureReference {
        uri: reference.uri,
        type_uri: reference.type_uri,
        transforms: reference.transforms,
        digest_method: reference
            .digest_method
            .ok_or_else(|| format_error("signature reference has no digest method"))?,
        digest_value: reference
            .digest_value
            .ok_or_else(|| format_error("signature reference has no digest value"))?,
    });
    Ok(())
}

fn finish_signature(value: SignatureBuilder) -> Result<OdfDigitalSignature> {
    if value.signed_info_depth.is_some() || value.reference.is_some() || value.references.is_empty()
    {
        return invalid("signature has incomplete or empty signed information");
    }
    Ok(OdfDigitalSignature {
        id: value.id,
        canonicalization_method: value
            .canonicalization_method
            .ok_or_else(|| format_error("signature has no canonicalization method"))?,
        signature_method: value
            .signature_method
            .ok_or_else(|| format_error("signature has no signature method"))?,
        references: value.references,
        signature_value: value
            .signature_value
            .ok_or_else(|| format_error("signature has no signature value"))?,
        x509_certificates: value.certificates,
        signing_time: value.signing_time.or(value.legacy_signing_time),
    })
}

fn start_text(target: &mut Option<TextTarget>, depth: usize, kind: TextTargetKind) -> Result<()> {
    if target.is_some() {
        return invalid("nested signature text value");
    }
    *target = Some(TextTarget {
        depth,
        kind,
        value: String::new(),
    });
    Ok(())
}

fn finish_text_target(target: TextTarget, signature: Option<&mut SignatureBuilder>) -> Result<()> {
    let signature = signature.ok_or_else(|| format_error("signature text outside a signature"))?;
    match target.kind {
        TextTargetKind::DigestValue => {
            let value = normalize_base64(&target.value, "digest value")?;
            let reference = signature
                .reference
                .as_mut()
                .ok_or_else(|| format_error("digest value outside a reference"))?;
            set_once(&mut reference.digest_value, value, "digest value")
        },
        TextTargetKind::SignatureValue => {
            let value = normalize_base64(&target.value, "signature value")?;
            set_once(&mut signature.signature_value, value, "signature value")
        },
        TextTargetKind::Certificate => {
            if signature.certificates.len() >= MAX_CERTIFICATES {
                return invalid("signature has too many X.509 certificates");
            }
            signature
                .certificates
                .push(normalize_base64(&target.value, "X.509 certificate")?);
            Ok(())
        },
        TextTargetKind::SigningTime => {
            let value = normalized_text(&target.value, "signing time")?;
            set_once(&mut signature.signing_time, value, "signing time")
        },
        TextTargetKind::LegacySigningTime => {
            let value = normalized_text(&target.value, "signature date")?;
            set_once(&mut signature.legacy_signing_time, value, "signature date")
        },
    }
}

fn attribute(decoder: Decoder, element: &BytesStart<'_>, name: &[u8]) -> Result<Option<String>> {
    let mut result = None;
    let mut count = 0usize;
    for value in element.attributes().with_checks(true) {
        count += 1;
        if count > 64 {
            return invalid("signature element has too many attributes");
        }
        let value = value.map_err(|error| format_error(format!("invalid attribute: {error}")))?;
        if value.key.as_ref() == name {
            if result.is_some() {
                return invalid("duplicate signature attribute");
            }
            let decoded = value
                .decoded_and_normalized_value(XmlVersion::Implicit1_0, decoder)
                .map_err(|error| format_error(format!("invalid attribute value: {error}")))?;
            validate_text(&decoded, "signature attribute", MAX_ATTRIBUTE_BYTES, true)?;
            result = Some(decoded.into_owned());
        }
    }
    Ok(result)
}

fn required_attribute(decoder: Decoder, element: &BytesStart<'_>, name: &[u8]) -> Result<String> {
    attribute(decoder, element, name)?.ok_or_else(|| {
        format_error(format!(
            "signature element requires {}",
            String::from_utf8_lossy(name)
        ))
    })
}

fn normalize_base64(value: &str, context: &str) -> Result<String> {
    let normalized: String = value
        .chars()
        .filter(|value| !value.is_whitespace())
        .collect();
    if normalized.is_empty()
        || normalized.len() > MAX_BASE64_BYTES
        || normalized.len() % 4 != 0
        || normalized
            .bytes()
            .any(|value| !(value.is_ascii_alphanumeric() || matches!(value, b'+' | b'/' | b'=')))
        || normalized
            .find('=')
            .is_some_and(|index| index < normalized.len().saturating_sub(2))
    {
        return invalid(format!("invalid base64 {context}"));
    }
    Ok(normalized)
}

fn normalized_text(value: &str, context: &str) -> Result<String> {
    let value = value.trim();
    validate_text(value, context, MAX_ATTRIBUTE_BYTES, false)?;
    Ok(value.to_owned())
}

fn validate_text(value: &str, context: &str, maximum: usize, allow_empty: bool) -> Result<()> {
    if (!allow_empty && value.is_empty())
        || value.len() > maximum
        || value
            .chars()
            .any(|value| matches!(value, '\0'..='\u{8}' | '\u{b}' | '\u{c}' | '\u{e}'..='\u{1f}'))
    {
        return invalid(format!("invalid {context}"));
    }
    Ok(())
}

fn append_bounded(target: &mut String, value: &str, maximum: usize) -> Result<()> {
    if target.len().saturating_add(value.len()) > maximum {
        return invalid("signature text exceeds its size limit");
    }
    target.push_str(value);
    Ok(())
}

fn set_once(target: &mut Option<String>, value: String, context: &str) -> Result<()> {
    if target.replace(value).is_some() {
        return invalid(format!("duplicate {context}"));
    }
    Ok(())
}

fn namespace_is(value: &ResolveResult<'_>, expected: &[u8]) -> bool {
    matches!(value, ResolveResult::Bound(found) if *found == Namespace(expected))
}

fn namespace_prefix_is(value: &ResolveResult<'_>, expected: &[u8]) -> bool {
    matches!(value, ResolveResult::Bound(found) if found.as_ref().starts_with(expected))
}

fn invalid<T>(message: impl Into<String>) -> Result<T> {
    Err(format_error(message))
}

fn format_error(message: impl Into<String>) -> Error {
    Error::InvalidFormat(message.into())
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::core::OwnedPackage;

    #[test]
    fn reads_libreoffice_document_signature_metadata_without_verification() {
        let bytes = include_bytes!(
            "../../../3rdparty/libreoffice-core/xmlsecurity/qa/unit/signing/data/signed_with_x509certificate_chain.odt"
        );
        let package = OwnedPackage::from_bytes(bytes.to_vec()).unwrap();
        let signatures = package.digital_signatures().unwrap();
        assert_eq!(signatures.document_signatures.len(), 1);
        assert!(signatures.macro_signatures.is_empty());
        let signature = &signatures.document_signatures[0];
        assert_eq!(
            signature.signature_method,
            "http://www.w3.org/2001/04/xmldsig-more#rsa-sha256"
        );
        assert_eq!(signature.references.len(), 10);
        assert_eq!(signature.x509_certificates.len(), 3);
        assert_eq!(
            signature.signing_time.as_deref(),
            Some("2021-03-02T12:06:44.336956519")
        );
        assert!(signature.references.iter().any(|reference| {
            reference.uri == "content.xml"
                && reference.digest_method == "http://www.w3.org/2001/04/xmlenc#sha256"
        }));
    }

    #[test]
    fn rejects_spoofed_incomplete_and_active_signature_xml() {
        let valid_start = concat!(
            r#"<document-signatures xmlns="urn:oasis:names:tc:opendocument:xmlns:digitalsignature:1.0">"#,
            r#"<Signature xmlns="http://www.w3.org/2000/09/xmldsig#"><SignedInfo>"#,
            r#"<CanonicalizationMethod Algorithm="c"/><SignatureMethod Algorithm="s"/>"#,
            r#"<Reference URI="content.xml"><DigestMethod Algorithm="d"/><DigestValue>AAAA</DigestValue></Reference>"#,
            r#"</SignedInfo>"#,
        );
        assert!(parse_signature_container(
            format!(r#"{valid_start}<SignatureValue>AAAA</SignatureValue></Signature></document-signatures>"#).as_bytes()
        ).is_ok());
        for xml in [
            format!(r#"{valid_start}</Signature></document-signatures>"#),
            r#"<document-signatures xmlns="urn:wrong"><Signature xmlns="http://www.w3.org/2000/09/xmldsig#"/></document-signatures>"#.to_owned(),
            format!(r#"<!DOCTYPE x>{valid_start}<SignatureValue>AAAA</SignatureValue></Signature></document-signatures>"#),
            format!(r#"{valid_start}<SignatureValue>not-base64</SignatureValue></Signature></document-signatures>"#),
        ] {
            assert!(parse_signature_container(xml.as_bytes()).is_err(), "accepted {xml}");
        }
    }
}
