//! Private canonical application-root classification for the legacy host.
//!
//! The focused detector owns package and directory ingress. The remaining
//! host readers only need to identify an already parsed `Document.iwa`
//! payload, so they use this private, schema-shaped classifier instead of a
//! detector dependency or a numeric message-type registry.

use crate::application::Application;
use litchi_iwa_common::wire::{WireField, parse_wire_fields};
use litchi_iwa_common::{Result, decode_varint_from_bytes, varint::encoded_len};

const PAGES_DOCUMENT_FIELD: u32 = 15;
const NUMBERS_REFERENCE_FIELDS: [u32; 3] = [4, 5, 6];
const NUMBERS_DOCUMENT_FIELD: u32 = 8;
const KEYNOTE_REFERENCE_FIELD: u32 = 2;
const KEYNOTE_DOCUMENT_FIELD: u32 = 3;
const REFERENCE_IDENTIFIER_FIELD: u32 = 1;
const MESSAGE_WIRE_TYPE: u8 = 2;
const VARINT_WIRE_TYPE: u8 = 0;

/// Classify one canonical root `DocumentArchive` payload.
///
/// This is deliberately crate-private. Numeric protobuf fields remain an
/// implementation detail of the host's root reader and never become a public
/// identifier or application-detection API. Shared and malformed roots fail
/// closed, matching the focused detector's compatibility classifier.
pub(crate) fn detect(payload: &[u8]) -> Option<Application> {
    classify(payload).ok().flatten()
}

fn classify(payload: &[u8]) -> Result<Option<Application>> {
    let fields = parse_canonical_fields(payload)?;
    let pages = unique_field(payload, &fields, PAGES_DOCUMENT_FIELD, MESSAGE_WIRE_TYPE)?
        .map(valid_shared_document)
        .transpose()?
        .unwrap_or(false);

    let numbers = NUMBERS_REFERENCE_FIELDS
        .into_iter()
        .try_fold(true, |matches, field| {
            Ok::<_, litchi_iwa_common::Error>(
                matches
                    && unique_field(payload, &fields, field, MESSAGE_WIRE_TYPE)?
                        .map(valid_reference)
                        .transpose()?
                        .unwrap_or(false),
            )
        })?
        && unique_field(payload, &fields, NUMBERS_DOCUMENT_FIELD, MESSAGE_WIRE_TYPE)?
            .map(valid_shared_document)
            .transpose()?
            .unwrap_or(false);

    let keynote = unique_field(payload, &fields, KEYNOTE_REFERENCE_FIELD, MESSAGE_WIRE_TYPE)?
        .map(valid_reference)
        .transpose()?
        .unwrap_or(false)
        && unique_field(payload, &fields, KEYNOTE_DOCUMENT_FIELD, MESSAGE_WIRE_TYPE)?
            .map(valid_shared_document)
            .transpose()?
            .unwrap_or(false);

    Ok(match (pages, numbers, keynote) {
        (true, false, false) => Some(Application::Pages),
        (false, true, false) => Some(Application::Numbers),
        (false, false, true) => Some(Application::Keynote),
        _ => None,
    })
}

fn parse_canonical_fields(payload: &[u8]) -> Result<Vec<WireField>> {
    let fields = parse_wire_fields(payload)?;
    for field in &fields {
        field.validate_canonical_framing(payload)?;
    }
    Ok(fields)
}

fn unique_field<'a>(
    payload: &'a [u8],
    fields: &[WireField],
    number: u32,
    wire_type: u8,
) -> Result<Option<&'a [u8]>> {
    let mut matches = fields.iter().filter(|field| field.number() == number);
    let Some(field) = matches.next() else {
        return Ok(None);
    };
    if matches.next().is_some() || field.wire_type() != wire_type {
        return Ok(None);
    }
    Ok(Some(field.checked_payload(payload)?))
}

fn valid_reference(payload: &[u8]) -> Result<bool> {
    let fields = parse_canonical_fields(payload)?;
    let Some(identifier) = unique_field(
        payload,
        &fields,
        REFERENCE_IDENTIFIER_FIELD,
        VARINT_WIRE_TYPE,
    )?
    else {
        return Ok(false);
    };
    Ok(is_canonical_reference_identifier(identifier))
}

fn is_canonical_reference_identifier(payload: &[u8]) -> bool {
    let Ok((identifier, width)) = decode_varint_from_bytes(payload) else {
        return false;
    };
    width == payload.len() && width == encoded_len(identifier)
}

fn valid_shared_document(payload: &[u8]) -> Result<bool> {
    let fields = parse_canonical_fields(payload)?;
    let Some(document) = unique_field(
        payload,
        &fields,
        REFERENCE_IDENTIFIER_FIELD,
        MESSAGE_WIRE_TYPE,
    )?
    else {
        return Ok(false);
    };
    let _ = parse_canonical_fields(document)?;
    Ok(true)
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::protobuf::{kn, tn, tp, tsa, tsk, tsp};
    use prost::Message;

    fn shared_document() -> tsa::DocumentArchive {
        tsa::DocumentArchive {
            super_: tsk::DocumentArchive::default(),
            ..Default::default()
        }
    }

    fn reference(identifier: u64) -> tsp::Reference {
        tsp::Reference {
            identifier,
            ..Default::default()
        }
    }

    fn payload(application: Application) -> Vec<u8> {
        match application {
            Application::Pages => tp::DocumentArchive {
                super_: shared_document(),
                ..Default::default()
            }
            .encode_to_vec(),
            Application::Numbers => tn::DocumentArchive {
                super_: shared_document(),
                stylesheet: reference(1),
                sidebar_order: reference(2),
                theme: reference(3),
                ..Default::default()
            }
            .encode_to_vec(),
            Application::Keynote => kn::DocumentArchive {
                super_: shared_document(),
                show: reference(1),
                ..Default::default()
            }
            .encode_to_vec(),
            Application::Common => Vec::new(),
        }
    }

    #[test]
    fn recognizes_canonical_application_roots() {
        for application in [
            Application::Pages,
            Application::Numbers,
            Application::Keynote,
        ] {
            assert_eq!(detect(&payload(application)), Some(application));
        }
    }

    #[test]
    fn rejects_malformed_and_ambiguous_roots() {
        assert_eq!(detect(&[0x80]), None);

        let mut ambiguous = payload(Application::Pages);
        ambiguous.extend_from_slice(&payload(Application::Numbers));
        assert_eq!(detect(&ambiguous), None);

        let mut duplicate = payload(Application::Pages);
        duplicate.extend_from_slice(&payload(Application::Pages));
        assert_eq!(detect(&duplicate), None);
    }
}
