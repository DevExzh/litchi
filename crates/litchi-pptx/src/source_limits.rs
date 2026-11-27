//! Physical XML-budget checks shared by deferred PPTX readers.
//!
//! The OPC package enforces member and part byte ceilings while materializing
//! a payload.  Semantic readers still need to enforce the caller's per-XML
//! event, depth, and attribute-byte ceilings before handing that payload to a
//! format-specific parser.  This module deliberately does not validate the
//! document vocabulary: format-specific structure remains for the
//! authoritative parser to classify. Tokenization and tag-balance errors are
//! refused here; document-root grammar belongs to the authoritative parser.

use litchi_ooxml_common::xml::attributes::BytesStartExt as _;
use litchi_opc::{OpcError, SourceLimitKind, SourceLimits};
use quick_xml::Reader;
use quick_xml::events::Event;

use crate::{Error, Result};

/// Validate physical XML resource dimensions selected for one source read.
///
/// The scan counts every non-EOF quick-XML event, including events that a
/// later MCE or semantic parser may discard. Attribute accounting follows
/// the shared OPC policy and measures each raw attribute value. Tokenization,
/// nesting and duplicate-attribute errors are refused with content-free text
/// before MCE can discard a branch. This is a physical-budget scan, not a
/// replacement for the authoritative document-root and namespace grammar.
pub(crate) fn validate_xml_limits(xml: &[u8], limits: SourceLimits) -> Result<()> {
    // Physical budgets do not require namespace resolution; the semantic
    // parser validates namespaces after this scan.
    let mut reader = Reader::from_reader(xml);
    reader.config_mut().check_end_names = true;
    let mut depth = 0usize;
    let mut events = 0usize;

    loop {
        let event = match reader.read_event() {
            Ok(event) => event,
            Err(_) => return Err(Error::Invalid("PPTX XML is malformed".into())),
        };
        if matches!(event, Event::Eof) {
            return if depth == 0 {
                Ok(())
            } else {
                Err(Error::Invalid("PPTX XML is malformed".into()))
            };
        }

        events = events.saturating_add(1);
        check(SourceLimitKind::Events, events, limits.max_events())?;

        match event {
            Event::Start(element) => {
                let next_depth = depth.saturating_add(1);
                check(SourceLimitKind::Depth, next_depth, limits.max_depth())?;
                for attribute in element.checked_attributes() {
                    let attribute =
                        attribute.map_err(|_| Error::Invalid("PPTX XML is malformed".into()))?;
                    check(
                        SourceLimitKind::XmlAttributeBytes,
                        attribute.value.len(),
                        limits.max_xml_attribute_bytes(),
                    )?;
                }
                depth = next_depth;
            },
            Event::Empty(element) => {
                let next_depth = depth.saturating_add(1);
                check(SourceLimitKind::Depth, next_depth, limits.max_depth())?;
                for attribute in element.checked_attributes() {
                    let attribute =
                        attribute.map_err(|_| Error::Invalid("PPTX XML is malformed".into()))?;
                    check(
                        SourceLimitKind::XmlAttributeBytes,
                        attribute.value.len(),
                        limits.max_xml_attribute_bytes(),
                    )?;
                }
            },
            Event::End(_) => {
                let Some(parent) = depth.checked_sub(1) else {
                    return Err(Error::Invalid("PPTX XML is malformed".into()));
                };
                depth = parent;
            },
            Event::Text(_)
            | Event::CData(_)
            | Event::Comment(_)
            | Event::Decl(_)
            | Event::PI(_)
            | Event::DocType(_)
            | Event::GeneralRef(_) => {},
            Event::Eof => unreachable!("EOF is handled above"),
        }
    }
}

fn check(kind: SourceLimitKind, observed: usize, maximum: usize) -> Result<()> {
    if observed > maximum {
        return Err(Error::Opc(OpcError::Limit {
            kind,
            observed: observed as u64,
            maximum: maximum as u64,
        }));
    }
    Ok(())
}

#[cfg(test)]
mod tests {
    use super::validate_xml_limits;
    use crate::Error;
    use litchi_opc::{OpcError, SourceLimitKind, SourceLimits};

    const PML: &str = "http://schemas.openxmlformats.org/presentationml/2006/main";

    fn limit_error(xml: &str, limits: SourceLimits, expected: SourceLimitKind) {
        assert!(matches!(
            validate_xml_limits(xml.as_bytes(), limits),
            Err(Error::Opc(OpcError::Limit { kind, .. })) if kind == expected
        ));
    }

    #[test]
    fn caller_event_depth_and_attribute_limits_are_checked() {
        let xml = format!(
            r#"<styleSheet xmlns="{PML}" marker="123"><fonts><font/></fonts></styleSheet>"#
        );
        limit_error(
            &xml,
            SourceLimits::builder()
                .max_events(2)
                .unwrap()
                .build()
                .unwrap(),
            SourceLimitKind::Events,
        );
        limit_error(
            &xml,
            SourceLimits::builder()
                .max_depth(1)
                .unwrap()
                .build()
                .unwrap(),
            SourceLimitKind::Depth,
        );
        limit_error(
            &xml,
            SourceLimits::builder()
                .max_xml_attribute_bytes(2)
                .unwrap()
                .max_relationship_target_bytes(2)
                .unwrap()
                .build()
                .unwrap(),
            SourceLimitKind::XmlAttributeBytes,
        );
    }

    #[test]
    fn malformed_xml_is_refused_before_authoritative_parser() {
        let xml = format!(r#"<styleSheet xmlns="{PML}"><fonts></styleSheet></fonts>"#);
        assert!(matches!(
            validate_xml_limits(xml.as_bytes(), SourceLimits::default()),
            Err(Error::Invalid(message)) if message == "PPTX XML is malformed"
        ));
    }

    #[test]
    fn duplicate_attributes_are_not_hidden_from_the_budget_gate() {
        let xml = format!(r#"<styleSheet xmlns="{PML}" marker="one" marker="two"/>"#);
        assert!(matches!(
            validate_xml_limits(xml.as_bytes(), SourceLimits::default()),
            Err(Error::Invalid(message)) if message == "PPTX XML is malformed"
        ));
    }
}
