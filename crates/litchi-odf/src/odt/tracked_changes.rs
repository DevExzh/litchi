//! Validation and deterministic serialization for text change declarations.

use super::parser::{ChangeType, TrackChange, TrackedChanges};
use base64::Engine as _;
use base64::engine::general_purpose::STANDARD as BASE64_STANDARD;
use litchi_core::{Error, Result};
use std::collections::HashSet;

const MAX_CHANGES: usize = 1_000_000;
const MAX_VALUE_BYTES: usize = 65_536;
const MAX_AGGREGATE_BYTES: usize = 16 * 1_048_576;

impl TrackedChanges {
    /// Validate a manually constructed tracked-change declaration table.
    ///
    /// This validates the declaration metadata only. Range markers remain part of
    /// the surrounding document and are correlated by [`Document::tracked_changes`]
    /// when parsing a complete document.
    pub fn validate(&self) -> Result<()> {
        if self.changes.len() > MAX_CHANGES {
            return invalid(format!(
                "tracked changes exceed the {MAX_CHANGES} declaration limit"
            ));
        }
        validate_policy(self)?;

        let mut ids = HashSet::with_capacity(self.changes.len());
        let mut aggregate = 0usize;
        for change in &self.changes {
            change.validate()?;
            if !ids.insert(change.id.as_str()) {
                return invalid(format!("duplicate tracked-change ID '{}'", change.id));
            }
            aggregate = aggregate
                .checked_add(change.content.len())
                .and_then(|value| value.checked_add(change.author.as_deref().map_or(0, str::len)))
                .and_then(|value| value.checked_add(change.date.as_deref().map_or(0, str::len)))
                .ok_or_else(|| make_error("tracked-change aggregate size overflow"))?;
            if aggregate > MAX_AGGREGATE_BYTES {
                return invalid("tracked-change values exceed 16 MiB");
            }
        }
        Ok(())
    }

    /// Serialize the `text:tracked-changes` declaration table deterministically.
    ///
    /// Insertion and format-change `content` is live document text discovered from
    /// range markers, so it is not duplicated in the declaration table. Deletion
    /// content is emitted as inert ODF text with tabs, line breaks, and repeated
    /// spaces represented by their standard text controls.
    pub fn to_xml_fragment(&self) -> Result<String> {
        self.validate()?;
        let mut xml = String::with_capacity(256 + self.changes.len().saturating_mul(192));
        xml.push_str(
            r#"<text:tracked-changes xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0" xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:dc="http://purl.org/dc/elements/1.1/""#,
        );
        if let Some(value) = self.track_changes {
            xml.push_str(r#" text:track-changes=""#);
            xml.push_str(if value { "true" } else { "false" });
            xml.push('"');
        }
        if let Some(value) = &self.protection_key {
            push_attribute(&mut xml, "text:protection-key", value);
        }
        if let Some(value) = &self.protection_key_digest_algorithm {
            push_attribute(
                &mut xml,
                "text:protection-key-digest-algorithm",
                value,
            );
        }
        if self.changes.is_empty() {
            xml.push_str("/>");
            return Ok(xml);
        }
        xml.push('>');
        for change in &self.changes {
            write_change(&mut xml, change);
        }
        xml.push_str("</text:tracked-changes>");
        Ok(xml)
    }
}

impl TrackChange {
    /// Validate one change declaration before serialization.
    pub fn validate(&self) -> Result<()> {
        validate_ncname(&self.id, "tracked-change ID")?;
        if let Some(xml_id) = &self.xml_id {
            validate_ncname(xml_id, "tracked-change xml:id")?;
            if xml_id != &self.id {
                return invalid("tracked-change text:id and xml:id must match");
            }
        }
        let author = self
            .author
            .as_deref()
            .ok_or_else(|| make_error("tracked change requires dc:creator"))?;
        let date = self
            .date
            .as_deref()
            .ok_or_else(|| make_error("tracked change requires dc:date"))?;
        validate_value(author, "tracked-change creator", false)?;
        validate_value(date, "tracked-change date", false)?;
        validate_value(&self.content, "tracked-change content", true)?;

        match self.change_type {
            ChangeType::Insertion => {
                if self.style_name.is_some() || self.merge_last_paragraph.is_some() {
                    return invalid("insertion cannot carry format-change or deletion attributes");
                }
            },
            ChangeType::Deletion => {
                if self.style_name.is_some() {
                    return invalid("deletion cannot carry text:style-name");
                }
            },
            ChangeType::FormatChange => {
                if self.merge_last_paragraph.is_some() {
                    return invalid("format change cannot carry text:merge-last-paragraph");
                }
                if let Some(style_name) = &self.style_name {
                    validate_value(style_name, "format-change style name", false)?;
                }
            },
        }
        Ok(())
    }
}

fn validate_policy(value: &TrackedChanges) -> Result<()> {
    match (
        value.protection_key.as_deref(),
        value.protection_key_digest_algorithm.as_deref(),
    ) {
        (None, Some(_)) => {
            return invalid("protection-key digest algorithm requires a protection key");
        },
        (Some(key), algorithm) => {
            validate_value(key, "tracked-change protection key", false)?;
            BASE64_STANDARD.decode(key).map_err(|error| {
                make_error(format!("invalid tracked-change protection key: {error}"))
            })?;
            if let Some(algorithm) = algorithm {
                validate_value(
                    algorithm,
                    "tracked-change protection-key digest algorithm",
                    false,
                )?;
            }
        },
        (None, None) => {},
    }
    Ok(())
}

fn write_change(xml: &mut String, change: &TrackChange) {
    xml.push_str(r#"<text:changed-region text:id=""#);
    push_escaped(xml, &change.id, true);
    xml.push('"');
    if let Some(xml_id) = &change.xml_id {
        push_attribute(xml, "xml:id", xml_id);
    }
    xml.push('>');

    let kind = match change.change_type {
        ChangeType::Insertion => "insertion",
        ChangeType::Deletion => "deletion",
        ChangeType::FormatChange => "format-change",
    };
    xml.push_str("<text:");
    xml.push_str(kind);
    if let Some(style_name) = &change.style_name {
        push_attribute(xml, "text:style-name", style_name);
    }
    if let Some(value) = change.merge_last_paragraph {
        xml.push_str(r#" text:merge-last-paragraph=""#);
        xml.push_str(if value { "true" } else { "false" });
        xml.push('"');
    }
    xml.push('>');
    xml.push_str("<office:change-info><dc:creator>");
    push_escaped(xml, change.author.as_deref().expect("validated creator"), false);
    xml.push_str("</dc:creator><dc:date>");
    push_escaped(xml, change.date.as_deref().expect("validated date"), false);
    xml.push_str("</dc:date></office:change-info>");
    if change.change_type == ChangeType::Deletion && !change.content.is_empty() {
        xml.push_str("<text:p>");
        write_text_content(xml, &change.content);
        xml.push_str("</text:p>");
    }
    xml.push_str("</text:");
    xml.push_str(kind);
    xml.push_str("></text:changed-region>");
}

fn write_text_content(xml: &mut String, value: &str) {
    let mut plain_start = 0usize;
    let mut chars = value.char_indices().peekable();
    while let Some((index, character)) = chars.next() {
        let special = matches!(character, ' ' | '\t' | '\n' | '\r');
        if !special {
            continue;
        }
        if plain_start < index {
            push_escaped(xml, &value[plain_start..index], false);
        }
        match character {
            ' ' => {
                let mut count = 1usize;
                let mut end = index + 1;
                while let Some(&(next_index, ' ')) = chars.peek() {
                    chars.next();
                    count += 1;
                    end = next_index + 1;
                }
                if count == 1 {
                    xml.push(' ');
                } else {
                    xml.push_str(r#"<text:s text:c=""#);
                    xml.push_str(&count.to_string());
                    xml.push_str("\"/>");
                }
                plain_start = end;
            },
            '\t' => {
                xml.push_str("<text:tab/>");
                plain_start = index + 1;
            },
            '\n' => {
                xml.push_str("<text:line-break/>");
                plain_start = index + 1;
            },
            '\r' => {
                if let Some(&(next_index, '\n')) = chars.peek() {
                    chars.next();
                    plain_start = next_index + 1;
                } else {
                    plain_start = index + 1;
                }
                xml.push_str("<text:line-break/>");
            },
            _ => unreachable!(),
        }
    }
    if plain_start < value.len() {
        push_escaped(xml, &value[plain_start..], false);
    }
}

fn push_attribute(xml: &mut String, name: &str, value: &str) {
    xml.push(' ');
    xml.push_str(name);
    xml.push_str("=\"");
    push_escaped(xml, value, true);
    xml.push('"');
}

fn push_escaped(xml: &mut String, value: &str, attribute: bool) {
    for character in value.chars() {
        match character {
            '&' => xml.push_str("&amp;"),
            '<' => xml.push_str("&lt;"),
            '>' => xml.push_str("&gt;"),
            '"' if attribute => xml.push_str("&quot;"),
            '\'' if attribute => xml.push_str("&apos;"),
            _ => xml.push(character),
        }
    }
}

fn validate_ncname(value: &str, context: &str) -> Result<()> {
    validate_value(value, context, false)?;
    let mut characters = value.chars();
    let first = characters.next().expect("non-empty value validated");
    if !(first == '_' || first.is_alphabetic())
        || characters.any(|character| {
            !(character == '_'
                || character == '-'
                || character == '.'
                || character.is_alphanumeric())
        })
    {
        return invalid(format!("{context} must be an XML NCName"));
    }
    Ok(())
}

fn validate_value(value: &str, context: &str, empty_allowed: bool) -> Result<()> {
    if !empty_allowed && value.is_empty() {
        return invalid(format!("{context} cannot be empty"));
    }
    if value.len() > MAX_VALUE_BYTES {
        return invalid(format!("{context} exceeds 64 KiB"));
    }
    if value.chars().any(|character| {
        matches!(character, '\0'..='\u{8}' | '\u{b}' | '\u{c}' | '\u{e}'..='\u{1f}')
    }) {
        return invalid(format!("{context} contains an XML-prohibited character"));
    }
    Ok(())
}

fn invalid<T>(message: impl Into<String>) -> Result<T> {
    Err(make_error(message))
}

fn make_error(message: impl Into<String>) -> Error {
    Error::InvalidFormat(message.into())
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::odt::parser::OdtParser;

    fn change(id: &str, change_type: ChangeType, content: &str) -> TrackChange {
        TrackChange {
            id: id.to_string(),
            xml_id: Some(id.to_string()),
            author: Some("A & B".to_string()),
            date: Some("2026-07-17T12:00:00+08:00".to_string()),
            change_type,
            style_name: (change_type == ChangeType::FormatChange)
                .then(|| "Changed Style".to_string()),
            merge_last_paragraph: (change_type == ChangeType::Deletion).then_some(false),
            content: content.to_string(),
        }
    }

    #[test]
    fn serializes_and_reparses_all_declaration_kinds() {
        let declarations = TrackedChanges {
            track_changes: Some(true),
            protection_key: Some("YWJj".to_string()),
            protection_key_digest_algorithm: Some("urn:example:sha256".to_string()),
            changes: vec![
                change("insert_1", ChangeType::Insertion, "live text"),
                change("delete_1", ChangeType::Deletion, "gone  text\t<&\nnext"),
                change("format_1", ChangeType::FormatChange, "formatted"),
            ],
        };
        let xml = declarations.to_xml_fragment().unwrap();
        assert!(xml.contains("<text:s text:c=\"2\"/>"));
        assert!(xml.contains("<text:tab/>"));
        assert!(xml.contains("&lt;&amp;"));
        assert!(!xml.contains("live text"));

        let parsed = OdtParser::parse_tracked_changes(&xml).unwrap();
        assert_eq!(parsed.track_changes, Some(true));
        assert_eq!(parsed.protection_key.as_deref(), Some("YWJj"));
        assert_eq!(parsed.changes.len(), 3);
        assert_eq!(parsed.changes[1].content, "gone  text\t<&\nnext");
        assert_eq!(parsed.changes[2].style_name.as_deref(), Some("Changed Style"));
    }

    #[test]
    fn rejects_invalid_constructed_declarations() {
        let mut declarations = TrackedChanges {
            changes: vec![change("same", ChangeType::Insertion, "")],
            ..TrackedChanges::default()
        };
        declarations.changes.push(declarations.changes[0].clone());
        assert!(declarations.to_xml_fragment().is_err());

        declarations.changes.truncate(1);
        declarations.changes[0].id = "bad:id".to_string();
        assert!(declarations.to_xml_fragment().is_err());

        declarations.changes[0].id = "valid".to_string();
        declarations.changes[0].xml_id = Some("other".to_string());
        assert!(declarations.to_xml_fragment().is_err());

        declarations.changes[0].xml_id = Some("valid".to_string());
        declarations.protection_key_digest_algorithm = Some("urn:sha256".to_string());
        assert!(declarations.to_xml_fragment().is_err());
    }

    #[test]
    fn parses_a_libreoffice_flat_document_before_serializing() {
        let path = std::path::Path::new(env!("CARGO_MANIFEST_DIR")).join(
            "../../3rdparty/libreoffice-core/sw/qa/uitest/data/redline-autocorrect.fodt",
        );
        let Ok(xml) = std::fs::read_to_string(path) else {
            return;
        };
        let parsed = OdtParser::parse_tracked_changes(&xml).unwrap();
        assert!(!parsed.changes.is_empty());
        let serialized = parsed.to_xml_fragment().unwrap();
        let reparsed = OdtParser::parse_tracked_changes(&serialized).unwrap();
        assert_eq!(reparsed.changes.len(), parsed.changes.len());
    }
}
