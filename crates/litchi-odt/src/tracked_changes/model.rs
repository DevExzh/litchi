//! Canonical semantic values and validation for tracked changes.

use super::{MAX_AGGREGATE_BYTES, MAX_CHANGES, MAX_VALUE_BYTES, invalid, make_error};
use crate::parser::{ChangeType, TrackChange, TrackedChanges};
use base64::Engine as _;
use base64::engine::general_purpose::STANDARD as BASE64_STANDARD;
use litchi_core::Result;
use std::collections::HashSet;

impl TrackedChanges {
    /// Validate a manually constructed tracked-change declaration table.
    ///
    /// This validates declaration metadata only. Range markers remain part of
    /// the surrounding document and are correlated when a complete document is
    /// parsed.
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
                .and_then(|value| value.checked_add(change.comment.as_deref().map_or(0, str::len)))
                .ok_or_else(|| make_error("tracked-change aggregate size overflow"))?;
            if aggregate > MAX_AGGREGATE_BYTES {
                return invalid("tracked-change values exceed 16 MiB");
            }
        }
        Ok(())
    }
}

impl TrackChange {
    /// Validate one change declaration before it is authored into ODT XML.
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
        if let Some(comment) = &self.comment {
            validate_value(comment, "tracked-change comment", true)?;
        }
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

pub(super) fn validate_ncname(value: &str, context: &str) -> Result<()> {
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

pub(super) fn validate_value(value: &str, context: &str, empty_allowed: bool) -> Result<()> {
    if !empty_allowed && value.is_empty() {
        return invalid(format!("{context} cannot be empty"));
    }
    if value.len() > MAX_VALUE_BYTES {
        return invalid(format!("{context} exceeds 64 KiB"));
    }
    if value.chars().any(
        |character| matches!(character, '\0'..='\u{8}' | '\u{b}' | '\u{c}' | '\u{e}'..='\u{1f}'),
    ) {
        return invalid(format!("{context} contains an XML-prohibited character"));
    }
    Ok(())
}

/// Stable text story used when placing tracked-change markers.
#[derive(Debug, Clone, PartialEq, Eq, Hash)]
pub enum Story {
    /// Body paragraph or heading outside tables, in document order.
    Paragraph(usize),
    /// Paragraph inside a table cell, addressed by lexical table/row/cell/paragraph order.
    TableCell {
        table: usize,
        row: usize,
        cell: usize,
        paragraph: usize,
    },
}

/// Unicode-scalar position inside one stable text story.
#[derive(Debug, Clone, PartialEq, Eq, Hash)]
pub struct Position {
    pub story: Story,
    pub character: usize,
}
