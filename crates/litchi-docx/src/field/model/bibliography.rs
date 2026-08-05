//! Bibliography field models.

use super::{Field, Switch};

use crate::error::{Error, Result};

use super::super::codec::{
    field_instruction_remainder, has_field_switch, parse_citation_operand_and_switches,
    parse_field_switches,
};

/// A typed, inert Word `CITATION` field.
///
/// The field stores one primary bibliography-source tag plus zero or more
/// multi-source tags introduced by `\m`. This model preserves that metadata
/// and cached display text only. It never accesses bibliography source XML,
/// formats citations, resolves locales, or executes field instructions.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Citation {
    instruction: String,
    cached_result: Option<String>,
    dirty: bool,
    locked: bool,
    source_tags: Vec<String>,
    switches: Vec<Switch>,
}

impl Citation {
    fn from_field(field: &Field) -> Result<Option<Self>> {
        let Some((primary_source_tag, switches)) =
            parse_citation_operand_and_switches(field.instruction())?
        else {
            return Ok(None);
        };
        let primary_source_tag = primary_source_tag.ok_or_else(|| {
            Error::Invalid("CITATION field is missing its source tag".to_string())
        })?;
        if primary_source_tag.is_empty() {
            return Err(Error::Invalid(
                "CITATION field source tag is empty".to_string(),
            ));
        }

        let mut source_tags = vec![primary_source_tag];
        for switch in &switches {
            if switch.name != 'm' {
                continue;
            }
            let source_tag = switch.argument.as_deref().ok_or_else(|| {
                Error::Invalid("CITATION \\m switch requires a source tag".to_string())
            })?;
            if source_tag.is_empty() {
                return Err(Error::Invalid(
                    "CITATION \\m source tag is empty".to_string(),
                ));
            }
            source_tags.push(source_tag.to_string());
        }

        Ok(Some(Self {
            instruction: field.instruction.clone(),
            cached_result: field.result.clone(),
            dirty: field.dirty,
            locked: field.locked,
            source_tags,
            switches,
        }))
    }

    /// Return the complete stored field instruction.
    pub fn instruction(&self) -> &str {
        &self.instruction
    }

    /// Return the cached formatted citation, if present.
    pub fn cached_result(&self) -> Option<&str> {
        self.cached_result.as_deref()
    }

    /// Whether a word processor has marked the cached result stale.
    pub fn is_dirty(&self) -> bool {
        self.dirty
    }

    /// Whether a word processor has locked this field against refresh.
    pub fn is_locked(&self) -> bool {
        self.locked
    }

    /// Return the primary source tag stored directly after `CITATION`.
    pub fn primary_source_tag(&self) -> &str {
        &self.source_tags[0]
    }

    /// Return primary and `\m` multi-source tags in instruction order.
    pub fn source_tags(&self) -> &[String] {
        &self.source_tags
    }

    /// Return the additional source tags introduced by `\m` switches.
    pub fn additional_source_tags(&self) -> &[String] {
        &self.source_tags[1..]
    }

    /// Return all stored switches in source order.
    ///
    /// Switch semantics can apply to the primary or a preceding `\m` source,
    /// so callers that need producer-specific interpretation should retain this
    /// source order instead of assuming a global setting.
    pub fn switches(&self) -> &[Switch] {
        &self.switches
    }

    /// Check whether a case-insensitive ASCII switch appears in this field.
    pub fn has_switch(&self, name: char) -> bool {
        has_field_switch(&self.switches, name)
    }
}

/// A typed, inert Word `BIBLIOGRAPHY` field.
///
/// This preserves only the stored field instruction, switches, and cached
/// result. It does not discover bibliography sources, apply a style, sort
/// entries, or generate a bibliography.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Bibliography {
    instruction: String,
    cached_result: Option<String>,
    dirty: bool,
    locked: bool,
    switches: Vec<Switch>,
}

impl Bibliography {
    fn from_field(field: &Field) -> Result<Option<Self>> {
        let Some(switches) = parse_field_switches(field.instruction(), "BIBLIOGRAPHY")? else {
            return Ok(None);
        };
        Ok(Some(Self {
            instruction: field.instruction.clone(),
            cached_result: field.result.clone(),
            dirty: field.dirty,
            locked: field.locked,
            switches,
        }))
    }

    /// Return the complete stored field instruction.
    pub fn instruction(&self) -> &str {
        &self.instruction
    }

    /// Return the cached visible bibliography result, if present.
    pub fn cached_result(&self) -> Option<&str> {
        self.cached_result.as_deref()
    }

    /// Whether a word processor has marked the cached result stale.
    pub fn is_dirty(&self) -> bool {
        self.dirty
    }

    /// Whether a word processor has locked this field against refresh.
    pub fn is_locked(&self) -> bool {
        self.locked
    }

    /// Return the field switches in source order.
    pub fn switches(&self) -> &[Switch] {
        &self.switches
    }

    /// Check whether a case-insensitive ASCII switch appears in this field.
    pub fn has_switch(&self, name: char) -> bool {
        has_field_switch(&self.switches, name)
    }
}

impl Field {
    /// Check whether this is a `CITATION` bibliography field.
    ///
    /// Recognition is limited to the stored field instruction. It never looks
    /// up bibliography sources, formats a citation, follows a data-store
    /// reference, or refreshes the cached result.
    pub fn is_citation(&self) -> bool {
        field_instruction_remainder(&self.instruction, "CITATION").is_some()
    }

    /// Parse this field as an inert typed bibliography citation.
    ///
    /// Returns `Ok(None)` for non-`CITATION` fields. The result exposes only
    /// stored source tags, switches, cached content, and dirty/lock state; it
    /// never resolves sources or formats a citation.
    pub fn citation(&self) -> Result<Option<Citation>> {
        Citation::from_field(self)
    }

    /// Check whether this is a `BIBLIOGRAPHY` field.
    ///
    /// This recognizes persisted configuration only. It does not enumerate
    /// sources, sort them, or generate bibliography text.
    pub fn is_bibliography(&self) -> bool {
        field_instruction_remainder(&self.instruction, "BIBLIOGRAPHY").is_some()
    }

    /// Parse this field as an inert typed bibliography field.
    ///
    /// Returns `Ok(None)` for non-`BIBLIOGRAPHY` fields. Stored switches and
    /// cached visible content remain data only; no bibliography is generated.
    pub fn bibliography(&self) -> Result<Option<Bibliography>> {
        Bibliography::from_field(self)
    }
}
