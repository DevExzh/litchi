//! Macro and navigation button field models.

use super::Field;

use crate::error::Result;

use super::super::codec::{
    field_instruction_remainder, parse_go_to_button_operands, parse_macro_button_operands,
};

/// A typed, inert Word `MACROBUTTON` field.
///
/// ECMA-376 Part 1 §17.16.5.34 defines two stored field arguments: a macro or
/// command name and the text or graphic used as its button. This type exposes
/// stored text only; it never resolves, loads, invokes, or otherwise executes
/// the named target.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct MacroButton {
    instruction: String,
    cached_result: Option<String>,
    dirty: bool,
    locked: bool,
    macro_name: String,
    display_text: String,
}

impl MacroButton {
    fn from_field(field: &Field) -> Result<Option<Self>> {
        let Some((macro_name, display_text)) = parse_macro_button_operands(field.instruction())?
        else {
            return Ok(None);
        };

        Ok(Some(Self {
            instruction: field.instruction.clone(),
            cached_result: field.result.clone(),
            dirty: field.dirty,
            locked: field.locked,
            macro_name,
            display_text,
        }))
    }

    /// Return the complete stored field instruction.
    pub fn instruction(&self) -> &str {
        &self.instruction
    }

    /// Return the stored macro or command name without resolving or invoking it.
    pub fn macro_name(&self) -> &str {
        &self.macro_name
    }

    /// Return the stored button text.
    ///
    /// This is source metadata, not a generated result.
    pub fn display_text(&self) -> &str {
        &self.display_text
    }

    /// Return the cached visible field result, if present.
    ///
    /// This is stored text only and is never regenerated from the named target.
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
}

/// A typed, inert Word `GOTOBUTTON` field.
///
/// ECMA-376 Part 1 §17.16.5.23 defines two stored field arguments: a
/// destination and the text or graphic used as its button. This type exposes
/// stored text only; it never resolves a destination, changes the insertion
/// point, or activates a jump.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct GoToButton {
    instruction: String,
    cached_result: Option<String>,
    dirty: bool,
    locked: bool,
    target: String,
    button_text: String,
}

impl GoToButton {
    fn from_field(field: &Field) -> Result<Option<Self>> {
        let Some((target, button_text)) = parse_go_to_button_operands(field.instruction())? else {
            return Ok(None);
        };

        Ok(Some(Self {
            instruction: field.instruction.clone(),
            cached_result: field.result.clone(),
            dirty: field.dirty,
            locked: field.locked,
            target,
            button_text,
        }))
    }

    /// Return the complete stored field instruction.
    pub fn instruction(&self) -> &str {
        &self.instruction
    }

    /// Return the stored destination without resolving or navigating to it.
    ///
    /// A destination can be a bookmark, page reference, annotation, footnote,
    /// line, page, or section expression.
    pub fn target(&self) -> &str {
        &self.target
    }

    /// Return the stored text or graphic-label expression for the button.
    ///
    /// This is source metadata, not an activated control.
    pub fn button_text(&self) -> &str {
        &self.button_text
    }

    /// Return the cached visible field result, if present.
    ///
    /// This is stored text only and is never regenerated from the destination.
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
}

impl Field {
    /// Check whether this is a `MACROBUTTON` field.
    ///
    /// Recognition is limited to the stored field instruction. It never
    /// resolves, loads, invokes, or otherwise executes a macro or command.
    pub fn is_macro_button(&self) -> bool {
        field_instruction_remainder(&self.instruction, "MACROBUTTON").is_some()
    }

    /// Parse this field as inert typed macro-button metadata.
    ///
    /// Returns `Ok(None)` for non-`MACROBUTTON` fields. The result exposes
    /// only the stored macro or command name, button text, cached content, and
    /// dirty/lock state; it never resolves, loads, invokes, or executes the
    /// named target.
    pub fn macro_button(&self) -> Result<Option<MacroButton>> {
        MacroButton::from_field(self)
    }

    /// Check whether this is a `GOTOBUTTON` field.
    ///
    /// Recognition is limited to the stored field instruction. It never
    /// resolves a destination, changes the insertion point, or refreshes the
    /// cached result.
    pub fn is_go_to_button(&self) -> bool {
        field_instruction_remainder(&self.instruction, "GOTOBUTTON").is_some()
    }

    /// Parse this field as inert typed `GOTOBUTTON` metadata.
    ///
    /// Returns `Ok(None)` for non-`GOTOBUTTON` fields. The result exposes
    /// only the stored target, button text, cached content, and dirty/lock
    /// state; it never resolves a bookmark, page, annotation, footnote, or
    /// other target, changes the insertion point, or refreshes a field.
    pub fn go_to_button(&self) -> Result<Option<GoToButton>> {
        GoToButton::from_field(self)
    }
}
