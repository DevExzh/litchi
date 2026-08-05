//! User-identity and placement field models.

use super::Field;

use crate::error::Result;

use super::super::codec::{
    field_instruction_remainder, parse_advance_field_adjustments, parse_user_identity_field_parts,
};

/// The stored kind of a user-identity field.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum UserIdentityKind {
    /// A `USERADDRESS` field.
    Address,
    /// A `USERINITIALS` field.
    Initials,
    /// A `USERNAME` field.
    Name,
}

/// A general-formatting request stored by a user-identity field.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum UserIdentityFormat {
    /// The `\\* Caps` formatting request.
    Caps,
    /// The `\\* FirstCap` formatting request.
    FirstCap,
    /// The `\\* Lower` formatting request.
    Lower,
    /// The `\\* Upper` formatting request.
    Upper,
}

/// A typed, inert Word `USERADDRESS`, `USERINITIALS`, or `USERNAME` field.
///
/// ECMA-376 Part 1 §§17.16.5.69–71 define these fields. This type exposes a
/// stored override, formatting request, and cached result only. It never reads
/// or modifies a host user's identity, applies formatting, or refreshes a
/// field.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct UserIdentity {
    instruction: String,
    kind: UserIdentityKind,
    override_value: Option<String>,
    formatting: Option<UserIdentityFormat>,
    cached_result: Option<String>,
    dirty: bool,
    locked: bool,
}

impl UserIdentity {
    fn from_field(field: &Field) -> Result<Option<Self>> {
        let Some((kind, override_value, formatting)) =
            parse_user_identity_field_parts(field.instruction())?
        else {
            return Ok(None);
        };

        Ok(Some(Self {
            instruction: field.instruction.clone(),
            kind,
            override_value,
            formatting,
            cached_result: field.result.clone(),
            dirty: field.dirty,
            locked: field.locked,
        }))
    }

    /// Return the complete stored field instruction.
    pub fn instruction(&self) -> &str {
        &self.instruction
    }

    /// Return whether this is an address, initials, or name field.
    pub fn kind(&self) -> UserIdentityKind {
        self.kind
    }

    /// Return the optional stored value that overrides the host user context.
    ///
    /// `Some("")` represents an explicitly supplied blank override. This
    /// stored text is never written to, read from, or compared with a host
    /// identity.
    pub fn override_value(&self) -> Option<&str> {
        self.override_value.as_deref()
    }

    /// Return the stored general-formatting request, if any.
    ///
    /// This request is metadata only and is never applied to an identity value.
    pub fn formatting(&self) -> Option<UserIdentityFormat> {
        self.formatting
    }

    /// Return the cached visible field result, if present.
    ///
    /// This is stored text only and is never regenerated from a host identity.
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

/// One stored point-based `ADVANCE` placement operation.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum AdvanceOperation {
    /// The `\\d` switch moves subsequent text down.
    Down,
    /// The `\\l` switch moves subsequent text left.
    Left,
    /// The `\\r` switch moves subsequent text right.
    Right,
    /// The `\\u` switch moves subsequent text up.
    Up,
    /// The `\\x` switch specifies a horizontal position from the left edge
    /// of the column, frame, or text box.
    HorizontalPosition,
    /// The `\\y` switch specifies a vertical position relative to the page.
    VerticalPosition,
}

/// One stored `ADVANCE` point adjustment.
///
/// This is an instruction for a word processor's layout engine only. It is
/// never applied by this library.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct AdvanceAdjustment {
    pub(in crate::field) operation: AdvanceOperation,
    pub(in crate::field) points: i64,
}

impl AdvanceAdjustment {
    /// Return the requested placement operation.
    pub fn operation(&self) -> AdvanceOperation {
        self.operation
    }

    /// Return the stored signed integral number of points.
    pub fn points(&self) -> i64 {
        self.points
    }
}

/// A typed, inert Word `ADVANCE` field.
///
/// ECMA-376 Part 1 §17.16.5.2 defines this field and its six point-based
/// placement switches. This type exposes stored adjustments and cached content
/// only. It never moves text, changes layout, reflows content, or refreshes a
/// field.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Advance {
    instruction: String,
    adjustments: Vec<AdvanceAdjustment>,
    cached_result: Option<String>,
    dirty: bool,
    locked: bool,
}

impl Advance {
    fn from_field(field: &Field) -> Result<Option<Self>> {
        let Some(adjustments) = parse_advance_field_adjustments(field.instruction())? else {
            return Ok(None);
        };

        Ok(Some(Self {
            instruction: field.instruction.clone(),
            adjustments,
            cached_result: field.result.clone(),
            dirty: field.dirty,
            locked: field.locked,
        }))
    }

    /// Return the complete stored field instruction.
    pub fn instruction(&self) -> &str {
        &self.instruction
    }

    /// Return the stored placement adjustments in source order.
    ///
    /// Repeated operations are preserved; this library does not resolve or
    /// apply them.
    pub fn adjustments(&self) -> &[AdvanceAdjustment] {
        &self.adjustments
    }

    /// Return the cached visible field result, if present.
    ///
    /// `ADVANCE` has no regenerated value here; any returned text is stored
    /// source content only.
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
    /// Check whether this is a `USERADDRESS` field.
    ///
    /// Recognition is limited to the stored field instruction. It never reads
    /// the current user's address or refreshes the cached result.
    pub fn is_user_address(&self) -> bool {
        field_instruction_remainder(&self.instruction, "USERADDRESS").is_some()
    }

    /// Check whether this is a `USERINITIALS` field.
    ///
    /// Recognition is limited to the stored field instruction. It never reads
    /// the current user's initials or refreshes the cached result.
    pub fn is_user_initials(&self) -> bool {
        field_instruction_remainder(&self.instruction, "USERINITIALS").is_some()
    }

    /// Check whether this is a `USERNAME` field.
    ///
    /// Recognition is limited to the stored field instruction. It never reads
    /// the current user's name or refreshes the cached result.
    pub fn is_user_name(&self) -> bool {
        field_instruction_remainder(&self.instruction, "USERNAME").is_some()
    }

    /// Check whether this is a `USERADDRESS`, `USERINITIALS`, or `USERNAME` field.
    pub fn is_user_identity_field(&self) -> bool {
        self.is_user_address() || self.is_user_initials() || self.is_user_name()
    }

    /// Parse this field as inert typed user-identity metadata.
    ///
    /// Returns `Ok(None)` for fields other than `USERADDRESS`, `USERINITIALS`, and
    /// `USERNAME`. The result exposes only stored override, formatting,
    /// cached-content, and dirty/lock metadata; it never reads or modifies a
    /// host user's identity or refreshes a field.
    pub fn user_identity_field(&self) -> Result<Option<UserIdentity>> {
        UserIdentity::from_field(self)
    }

    /// Check whether this is an `ADVANCE` placement field.
    ///
    /// Recognition is limited to the stored field instruction. It never moves
    /// text, changes layout, reflows content, or refreshes a cached result.
    pub fn is_advance_field(&self) -> bool {
        field_instruction_remainder(&self.instruction, "ADVANCE").is_some()
    }

    /// Parse this field as inert typed `ADVANCE` placement metadata.
    ///
    /// Returns `Ok(None)` for fields other than `ADVANCE`. The returned
    /// values expose only stored point adjustments, cached content, and
    /// dirty/lock state. This method never moves text, changes layout, reflows
    /// content, or refreshes a field.
    pub fn advance_field(&self) -> Result<Option<Advance>> {
        Advance::from_field(self)
    }
}
