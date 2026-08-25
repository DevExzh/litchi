//! Neutral strict Numbers interactive-cell control codec.
//!
//! The implementation is shared with the popup-menu codec so CellSpec wire
//! parsing, fixed64 range validation, unknown-field policy, prepared execution
//! limits, and control-list refcount authority remain one audited route. This module is the stable
//! hidden seam for package owners; no generated Buffa/Prost values or native
//! archive identifiers cross it.

#![allow(clippy::module_name_repetitions)]

pub use crate::numbers_table_cell_pop_up_menu_codec::{
    CHECKBOX_INTERACTION_TYPE, ControlCellSpecSnapshot, ControlFormatSnapshot, ControlFormatWrite,
    DecodeError, DecodeLimit, DecodeOptions, DecodeReport, RewriteExecutionLimits,
    RewriteExecutionRequirements, RewriteOutput, SLIDER_INTERACTION_TYPE,
    STAR_RATING_INTERACTION_TYPE, STEPPER_INTERACTION_TYPE,
};

pub use crate::numbers_table_cell_pop_up_menu_codec::{
    PreparedControlCellSpecWrite, PreparedControlCellSpecWrite as PreparedCellSpecWrite,
    PreparedControlFormatWrite, PreparedControlFormatWrite as PreparedFormatWrite,
};

/// The strict projection selected for one `CellSpecArchive` payload.
///
/// Popup menus and the four scalar controls share the native envelope, but
/// their known fields are deliberately disjoint.  Keeping the dispatch enum
/// here lets table-storage callers validate a mixed control list without
/// treating non-popup range fields as opaque (or routing them through the
/// popup-only validator).
#[derive(Debug, Clone, Copy, PartialEq)]
pub enum CellSpecSnapshot<'source> {
    /// A popup-menu interaction (`interaction_type == 7`).
    Popup(crate::numbers_table_cell_pop_up_menu_codec::CellSpecSnapshot<'source>),
    /// A checkbox, star-rating, slider, or stepper interaction.
    Control(ControlCellSpecSnapshot<'source>),
}

impl<'source> CellSpecSnapshot<'source> {
    /// Borrow the original payload without normalizing unknown fields.
    #[must_use]
    pub const fn raw(self) -> &'source [u8] {
        match self {
            Self::Popup(snapshot) => snapshot.raw(),
            Self::Control(snapshot) => snapshot.raw(),
        }
    }

    /// Return the native interaction discriminator.
    #[must_use]
    pub const fn interaction_type(self) -> u32 {
        match self {
            Self::Popup(snapshot) => snapshot.interaction_type(),
            Self::Control(snapshot) => snapshot.interaction_type(),
        }
    }

    /// Whether this projection is the popup-menu route.
    #[must_use]
    pub const fn is_popup(self) -> bool {
        matches!(self, Self::Popup(_))
    }
}

/// Strictly decode either a popup or one of the neutral scalar-control
/// `CellSpecArchive` variants.
///
/// The control parser is attempted first so interactions 4, 5, 6, and 8
/// receive range/refusal validation.  A popup payload is then attempted when
/// that parser rejects the popup discriminator.  A typed limit is retained
/// only when both projections fail; this prevents a failed speculative parse
/// from masking a valid alternate projection while preserving max-minus-one
/// behavior for malformed or oversized payloads.
pub fn decode_cell_spec_with_report(
    source: &[u8],
    options: DecodeOptions,
) -> Result<(CellSpecSnapshot<'_>, DecodeReport), DecodeError> {
    let control_error = match decode_control_cell_spec_with_report(source, options) {
        Ok((snapshot, report)) => return Ok((CellSpecSnapshot::Control(snapshot), report)),
        Err(error) => error,
    };
    match crate::numbers_table_cell_pop_up_menu_codec::decode_cell_spec_with_report(source, options)
    {
        Ok((snapshot, report)) => Ok((CellSpecSnapshot::Popup(snapshot), report)),
        Err(popup_error) => {
            if control_error.resource_limit().is_some() {
                Err(control_error)
            } else {
                Err(popup_error)
            }
        },
    }
}

/// Compatibility spelling for storage and package owners that want to make
/// the projection-vs-popup distinction explicit at the call site.
pub use decode_cell_spec_with_report as decode_any_cell_spec_with_report;

pub use crate::numbers_table_cell_pop_up_menu_codec::{
    canonical_control_cell_spec, canonical_control_cell_spec as canonical_cell_spec,
    canonical_control_format, canonical_control_format as canonical_format,
    canonical_control_format_fields, canonical_control_format_fields as canonical_format_fields,
    decode_control_cell_spec, decode_control_cell_spec_with_report, decode_control_format,
    decode_control_format_with_report, prepare_control_cell_spec_write,
    prepare_control_cell_spec_write as prepare_cell_spec_write, prepare_control_format_write,
    prepare_control_format_write as prepare_format_write,
    prepare_control_format_write_fields as prepare_format_write_fields, rewrite_control_cell_spec,
    rewrite_control_cell_spec as rewrite_cell_spec, rewrite_control_format,
    rewrite_control_format as rewrite_format,
    rewrite_control_format_fields as rewrite_format_fields,
};
