//! Bounded, inert SpreadsheetML 2009/9 form-control properties.
//!
//! This module owns the XML payload of an Office 2010 `ctrlProp` part and a
//! bounded read-only worksheet owner that composes that leaf with its admitted
//! DrawingML/VML identity closure.  Package mutation, ActiveX ownership,
//! workbook transactions, and mirror-writing lifecycle remain outside this
//! batch.  Parsing and detached writing are source preserving and bounded so
//! the package owner can compose them without a second XML safety policy.

mod codec;
mod model;
mod owner;

pub use codec::{
    SourceProperties, SourceView, SourceWritable, insert_item, inspect, inspect_with_limits, parse,
    parse_with_limits, remove_item, replace_items, replace_scalar, set_scalar, write,
    write_with_limits,
};
pub use model::{
    Checked, ControlSelector, DropStyle, EditValidation, FormControl, FormControlDraft,
    FormControlFormula, Item, ItemList, KnownOrUnknown, NamespaceBinding, ObjectType,
    OpaqueAttribute, OpaqueXml, Properties, ScalarField, ScalarValue, SelectionType, TextHAlign,
    TextVAlign,
};
pub use owner::{
    FormControlCollection, FormControlDiagnostic, FormControlDiagnosticCode, FormControlOwnerError,
    FormControlPartRead, FormControlReadSet, FormControlView, MceProvenance, OwnerLimits,
    OwnerProfile, OwnerResult, RelationshipFingerprint, ShapeClosure, SourceBackedFormControlOwner,
};
pub(crate) use owner::{
    eager_form_controls_for_sheet, owner_to_xlsx, source_form_controls_for_sheet,
};
pub(crate) use owner::{
    eager_form_controls_for_sheet_with_limits, source_form_controls_for_sheet_with_limits,
};

/// The Office 2010 SpreadsheetML form-control-properties namespace.
pub const FORM_CONTROL_NAMESPACE: &str =
    "http://schemas.microsoft.com/office/spreadsheetml/2009/9/main";
/// The OPC content type owned by a control-properties part.
pub const CONTROL_PROPERTIES_CONTENT_TYPE: &str = "application/vnd.ms-excel.controlproperties+xml";
/// The worksheet relationship type which owns a control-properties part.
pub const CONTROL_PROPERTIES_RELATIONSHIP_TYPE: &str =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/ctrlProp";

/// The maximum source part accepted by the focused codec.
pub const MAX_PART_BYTES: usize = 16 * 1024 * 1024;
/// The maximum generated part accepted by the focused codec.
pub const MAX_GENERATED_BYTES: usize = 32 * 1024 * 1024;
/// The maximum nesting depth inspected by the focused codec.
pub const MAX_XML_DEPTH: usize = 256;
/// The maximum number of XML events inspected for one part.
pub const MAX_XML_EVENTS: usize = 400_000;
/// The maximum number of list items in one properties part.
pub const MAX_ITEMS: usize = 65_536;
/// The maximum source/decoded size of one item value.
pub const MAX_ITEM_VALUE_BYTES: usize = 1024 * 1024;
/// The maximum authored formula/reference lexical size.
pub const MAX_FORMULA_BYTES: usize = 8192;
/// The maximum attributes retained on one XML element.
pub const MAX_ATTRIBUTES: usize = 256;
/// The maximum opaque extension bytes retained by one properties part.
pub const MAX_OPAQUE_BYTES: usize = 16 * 1024 * 1024;
/// The maximum aggregate bytes retained by one parsed/detached model.
pub const MAX_RETAINED_BYTES: usize = MAX_PART_BYTES + MAX_OPAQUE_BYTES;

/// A bounded local policy for this leaf XML codec.
///
/// Package owners should derive this policy from their retained OPC limits.
/// The setters only lower the hard ceilings; they cannot raise the safety
/// limits above the constants exported by this module.
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub struct Limits {
    max_part_bytes: usize,
    max_output_bytes: usize,
    max_depth: usize,
    max_events: usize,
    max_items: usize,
    max_item_value_bytes: usize,
    max_attributes: usize,
    max_opaque_bytes: usize,
    max_retained_bytes: usize,
}

impl Default for Limits {
    fn default() -> Self {
        Self {
            max_part_bytes: MAX_PART_BYTES,
            max_output_bytes: MAX_GENERATED_BYTES,
            max_depth: MAX_XML_DEPTH,
            max_events: MAX_XML_EVENTS,
            max_items: MAX_ITEMS,
            max_item_value_bytes: MAX_ITEM_VALUE_BYTES,
            max_attributes: MAX_ATTRIBUTES,
            max_opaque_bytes: MAX_OPAQUE_BYTES,
            max_retained_bytes: MAX_RETAINED_BYTES,
        }
    }
}

impl Limits {
    /// Create the standard bounded codec policy.
    #[must_use]
    pub const fn new() -> Self {
        Self {
            max_part_bytes: MAX_PART_BYTES,
            max_output_bytes: MAX_GENERATED_BYTES,
            max_depth: MAX_XML_DEPTH,
            max_events: MAX_XML_EVENTS,
            max_items: MAX_ITEMS,
            max_item_value_bytes: MAX_ITEM_VALUE_BYTES,
            max_attributes: MAX_ATTRIBUTES,
            max_opaque_bytes: MAX_OPAQUE_BYTES,
            max_retained_bytes: MAX_RETAINED_BYTES,
        }
    }

    /// Set a lower source-part ceiling.
    #[must_use]
    pub fn with_max_part_bytes(mut self, value: usize) -> Self {
        self.max_part_bytes = value.min(MAX_PART_BYTES);
        self
    }

    /// Set a lower candidate-output ceiling.
    #[must_use]
    pub fn with_max_output_bytes(mut self, value: usize) -> Self {
        self.max_output_bytes = value.min(MAX_GENERATED_BYTES);
        self
    }

    /// Set a lower XML-depth ceiling.
    #[must_use]
    pub fn with_max_depth(mut self, value: usize) -> Self {
        self.max_depth = value.min(MAX_XML_DEPTH);
        self
    }

    /// Set a lower XML-event ceiling.
    #[must_use]
    pub fn with_max_events(mut self, value: usize) -> Self {
        self.max_events = value.min(MAX_XML_EVENTS);
        self
    }

    /// Set a lower item-count ceiling.
    #[must_use]
    pub fn with_max_items(mut self, value: usize) -> Self {
        self.max_items = value.min(MAX_ITEMS);
        self
    }

    /// Set a lower item-value ceiling.
    #[must_use]
    pub fn with_max_item_value_bytes(mut self, value: usize) -> Self {
        self.max_item_value_bytes = value.min(MAX_ITEM_VALUE_BYTES);
        self
    }

    /// Set a lower per-element attribute ceiling.
    #[must_use]
    pub fn with_max_attributes(mut self, value: usize) -> Self {
        self.max_attributes = value.min(MAX_ATTRIBUTES);
        self
    }

    /// Set a lower opaque-extension ceiling.
    #[must_use]
    pub fn with_max_opaque_bytes(mut self, value: usize) -> Self {
        self.max_opaque_bytes = value.min(MAX_OPAQUE_BYTES);
        self
    }

    /// Set a lower aggregate retained-memory ceiling.
    #[must_use]
    pub fn with_max_retained_bytes(mut self, value: usize) -> Self {
        self.max_retained_bytes = value.min(MAX_RETAINED_BYTES);
        self
    }

    pub(crate) fn max_part_bytes(self) -> usize {
        self.max_part_bytes
    }

    pub(crate) fn max_output_bytes(self) -> usize {
        self.max_output_bytes.min(MAX_GENERATED_BYTES)
    }

    pub(crate) fn max_depth(self) -> usize {
        self.max_depth.min(MAX_XML_DEPTH)
    }

    pub(crate) fn max_events(self) -> usize {
        self.max_events.min(MAX_XML_EVENTS)
    }

    pub(crate) fn max_items(self) -> usize {
        self.max_items.min(MAX_ITEMS)
    }

    pub(crate) fn max_item_value_bytes(self) -> usize {
        self.max_item_value_bytes.min(MAX_ITEM_VALUE_BYTES)
    }

    pub(crate) fn max_attributes(self) -> usize {
        self.max_attributes.min(MAX_ATTRIBUTES)
    }

    pub(crate) fn max_opaque_bytes(self) -> usize {
        self.max_opaque_bytes.min(MAX_OPAQUE_BYTES)
    }

    pub(crate) fn max_retained_bytes(self) -> usize {
        self.max_retained_bytes.min(MAX_RETAINED_BYTES)
    }
}

/// A bounded failure from the form-control leaf codec.
#[derive(Debug, thiserror::Error)]
#[non_exhaustive]
pub enum FormControlError {
    /// The caller's execution policy refused parsing or was cancelled.
    #[error(transparent)]
    Execution(#[from] litchi_core::ExecutionError),
    /// The XML or typed model violates the admitted schema slice.
    #[error("invalid form-control properties: {0}")]
    Invalid(String),
    /// The source or candidate exceeds one of the retained safety ceilings.
    #[error("form-control properties {resource} exceeds {maximum} (observed {observed})")]
    Limit {
        /// Bounded resource name.
        resource: &'static str,
        /// Observed resource use.
        observed: usize,
        /// Caller/effective maximum.
        maximum: usize,
    },
    /// A fallible source or candidate reservation failed.
    #[error("could not reserve memory for form-control properties {resource}: {source}")]
    Allocation {
        /// Allocation subject.
        resource: &'static str,
        /// The allocator's recoverable reservation error.
        #[source]
        source: std::collections::TryReserveError,
    },
}

/// Result type used by the focused form-control API.
pub type Result<T> = std::result::Result<T, FormControlError>;

pub(crate) fn invalid(message: impl Into<String>) -> FormControlError {
    FormControlError::Invalid(message.into())
}

pub(crate) fn limit(resource: &'static str, observed: usize, maximum: usize) -> FormControlError {
    FormControlError::Limit {
        resource,
        observed,
        maximum,
    }
}

pub(crate) fn allocation(
    resource: &'static str,
    source: std::collections::TryReserveError,
) -> FormControlError {
    FormControlError::Allocation { resource, source }
}
