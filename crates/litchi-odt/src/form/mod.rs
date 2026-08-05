//! Inert, bounded semantic inventory of classic OpenDocument forms.
//!
//! The owner is intentionally layered: [`model`] contains the public semantic
//! vocabulary, [`codec`] owns bounded XML inspection, and the focused writing
//! modules provide typed authoring facades for individual control families.

mod codec;
mod model;

#[cfg(test)]
mod tests;

mod connection_resource_writing;
mod control_writing;
mod generic_writing;
mod grid_writing;
mod image_frame_writing;
mod input_writing;
mod interactive_writing;
mod selection_writing;
mod typed_value_writing;
mod value_range_writing;
mod visual_writing;
mod writing;

pub use model::*;

pub(crate) use codec::parse_form_parts;
pub(crate) use writing::property_xml;

pub use connection_resource_writing::{
    ConnectionResourceForm, FormConnectionResource, OwnedFormConnectionResource,
    form_connection_resources, insert_form_connection_resource_xml,
    remove_form_connection_resource_xml, replace_form_connection_resource_xml,
};
pub use control_writing::{
    ControlForm, TextControl, TextControlKind, insert_text_control_xml, remove_text_control_xml,
    replace_text_control_xml, text_controls,
};
pub use generic_writing::{
    FixedTextControl, GenericControl, GenericControlMetadata, GenericForm, GenericFormControl,
    HiddenControl, generic_form_controls, insert_generic_form_control_xml,
    remove_generic_form_control_xml, replace_generic_form_control_xml,
};
pub use grid_writing::{
    GridColumn, GridColumnControl, GridColumnControlKind, GridControl, GridForm,
    GridNonNegativeInteger, grid_controls, insert_grid_control_xml, remove_grid_control_xml,
    replace_grid_control_xml,
};
pub use image_frame_writing::{
    ImageFrameControl, ImageFrameForm, image_frame_controls, insert_image_frame_control_xml,
    remove_image_frame_control_xml, replace_image_frame_control_xml,
};
pub use input_writing::{
    FileControl, PasswordControl, PasswordFileControl, PasswordFileForm,
    insert_password_file_control_xml, password_file_controls, remove_password_file_control_xml,
    replace_password_file_control_xml,
};
pub use interactive_writing::{
    ButtonControl, ButtonType, CheckboxControl, CheckboxState, InteractiveControl, InteractiveForm,
    insert_interactive_control_xml, interactive_controls, remove_interactive_control_xml,
    replace_interactive_control_xml,
};
pub use selection_writing::{
    ComboItem, ComboboxControl, ListLinkageType, ListOption, ListSourceType, ListboxControl,
    SelectionControl, SelectionForm, insert_selection_control_xml, remove_selection_control_xml,
    replace_selection_control_xml, selection_controls,
};
pub use typed_value_writing::{
    FormDate, FormDouble, TypedValueBound, TypedValueControl, TypedValueControlKind,
    TypedValueDuration, TypedValueForm, TypedValueNonNegativeInteger,
    insert_typed_value_control_xml, remove_typed_value_control_xml,
    replace_typed_value_control_xml, typed_value_controls,
};
pub use value_range_writing::{
    ValueRangeControl, ValueRangeDuration, ValueRangeForm, ValueRangeInteger,
    ValueRangeNonNegativeInteger, ValueRangeOrientation, ValueRangePositiveInteger,
    insert_value_range_control_xml, remove_value_range_control_xml,
    replace_value_range_control_xml, value_range_controls,
};
pub use visual_writing::{
    FrameControl, ImageButtonType, ImageControl, RadioControl, RadioVisualEffect,
    RelativeImageAlign, RelativeImagePosition, VisualControl, VisualForm,
    insert_visual_control_xml, remove_visual_control_xml, replace_visual_control_xml,
    visual_controls,
};
pub use writing::{
    PropertyForm, form_properties, insert_form_property_xml, remove_form_property_xml,
    replace_form_property_xml,
};

pub(super) const OFFICE: &str = "urn:oasis:names:tc:opendocument:xmlns:office:1.0";
pub(super) const FORM: &str = "urn:oasis:names:tc:opendocument:xmlns:form:1.0";
pub(super) const DRAW: &str = "urn:oasis:names:tc:opendocument:xmlns:drawing:1.0";
pub(super) const TABLE: &str = "urn:oasis:names:tc:opendocument:xmlns:table:1.0";
pub(super) const PRESENTATION: &str = "urn:oasis:names:tc:opendocument:xmlns:presentation:1.0";
pub(super) const STYLE: &str = "urn:oasis:names:tc:opendocument:xmlns:style:1.0";
pub(super) const SVG: &str = "urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0";
pub(super) const XML: &str = "http://www.w3.org/XML/1998/namespace";
pub(super) const XFORMS: &str = "http://www.w3.org/2002/xforms";
pub(super) const SCRIPT: &str = "urn:oasis:names:tc:opendocument:xmlns:script:1.0";
pub(super) const XLINK: &str = "http://www.w3.org/1999/xlink";
pub(super) const MAX_RAW: usize = 64 * 1024 * 1024;
pub(super) const MAX_DECODED: usize = 16 * 1024 * 1024;
pub(super) const MAX_SCALAR: usize = 64 * 1024;
pub(super) const MAX_TEXT: usize = 4 * 1024 * 1024;
pub(super) const MAX_NODES: usize = 65_536;
pub(super) const MAX_SHAPES: usize = 65_536;
pub(super) const MAX_DEPTH: usize = 128;
pub(super) const MAX_ATTRIBUTES: usize = 256;
