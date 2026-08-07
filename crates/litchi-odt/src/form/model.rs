//! Semantic model for inert, classic `OpenDocument` form controls.

use litchi_core::Result;

#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
#[non_exhaustive]
pub enum Part {
    Content,
    Styles,
    Flat,
}

#[derive(Debug, Clone, PartialEq, Eq, Hash)]
#[non_exhaustive]
pub enum Scope {
    Document,
    Text,
    Sheet { index: usize, name: Option<String> },
    DrawPage { index: usize, name: Option<String> },
    Notes { index: usize },
    MasterPage { index: usize, name: Option<String> },
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Attribute {
    pub namespace_uri: Option<String>,
    pub local_name: String,
    pub value: String,
}

#[derive(Debug, Clone, Default, PartialEq)]
pub struct Forms {
    pub groups: Vec<Group>,
    pub control_shapes: Vec<Shape>,
    /// Event declarations in document order. They are retained but never executed.
    pub event_listeners: Vec<Listener>,
    pub has_xforms: bool,
    pub has_event_listeners: bool,
}

/// Stable identity snapshot for the element that owns an event declaration.
#[derive(Debug, Clone, PartialEq)]
#[non_exhaustive]
pub enum Target {
    Forms {
        part: Part,
        scope: Scope,
    },
    Form {
        xml_id: Option<String>,
        form_id: Option<String>,
        name: Option<String>,
    },
    Control {
        kind: ControlKind,
        xml_id: Option<String>,
        form_id: Option<String>,
        name: Option<String>,
    },
}

/// `XLink` activation mode retained for an event declaration.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum Actuate {
    OnLoad,
    OnRequest,
    Other,
    None,
}

/// An inert `script:event-listener` declaration.
#[derive(Debug, Clone, PartialEq)]
pub struct Listener {
    pub target: Target,
    /// Required by conforming ODF, but optional here for tolerant producer inspection.
    pub event_name: Option<String>,
    /// Required by conforming ODF, but optional here for tolerant producer inspection.
    pub language: Option<String>,
    pub macro_name: Option<String>,
    /// URI retained verbatim; it is never fetched or resolved.
    pub href: Option<String>,
    pub actuate: Option<Actuate>,
    /// Whether the optional `XLink` type was explicitly `simple`.
    pub simple_link: bool,
    pub attributes: Vec<Attribute>,
}

#[derive(Debug, Clone, PartialEq)]
pub struct Group {
    pub part: Part,
    pub scope: Scope,
    pub automatic_focus: Option<bool>,
    pub apply_design_mode: Option<bool>,
    pub forms: Vec<Form>,
    pub has_xforms: bool,
    pub has_event_listeners: bool,
    pub attributes: Vec<Attribute>,
}

#[derive(Debug, Clone, PartialEq)]
pub struct Form {
    pub xml_id: Option<String>,
    pub form_id: Option<String>,
    pub name: Option<String>,
    pub control_implementation: Option<String>,
    pub properties: Vec<Property>,
    pub children: Vec<Node>,
    pub attributes: Vec<Attribute>,
}

#[derive(Debug, Clone, PartialEq)]
#[allow(clippy::large_enum_variant)] // public API; boxing would break callers
pub enum Node {
    Form(Form),
    Control(Control),
}

#[derive(Debug, Clone, PartialEq, Eq)]
#[non_exhaustive]
pub enum ControlKind {
    Text,
    TextArea,
    Password,
    File,
    FormattedText,
    Number,
    Date,
    Time,
    FixedText,
    ComboBox,
    Item,
    ListBox,
    Option,
    Button,
    Image,
    CheckBox,
    Radio,
    Frame,
    ImageFrame,
    Hidden,
    Grid,
    Column,
    ValueRange,
    GenericControl,
    Other(String),
}

impl ControlKind {
    pub(super) fn parse(name: &str) -> Option<Self> {
        Some(match name {
            "text" => Self::Text,
            "textarea" => Self::TextArea,
            "password" => Self::Password,
            "file" => Self::File,
            "formatted-text" => Self::FormattedText,
            "number" => Self::Number,
            "date" => Self::Date,
            "time" => Self::Time,
            "fixed-text" => Self::FixedText,
            "combobox" => Self::ComboBox,
            "item" => Self::Item,
            "listbox" => Self::ListBox,
            "option" => Self::Option,
            "button" => Self::Button,
            "image" => Self::Image,
            "checkbox" => Self::CheckBox,
            "radio" => Self::Radio,
            "frame" => Self::Frame,
            "image-frame" => Self::ImageFrame,
            "hidden" => Self::Hidden,
            "grid" => Self::Grid,
            "column" => Self::Column,
            "value-range" => Self::ValueRange,
            "generic-control" => Self::GenericControl,
            "form"
            | "properties"
            | "property"
            | "list-property"
            | "list-value"
            | "connection-resource" => return None,
            other => Self::Other(other.to_string()),
        })
    }
}

#[derive(Debug, Clone, PartialEq)]
pub struct Control {
    pub kind: ControlKind,
    pub xml_id: Option<String>,
    pub form_id: Option<String>,
    pub name: Option<String>,
    pub label: Option<String>,
    pub title: Option<String>,
    pub value: Option<String>,
    pub current_value: Option<String>,
    pub state: Option<String>,
    pub current_state: Option<String>,
    pub linked_cell: Option<String>,
    pub source_cell_range: Option<String>,
    pub image_data: Option<String>,
    pub disabled: Option<bool>,
    pub read_only: Option<bool>,
    pub selected: Option<bool>,
    pub current_selected: Option<bool>,
    pub text: String,
    pub properties: Vec<Property>,
    pub children: Vec<Node>,
    pub attributes: Vec<Attribute>,
}

#[derive(Debug, Clone, PartialEq)]
pub struct Property {
    pub name: String,
    pub value: PropertyValue,
}

impl Property {
    pub fn boolean(name: impl Into<String>, value: bool) -> Self {
        Self {
            name: name.into(),
            value: PropertyValue::Scalar(ScalarValue::Boolean(value)),
        }
    }

    pub fn number(
        name: impl Into<String>,
        value_type: impl Into<String>,
        lexical: impl Into<String>,
        currency: Option<String>,
    ) -> Self {
        Self {
            name: name.into(),
            value: PropertyValue::Scalar(ScalarValue::Number {
                value_type: value_type.into(),
                lexical: lexical.into(),
                currency,
            }),
        }
    }

    pub fn text(name: impl Into<String>, value: impl Into<String>) -> Self {
        Self {
            name: name.into(),
            value: PropertyValue::Scalar(ScalarValue::Text(value.into())),
        }
    }

    pub fn date(name: impl Into<String>, value: impl Into<String>) -> Self {
        Self {
            name: name.into(),
            value: PropertyValue::Scalar(ScalarValue::Date(value.into())),
        }
    }

    pub fn time(name: impl Into<String>, value: impl Into<String>) -> Self {
        Self {
            name: name.into(),
            value: PropertyValue::Scalar(ScalarValue::Time(value.into())),
        }
    }

    pub fn void(name: impl Into<String>) -> Self {
        Self {
            name: name.into(),
            value: PropertyValue::Scalar(ScalarValue::Void),
        }
    }

    pub fn list(
        name: impl Into<String>,
        value_type: impl Into<String>,
        values: Vec<ScalarValue>,
    ) -> Self {
        Self {
            name: name.into(),
            value: PropertyValue::List {
                value_type: Some(value_type.into()),
                values,
            },
        }
    }

    pub fn to_xml_fragment(&self) -> Result<String> {
        super::property_xml(self)
    }
}

#[derive(Debug, Clone, PartialEq)]
pub enum PropertyValue {
    Scalar(ScalarValue),
    List {
        value_type: Option<String>,
        values: Vec<ScalarValue>,
    },
}

#[derive(Debug, Clone, PartialEq)]
#[non_exhaustive]
pub enum ScalarValue {
    Boolean(bool),
    Number {
        value_type: String,
        lexical: String,
        currency: Option<String>,
    },
    Text(String),
    Date(String),
    Time(String),
    Void,
    Other {
        value_type: String,
        lexical: Option<String>,
    },
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct ControlRef {
    pub group_index: usize,
    pub form_index: usize,
    pub node_path: Vec<usize>,
}

#[derive(Debug, Clone, PartialEq)]
pub struct Shape {
    pub part: Part,
    pub scope: Scope,
    pub control_id: String,
    pub resolved_control: Option<ControlRef>,
    pub draw_name: Option<String>,
    pub style_name: Option<String>,
    pub text_style_name: Option<String>,
    pub z_index: Option<u32>,
    pub x: Option<String>,
    pub y: Option<String>,
    pub width: Option<String>,
    pub height: Option<String>,
    pub attributes: Vec<Attribute>,
}
