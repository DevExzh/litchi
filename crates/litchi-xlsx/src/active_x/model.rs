//! ActiveX and worksheet-control value models.

use litchi_opc::PackURI;

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum Persistence {
    PropertyBag,
    Stream,
    StreamInit,
    Storage,
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Marker {
    pub column: i32,
    pub column_offset: i64,
    pub row: i32,
    pub row_offset: i64,
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct ObjectAnchor {
    pub from: Marker,
    pub to: Marker,
    pub move_with_cells: Option<bool>,
    pub size_with_cells: Option<bool>,
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct ControlProperties {
    pub anchor: ObjectAnchor,
    pub locked: Option<bool>,
    pub default_size: Option<bool>,
    pub print: Option<bool>,
    pub disabled: Option<bool>,
    pub recalc_always: Option<bool>,
    pub ui_object: Option<bool>,
    pub auto_fill: Option<bool>,
    pub auto_line: Option<bool>,
    pub auto_picture: Option<bool>,
    pub macro_name: Option<String>,
    pub alternate_text: Option<String>,
    pub preview_relationship_id: Option<String>,
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Control {
    pub shape_id: u32,
    pub relationship_id: String,
    pub name: Option<String>,
    pub properties: Option<ControlProperties>,
}

#[derive(Debug, Clone, Default, PartialEq, Eq)]
pub struct Controls {
    pub controls: Vec<Control>,
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Picture {
    pub relationship_id: Option<String>,
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Font {
    pub persistence: Option<Persistence>,
    pub relationship_id: Option<String>,
    pub properties: Vec<Property>,
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub enum PropertyObject {
    Font(Font),
    Picture(Picture),
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Property {
    pub name: String,
    pub value: Option<String>,
    pub object: Option<PropertyObject>,
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Descriptor {
    pub class_id: String,
    pub license: Option<String>,
    pub persistence: Persistence,
    pub relationship_id: Option<String>,
    pub properties: Vec<Property>,
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Binary {
    pub relationship_id: String,
    pub part_uri: PackURI,
    pub bytes: Vec<u8>,
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct PreviewImage {
    pub relationship_id: String,
    pub part_uri: PackURI,
    pub content_type: String,
    pub bytes: Vec<u8>,
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct LoadedControl {
    pub control: Control,
    pub descriptor_uri: PackURI,
    pub descriptor: Descriptor,
    pub binaries: Vec<Binary>,
    pub preview: Option<PreviewImage>,
}

#[derive(Debug, Clone, Default, PartialEq, Eq)]
pub struct ControlSet {
    pub controls: Vec<LoadedControl>,
}
