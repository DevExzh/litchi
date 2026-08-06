//! Opaque ActiveX payloads and the loaded worksheet graph.

use super::{Control, Descriptor};
use litchi_opc::PackURI;

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
