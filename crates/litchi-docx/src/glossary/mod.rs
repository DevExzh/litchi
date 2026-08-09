#![expect(
    clippy::arbitrary_source_item_ordering,
    reason = "items remain grouped by OOXML schema family and package lifecycle"
)]
//! Layered `WordprocessingML` glossary / `AutoText` catalog owner.

mod codec;
mod graph;
mod model;
mod package;
mod transaction;

#[cfg(test)]
mod tests;

pub(super) use crate::{Error, Result};
pub(super) use caseless::Caseless;
pub(super) use litchi_opc::constants::content_type as ct;
pub(super) use litchi_opc::part::{BlobPart, Part};
pub(super) use litchi_opc::{ContentType, OpcPackage, PackURI};
pub(super) use quick_xml::{
    Reader, XmlVersion,
    encoding::Decoder,
    events::{BytesStart, Event},
};
pub(super) use std::collections::{HashMap, HashSet, VecDeque};
pub(super) use std::sync::Arc;
pub(super) use unicode_normalization::UnicodeNormalization;

const W: &str = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
const WS: &str = "http://purl.oclc.org/ooxml/wordprocessingml/main";
const R: &str = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const RS: &str = "http://purl.oclc.org/ooxml/officeDocument/relationships";
const MC: &str = "http://schemas.openxmlformats.org/markup-compatibility/2006";
const VML: &str = "urn:schemas-microsoft-com:vml";
const REL: &str =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/glossaryDocument";
const STRICT_REL: &str = "http://purl.oclc.org/ooxml/officeDocument/relationships/glossaryDocument";
const STYLES_EFFECTS_REL: &str =
    "http://schemas.microsoft.com/office/2007/relationships/stylesWithEffects";
const CUSTOMIZATIONS_REL: &str =
    "http://schemas.microsoft.com/office/2006/relationships/keyMapCustomizations";
const CUSTOMIZATIONS_CT: &str = "application/vnd.ms-word.keyMapCustomizations+xml";
const ATTACHED_TOOLBARS_REL: &str =
    "http://schemas.microsoft.com/office/2006/relationships/attachedToolbars";
const ATTACHED_TOOLBARS_CT: &str = "application/vnd.ms-word.attachedToolbars";
const DIAGRAM_DRAWING_REL: &str =
    "http://schemas.microsoft.com/office/2007/relationships/diagramDrawing";
const CHART_STYLE_REL: &str = "http://schemas.microsoft.com/office/2011/relationships/chartStyle";
const CHART_COLOR_STYLE_REL: &str =
    "http://schemas.microsoft.com/office/2011/relationships/chartColorStyle";
const CHART_STYLE_REL_2012: &str =
    "http://schemas.microsoft.com/office/2012/relationships/chartStyle";
const CHART_COLOR_STYLE_REL_2012: &str =
    "http://schemas.microsoft.com/office/2012/relationships/chartColorStyle";
const ACTIVE_X_BINARY_REL: &str =
    "http://schemas.microsoft.com/office/2006/relationships/activeXControlBinary";
const STYLES_EFFECTS_CT: &str = "application/vnd.ms-word.stylesWithEffects+xml";
const CHART_STYLE_CT: &str = "application/vnd.ms-office.chartstyle+xml";
const CHART_COLOR_STYLE_CT: &str = "application/vnd.ms-office.chartcolorstyle+xml";
const ACTIVE_X_DESCRIPTOR_CT: &str = "application/vnd.ms-office.activeX+xml";
const ACTIVE_X_BINARY_CT: &str = "application/vnd.ms-office.activeX";
const RECIPIENT_DATA_CT: &str =
    "application/vnd.openxmlformats-officedocument.wordprocessingml.mailMergeRecipientData+xml";
const OBFUSCATED_FONT_CT: &str = "application/vnd.openxmlformats-officedocument.obfuscatedFont";
const FONT_DATA_CT: &str = "application/x-fontdata";
const FONT_TTF_CT: &str = "application/x-font-ttf";
const PRINTER_SETTINGS_CT: &str =
    "application/vnd.openxmlformats-officedocument.wordprocessingml.printerSettings";
const CT: &str =
    "application/vnd.openxmlformats-officedocument.wordprocessingml.document.glossary+xml";
const MAX: usize = 32 * 1024 * 1024;
const MAX_DEPTH: usize = 256;
const MAX_NODES: usize = 262_144;
const MAX_DOM_ATTRIBUTES: usize = 262_144;
const MAX_DOM_CONTENT: usize = 262_144;
const MAX_DOM_TOKENS: usize = 1_000_000;
const MAX_DOM_BYTES: usize = 64 * 1024 * 1024;
const MAX_PARTS: usize = 100_000;
const MAX_VALUES: usize = 4096;
const MAX_STRING: usize = 1024 * 1024;
const MAX_NAME_KEY: usize = 4 * MAX_STRING;
/// Conservative aggregate ceiling for inert glossary-owned auxiliary payloads.
const MAX_GRAPH_BYTES: usize = 256 * 1024 * 1024;
/// Aggregate ceiling for raw graph names, content types, and relationship metadata.
const MAX_GRAPH_METADATA_BYTES: usize = 32 * 1024 * 1024;

pub use codec::{read, write};
pub use graph::raw;
pub use model::{Catalog, Category, Conformance, Entry, Gallery, Id, Insert, Kind, Name, Props};
pub use package::{load, put, remove};
pub use transaction::{Commit, Patch, Revision, Snapshot, Transaction};
