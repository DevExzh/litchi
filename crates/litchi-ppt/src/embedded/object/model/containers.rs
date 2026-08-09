//! Typed object-container models and the ergonomic external-object enum.

use super::metadata::{EmbedPreferences, LinkInfo, Metadata};

/// Container-specific metadata preceding the shared OLE object atom.
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub enum ContainerKind {
    Embedded(EmbedPreferences),
    Linked(LinkInfo),
}

/// A strict, inert `ExOleEmbedContainer` or `ExOleLinkContainer`.
#[derive(Clone, Debug, PartialEq, Eq)]
pub struct Definition {
    pub kind: ContainerKind,
    pub object: Metadata,
    pub menu_name: Option<String>,
    pub program_id: Option<String>,
    pub clipboard_name: Option<String>,
    /// Opaque icon bytes. They are retained but never decoded or rendered here.
    pub metafile: Option<Vec<u8>>,
}

/// Inert metadata for an `ExControlContainer` `ActiveX` definition.
#[derive(Clone, Debug, PartialEq, Eq)]
pub struct Control {
    pub slide_id: Option<u32>,
    pub object: Metadata,
    pub menu_name: Option<String>,
    pub program_id: Option<String>,
    pub clipboard_name: Option<String>,
    /// Opaque icon bytes. Control storage is not loaded or executed.
    pub metafile: Option<Vec<u8>>,
}

/// Strict embedded and linked OLE definitions in document order.
#[derive(Clone, Debug, PartialEq, Eq)]
pub enum ExternalObject {
    Object(Definition),
    ActiveXControl(Control),
}
