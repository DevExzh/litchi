use litchi_opc::PackURI;

/// Whether an OLE shape declares an embedded payload or an external link.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum Mode {
    Embedded,
    Linked,
}

/// OPC payload family declared by an OLE relationship.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum Kind {
    OleObject,
    Package,
}

/// Inert internal or external target metadata.
#[derive(Debug, Clone, PartialEq, Eq)]
pub enum Target {
    Internal {
        part_name: PackURI,
        content_type: String,
        relationship_type: String,
    },
    External {
        target: String,
        relationship_type: String,
    },
}

impl Target {
    #[must_use]
    pub fn relationship_type(&self) -> &str {
        match self {
            Self::Internal {
                relationship_type, ..
            }
            | Self::External {
                relationship_type, ..
            } => relationship_type,
        }
    }

    #[must_use]
    pub fn part_name(&self) -> Option<&PackURI> {
        match self {
            Self::Internal { part_name, .. } => Some(part_name),
            Self::External { .. } => None,
        }
    }

    #[must_use]
    pub fn content_type(&self) -> Option<&str> {
        match self {
            Self::Internal { content_type, .. } => Some(content_type),
            Self::External { .. } => None,
        }
    }

    #[must_use]
    pub fn external_target(&self) -> Option<&str> {
        match self {
            Self::Internal { .. } => None,
            Self::External { target, .. } => Some(target),
        }
    }
}

/// Bounded, inert metadata for one OLE graphic frame.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Object {
    pub(crate) slide_index: usize,
    pub(crate) index: usize,
    pub(crate) shape_id: Option<u32>,
    pub(crate) shape_name: Option<String>,
    pub(crate) legacy_shape_id: Option<String>,
    pub(crate) name: Option<String>,
    pub(crate) program_id: Option<String>,
    pub(crate) show_as_icon: Option<bool>,
    pub(crate) preview_width: Option<u32>,
    pub(crate) preview_height: Option<u32>,
    pub(crate) anchor: Option<Frame>,
    pub(crate) mode: Mode,
    pub(crate) relationship_id: Option<String>,
    pub(crate) kind: Option<Kind>,
    pub(crate) target: Option<Target>,
    pub(crate) preview_relationship_id: Option<String>,
}

impl Object {
    #[must_use]
    pub fn slide_index(&self) -> usize {
        self.slide_index
    }
    #[must_use]
    pub fn index(&self) -> usize {
        self.index
    }
    #[must_use]
    pub fn shape_id(&self) -> Option<u32> {
        self.shape_id
    }
    #[must_use]
    pub fn shape_name(&self) -> Option<&str> {
        self.shape_name.as_deref()
    }
    #[must_use]
    pub fn legacy_shape_id(&self) -> Option<&str> {
        self.legacy_shape_id.as_deref()
    }
    #[must_use]
    pub fn name(&self) -> Option<&str> {
        self.name.as_deref()
    }
    #[must_use]
    pub fn program_id(&self) -> Option<&str> {
        self.program_id.as_deref()
    }
    #[must_use]
    pub fn show_as_icon(&self) -> Option<bool> {
        self.show_as_icon
    }
    #[must_use]
    pub fn preview_width(&self) -> Option<u32> {
        self.preview_width
    }
    #[must_use]
    pub fn preview_height(&self) -> Option<u32> {
        self.preview_height
    }
    /// `DrawingML` position and extent of the owning graphic frame, in EMUs.
    #[must_use]
    pub fn anchor(&self) -> Option<Frame> {
        self.anchor
    }
    #[must_use]
    pub fn mode(&self) -> Mode {
        self.mode
    }
    #[must_use]
    pub fn relationship_id(&self) -> Option<&str> {
        self.relationship_id.as_deref()
    }
    #[must_use]
    pub fn kind(&self) -> Option<Kind> {
        self.kind
    }
    #[must_use]
    pub fn target(&self) -> Option<&Target> {
        self.target.as_ref()
    }
    #[must_use]
    pub fn preview_relationship_id(&self) -> Option<&str> {
        self.preview_relationship_id.as_deref()
    }
}

/// Position and extent of an authored OLE frame, in EMUs.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct Frame {
    pub x: i64,
    pub y: i64,
    pub cx: i64,
    pub cy: i64,
}

impl Frame {
    #[must_use]
    pub const fn new(x: i64, y: i64, cx: i64, cy: i64) -> Self {
        Self { x, y, cx, cy }
    }
}

/// Result of adding one inert OLE payload to a slide.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Authored {
    pub part_name: PackURI,
    pub relationship_id: String,
    pub shape_id: u32,
}
