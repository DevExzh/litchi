//! Semantic diagram-build values.

use crate::package::{Error, Result};

/// The `BuildTypeEnum` value stored by the shared build atom.
///
/// `Unknown` is safe because the field is a fixed four-byte scalar.  It is
/// retained so a producer extension can be read and written without
/// changing the record boundary.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
#[non_exhaustive]
pub enum Kind {
    /// A paragraph build.
    Paragraph,
    /// A chart build.
    Chart,
    /// The diagram build kind required by this container.
    Diagram,
    /// A bounded producer extension.
    Unknown(u32),
}

impl Kind {
    /// Decode the raw `BuildTypeEnum` value without losing unknown values.
    #[must_use]
    pub const fn from_raw(value: u32) -> Self {
        match value {
            1 => Self::Paragraph,
            2 => Self::Chart,
            3 => Self::Diagram,
            other => Self::Unknown(other),
        }
    }

    /// Return the exact wire value.
    #[must_use]
    pub const fn raw(self) -> u32 {
        match self {
            Self::Paragraph => 1,
            Self::Chart => 2,
            Self::Diagram => 3,
            Self::Unknown(value) => value,
        }
    }
}

/// The `DiagramBuildEnum` value stored by [`Atom`].
///
/// Unknown values remain lossless because this enum occupies one fixed
/// four-byte field and cannot affect child-record boundaries.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
#[non_exhaustive]
pub enum BuildType {
    AsOneObject,
    DepthByNode,
    DepthByBranch,
    BreadthByNode,
    BreadthByLevel,
    Clockwise,
    ClockwiseIn,
    ClockwiseOut,
    CounterClockwise,
    CounterClockwiseIn,
    CounterClockwiseOut,
    InByRing,
    OutByRing,
    Up,
    Down,
    AllAtOnce,
    Custom,
    Unknown(u32),
}

impl BuildType {
    /// Decode the exact MS-PPT value, retaining future values.
    #[must_use]
    pub const fn from_raw(value: u32) -> Self {
        match value {
            0x00 => Self::AsOneObject,
            0x01 => Self::DepthByNode,
            0x02 => Self::DepthByBranch,
            0x03 => Self::BreadthByNode,
            0x04 => Self::BreadthByLevel,
            0x05 => Self::Clockwise,
            0x06 => Self::ClockwiseIn,
            0x07 => Self::ClockwiseOut,
            0x08 => Self::CounterClockwise,
            0x09 => Self::CounterClockwiseIn,
            0x0A => Self::CounterClockwiseOut,
            0x0B => Self::InByRing,
            0x0C => Self::OutByRing,
            0x0D => Self::Up,
            0x0E => Self::Down,
            0x0F => Self::AllAtOnce,
            0x10 => Self::Custom,
            other => Self::Unknown(other),
        }
    }

    /// Return the exact wire value.
    #[must_use]
    pub const fn raw(self) -> u32 {
        match self {
            Self::AsOneObject => 0x00,
            Self::DepthByNode => 0x01,
            Self::DepthByBranch => 0x02,
            Self::BreadthByNode => 0x03,
            Self::BreadthByLevel => 0x04,
            Self::Clockwise => 0x05,
            Self::ClockwiseIn => 0x06,
            Self::ClockwiseOut => 0x07,
            Self::CounterClockwise => 0x08,
            Self::CounterClockwiseIn => 0x09,
            Self::CounterClockwiseOut => 0x0A,
            Self::InByRing => 0x0B,
            Self::OutByRing => 0x0C,
            Self::Up => 0x0D,
            Self::Down => 0x0E,
            Self::AllAtOnce => 0x0F,
            Self::Custom => 0x10,
            Self::Unknown(value) => value,
        }
    }
}

/// The shared `BuildAtom` payload carried by a diagram build.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct Build {
    /// Build identifier, unique with `shape_id_ref` on a slide.
    pub build_id: u32,
    /// Shape targeted by this build.
    pub shape_id_ref: u32,
    /// Whether the build has been expanded into time nodes.
    pub expanded: bool,
    /// Whether the build is expanded in the authoring UI.
    pub ui_expanded: bool,
    kind: Kind,
    reserved: [u8; 2],
}

impl Build {
    /// Create a diagram build atom with normalized reserved bytes.
    #[must_use]
    pub const fn new(build_id: u32, shape_id_ref: u32, expanded: bool, ui_expanded: bool) -> Self {
        Self {
            build_id,
            shape_id_ref,
            expanded,
            ui_expanded,
            kind: Kind::Diagram,
            reserved: [0; 2],
        }
    }

    /// Return the shared build kind, including an unknown raw value when read.
    #[must_use]
    pub const fn kind(self) -> Kind {
        self.kind
    }

    /// Return the two undefined bytes retained from the source record.
    #[must_use]
    pub const fn reserved(self) -> [u8; 2] {
        self.reserved
    }

    /// Preserve undefined bytes when authoring a lossless edit.
    #[must_use]
    pub const fn with_reserved(mut self, reserved: [u8; 2]) -> Self {
        self.reserved = reserved;
        self
    }

    pub(crate) const fn from_parts(
        kind: Kind,
        build_id: u32,
        shape_id_ref: u32,
        expanded: bool,
        ui_expanded: bool,
        reserved: [u8; 2],
    ) -> Self {
        Self {
            build_id,
            shape_id_ref,
            expanded,
            ui_expanded,
            kind,
            reserved,
        }
    }
}

/// The fixed `DiagramBuildAtom` payload.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct Atom {
    /// Diagram animation order.
    pub build_type: BuildType,
}

impl Atom {
    /// Create a typed diagram-build atom.
    #[must_use]
    pub const fn new(build_type: BuildType) -> Self {
        Self { build_type }
    }
}

/// The fixed `DiagramBuildContainer` record.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct Container {
    build: Build,
    atom: Atom,
}

impl Container {
    /// Construct a diagram container from its two typed child records.
    ///
    /// Known paragraph/chart build kinds are rejected because they would
    /// make this container semantically incorrect. Unknown fixed-width kinds
    /// are retained for forward-compatible metadata round trips.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation fails.
    pub fn new(build: Build, atom: Atom) -> Result<Self> {
        if matches!(build.kind, Kind::Paragraph | Kind::Chart) {
            return Err(Error::InvalidFormat(
                "diagram build container requires a diagram BuildAtom".to_string(),
            ));
        }
        Ok(Self { build, atom })
    }

    /// Return the shared build child.
    #[must_use]
    pub const fn build(self) -> Build {
        self.build
    }

    /// Return the diagram-specific atom child.
    #[must_use]
    pub const fn atom(self) -> Atom {
        self.atom
    }
}

/// Resource bounds for one complete diagram-build container.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct Limits {
    /// Maximum accepted container size, including its record header.
    pub max_record_bytes: usize,
}

impl Default for Limits {
    fn default() -> Self {
        Self {
            max_record_bytes: Container::RECORD_LEN,
        }
    }
}
