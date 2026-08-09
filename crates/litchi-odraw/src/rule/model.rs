//! Semantic values for `OfficeArt` solver rules.

use crate::Record;

/// The supported `OfficeArt` solver-rule record family.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub enum Kind {
    /// A connector joining two shapes through a third connector shape.
    Connector,
    /// An arc shape rule.
    Arc,
    /// A callout shape rule.
    Callout,
    /// An unrecognized `OfficeArt` record type, retaining its wire value.
    Opaque(u16),
}

impl Kind {
    /// Returns the canonical `OfficeArt` record type for a known rule kind.
    #[must_use]
    pub const fn raw(self) -> u16 {
        match self {
            Self::Connector => 0xF012,
            Self::Arc => 0xF014,
            Self::Callout => 0xF017,
            Self::Opaque(raw) => raw,
        }
    }
}

/// The fixed-layout `OfficeArtFConnectorRule` record.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
#[allow(
    clippy::struct_field_names,
    reason = "fields mirror the [MS-ODRAW] OfficeArtFConnectorRule layout (spidA/spidB/spidC) and the public accessors"
)]
pub struct Connector {
    rule_id: u32,
    start_shape_id: u32,
    end_shape_id: u32,
    connector_shape_id: u32,
    start_connection_site: u32,
    end_connection_site: u32,
}

impl Connector {
    /// Creates a connector rule from its `[MS-ODRAW]` fields.
    #[must_use]
    pub const fn new(
        rule_id: u32,
        start_shape_id: u32,
        end_shape_id: u32,
        connector_shape_id: u32,
        start_connection_site: u32,
        end_connection_site: u32,
    ) -> Self {
        Self {
            rule_id,
            start_shape_id,
            end_shape_id,
            connector_shape_id,
            start_connection_site,
            end_connection_site,
        }
    }

    /// Returns the connector-rule identifier (`ruid`).
    #[must_use]
    pub const fn rule_id(self) -> u32 {
        self.rule_id
    }

    /// Returns the shape at which the connector starts (`spidA`).
    #[must_use]
    pub const fn start_shape_id(self) -> u32 {
        self.start_shape_id
    }

    /// Returns the shape at which the connector ends (`spidB`).
    #[must_use]
    pub const fn end_shape_id(self) -> u32 {
        self.end_shape_id
    }

    /// Returns the connector shape (`spidC`).
    #[must_use]
    pub const fn connector_shape_id(self) -> u32 {
        self.connector_shape_id
    }

    /// Returns the start-shape connection-site index (`cptiA`).
    #[must_use]
    pub const fn start_connection_site(self) -> u32 {
        self.start_connection_site
    }

    /// Returns the end-shape connection-site index (`cptiB`).
    #[must_use]
    pub const fn end_connection_site(self) -> u32 {
        self.end_connection_site
    }
}

/// The fixed-layout `OfficeArtFArcRule` record.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub struct Arc {
    rule_id: u32,
    shape_id: u32,
}

impl Arc {
    /// Creates an arc rule from its `[MS-ODRAW]` fields.
    #[must_use]
    pub const fn new(rule_id: u32, shape_id: u32) -> Self {
        Self { rule_id, shape_id }
    }

    /// Returns the arc-rule identifier (`ruid`).
    #[must_use]
    pub const fn rule_id(self) -> u32 {
        self.rule_id
    }

    /// Returns the arc shape identifier (`spid`).
    #[must_use]
    pub const fn shape_id(self) -> u32 {
        self.shape_id
    }
}

/// The fixed-layout `OfficeArtFCalloutRule` record.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub struct Callout {
    rule_id: u32,
    shape_id: u32,
}

impl Callout {
    /// Creates a callout rule from its `[MS-ODRAW]` fields.
    #[must_use]
    pub const fn new(rule_id: u32, shape_id: u32) -> Self {
        Self { rule_id, shape_id }
    }

    /// Returns the callout-rule identifier (`ruid`).
    #[must_use]
    pub const fn rule_id(self) -> u32 {
        self.rule_id
    }

    /// Returns the callout shape identifier (`spid`).
    #[must_use]
    pub const fn shape_id(self) -> u32 {
        self.shape_id
    }
}

/// A borrowed `OfficeArt` record outside the supported rule layouts.
///
/// The complete record header and body are retained through the existing
/// zero-copy [`Record`] view.  This permits a caller to inspect or replay a
/// producer extension without the common layer guessing at its schema.
#[derive(Debug, Clone)]
pub struct Opaque<'data> {
    pub(super) record: Record<'data>,
}

impl<'data> Opaque<'data> {
    /// Returns the exact record type read from the wire.
    #[must_use]
    pub const fn raw_kind(&self) -> u16 {
        self.record.raw_kind()
    }

    /// Returns the record version.
    #[must_use]
    pub const fn version(&self) -> u8 {
        self.record.version()
    }

    /// Returns the record instance.
    #[must_use]
    pub const fn instance(&self) -> u16 {
        self.record.instance()
    }

    /// Returns the declared body length.
    #[must_use]
    pub const fn len(&self) -> u32 {
        self.record.len()
    }

    /// Returns whether the declared body length is zero.
    #[must_use]
    pub const fn is_empty(&self) -> bool {
        self.record.len() == 0
    }

    /// Returns the borrowed, uninterpreted record body.
    #[must_use]
    pub const fn data(&self) -> &'data [u8] {
        self.record.data()
    }

    /// Returns the underlying record view.
    #[must_use]
    pub const fn record(&self) -> &Record<'data> {
        &self.record
    }
}

/// One decoded `OfficeArt` solver rule.
#[derive(Debug, Clone)]
pub enum Rule<'data> {
    /// Connector rule (`0xF012`).
    Connector(Connector),
    /// Arc rule (`0xF014`).
    Arc(Arc),
    /// Callout rule (`0xF017`).
    Callout(Callout),
    /// An unrecognized record retained for lossless replay.
    Opaque(Opaque<'data>),
}

impl Rule<'_> {
    /// Returns the semantic rule kind.
    #[must_use]
    pub const fn kind(&self) -> Kind {
        match self {
            Self::Connector(_) => Kind::Connector,
            Self::Arc(_) => Kind::Arc,
            Self::Callout(_) => Kind::Callout,
            Self::Opaque(record) => Kind::Opaque(record.raw_kind()),
        }
    }

    /// Returns the exact `OfficeArt` record type.
    #[must_use]
    pub const fn raw_kind(&self) -> u16 {
        self.kind().raw()
    }

    /// Returns the rule's wire version.
    #[must_use]
    pub const fn version(&self) -> u8 {
        match self {
            Self::Connector(_) => 1,
            Self::Arc(_) | Self::Callout(_) => 0,
            Self::Opaque(record) => record.version(),
        }
    }

    /// Returns the rule's wire instance.
    #[must_use]
    pub const fn instance(&self) -> u16 {
        match self {
            Self::Connector(_) | Self::Arc(_) | Self::Callout(_) => 0,
            Self::Opaque(record) => record.instance(),
        }
    }

    /// Returns the connector value when this is a connector rule.
    #[must_use]
    pub const fn as_connector(&self) -> Option<&Connector> {
        match self {
            Self::Connector(rule) => Some(rule),
            Self::Arc(_) | Self::Callout(_) | Self::Opaque(_) => None,
        }
    }

    /// Returns the arc value when this is an arc rule.
    #[must_use]
    pub const fn as_arc(&self) -> Option<&Arc> {
        match self {
            Self::Arc(rule) => Some(rule),
            Self::Connector(_) | Self::Callout(_) | Self::Opaque(_) => None,
        }
    }

    /// Returns the callout value when this is a callout rule.
    #[must_use]
    pub const fn as_callout(&self) -> Option<&Callout> {
        match self {
            Self::Callout(rule) => Some(rule),
            Self::Connector(_) | Self::Arc(_) | Self::Opaque(_) => None,
        }
    }

    /// Returns the opaque value when this record is not a known rule.
    #[must_use]
    pub const fn as_opaque(&self) -> Option<&Opaque<'_>> {
        match self {
            Self::Opaque(record) => Some(record),
            Self::Connector(_) | Self::Arc(_) | Self::Callout(_) => None,
        }
    }
}
