//! Package-independent classification values and bounded snapshots.

use crate::{Error, Result};

const MAX_UNKNOWN_EXTENSIONS: usize = 1_024;
const MAX_UNKNOWN_EXTENSION_BYTES: usize = 1 << 20;
const MAX_UNKNOWN_BYTES: usize = 8 << 20;

/// The four classification outcomes defined by [MS-PPTX] 2.15.4.1.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
#[non_exhaustive]
pub enum Outcome {
    /// No classification outcome.
    None,
    /// Header classification outcome.
    Header,
    /// Footer classification outcome.
    Footer,
    /// Watermark classification outcome.
    Watermark,
}

impl Outcome {
    /// Parse the exact wire token from `p184:classification/@val`.
    pub(crate) fn parse(value: &str) -> Result<Self> {
        match value {
            "none" => Ok(Self::None),
            "hdr" => Ok(Self::Header),
            "ftr" => Ok(Self::Footer),
            "watermark" => Ok(Self::Watermark),
            _ => Err(Error::Invalid(format!(
                "unsupported PowerPoint classification outcome '{value}'"
            ))),
        }
    }

    /// Return the exact schema token used in `@val`.
    pub const fn wire(self) -> &'static str {
        match self {
            Self::None => "none",
            Self::Header => "hdr",
            Self::Footer => "ftr",
            Self::Watermark => "watermark",
        }
    }
}

/// An extension element that this release does not interpret.
///
/// The bytes are borrowed only through this owned, bounded value so a
/// snapshot can be edited and later written without normalizing a producer's
/// prefix, attribute order, whitespace, or child markup.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Opaque {
    pub(super) xml: Vec<u8>,
}

impl Opaque {
    /// Borrow the exact extension element bytes.
    #[inline]
    pub fn xml(&self) -> &[u8] {
        &self.xml
    }

    pub(crate) fn from_wire(xml: Vec<u8>) -> Result<Self> {
        if xml.is_empty() {
            return Err(Error::Invalid(
                "classification opaque extension is empty".into(),
            ));
        }
        if xml.len() > MAX_UNKNOWN_EXTENSION_BYTES {
            return Err(Error::Limit {
                resource: "classification opaque extension bytes",
                limit: MAX_UNKNOWN_EXTENSION_BYTES,
            });
        }
        Ok(Self { xml })
    }
}

/// A detached, lossless classification snapshot for one shape.
///
/// `None` from [`Snapshot::outcome`] represents a schema-valid extension whose
/// `classification` element omitted its optional `val` attribute. It is kept
/// distinct from [`Outcome::None`], which is the explicit `val="none"` token.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Snapshot {
    pub(super) outcome: Option<Outcome>,
    pub(super) unknown_extensions: Vec<Opaque>,
}

impl Snapshot {
    /// Construct a new snapshot with an explicit classification outcome.
    #[inline]
    pub fn new(outcome: Outcome) -> Self {
        Self {
            outcome: Some(outcome),
            unknown_extensions: Vec::new(),
        }
    }

    /// Return the optional wire outcome.
    #[inline]
    pub const fn outcome(&self) -> Option<Outcome> {
        self.outcome
    }

    /// Borrow extension elements retained without interpretation.
    #[inline]
    pub fn unknown_extensions(&self) -> &[Opaque] {
        &self.unknown_extensions
    }

    /// Start a detached atomic edit of this snapshot.
    #[inline]
    pub fn edit(&self) -> crate::shape::classification::transaction::Editor {
        crate::shape::classification::transaction::Editor::new(self.clone())
    }

    pub(crate) fn from_wire(
        outcome: Option<Outcome>,
        unknown_extensions: Vec<Opaque>,
    ) -> Result<Self> {
        let snapshot = Self {
            outcome,
            unknown_extensions,
        };
        snapshot.validate()?;
        Ok(snapshot)
    }

    pub(crate) fn validate(&self) -> Result<()> {
        if self.unknown_extensions.len() > MAX_UNKNOWN_EXTENSIONS {
            return Err(Error::Limit {
                resource: "classification opaque extensions",
                limit: MAX_UNKNOWN_EXTENSIONS,
            });
        }
        let mut total = 0usize;
        for extension in &self.unknown_extensions {
            if extension.xml.is_empty() || extension.xml.len() > MAX_UNKNOWN_EXTENSION_BYTES {
                return Err(Error::Limit {
                    resource: "classification opaque extension bytes",
                    limit: MAX_UNKNOWN_EXTENSION_BYTES,
                });
            }
            total = total.checked_add(extension.xml.len()).ok_or(Error::Limit {
                resource: "classification opaque extension bytes",
                limit: MAX_UNKNOWN_BYTES,
            })?;
            if total > MAX_UNKNOWN_BYTES {
                return Err(Error::Limit {
                    resource: "classification opaque extension bytes",
                    limit: MAX_UNKNOWN_BYTES,
                });
            }
        }
        Ok(())
    }
}
