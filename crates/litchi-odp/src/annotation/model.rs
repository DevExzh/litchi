//! Semantic presentation annotation anchors and inventory values.

use litchi_core::{Error, Result};

const MAX_SHAPE_NAME_BYTES: usize = 16 * 1024;

/// A presentation location that can own an ODF annotation.
#[derive(Clone, Debug, PartialEq, Eq, Hash)]
#[non_exhaustive]
pub enum Position {
    /// The annotation is a child of the selected `draw:page`.
    Page { index: usize },
    /// The annotation is a child of the uniquely named drawing shape.
    Shape { page_index: usize, name: String },
}

/// A checked presentation annotation anchor.
#[derive(Clone, Debug, PartialEq, Eq, Hash)]
pub struct Anchor {
    position: Position,
}

impl Anchor {
    /// Construct an anchor at a page index.
    #[must_use]
    pub const fn page(index: usize) -> Self {
        Self {
            position: Position::Page { index },
        }
    }

    /// Construct an anchor at a uniquely named shape on a page.
    pub fn shape(page_index: usize, name: impl Into<String>) -> Result<Self> {
        let name = name.into();
        validate_shape_name(&name)?;
        Ok(Self {
            position: Position::Shape { page_index, name },
        })
    }

    /// Return the semantic anchor position.
    #[must_use]
    pub const fn position(&self) -> &Position {
        &self.position
    }

    pub(crate) const fn from_position(position: Position) -> Self {
        Self { position }
    }
}

/// One annotation in presentation document order.
#[derive(Clone, Debug, PartialEq, Eq)]
pub struct Info {
    /// Stable only for this inventory snapshot.
    pub index: usize,
    /// Shared rich ODF annotation content and metadata.
    pub annotation: super::Annotation,
    /// The page or uniquely named shape that contains the annotation.
    pub anchor: Anchor,
}

pub(crate) fn validate_shape_name(name: &str) -> Result<()> {
    if name.is_empty() {
        return Err(Error::InvalidFormat(
            "presentation shape annotation anchor requires a non-empty draw:name".to_string(),
        ));
    }
    if name.len() > MAX_SHAPE_NAME_BYTES {
        return Err(Error::InvalidFormat(
            "presentation shape annotation anchor name exceeds the size limit".to_string(),
        ));
    }
    if name.chars().any(char::is_control) {
        return Err(Error::InvalidFormat(
            "presentation shape annotation anchor contains an XML control character".to_string(),
        ));
    }
    Ok(())
}
