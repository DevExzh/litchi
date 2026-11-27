//! Named drawing resources referenced by automatic styles.

use std::collections::BTreeMap;

/// Supported named drawing-resource element kind.
#[derive(Clone, Copy, Debug, PartialEq, Eq, PartialOrd, Ord)]
#[non_exhaustive]
pub enum StyleResourceKind {
    Gradient,
    Hatch,
    FillImage,
    Marker,
    Opacity,
    StrokeDash,
}

impl StyleResourceKind {
    /// ODF `draw:*` element local name.
    #[must_use]
    pub const fn element(self) -> &'static str {
        match self {
            Self::Gradient => "gradient",
            Self::Hatch => "hatch",
            Self::FillImage => "fill-image",
            Self::Marker => "marker",
            Self::Opacity => "opacity",
            Self::StrokeDash => "stroke-dash",
        }
    }
}

/// One bounded inert named drawing resource and its qualified attributes.
#[derive(Clone, Debug, PartialEq, Eq)]
pub struct StyleResource {
    kind: StyleResourceKind,
    name: String,
    attributes: BTreeMap<String, String>,
}

impl StyleResource {
    /// Creates a detached named drawing resource.
    #[must_use]
    pub fn new(kind: StyleResourceKind, name: impl Into<String>) -> Self {
        Self {
            kind,
            name: name.into(),
            attributes: BTreeMap::new(),
        }
    }

    /// Adds or replaces one inert qualified attribute.
    #[must_use]
    pub fn with_attribute(mut self, name: impl Into<String>, value: impl Into<String>) -> Self {
        self.attributes.insert(name.into(), value.into());
        self
    }

    pub(crate) fn parsed(
        kind: StyleResourceKind,
        name: String,
        attributes: BTreeMap<String, String>,
    ) -> Self {
        Self {
            kind,
            name,
            attributes,
        }
    }

    /// Resource element kind.
    #[must_use]
    pub const fn kind(&self) -> StyleResourceKind {
        self.kind
    }

    /// Name referenced by automatic-style properties.
    #[must_use]
    pub fn name(&self) -> &str {
        &self.name
    }

    /// Stable inert qualified attributes excluding `draw:name`.
    #[must_use]
    pub const fn attributes(&self) -> &BTreeMap<String, String> {
        &self.attributes
    }
}
