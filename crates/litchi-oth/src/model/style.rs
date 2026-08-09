//! Named style inventory.

/// XML part in which a style was declared.
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub enum Origin {
    /// `content.xml` automatic styles.
    Content,
    /// `styles.xml` common, automatic, or master styles.
    Styles,
}

/// One named ODF style declaration.
#[derive(Clone, Debug, PartialEq, Eq)]
pub struct Style {
    family: Option<String>,
    name: String,
    origin: Origin,
    parent_name: Option<String>,
}

impl Style {
    pub(crate) const fn projected(
        name: String,
        family: Option<String>,
        parent_name: Option<String>,
        origin: Origin,
    ) -> Self {
        Self {
            family,
            name,
            origin,
            parent_name,
        }
    }

    /// Style name used by content references.
    #[must_use]
    pub fn name(&self) -> &str {
        &self.name
    }

    /// ODF style family.
    #[must_use]
    pub fn family(&self) -> Option<&str> {
        self.family.as_deref()
    }

    /// Parent style reference.
    #[must_use]
    pub fn parent_name(&self) -> Option<&str> {
        self.parent_name.as_deref()
    }

    /// Declaring package part.
    #[must_use]
    pub const fn origin(&self) -> Origin {
        self.origin
    }
}
