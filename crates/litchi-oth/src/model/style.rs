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
    /// Creates a detached named style.
    ///
    /// # Errors
    ///
    /// Returns an error when the name is empty or either value contains NUL.
    pub fn new(name: impl Into<String>, family: impl Into<String>) -> litchi_core::Result<Self> {
        let style_name = name.into();
        let style_family = family.into();
        if style_name.is_empty() || style_name.contains('\0') || style_family.contains('\0') {
            return Err(litchi_core::Error::InvalidFormat(
                "invalid OTH detached style name or family".to_string(),
            ));
        }
        Ok(Self {
            family: Some(style_family),
            name: style_name,
            origin: Origin::Styles,
            parent_name: None,
        })
    }

    /// Sets a parent style reference.
    ///
    /// # Errors
    ///
    /// Returns an error when the parent name contains NUL.
    pub fn with_parent(mut self, parent_name: impl Into<String>) -> litchi_core::Result<Self> {
        let converted_parent_name = parent_name.into();
        if converted_parent_name.contains('\0') {
            return Err(litchi_core::Error::InvalidFormat(
                "invalid OTH parent style name".to_string(),
            ));
        }
        self.parent_name = Some(converted_parent_name);
        Ok(self)
    }

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
