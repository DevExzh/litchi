//! Named style inventory.

/// XML part in which a style was declared.
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub enum Origin {
    /// `content.xml` automatic styles.
    Content,
    /// `styles.xml` common, automatic, or master styles.
    Styles,
}

/// Supported font weight.
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub enum Weight {
    /// Normal weight.
    Normal,
    /// Bold weight.
    Bold,
}

/// Supported font posture.
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub enum Slant {
    /// Upright text.
    Normal,
    /// Italic text.
    Italic,
}

/// A bounded typed subset of ODF `style:text-properties`.
#[derive(Clone, Debug, Default, PartialEq, Eq)]
pub struct TextProperties {
    background_color: Option<String>,
    color: Option<String>,
    slant: Option<Slant>,
    weight: Option<Weight>,
}

impl TextProperties {
    /// Creates empty text properties.
    #[must_use]
    pub const fn new() -> Self {
        Self {
            background_color: None,
            color: None,
            slant: None,
            weight: None,
        }
    }

    /// Sets an RGB color written as `#RRGGBB`.
    ///
    /// # Errors
    ///
    /// Returns an error for a non-canonical color.
    pub fn with_color(mut self, color: impl Into<String>) -> litchi_core::Result<Self> {
        let owned_color = color.into();
        self.color = Some(validate_color(&owned_color)?);
        Ok(self)
    }

    /// Sets an RGB background color written as `#RRGGBB`.
    ///
    /// # Errors
    ///
    /// Returns an error for a non-canonical color.
    pub fn with_background_color(mut self, color: impl Into<String>) -> litchi_core::Result<Self> {
        let owned_color = color.into();
        self.background_color = Some(validate_color(&owned_color)?);
        Ok(self)
    }

    /// Sets font weight.
    #[must_use]
    pub const fn with_weight(mut self, weight: Weight) -> Self {
        self.weight = Some(weight);
        self
    }

    /// Sets font posture.
    #[must_use]
    pub const fn with_slant(mut self, slant: Slant) -> Self {
        self.slant = Some(slant);
        self
    }

    /// Foreground color.
    #[must_use]
    pub fn color(&self) -> Option<&str> {
        self.color.as_deref()
    }

    /// Background color.
    #[must_use]
    pub fn background_color(&self) -> Option<&str> {
        self.background_color.as_deref()
    }

    /// Font weight.
    #[must_use]
    pub const fn weight(&self) -> Option<Weight> {
        self.weight
    }

    /// Font posture.
    #[must_use]
    pub const fn slant(&self) -> Option<Slant> {
        self.slant
    }

    pub(crate) fn projected(
        color: Option<String>,
        background_color: Option<String>,
        weight: Option<Weight>,
        slant: Option<Slant>,
    ) -> Self {
        Self {
            background_color: background_color.and_then(|value| validate_color(&value).ok()),
            color: color.and_then(|value| validate_color(&value).ok()),
            slant,
            weight,
        }
    }

    pub(crate) const fn is_empty(&self) -> bool {
        self.background_color.is_none()
            && self.color.is_none()
            && self.slant.is_none()
            && self.weight.is_none()
    }
}

/// One named ODF style declaration.
#[derive(Clone, Debug, PartialEq, Eq)]
pub struct Style {
    family: Option<String>,
    name: String,
    origin: Origin,
    parent_name: Option<String>,
    text_properties: Option<TextProperties>,
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
            text_properties: None,
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

    /// Sets typed text properties.
    #[must_use]
    pub fn with_text_properties(mut self, properties: TextProperties) -> Self {
        self.text_properties = Some(properties);
        self
    }

    pub(crate) const fn projected(
        name: String,
        family: Option<String>,
        parent_name: Option<String>,
        origin: Origin,
        text_properties: Option<TextProperties>,
    ) -> Self {
        Self {
            family,
            name,
            origin,
            parent_name,
            text_properties,
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

    /// Typed text properties declared directly by this style.
    #[must_use]
    pub const fn text_properties(&self) -> Option<&TextProperties> {
        self.text_properties.as_ref()
    }
}

fn validate_color(color: &str) -> litchi_core::Result<String> {
    if color.len() == 7
        && color.starts_with('#')
        && color.as_bytes()[1..].iter().all(u8::is_ascii_hexdigit)
    {
        return Ok(color.to_ascii_uppercase());
    }
    Err(litchi_core::Error::InvalidFormat(
        "OTH style color must be canonical #RRGGBB".to_string(),
    ))
}
