//! Strict references used by Pages' following-paragraph-style control.

use std::num::NonZeroU64;

/// Validated user-visible name for a named paragraph style.
#[derive(Debug, Clone, PartialEq, Eq, Hash)]
pub struct ParagraphStyleName(Box<str>);

impl ParagraphStyleName {
    /// Construct a nonempty printable name without surrounding whitespace.
    pub fn new(name: impl Into<String>) -> crate::Result<Self> {
        let name = name.into();
        if name.is_empty()
            || name.trim() != name
            || name.chars().any(|character| character.is_control())
        {
            return Err(crate::Error::InvalidFormat(
                "named iWork paragraph styles require a nonempty printable name without surrounding whitespace"
                    .to_owned(),
            ));
        }
        Ok(Self(name.into_boxed_str()))
    }

    /// Return the validated user-visible name.
    pub fn as_str(&self) -> &str {
        &self.0
    }
}

/// Package-local identifier of a named iWork paragraph style.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub struct ParagraphStyleId(NonZeroU64);

impl ParagraphStyleId {
    /// Construct a nonzero package-local paragraph-style identifier.
    pub fn new(identifier: u64) -> crate::Result<Self> {
        NonZeroU64::new(identifier).map(Self).ok_or_else(|| {
            crate::Error::InvalidFormat(
                "iWork paragraph-style identifiers must be nonzero".to_owned(),
            )
        })
    }

    /// Return the package-local object identifier.
    pub const fn get(self) -> u64 {
        self.0.get()
    }
}

/// One named paragraph-style preset selectable in the current document.
#[derive(Debug, Clone, PartialEq, Eq, Hash)]
pub struct NamedParagraphStyle {
    id: ParagraphStyleId,
    name: Box<str>,
}

impl NamedParagraphStyle {
    pub(crate) fn new(id: ParagraphStyleId, name: String) -> crate::Result<Self> {
        if name.is_empty() || name.chars().any(char::is_control) {
            return Err(crate::Error::InvalidFormat(
                "named iWork paragraph styles require a nonempty printable name".to_owned(),
            ));
        }
        Ok(Self {
            id,
            name: name.into_boxed_str(),
        })
    }

    /// Return the package-local identifier used to target this style.
    pub const fn id(&self) -> ParagraphStyleId {
        self.id
    }

    /// Return the user-visible paragraph-style name.
    pub fn name(&self) -> &str {
        &self.name
    }
}

/// Paragraph style Pages should apply after the current paragraph.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash, Default)]
pub enum ParagraphFollowingStyle {
    /// Use Pages' native “Same” behavior.
    #[default]
    Same,
    /// Apply one named paragraph style from the current document.
    Named(ParagraphStyleId),
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn paragraph_style_identifiers_and_names_are_strict() {
        assert!(ParagraphStyleId::new(0).is_err());
        assert!(ParagraphStyleName::new("").is_err());
        assert!(ParagraphStyleName::new(" Heading").is_err());
        assert!(ParagraphStyleName::new("Bad\nName").is_err());
        assert_eq!(
            ParagraphStyleName::new("Heading").unwrap().as_str(),
            "Heading"
        );
        let identifier = ParagraphStyleId::new(42).unwrap();
        assert_eq!(identifier.get(), 42);
        assert_eq!(
            NamedParagraphStyle::new(identifier, "Heading".to_owned())
                .unwrap()
                .name(),
            "Heading"
        );
        assert!(NamedParagraphStyle::new(identifier, String::new()).is_err());
        assert!(NamedParagraphStyle::new(identifier, "Bad\nName".to_owned()).is_err());
    }
}
