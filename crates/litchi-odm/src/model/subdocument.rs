//! Inert referenced-subdocument semantics.

/// A classified, inert subdocument target.
///
/// The library never opens, resolves, fetches, or executes either variant.
/// `Package` is restricted to a relative package path; all other URI-like or
/// unsafe paths stay explicitly external.
#[derive(Clone, Debug, PartialEq, Eq)]
#[non_exhaustive]
pub enum Target {
    /// A safe relative path in the master document's package namespace.
    Package(String),
    /// An external, absolute, URI-like, or otherwise unsafe target.
    External(String),
}

impl Target {
    /// Returns the original reference text without resolving it.
    #[must_use]
    pub fn href(&self) -> &str {
        match self {
            Self::Package(href) | Self::External(href) => href,
        }
    }

    /// Returns whether this target is deliberately classified as external.
    #[must_use]
    pub const fn is_external(&self) -> bool {
        matches!(self, Self::External(_))
    }
}

/// An ordered subdocument reference bound to its containing section.
#[derive(Clone, Debug, PartialEq, Eq)]
pub struct Reference {
    section: String,
    target: Target,
    source_section: Option<String>,
    filter_name: Option<String>,
}

impl Reference {
    pub(crate) fn new(
        section: String,
        href: String,
        source_section: Option<String>,
        filter_name: Option<String>,
    ) -> Self {
        Self {
            section,
            target: classify_target(href),
            source_section,
            filter_name,
        }
    }

    /// Returns the containing `text:section` name.
    #[must_use]
    pub fn section(&self) -> &str {
        &self.section
    }

    /// Returns the classified inert target.
    #[must_use]
    pub const fn target(&self) -> &Target {
        &self.target
    }

    /// Returns the original target text without resolving it.
    #[must_use]
    pub fn href(&self) -> &str {
        self.target.href()
    }

    /// Returns the named section selected within the linked document.
    #[must_use]
    pub fn source_section(&self) -> Option<&str> {
        self.source_section.as_deref()
    }

    /// Returns the producer filter name attached to the linked section.
    #[must_use]
    pub fn filter_name(&self) -> Option<&str> {
        self.filter_name.as_deref()
    }
}

/// A referenced master-document subdocument.
#[derive(Clone, Debug, PartialEq, Eq)]
pub struct Subdocument {
    href: String,
}

impl Subdocument {
    pub fn new(href: impl Into<String>) -> Self {
        Self { href: href.into() }
    }

    /// Returns the subdocument reference target.
    #[must_use]
    pub fn href(&self) -> &str {
        &self.href
    }
}

fn classify_target(href: String) -> Target {
    if is_safe_package_path(&href) {
        Target::Package(href)
    } else {
        Target::External(href)
    }
}

fn is_safe_package_path(href: &str) -> bool {
    !href.is_empty()
        && !href.starts_with('/')
        && !href.starts_with('\\')
        && !href.starts_with("//")
        && !href.contains('\\')
        && !href.contains(':')
        && href
            .split('/')
            .all(|segment| !segment.is_empty() && segment != "." && segment != "..")
}

#[cfg(test)]
mod tests {
    use super::{Reference, Target};

    #[test]
    fn classifies_only_safe_relative_paths_as_package_targets() {
        assert!(matches!(
            Reference::new("A".to_string(), "Chapters/a.odt".to_string(), None, None).target(),
            Target::Package(_)
        ));
        for href in [
            "../a.odt",
            "/a.odt",
            "https://example.test/a.odt",
            "file:a.odt",
        ] {
            assert!(matches!(
                Reference::new("A".to_string(), href.to_string(), None, None).target(),
                Target::External(_)
            ));
        }
    }
}
