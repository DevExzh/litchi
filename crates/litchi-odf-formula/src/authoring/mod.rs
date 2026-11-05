//! Typed `MathML` constructors and Formula package authoring.

mod constructors;
pub mod content;

use crate::model::Element;

pub use constructors::{
    Display, Variant, document_root, fenced, fraction, identifier, identifier_with_variant,
    literal_text, number, operator, over, root, row, semantics, square_root, string_literal,
    sub_superscript, subscript, superscript, table, under, under_over,
};

/// The `MathML` source a [`Builder`] packages.
enum Source {
    /// A typed tree, serialized at build time.
    Root(Element),
    /// Exact `MathML` text, validated and stored verbatim as `content.xml`.
    Text(String),
}

/// Build a Formula [`crate::Package`] from inert `MathML`.
pub struct Builder {
    source: Source,
    template: bool,
    limits: crate::codec::Limits,
}

impl Builder {
    /// Start authoring a standard `.odf` package from a typed root.
    #[must_use]
    pub fn new(root: Element) -> Self {
        Self::with_source(Source::Root(root), false)
    }

    /// Start authoring a Formula template `.otf` package from a typed root.
    #[must_use]
    pub fn template(root: Element) -> Self {
        Self::with_source(Source::Root(root), true)
    }

    /// Start authoring a standard `.odf` package whose `content.xml` is the
    /// exact `MathML` text, validated at [`Self::build`].
    #[must_use]
    pub fn from_mathml(mathml: impl Into<String>) -> Self {
        Self::with_source(Source::Text(mathml.into()), false)
    }

    /// Start authoring a Formula template `.otf` package whose `content.xml`
    /// is the exact `MathML` text, validated at [`Self::build`].
    #[must_use]
    pub fn template_from_mathml(mathml: impl Into<String>) -> Self {
        Self::with_source(Source::Text(mathml.into()), true)
    }

    fn with_source(source: Source, template: bool) -> Self {
        Self {
            source,
            template,
            limits: crate::codec::Limits::default(),
        }
    }

    /// Select the finite `MathML` parser limits validated at build time and
    /// retained by the package for later edits.
    #[must_use]
    pub const fn limits(mut self, limits: crate::codec::Limits) -> Self {
        self.limits = limits;
        self
    }

    /// Build and validate the package.
    ///
    /// # Errors
    ///
    /// Returns an error when the `MathML` exceeds a selected limit, fails
    /// validation, or the package cannot be created.
    pub fn build(self) -> Result<crate::Package, crate::Error> {
        let xml = match self.source {
            Source::Root(root) => root.to_xml(&self.limits)?,
            Source::Text(text) => text,
        };
        let formula = crate::Formula::from_mathml(&xml, self.template, self.limits)?;
        Ok(crate::Package::from_formula(formula))
    }
}
