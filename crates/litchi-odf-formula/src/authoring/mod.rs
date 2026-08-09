//! Typed `MathML` constructors and Formula package authoring.

mod constructors;

use litchi_core::Result;

use crate::model::Element;

pub use constructors::{
    Display, Variant, document_root, fenced, fraction, identifier, identifier_with_variant,
    literal_text, number, operator, over, root, row, semantics, semantics_with_opaque_starmath,
    semantics_with_starmath, square_root, string_literal, sub_superscript, subscript, superscript,
    table, under, under_over,
};

/// Build a Formula package from an inert `MathML` root.
pub struct Builder {
    root: Element,
    template: bool,
}

impl Builder {
    /// Start authoring a standard `.odf` package.
    #[must_use]
    pub fn new(root: Element) -> Self {
        Self {
            root,
            template: false,
        }
    }

    /// Start authoring a Formula template `.otf` package.
    #[must_use]
    pub fn template(root: Element) -> Self {
        Self {
            root,
            template: true,
        }
    }

    /// Build and validate the package.
    ///
    /// # Errors
    ///
    /// Returns an error when the serialized `MathML` fails validation or the
    /// package cannot be created.
    pub fn build(self) -> Result<crate::Formula> {
        self.build_with_limits(crate::Limits::default())
    }

    /// Build and validate the package under caller-selected finite limits.
    ///
    /// # Errors
    ///
    /// Returns an error when the serialized `MathML` exceeds a limit, fails
    /// validation, or the package cannot be created.
    pub fn build_with_limits(self, limits: crate::Limits) -> Result<crate::Formula> {
        let xml = self.root.to_xml();
        if self.template {
            crate::Formula::create_template_with_limits(xml, limits)
        } else {
            crate::Formula::create_with_limits(xml, limits)
        }
    }
}
