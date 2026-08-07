//! Typed `MathML` constructors and Formula package authoring.

mod constructors;

use litchi_core::Result;

use crate::model::Element;

pub use constructors::{
    Display, Variant, document_root, fenced, fraction, identifier, identifier_with_variant,
    literal_text, number, operator, over, root, row, semantics, square_root, string_literal,
    sub_superscript, subscript, superscript, table, under, under_over,
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
        let xml = self.root.to_xml();
        if self.template {
            crate::Formula::create_template(xml)
        } else {
            crate::Formula::create(xml)
        }
    }
}
