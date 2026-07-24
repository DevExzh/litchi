//! Mutation support for the inert MathML element tree.
//!
//! These methods keep the tree well-formed by construction: element and
//! attribute names are validated as XML names and expanded attributes stay
//! unique. Structural validation (which elements may contain which children)
//! is deliberately left to the caller; the tree remains inert data and is
//! never evaluated.

use super::document::{MATHML_NAMESPACE, MathAttribute, MathContent, MathElement};
use litchi_core::{Error, Result};

fn invalid(message: impl Into<String>) -> Error {
    Error::InvalidFormat(message.into())
}

/// Validate an XML element or attribute local name (NCName subset).
fn validate_name(name: &str, kind: &str) -> Result<()> {
    let mut characters = name.chars();
    let Some(first) = characters.next() else {
        return Err(invalid(format!("MathML {kind} name is empty")));
    };
    if !(first.is_alphabetic() || first == '_')
        || !characters.all(|character| {
            character.is_alphanumeric() || matches!(character, '_' | '-' | '.')
        })
    {
        return Err(invalid(format!("invalid MathML {kind} name '{name}'")));
    }
    Ok(())
}

impl MathElement {
    /// Create an empty MathML-namespace element.
    ///
    /// The local name is validated so the tree always serializes to
    /// well-formed XML.
    pub fn new(local_name: &str) -> Result<Self> {
        Self::with_namespace(Some(MATHML_NAMESPACE), local_name)
    }

    /// Create an empty element in an explicit namespace.
    ///
    /// Names are validated so the tree always serializes to well-formed XML.
    pub fn with_namespace(namespace_uri: Option<&str>, local_name: &str) -> Result<Self> {
        validate_name(local_name, "element")?;
        if namespace_uri.is_some_and(str::is_empty) {
            return Err(invalid("MathML element namespace URI is empty"));
        }
        Ok(Self::from_parts(
            namespace_uri.map(str::to_string),
            local_name.to_string(),
            Vec::new(),
            Vec::new(),
        ))
    }

    /// Set or replace an attribute by expanded name.
    pub fn set_attribute(
        &mut self,
        namespace_uri: Option<&str>,
        local_name: &str,
        value: &str,
    ) -> Result<()> {
        validate_name(local_name, "attribute")?;
        if namespace_uri.is_some_and(str::is_empty) {
            return Err(invalid("MathML attribute namespace URI is empty"));
        }
        if let Some(existing) = self.attributes_mut().iter_mut().find(|attribute| {
            attribute.namespace_uri() == namespace_uri && attribute.local_name() == local_name
        }) {
            *existing = MathAttribute::from_parts(
                namespace_uri.map(str::to_string),
                local_name.to_string(),
                value.to_string(),
            );
        } else {
            self.attributes_mut().push(MathAttribute::from_parts(
                namespace_uri.map(str::to_string),
                local_name.to_string(),
                value.to_string(),
            ));
        }
        Ok(())
    }

    /// Remove an attribute by expanded name; returns whether it existed.
    pub fn remove_attribute(&mut self, namespace_uri: Option<&str>, local_name: &str) -> bool {
        let position = self.attributes().iter().position(|attribute| {
            attribute.namespace_uri() == namespace_uri && attribute.local_name() == local_name
        });
        if let Some(position) = position {
            self.attributes_mut().remove(position);
            true
        } else {
            false
        }
    }

    /// Append a child element after any existing content.
    pub fn push_child(&mut self, child: MathElement) {
        self.content_mut().push(MathContent::Element(child));
    }

    /// Append character content, merging with a trailing text run.
    pub fn push_text(&mut self, text: &str) {
        if text.is_empty() {
            return;
        }
        if let Some(MathContent::Text(existing)) = self.content_mut().last_mut() {
            existing.push_str(text);
        } else {
            self.content_mut().push(MathContent::Text(text.to_string()));
        }
    }

    /// Insert a child element at `index` among the element children only.
    ///
    /// Text runs do not count toward the index. Returns an error when the
    /// index exceeds the current child-element count.
    pub fn insert_child(&mut self, index: usize, child: MathElement) -> Result<()> {
        let position = self
            .content()
            .iter()
            .enumerate()
            .filter_map(|(position, content)| {
                matches!(content, MathContent::Element(_)).then_some(position)
            })
            .nth(index);
        match position {
            Some(position) => self.content_mut().insert(position, MathContent::Element(child)),
            None if index == self.children().count() => {
                self.content_mut().push(MathContent::Element(child));
            },
            None => {
                return Err(invalid(format!(
                    "MathML child insertion index {index} is out of range"
                )));
            },
        }
        Ok(())
    }

    /// Remove and return the child element at `index` among element children.
    pub fn remove_child(&mut self, index: usize) -> Option<MathElement> {
        let position = self
            .content()
            .iter()
            .enumerate()
            .filter_map(|(position, content)| {
                matches!(content, MathContent::Element(_)).then_some(position)
            })
            .nth(index)?;
        match self.content_mut().remove(position) {
            MathContent::Element(element) => Some(element),
            MathContent::Text(_) => unreachable!("position selects an element"),
        }
    }

    /// Replace the child element at `index`, returning the old element.
    pub fn replace_child(&mut self, index: usize, child: MathElement) -> Option<MathElement> {
        let removed = self.remove_child(index)?;
        self.insert_child(index, child)
            .expect("removal vacated the index");
        Some(removed)
    }

    /// Remove all content (elements and text) from the element.
    pub fn clear_content(&mut self) {
        self.content_mut().clear();
    }

    /// Serialize the subtree to a well-formed, self-contained XML string.
    ///
    /// MathML-namespace elements share the default namespace declared on the
    /// subtree root; foreign namespaces receive generated `ns1..nsN`
    /// prefixes in first-use order.
    pub fn to_xml(&self) -> String {
        super::serialize::write_mathml(self)
    }
}

#[cfg(test)]
mod tests {
    use super::super::document::parse_mathml;
    use super::*;

    #[test]
    fn builds_and_edits_a_tree() {
        let mut fraction = MathElement::new("mfrac").unwrap();
        assert!(fraction.insert_child(1, MathElement::new("mn").unwrap()).is_err());
        fraction.push_child(MathElement::new("mn").unwrap());
        fraction.insert_child(0, MathElement::new("mi").unwrap()).unwrap();
        assert_eq!(
            fraction.children().map(MathElement::local_name).collect::<Vec<_>>(),
            ["mi", "mn"]
        );
        let old = fraction.replace_child(1, MathElement::new("mrow").unwrap()).unwrap();
        assert_eq!(old.local_name(), "mn");
        assert_eq!(fraction.remove_child(0).unwrap().local_name(), "mi");
        assert!(fraction.remove_child(3).is_none());
        assert_eq!(fraction.children().count(), 1);
    }

    #[test]
    fn attribute_mutation_keeps_expanded_names_unique() {
        let mut element = MathElement::new("mi").unwrap();
        element
            .set_attribute(None, "mathvariant", "italic")
            .unwrap();
        element.set_attribute(None, "mathvariant", "bold").unwrap();
        assert_eq!(element.attributes().len(), 1);
        assert_eq!(element.attribute(None, "mathvariant"), Some("bold"));
        assert!(element.remove_attribute(None, "mathvariant"));
        assert!(!element.remove_attribute(None, "mathvariant"));
        assert!(element.set_attribute(None, "bad name", "x").is_err());
        assert!(element.set_attribute(Some(""), "x", "y").is_err());
    }

    #[test]
    fn text_pushes_merge_and_names_are_validated() {
        let mut element = MathElement::new("mtext").unwrap();
        element.push_text("a");
        element.push_text("");
        element.push_text("b");
        assert_eq!(element.content().len(), 1);
        assert_eq!(element.all_text(), "ab");
        assert!(MathElement::with_namespace(None, "1bad").is_err());
        assert!(MathElement::with_namespace(Some(""), "x").is_err());
        element.clear_content();
        assert!(element.content().is_empty());
    }

    #[test]
    fn edited_tree_round_trips_through_the_parser() {
        let mut root = MathElement::new("math").unwrap();
        let mut row = MathElement::new("mrow").unwrap();
        let mut identifier = MathElement::new("mi").unwrap();
        identifier.push_text("x");
        let mut operator = MathElement::new("mo").unwrap();
        operator.push_text("+");
        row.push_child(identifier);
        row.push_child(operator);
        root.push_child(row);
        let reparsed = parse_mathml(&root.to_xml()).unwrap();
        assert_eq!(reparsed, root);
    }
}
