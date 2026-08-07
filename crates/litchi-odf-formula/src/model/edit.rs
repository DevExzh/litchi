//! Mutation support for the inert `MathML` element tree.
//!
//! These methods keep the tree well-formed by construction: element and
//! attribute names are validated as XML names and expanded attributes stay
//! unique. Structural validation (which elements may contain which children)
//! is deliberately left to the caller; the tree remains inert data and is
//! never evaluated.

use super::{Attribute, Content, Element, MATHML_NAMESPACE};
use litchi_core::{Error, Result};

impl Element {
    /// Create an empty MathML-namespace element.
    ///
    /// The local name is validated so the tree always serializes to
    /// well-formed XML.
    ///
    /// # Errors
    ///
    /// Returns an error when `local_name` is not a valid XML name.
    pub fn new(local_name: &str) -> Result<Self> {
        Self::with_namespace(Some(MATHML_NAMESPACE), local_name)
    }

    /// Create an empty element in an explicit namespace.
    ///
    /// Names are validated so the tree always serializes to well-formed XML.
    ///
    /// # Errors
    ///
    /// Returns an error when `local_name` is not a valid XML name or the
    /// namespace URI is empty.
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
    ///
    /// # Errors
    ///
    /// Returns an error when `local_name` is not a valid XML name or the
    /// namespace URI is empty.
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
            *existing = Attribute::from_parts(
                namespace_uri.map(str::to_string),
                local_name.to_string(),
                value.to_string(),
            );
        } else {
            self.attributes_mut().push(Attribute::from_parts(
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
        if let Some(index) = position {
            self.attributes_mut().remove(index);
            true
        } else {
            false
        }
    }

    /// Append a child element after any existing content.
    pub fn push_child(&mut self, child: Element) {
        self.content_mut().push(Content::Element(child));
    }

    /// Append character content, merging with a trailing text run.
    pub fn push_text(&mut self, text: &str) {
        if text.is_empty() {
            return;
        }
        if let Some(Content::Text(existing)) = self.content_mut().last_mut() {
            existing.push_str(text);
        } else {
            self.content_mut().push(Content::Text(text.to_string()));
        }
    }

    /// Insert a child element at `index` among the element children only.
    ///
    /// Text runs do not count toward the index. Returns an error when the
    /// index exceeds the current child-element count.
    ///
    /// # Errors
    ///
    /// Returns an error when `index` exceeds the current child-element count.
    pub fn insert_child(&mut self, index: usize, child: Element) -> Result<()> {
        let position = self
            .content()
            .iter()
            .enumerate()
            .filter_map(|(position, content)| {
                matches!(content, Content::Element(_)).then_some(position)
            })
            .nth(index);
        match position {
            Some(content_position) => self
                .content_mut()
                .insert(content_position, Content::Element(child)),
            None if index == self.children().count() => {
                self.content_mut().push(Content::Element(child));
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
    pub fn remove_child(&mut self, index: usize) -> Option<Element> {
        let position = self
            .content()
            .iter()
            .enumerate()
            .filter_map(|(position, content)| {
                matches!(content, Content::Element(_)).then_some(position)
            })
            .nth(index)?;
        match self.content_mut().remove(position) {
            Content::Element(element) => Some(element),
            Content::Text(_) => unreachable!("position selects an element"),
        }
    }

    /// Replace the child element at `index`, returning the old element.
    ///
    /// # Panics
    ///
    /// Panics only on an internal invariant violation: reinsertion at the
    /// just-vacated index cannot fail.
    pub fn replace_child(&mut self, index: usize, child: Element) -> Option<Element> {
        let removed = self.remove_child(index)?;
        self.insert_child(index, child)
            .unwrap_or_else(|_| unreachable!("removal vacated the index"));
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
    #[must_use]
    pub fn to_xml(&self) -> String {
        crate::codec::serialize::write_mathml(self)
    }
}

fn invalid(message: impl Into<String>) -> Error {
    Error::InvalidFormat(message.into())
}

/// Validate an XML element or attribute local name (`NCName` subset).
fn validate_name(name: &str, kind: &str) -> Result<()> {
    let mut characters = name.chars();
    let Some(first) = characters.next() else {
        return Err(invalid(format!("MathML {kind} name is empty")));
    };
    if !(first.is_alphabetic() || first == '_')
        || !characters
            .all(|character| character.is_alphanumeric() || matches!(character, '_' | '-' | '.'))
    {
        return Err(invalid(format!("invalid MathML {kind} name '{name}'")));
    }
    Ok(())
}
