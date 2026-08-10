//! Typed semantic models for generated `OpenDocument` text indexes.

const TEXT_NAMESPACE_STR: &str = "urn:oasis:names:tc:opendocument:xmlns:text:1.0";

/// The seven generated-index families defined by `OpenDocument` Text.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
#[non_exhaustive]
pub enum TextIndexKind {
    TableOfContents,
    Illustration,
    Table,
    Object,
    User,
    Alphabetical,
    Bibliography,
}

/// A decoded attribute identified by its expanded XML name.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct TextIndexAttribute {
    pub(crate) namespace_uri: Option<String>,
    pub(crate) local_name: String,
    pub(crate) value: String,
}

impl TextIndexAttribute {
    pub fn namespace_uri(&self) -> Option<&str> {
        self.namespace_uri.as_deref()
    }

    pub fn local_name(&self) -> &str {
        &self.local_name
    }

    pub fn value(&self) -> &str {
        &self.value
    }
}

/// Ordered mixed content within an index element.
#[derive(Debug, Clone, PartialEq, Eq)]
pub enum TextIndexContent {
    Text(String),
    Element(TextIndexElement),
}

/// One namespace-aware element in an index source, template, or cached body.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct TextIndexElement {
    pub(crate) namespace_uri: Option<String>,
    pub(crate) local_name: String,
    pub(crate) attributes: Vec<TextIndexAttribute>,
    pub(crate) content: Vec<TextIndexContent>,
}

impl TextIndexElement {
    pub fn namespace_uri(&self) -> Option<&str> {
        self.namespace_uri.as_deref()
    }

    pub fn local_name(&self) -> &str {
        &self.local_name
    }

    pub fn attributes(&self) -> &[TextIndexAttribute] {
        &self.attributes
    }

    pub fn attribute(&self, namespace_uri: Option<&str>, local_name: &str) -> Option<&str> {
        self.attributes
            .iter()
            .find(|attribute| {
                attribute.namespace_uri() == namespace_uri && attribute.local_name() == local_name
            })
            .map(TextIndexAttribute::value)
    }

    pub fn content(&self) -> &[TextIndexContent] {
        &self.content
    }

    pub fn child_elements(&self) -> impl Iterator<Item = &TextIndexElement> {
        self.content.iter().filter_map(|content| match content {
            TextIndexContent::Element(element) => Some(element),
            TextIndexContent::Text(_) => None,
        })
    }

    /// Compose character content in exact document order.
    pub fn all_text(&self) -> String {
        let mut output = String::new();
        append_all_text(self, &mut output);
        output
    }
}

/// A generated index declaration and its stored, inert cached body.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct TextIndex {
    pub(crate) kind: TextIndexKind,
    pub(crate) root: TextIndexElement,
}

impl TextIndex {
    pub fn kind(&self) -> TextIndexKind {
        self.kind
    }

    pub fn root(&self) -> &TextIndexElement {
        &self.root
    }

    pub fn name(&self) -> &str {
        self.root
            .attribute(Some(TEXT_NAMESPACE_STR), "name")
            .unwrap_or_default()
    }

    pub fn protected(&self) -> bool {
        matches!(
            self.root.attribute(Some(TEXT_NAMESPACE_STR), "protected"),
            Some("true" | "1")
        )
    }

    pub fn source(&self) -> Option<&TextIndexElement> {
        self.root.child_elements().find(|element| {
            element.namespace_uri() == Some(TEXT_NAMESPACE_STR)
                && element.local_name().ends_with("-source")
        })
    }

    pub fn body(&self) -> Option<&TextIndexElement> {
        self.root.child_elements().find(|element| {
            element.namespace_uri() == Some(TEXT_NAMESPACE_STR)
                && element.local_name() == "index-body"
        })
    }
}

fn append_all_text(element: &TextIndexElement, output: &mut String) {
    for content in &element.content {
        match content {
            TextIndexContent::Text(text) => output.push_str(text),
            TextIndexContent::Element(child) => append_all_text(child, output),
        }
    }
}
