//! Retained vendor extensions for typed ODF chart authoring.

use litchi_odf_common::chart::Element;

/// An extension attribute retained by expanded XML name.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct ExtensionAttribute {
    pub namespace_uri: Option<String>,
    pub local_name: String,
    pub value: String,
}

/// An extension subtree retained without interpreting vendor behavior.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct ExtensionElement {
    pub namespace_uri: Option<String>,
    pub local_name: String,
    pub attributes: Vec<ExtensionAttribute>,
    pub text: String,
    pub children: Vec<ExtensionElement>,
}

impl ExtensionElement {
    /// Clone a retained read-only element into an owned extension subtree.
    pub fn from_retained(element: &Element) -> Self {
        Self {
            namespace_uri: element.namespace_uri().map(str::to_string),
            local_name: element.local_name().to_string(),
            attributes: element
                .attributes()
                .iter()
                .map(|attribute| ExtensionAttribute {
                    namespace_uri: attribute.namespace_uri().map(str::to_string),
                    local_name: attribute.local_name().to_string(),
                    value: attribute.value().to_string(),
                })
                .collect(),
            text: element.text().to_string(),
            children: element.children().iter().map(Self::from_retained).collect(),
        }
    }
}

/// Unknown attributes and child elements attached to a typed chart node.
#[derive(Debug, Clone, Default, PartialEq, Eq)]
pub struct Extensions {
    pub attributes: Vec<ExtensionAttribute>,
    pub children: Vec<ExtensionElement>,
}
