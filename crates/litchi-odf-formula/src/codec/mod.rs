//! Bounded MathML parsing and canonical XML serialization.

use litchi_core::Result;

use crate::model::Element;

/// Parse one MathML `math` root into the inert formula model.
pub fn parse(xml: &str) -> Result<Element> {
    crate::migration::document::parse_mathml(xml)
}

/// Serialize a formula model as self-contained, well-formed MathML XML.
pub fn serialize(root: &Element) -> String {
    root.to_xml()
}
