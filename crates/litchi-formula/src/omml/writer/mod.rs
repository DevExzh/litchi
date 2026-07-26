// OMML writer: serializes the formula AST to Office Math Markup Language
//
// This is the inverse of `OmmlParser`. Output follows ECMA-376 Part 1, §22.1
// (m:oMath, m:f, m:sSup/m:sSub/m:sSubSup/m:sPre, m:rad, m:nary, m:d, m:func,
// m:m, m:eqArr, m:acc, m:bar, m:groupChr, m:limLow/m:limUpp, m:r/m:t) so
// serialized formulas can be embedded in OOXML documents and re-parsed with
// `OmmlParser` without losing structure.

/// AST-to-OMML character and value mappings
mod chars;
/// OMML element and attribute name constants
mod names;
/// Per-node serialization logic
mod node;

#[cfg(test)]
mod tests;

use crate::ast::{Formula, MathNode};
use crate::omml::error::OmmlError;
use names::{ATTR_XMLNS_M, EL_OMATH, OMML_NAMESPACE};

/// Initial output buffer capacity (typical formulas serialize well below this)
const INITIAL_BUFFER_CAPACITY: usize = 1024;

/// Serializer that converts a formula AST to an OMML XML string
///
/// # Example
/// ```ignore
/// let formula = Formula::new();
/// let parser = OmmlParser::new(formula.arena());
/// let nodes = parser.parse("<m:oMath><m:r><m:t>x</m:t></m:r></m:oMath>")?;
///
/// let mut writer = OmmlWriter::new();
/// let xml = writer.write_nodes(&nodes)?;
/// ```
pub struct OmmlWriter {
    /// Buffer holding the serialized XML
    buffer: String,
}

impl OmmlWriter {
    /// Create a new OMML writer
    pub fn new() -> Self {
        Self {
            buffer: String::with_capacity(INITIAL_BUFFER_CAPACITY),
        }
    }

    /// Serialize a formula to an `m:oMath` fragment
    ///
    /// The returned string borrows the writer's internal buffer and is valid
    /// until the next call to a `write*` method.
    pub fn write(&mut self, formula: &Formula) -> Result<&str, OmmlError> {
        self.write_nodes(formula.root())
    }

    /// Serialize a slice of AST nodes to an `m:oMath` fragment
    pub fn write_nodes(&mut self, nodes: &[MathNode]) -> Result<&str, OmmlError> {
        self.buffer.clear();

        self.buffer.push('<');
        self.buffer.push_str(EL_OMATH);
        self.push_attr(ATTR_XMLNS_M, OMML_NAMESPACE);
        self.buffer.push('>');

        self.write_all(nodes)?;

        self.close_element(EL_OMATH);
        Ok(&self.buffer)
    }

    // ------------------------------------------------------------------
    // XML emission primitives
    // ------------------------------------------------------------------

    /// Write a sequence of nodes
    pub(super) fn write_all(&mut self, nodes: &[MathNode]) -> Result<(), OmmlError> {
        for node in nodes {
            self.write_node(node)?;
        }
        Ok(())
    }

    /// Emit `<name>`
    fn open_element(&mut self, name: &str) {
        self.buffer.push('<');
        self.buffer.push_str(name);
        self.buffer.push('>');
    }

    /// Emit `</name>`
    fn close_element(&mut self, name: &str) {
        self.buffer.push_str("</");
        self.buffer.push_str(name);
        self.buffer.push('>');
    }

    /// Emit `<name/>`
    fn empty_element(&mut self, name: &str) {
        self.buffer.push('<');
        self.buffer.push_str(name);
        self.buffer.push_str("/>");
    }

    /// Emit `<name m:val="value"/>`
    fn val_element(&mut self, name: &str, value: &str) {
        self.buffer.push('<');
        self.buffer.push_str(name);
        self.push_attr(names::ATTR_VAL, value);
        self.buffer.push_str("/>");
    }

    /// Emit ` key="escaped-value"` (attribute inside an open start tag)
    fn push_attr(&mut self, key: &str, value: &str) {
        self.buffer.push(' ');
        self.buffer.push_str(key);
        self.buffer.push_str("=\"");
        push_escaped(&mut self.buffer, value, true);
        self.buffer.push('"');
    }

    /// Emit escaped character data
    fn push_text(&mut self, text: &str) {
        push_escaped(&mut self.buffer, text, false);
    }

    /// Emit an element wrapping the output of `body`
    fn element<F>(&mut self, name: &str, body: F) -> Result<(), OmmlError>
    where
        F: FnOnce(&mut Self) -> Result<(), OmmlError>,
    {
        self.open_element(name);
        body(self)?;
        self.close_element(name);
        Ok(())
    }

    /// Emit `<m:e>...</m:e>` around a node sequence (self-closing when empty)
    fn base_element(&mut self, nodes: &[MathNode]) -> Result<(), OmmlError> {
        self.wrapped_nodes(names::EL_ELEMENT, nodes)
    }

    /// Emit `<name>...</name>` around a node sequence (self-closing when empty)
    fn wrapped_nodes(&mut self, name: &str, nodes: &[MathNode]) -> Result<(), OmmlError> {
        if nodes.is_empty() {
            self.empty_element(name);
            Ok(())
        } else {
            self.element(name, |w| w.write_all(nodes))
        }
    }

    /// Emit `<m:r><m:t>text</m:t></m:r>`
    fn text_run(&mut self, text: &str) {
        self.open_element(names::EL_RUN);
        self.open_element(names::EL_TEXT);
        self.push_text(text);
        self.close_element(names::EL_TEXT);
        self.close_element(names::EL_RUN);
    }
}

impl Default for OmmlWriter {
    fn default() -> Self {
        Self::new()
    }
}

/// Escape XML character data; additionally escapes quotes in attribute values
fn push_escaped(buffer: &mut String, value: &str, is_attribute: bool) {
    for ch in value.chars() {
        match ch {
            '&' => buffer.push_str("&amp;"),
            '<' => buffer.push_str("&lt;"),
            '>' => buffer.push_str("&gt;"),
            '"' if is_attribute => buffer.push_str("&quot;"),
            '\'' if is_attribute => buffer.push_str("&apos;"),
            _ => buffer.push(ch),
        }
    }
}
