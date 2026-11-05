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
use crate::omml::limits::{LimitKind, Limits};
use names::{ATTR_XMLNS_M, EL_OMATH, OMML_NAMESPACE};

/// Serializer that converts a formula AST to an OMML XML string
///
/// # Example
/// ```ignore
/// let formula = Formula::new();
/// let parser = OmmlParser::new(formula.arena(), crate::omml::Limits::DEFAULT);
/// let nodes = parser.parse("<m:oMath><m:r><m:t>x</m:t></m:r></m:oMath>")?;
///
/// let writer = OmmlWriter::new(crate::omml::Limits::DEFAULT);
/// let xml = writer.write_nodes(&nodes)?;
/// ```
pub struct OmmlWriter {
    /// Caller-selected serialization policy.
    limits: Limits,
}

impl OmmlWriter {
    /// Create a new OMML writer under `limits`.
    pub fn new(limits: Limits) -> Self {
        Self { limits }
    }

    /// The resource policy retained by this writer.
    pub fn limits(&self) -> Limits {
        self.limits
    }

    /// Serialize a formula to an owned `m:oMath` fragment.
    pub fn write(&self, formula: &Formula) -> Result<String, OmmlError> {
        self.write_nodes(formula.root())
    }

    /// Serialize a slice of AST nodes to an owned `m:oMath` fragment.
    pub fn write_nodes(&self, nodes: &[MathNode]) -> Result<String, OmmlError> {
        let estimate = preflight_nodes(nodes, self.limits)?;
        let mut output = Output {
            buffer: String::new(),
        };
        output
            .buffer
            .try_reserve(estimate)
            .map_err(|_| OmmlError::LimitExceeded {
                kind: LimitKind::OutputBytes,
                limit: self.limits.max_output_bytes(),
                actual: estimate,
            })?;

        output.buffer.push('<');
        output.buffer.push_str(EL_OMATH);
        output.push_attr(ATTR_XMLNS_M, OMML_NAMESPACE);
        output.buffer.push('>');

        output.write_all(nodes)?;

        output.close_element(EL_OMATH);
        Ok(output.buffer)
    }
}

struct Output {
    buffer: String,
}

impl Output {
    // ------------------------------------------------------------------
    // XML emission primitives
    // ------------------------------------------------------------------

    /// Write a sequence of nodes
    fn write_all(&mut self, nodes: &[MathNode]) -> Result<(), OmmlError> {
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

/// Preflight an AST before allocating or emitting XML.
fn preflight_nodes(nodes: &[MathNode<'_>], limits: Limits) -> Result<usize, OmmlError> {
    let mut state = Preflight {
        limits,
        work: 0,
        items: 0,
        bytes: 128, // root element, namespace declaration and closing tag
    };
    state.sequence(nodes, 0)?;
    Ok(state.bytes)
}

struct Preflight {
    limits: Limits,
    work: usize,
    items: usize,
    bytes: usize,
}

impl Preflight {
    fn add(&mut self, amount: usize) -> Result<(), OmmlError> {
        self.bytes = self.bytes.saturating_add(amount);
        if self.bytes > self.limits.max_output_bytes() {
            return Err(OmmlError::LimitExceeded {
                kind: LimitKind::OutputBytes,
                limit: self.limits.max_output_bytes(),
                actual: self.bytes,
            });
        }
        Ok(())
    }

    fn charge(&mut self) -> Result<(), OmmlError> {
        self.charge_work(1)?;
        self.items = self.items.saturating_add(1);
        if self.items > self.limits.max_items() {
            return Err(OmmlError::LimitExceeded {
                kind: LimitKind::Items,
                limit: self.limits.max_items(),
                actual: self.items,
            });
        }
        Ok(())
    }

    fn charge_work(&mut self, amount: usize) -> Result<(), OmmlError> {
        self.work = self.work.saturating_add(amount);
        if self.work > self.limits.max_work() {
            return Err(OmmlError::LimitExceeded {
                kind: LimitKind::Work,
                limit: self.limits.max_work(),
                actual: self.work,
            });
        }
        Ok(())
    }

    fn string(&mut self, value: &str) -> Result<(), OmmlError> {
        if value.len() > self.limits.max_string_bytes() {
            return Err(OmmlError::LimitExceeded {
                kind: LimitKind::StringBytes,
                limit: self.limits.max_string_bytes(),
                actual: value.len(),
            });
        }
        self.charge_work(value.chars().count())?;
        // XML escaping expands each input byte by at most six bytes.
        self.add(value.len().saturating_mul(6))
    }

    fn sequence(&mut self, nodes: &[MathNode<'_>], depth: usize) -> Result<(), OmmlError> {
        if depth > self.limits.max_depth() {
            return Err(OmmlError::LimitExceeded {
                kind: LimitKind::Depth,
                limit: self.limits.max_depth(),
                actual: depth,
            });
        }
        for node in nodes {
            self.charge()?;
            self.node(node, depth)?;
        }
        Ok(())
    }

    fn child(&mut self, nodes: &[MathNode<'_>], depth: usize) -> Result<(), OmmlError> {
        self.add(64)?;
        self.sequence(nodes, depth.saturating_add(1))
    }

    fn node(&mut self, node: &MathNode<'_>, depth: usize) -> Result<(), OmmlError> {
        match node {
            MathNode::Text(value) | MathNode::Number(value) => {
                self.add(48)?;
                self.string(value)?;
            },
            MathNode::Operator(_) | MathNode::PredefinedSymbol(_) | MathNode::Space(_) => {
                self.add(64)?;
            },
            MathNode::Symbol(symbol) => {
                self.add(48)?;
                if let Some(unicode) = symbol.unicode {
                    self.string(unicode.encode_utf8(&mut [0; 4]))?;
                } else {
                    self.string(&symbol.name)?;
                }
            },
            MathNode::Frac {
                numerator,
                denominator,
                ..
            } => {
                self.add(128)?;
                self.child(numerator, depth)?;
                self.child(denominator, depth)?;
            },
            MathNode::Root { base, index } => {
                self.add(128)?;
                self.child(base, depth)?;
                if let Some(index) = index {
                    self.child(index, depth)?;
                }
            },
            MathNode::Power { base, exponent }
            | MathNode::Sub {
                base,
                subscript: exponent,
            } => {
                self.add(96)?;
                self.child(base, depth)?;
                self.child(exponent, depth)?;
            },
            MathNode::SubSup {
                base,
                subscript,
                superscript,
            } => {
                self.add(128)?;
                self.child(base, depth)?;
                self.child(subscript, depth)?;
                self.child(superscript, depth)?;
            },
            MathNode::PreSub {
                base,
                pre_subscript,
            }
            | MathNode::PreSup {
                base,
                pre_superscript: pre_subscript,
            } => {
                self.add(128)?;
                self.child(base, depth)?;
                self.child(pre_subscript, depth)?;
            },
            MathNode::PreSubSup {
                base,
                pre_subscript,
                pre_superscript,
            } => {
                self.add(160)?;
                self.child(base, depth)?;
                self.child(pre_subscript, depth)?;
                self.child(pre_superscript, depth)?;
            },
            MathNode::Under { base, under, .. }
            | MathNode::Over {
                base, over: under, ..
            } => {
                self.add(128)?;
                self.child(base, depth)?;
                self.child(under, depth)?;
            },
            MathNode::UnderOver {
                base, under, over, ..
            } => {
                self.add(160)?;
                self.child(base, depth)?;
                self.child(under, depth)?;
                self.child(over, depth)?;
            },
            MathNode::Fenced {
                content, separator, ..
            } => {
                self.add(128)?;
                if let Some(separator) = separator {
                    self.string(separator)?;
                }
                self.child(content, depth)?;
            },
            MathNode::LargeOp {
                lower_limit,
                upper_limit,
                integrand,
                ..
            } => {
                self.add(192)?;
                if let Some(nodes) = lower_limit {
                    self.child(nodes, depth)?;
                }
                if let Some(nodes) = upper_limit {
                    self.child(nodes, depth)?;
                }
                if let Some(nodes) = integrand {
                    self.child(nodes, depth)?;
                }
            },
            MathNode::Function { name, argument } => {
                self.add(64)?;
                self.string(name)?;
                self.child(argument, depth)?;
            },
            MathNode::PredefinedFunction { argument, .. } => {
                self.add(64)?;
                self.child(argument, depth)?;
            },
            MathNode::Matrix {
                rows, properties, ..
            } => {
                self.add(256)?;
                if let Some(properties) = properties
                    && let Some(rule) = properties.row_spacing
                {
                    // Numeric matrix properties are formatted into a short
                    // XML value by the emitter; charge their source width.
                    self.string(&rule.to_string())?;
                }
                for row in rows {
                    for cell in row {
                        self.child(cell, depth)?;
                    }
                }
            },
            MathNode::EqArray { rows, properties } => {
                self.add(320)?;
                if let Some(properties) = properties
                    && let Some(rule) = properties.row_spacing_rule.as_deref()
                {
                    self.string(rule)?;
                }
                for row in rows {
                    self.child(row, depth)?;
                }
            },
            MathNode::Accent { base, .. }
            | MathNode::Bar { base, .. }
            | MathNode::GroupChar { base, .. } => {
                self.add(128)?;
                self.child(base, depth)?;
            },
            MathNode::BorderBox { content, .. } => {
                // A BorderBox can carry eight independent property elements;
                // reserve enough static space for all of them before output.
                self.add(320)?;
                self.child(content, depth)?;
            },
            MathNode::Style { content, .. } => {
                self.add(64)?;
                self.sequence(content, depth.saturating_add(1))?;
            },
            MathNode::Run {
                content,
                font,
                color,
                ..
            } => {
                self.add(96)?;
                if let Some(font) = font {
                    self.string(font)?;
                }
                if let Some(color) = color {
                    self.string(color)?;
                }
                self.sequence(content, depth.saturating_add(1))?;
            },
            MathNode::Row(content) => {
                self.add(64)?;
                self.sequence(content, depth.saturating_add(1))?;
            },
            MathNode::Phantom(content)
            | MathNode::Limit { content, .. }
            | MathNode::Degree(content)
            | MathNode::Base(content)
            | MathNode::Argument(content)
            | MathNode::Numerator(content)
            | MathNode::Denominator(content)
            | MathNode::Integrand(content)
            | MathNode::LowerLimit(content)
            | MathNode::UpperLimit(content) => {
                self.add(64)?;
                self.sequence(content, depth.saturating_add(1))?;
            },
            MathNode::LineBreak => self.add(64)?,
            MathNode::Error(message) => {
                self.string(message)?;
                self.add(64)?;
            },
        }
        Ok(())
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
