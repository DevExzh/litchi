// MTEF writer: serializes the formula AST to the binary MathType Equation Format
//
// This is the inverse of `MtefParser`. Output is MTEF 5, the version MathType 5
// and later write, optionally prefixed with the 28-byte OLE equation header that
// introduces an `Equation.3` object's `Equation Native` stream.
//
// The emitted stream has the shape MathType produces:
//
//     [OLE header]  MTEF header  FULL  LINE <records> END  END
//
// References:
// - http://rtf2latex2e.sourceforge.net/MTEF5.html
// - LibreOffice `starmath/source/mathtype.cxx` (MathType import and export)

// Only `mtef::mod` can name these items until `lib.rs` re-exports them at the
// crate root, so the compiler cannot yet see a user outside the module tree.
#![allow(dead_code)]

/// AST character, symbol and operator to typeface/MTCode mapping
mod charmap;
/// Writer error type
mod error;
/// OLE and MTEF header emission
mod header;
/// Per-node record lowering
mod node;
/// Low-level record encoders
mod records;
/// AST construct to template selector/variation mapping
mod templates;

#[cfg(test)]
mod tests;

pub use error::MtefWriteError;

use crate::ast::{Formula, MathNode};
use header::{patch_ole_header, reserve_ole_header, write_mtef_header};
use node::NodeWriter;
use records::{write_end, write_size_full};

/// Initial output buffer capacity (most equations encode well below this)
const INITIAL_BUFFER_CAPACITY: usize = 512;

/// Options controlling how an equation is serialized
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct MtefWriteOptions {
    /// Emit the 28-byte OLE equation header before the MTEF data
    ///
    /// Enable this for the `Equation Native` stream of an embedded
    /// `Equation.3` object (and for `\x01Ole10Native`-style payloads); disable
    /// it to obtain a bare MTEF stream.
    pub ole_header: bool,

    /// Mark the equation as inline rather than display material
    pub inline: bool,
}

impl Default for MtefWriteOptions {
    fn default() -> Self {
        Self {
            ole_header: true,
            inline: false,
        }
    }
}

impl MtefWriteOptions {
    /// Options for a bare MTEF stream with no OLE equation header
    pub fn without_ole_header() -> Self {
        Self {
            ole_header: false,
            ..Self::default()
        }
    }

    /// Return these options with the OLE header setting replaced
    #[must_use]
    pub fn with_ole_header(mut self, ole_header: bool) -> Self {
        self.ole_header = ole_header;
        self
    }

    /// Return these options with the inline flag replaced
    #[must_use]
    pub fn with_inline(mut self, inline: bool) -> Self {
        self.inline = inline;
        self
    }
}

/// Serializer that converts a formula AST to MTEF 5
///
/// Constructs that MTEF cannot express degrade to their content rather than
/// failing, so serialization only errors on hard format limits (see
/// [`MtefWriteError`]).
///
/// # Example
/// ```ignore
/// let formula = Formula::new();
/// let parser = OmmlParser::new(formula.arena());
/// let nodes = parser.parse("<m:oMath><m:r><m:t>x</m:t></m:r></m:oMath>")?;
///
/// let writer = MtefWriter::new();
/// let bytes = writer.write_nodes(&nodes)?;
/// ```
#[derive(Debug, Clone, Copy, Default)]
pub struct MtefWriter {
    /// Options applied to every equation this writer serializes
    options: MtefWriteOptions,
}

impl MtefWriter {
    /// Create a writer that emits the OLE equation header
    pub fn new() -> Self {
        Self::default()
    }

    /// Create a writer with explicit options
    pub fn with_options(options: MtefWriteOptions) -> Self {
        Self { options }
    }

    /// The options this writer applies
    pub fn options(&self) -> MtefWriteOptions {
        self.options
    }

    /// Serialize a formula into a new buffer
    pub fn write(&self, formula: &Formula<'_>) -> Result<Vec<u8>, MtefWriteError> {
        self.write_nodes(formula.root())
    }

    /// Serialize a slice of AST nodes into a new buffer
    pub fn write_nodes(&self, nodes: &[MathNode<'_>]) -> Result<Vec<u8>, MtefWriteError> {
        let mut out = Vec::with_capacity(INITIAL_BUFFER_CAPACITY);
        self.write_nodes_into(nodes, &mut out)?;
        Ok(out)
    }

    /// Serialize a slice of AST nodes by appending to `out`
    ///
    /// Reusing one buffer across equations avoids re-allocating for each of
    /// them; the equation always starts at the buffer's current end.
    pub fn write_nodes_into(
        &self,
        nodes: &[MathNode<'_>],
        out: &mut Vec<u8>,
    ) -> Result<(), MtefWriteError> {
        let header_offset = self.options.ole_header.then(|| reserve_ole_header(out));
        let payload_start = out.len();

        write_mtef_header(out, self.options.inline);

        // A size record precedes the equation body: MathType expects the
        // equation's base size to be selected before any character is drawn.
        write_size_full(out);
        self.write_body(nodes, out)?;
        write_end(out);

        if let Some(offset) = header_offset {
            let payload_len = out.len() - payload_start;
            patch_ole_header(out, offset, payload_len)?;
        }
        Ok(())
    }

    /// Write the equation body, splitting it into a pile at explicit breaks
    ///
    /// MTEF has no in-line break: a multi-line equation is a pile of lines, so
    /// `MathNode::LineBreak` becomes a line boundary.
    fn write_body(&self, nodes: &[MathNode<'_>], out: &mut Vec<u8>) -> Result<(), MtefWriteError> {
        let mut writer = NodeWriter::new(out);
        let lines = split_lines(nodes);
        match lines.len() {
            0 | 1 => writer.write_line(lines.first().copied().unwrap_or(&[])),
            _ => writer.write_pile(&lines),
        }
    }
}

/// Split a node sequence at `MathNode::LineBreak` boundaries
fn split_lines<'n, 'a>(nodes: &'n [MathNode<'a>]) -> Vec<&'n [MathNode<'a>]> {
    if !nodes.iter().any(|node| matches!(node, MathNode::LineBreak)) {
        return vec![nodes];
    }
    nodes
        .split(|node| matches!(node, MathNode::LineBreak))
        .collect()
}
