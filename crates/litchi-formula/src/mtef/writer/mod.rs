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

use crate::ast::{Fence, Formula, MathNode, MatrixFence};
use crate::mtef::{LimitKind, Limits};
use header::{patch_ole_header, reserve_ole_header, write_mtef_header};
use node::NodeWriter;
use records::{write_end, write_size_full};
use templates::{fence_template, large_op_template};

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
/// let parser = OmmlParser::new(formula.arena(), crate::omml::Limits::DEFAULT);
/// let nodes = parser.parse("<m:oMath><m:r><m:t>x</m:t></m:r></m:oMath>")?;
///
/// let writer = MtefWriter::new(Limits::DEFAULT);
/// let bytes = writer.write_nodes(&nodes)?;
/// ```
#[derive(Debug, Clone, Copy)]
pub struct MtefWriter {
    /// Options applied to every equation this writer serializes
    options: MtefWriteOptions,
    limits: Limits,
}

impl MtefWriter {
    /// Create a writer that emits the OLE equation header under `limits`.
    pub fn new(limits: Limits) -> Self {
        Self {
            options: MtefWriteOptions::default(),
            limits,
        }
    }

    /// Create a writer with explicit output options and resource policy.
    pub fn with_options(limits: Limits, options: MtefWriteOptions) -> Self {
        Self { options, limits }
    }

    /// The options this writer applies
    pub fn options(&self) -> MtefWriteOptions {
        self.options
    }

    /// The resource policy retained by this writer.
    pub fn limits(&self) -> Limits {
        self.limits
    }

    /// Serialize a formula into a new buffer
    pub fn write(&self, formula: &Formula<'_>) -> Result<Vec<u8>, MtefWriteError> {
        self.write_nodes(formula.root())
    }

    /// Serialize a slice of AST nodes into a new buffer
    pub fn write_nodes(&self, nodes: &[MathNode<'_>]) -> Result<Vec<u8>, MtefWriteError> {
        let estimate = preflight_nodes(nodes, self.limits, self.options.ole_header)?;
        let mut out = Vec::new();
        out.try_reserve(estimate)
            .map_err(|_| MtefWriteError::LimitExceeded {
                kind: LimitKind::OutputBytes,
                limit: self.limits.max_output_bytes(),
                actual: estimate,
            })?;
        self.write_nodes_into_admitted(nodes, &mut out)?;
        Ok(out)
    }

    /// Serialize a slice of AST nodes by appending to `out`
    ///
    /// Reusing one buffer across equations avoids re-allocating for each of
    /// them; the equation always starts at the buffer's current end. If
    /// encoding fails, the caller's original prefix is restored.
    pub fn write_nodes_into(
        &self,
        nodes: &[MathNode<'_>],
        out: &mut Vec<u8>,
    ) -> Result<(), MtefWriteError> {
        let estimate = preflight_nodes(nodes, self.limits, self.options.ole_header)?;
        out.try_reserve(estimate)
            .map_err(|_| MtefWriteError::LimitExceeded {
                kind: LimitKind::OutputBytes,
                limit: self.limits.max_output_bytes(),
                actual: estimate,
            })?;
        self.write_nodes_into_admitted(nodes, out)
    }

    fn write_nodes_into_admitted(
        &self,
        nodes: &[MathNode<'_>],
        out: &mut Vec<u8>,
    ) -> Result<(), MtefWriteError> {
        let original_len = out.len();
        let result = self.write_nodes_into_admitted_unchecked(nodes, out);
        if let Err(error) = result {
            // Encoding errors can occur after a low-level record has started
            // (for example a later unsupported character or a font C string
            // containing NUL).  A caller-owned buffer is an append target, so
            // preserve its original prefix on failure.
            out.truncate(original_len);
            return Err(error);
        }
        Ok(())
    }

    fn write_nodes_into_admitted_unchecked(
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
        let mut writer = NodeWriter::new(out, self.limits);
        let lines = split_lines(nodes);
        match lines.len() {
            0 | 1 => writer.write_line(lines.first().copied().unwrap_or(&[])),
            _ => writer.write_pile(&lines),
        }
    }
}

/// Admit one encoded MTEF nesting level.
///
/// Both the real writer and the preflight depth walk use this helper so the
/// selected boundary cannot drift between reservation and emission.
fn admit_depth(depth: usize, limits: Limits) -> Result<usize, MtefWriteError> {
    if depth >= limits.max_depth() {
        return Err(MtefWriteError::LimitExceeded {
            kind: LimitKind::Depth,
            limit: limits.max_depth(),
            actual: depth.saturating_add(1),
        });
    }
    Ok(depth.saturating_add(1))
}

/// Preflight an AST before allocating or emitting a byte.
///
/// The estimate is deliberately conservative: every AST item gets enough
/// room for the largest MTEF record it can lower to, and text is charged at
/// the exact five-byte CHAR-record width. This keeps `write_nodes` bounded
/// without adding a checked branch to every low-level record append.
fn preflight_nodes(
    nodes: &[MathNode<'_>],
    limits: Limits,
    ole_header: bool,
) -> Result<usize, MtefWriteError> {
    let mut state = Preflight {
        limits,
        work: 0,
        items: 0,
        bytes: (if ole_header { 28 } else { 0 }) + 7 + 1 + 2 + 1,
    };
    // Mirror the NodeWriter enter points before the byte/work preflight can
    // reserve output.  This keeps transparent AST wrappers from consuming a
    // fake depth while catching every encoded LINE/TMPL/PILE level.
    state.depth_root(nodes)?;
    state.sequence(nodes, 0)?;
    state.add(1)?; // the outer END record
    Ok(state.bytes)
}

struct Preflight {
    limits: Limits,
    work: usize,
    items: usize,
    bytes: usize,
}

impl Preflight {
    fn add(&mut self, amount: usize) -> Result<(), MtefWriteError> {
        self.bytes = self.bytes.saturating_add(amount);
        if self.bytes > self.limits.max_output_bytes() {
            return Err(MtefWriteError::LimitExceeded {
                kind: LimitKind::OutputBytes,
                limit: self.limits.max_output_bytes(),
                actual: self.bytes,
            });
        }
        Ok(())
    }

    fn charge(&mut self) -> Result<(), MtefWriteError> {
        self.charge_work(1)?;
        self.items = self.items.saturating_add(1);
        if self.items > self.limits.max_items() {
            return Err(MtefWriteError::LimitExceeded {
                kind: LimitKind::Items,
                limit: self.limits.max_items(),
                actual: self.items,
            });
        }
        Ok(())
    }

    fn charge_work(&mut self, amount: usize) -> Result<(), MtefWriteError> {
        self.work = self.work.saturating_add(amount);
        if self.work > self.limits.max_work() {
            return Err(MtefWriteError::LimitExceeded {
                kind: LimitKind::Work,
                limit: self.limits.max_work(),
                actual: self.work,
            });
        }
        Ok(())
    }

    fn string(&mut self, value: &str) -> Result<(), MtefWriteError> {
        if value.len() > self.limits.max_string_bytes() {
            return Err(MtefWriteError::LimitExceeded {
                kind: LimitKind::StringBytes,
                limit: self.limits.max_string_bytes(),
                actual: value.len(),
            });
        }
        self.charge_work(value.chars().count())?;
        Ok(())
    }

    /// Admit a string that the node writer emits one character at a time.
    fn text(&mut self, value: &str) -> Result<(), MtefWriteError> {
        if value.len() > self.limits.max_string_bytes() {
            return Err(MtefWriteError::LimitExceeded {
                kind: LimitKind::StringBytes,
                limit: self.limits.max_string_bytes(),
                actual: value.len(),
            });
        }
        let mut chars = 0;
        for ch in value.chars() {
            self.character(ch)?;
            chars += 1;
        }
        self.charge_work(chars)
    }

    /// Validate the character-width refusal that `char_code`/`char_code_in`
    /// would return, without constructing a temporary CharCode.
    fn character(&self, ch: char) -> Result<(), MtefWriteError> {
        if u32::from(ch) > u32::from(u16::MAX) {
            return Err(MtefWriteError::UnsupportedCharacter(ch));
        }
        Ok(())
    }

    /// Admit a font name before `write_font` starts its record.
    fn font(&mut self, value: &str) -> Result<(), MtefWriteError> {
        if value.as_bytes().contains(&0) {
            return Err(MtefWriteError::InvalidFontName);
        }
        self.string(value)
    }

    // ------------------------------------------------------------------
    // Exact encoded-depth preflight
    // ------------------------------------------------------------------

    /// Admit one of the same nesting levels that `NodeWriter::enter` admits.
    ///
    /// This walk is intentionally separate from the conservative byte walk
    /// below: AST wrappers such as Row and Style are transparent to the
    /// emitter, while every encoded LINE/TMPL/PILE/MATRIX call enters a real
    /// level.  Keeping the two walks separate avoids making a byte estimate's
    /// convenience depth parameter part of the resource contract.
    fn depth_enter(&self, depth: usize) -> Result<usize, MtefWriteError> {
        admit_depth(depth, self.limits)
    }

    /// Mirror `MtefWriter::write_body`, including its direct-break pile.
    fn depth_root(&self, nodes: &[MathNode<'_>]) -> Result<(), MtefWriteError> {
        if !nodes.iter().any(|node| matches!(node, MathNode::LineBreak)) {
            return self.depth_line(nodes, 0);
        }

        let pile_depth = self.depth_enter(0)?;
        let mut start = 0;
        for (index, node) in nodes.iter().enumerate() {
            if matches!(node, MathNode::LineBreak) {
                self.depth_line(&nodes[start..index], pile_depth)?;
                start = index.saturating_add(1);
            }
        }
        self.depth_line(&nodes[start..], pile_depth)
    }

    /// Mirror `NodeWriter::write_line` at the current encoded depth.
    fn depth_line(&self, nodes: &[MathNode<'_>], depth: usize) -> Result<(), MtefWriteError> {
        let line_depth = self.depth_enter(depth)?;
        self.depth_sequence(nodes, line_depth)
    }

    /// Mirror `NodeWriter::write_template` and its one LINE per slot.
    fn depth_template(
        &self,
        slots: &[&[MathNode<'_>]],
        depth: usize,
    ) -> Result<(), MtefWriteError> {
        let template_depth = self.depth_enter(depth)?;
        for slot in slots {
            self.depth_line(slot, template_depth)?;
        }
        Ok(())
    }

    /// Mirror `NodeWriter::write_matrix_record`.
    fn depth_matrix(
        &self,
        rows: &[Vec<Vec<MathNode<'_>>>],
        depth: usize,
    ) -> Result<(), MtefWriteError> {
        let matrix_depth = self.depth_enter(depth)?;
        let col_count = rows.iter().map(Vec::len).max().unwrap_or(0);
        for row in rows {
            for column in 0..col_count {
                self.depth_line(row.get(column).map_or(&[], Vec::as_slice), matrix_depth)?;
            }
        }
        Ok(())
    }

    /// Mirror the fallback used by `NodeWriter::write_word_operator`.
    fn depth_word_operator(
        &self,
        lower: &[MathNode<'_>],
        upper: &[MathNode<'_>],
        body: &[MathNode<'_>],
        depth: usize,
    ) -> Result<(), MtefWriteError> {
        let template_depth = self.depth_enter(depth)?;
        // The function name occupies the first LINE slot.
        let _name_depth = self.depth_enter(template_depth)?;
        self.depth_line(lower, template_depth)?;
        self.depth_line(upper, template_depth)?;
        // The body is written after the limit template has left its level.
        self.depth_sequence(body, depth)
    }

    /// Mirror `NodeWriter::write_ruled`/`write_run` without emitting bytes.
    fn depth_run(
        &self,
        content: &[MathNode<'_>],
        underline: Option<crate::ast::LineStyle>,
        overline: Option<crate::ast::LineStyle>,
        depth: usize,
    ) -> Result<(), MtefWriteError> {
        if overline.is_some() {
            self.depth_ruled(content, depth, underline)
        } else if underline.is_some() {
            self.depth_ruled(content, depth, None)
        } else {
            self.depth_sequence(content, depth)
        }
    }

    fn depth_ruled(
        &self,
        content: &[MathNode<'_>],
        depth: usize,
        remaining_underline: Option<crate::ast::LineStyle>,
    ) -> Result<(), MtefWriteError> {
        let template_depth = self.depth_enter(depth)?;
        let line_depth = self.depth_enter(template_depth)?;
        self.depth_run(content, remaining_underline, None, line_depth)
    }

    fn depth_sequence(&self, nodes: &[MathNode<'_>], depth: usize) -> Result<(), MtefWriteError> {
        for node in nodes {
            self.depth_node(node, depth)?;
        }
        Ok(())
    }

    fn depth_node(&self, node: &MathNode<'_>, depth: usize) -> Result<(), MtefWriteError> {
        match node {
            MathNode::Text(_)
            | MathNode::Number(_)
            | MathNode::Operator(_)
            | MathNode::Symbol(_)
            | MathNode::PredefinedSymbol(_)
            | MathNode::Space(_)
            | MathNode::LineBreak
            | MathNode::Error(_) => Ok(()),
            MathNode::Frac {
                numerator,
                denominator,
                ..
            } => self.depth_template(&[numerator.as_slice(), denominator.as_slice()], depth),
            MathNode::Root { base, index } => {
                if let Some(index) = index {
                    self.depth_template(&[base.as_slice(), index.as_slice()], depth)
                } else {
                    self.depth_template(&[base.as_slice()], depth)
                }
            },
            MathNode::Power { base, exponent } => {
                self.depth_sequence(base, depth)?;
                self.depth_template(&[&[], exponent.as_slice()], depth)
            },
            MathNode::Sub { base, subscript } => {
                self.depth_sequence(base, depth)?;
                self.depth_template(&[subscript.as_slice(), &[]], depth)
            },
            MathNode::SubSup {
                base,
                subscript,
                superscript,
            } => {
                self.depth_sequence(base, depth)?;
                self.depth_template(&[subscript.as_slice(), superscript.as_slice()], depth)
            },
            MathNode::PreSub {
                base,
                pre_subscript,
            }
            | MathNode::PreSup {
                base,
                pre_superscript: pre_subscript,
            } => {
                self.depth_sequence(pre_subscript, depth)?;
                self.depth_sequence(base, depth)
            },
            MathNode::PreSubSup {
                base,
                pre_subscript,
                pre_superscript,
            } => {
                self.depth_sequence(pre_subscript, depth)?;
                self.depth_sequence(pre_superscript, depth)?;
                self.depth_sequence(base, depth)
            },
            MathNode::Under { base, under, .. }
            | MathNode::Over {
                base, over: under, ..
            } => self.depth_template(&[base.as_slice(), under.as_slice(), &[]], depth),
            MathNode::UnderOver {
                base, under, over, ..
            } => self.depth_template(&[base.as_slice(), under.as_slice(), over.as_slice()], depth),
            MathNode::Fenced {
                open,
                content,
                close,
                ..
            } => match fence_template(*open, *close) {
                Some(_) => self.depth_template(&[content.as_slice()], depth),
                None => self.depth_sequence(content, depth),
            },
            MathNode::LargeOp {
                operator,
                lower_limit,
                upper_limit,
                integrand,
                hide_lower,
                hide_upper,
            } => {
                let lower = match (lower_limit.as_deref(), *hide_lower) {
                    (Some(nodes), false) => nodes,
                    _ => &[],
                };
                let upper = match (upper_limit.as_deref(), *hide_upper) {
                    (Some(nodes), false) => nodes,
                    _ => &[],
                };
                let body = integrand.as_deref().unwrap_or(&[]);
                if large_op_template(*operator, !lower.is_empty() || !upper.is_empty()).is_some() {
                    self.depth_template(&[body, lower, upper], depth)
                } else {
                    self.depth_word_operator(lower, upper, body, depth)
                }
            },
            MathNode::Function { argument, .. } | MathNode::PredefinedFunction { argument, .. } => {
                self.depth_sequence(argument, depth)
            },
            MathNode::Matrix {
                rows, fence_type, ..
            } => {
                let fence = match fence_type {
                    MatrixFence::None => Fence::None,
                    MatrixFence::Paren => Fence::Paren,
                    MatrixFence::Bracket => Fence::Bracket,
                    MatrixFence::Brace => Fence::Brace,
                    MatrixFence::Pipe => Fence::Pipe,
                    MatrixFence::DoublePipe => Fence::DoublePipe,
                };
                if fence_template(fence, fence).is_some() {
                    let fence_depth = self.depth_enter(depth)?;
                    let line_depth = self.depth_enter(fence_depth)?;
                    self.depth_matrix(rows, line_depth)
                } else {
                    self.depth_matrix(rows, depth)
                }
            },
            MathNode::EqArray { rows, .. } => {
                let pile_depth = self.depth_enter(depth)?;
                for row in rows {
                    self.depth_line(row, pile_depth)?;
                }
                Ok(())
            },
            MathNode::Accent { base, .. } | MathNode::BorderBox { content: base, .. } => {
                self.depth_sequence(base, depth)
            },
            MathNode::Bar { base, .. } => self.depth_template(&[base.as_slice()], depth),
            MathNode::GroupChar { base, .. } => self.depth_template(&[base.as_slice(), &[]], depth),
            MathNode::Style { content, .. } => self.depth_sequence(content, depth),
            MathNode::Run {
                content,
                underline,
                overline,
                ..
            } => self.depth_run(content, *underline, *overline, depth),
            MathNode::Row(content) => self.depth_sequence(content, depth),
            MathNode::Phantom(content)
            | MathNode::Degree(content)
            | MathNode::Base(content)
            | MathNode::Argument(content)
            | MathNode::Numerator(content)
            | MathNode::Denominator(content)
            | MathNode::Integrand(content)
            | MathNode::LowerLimit(content)
            | MathNode::UpperLimit(content)
            | MathNode::Limit { content, .. } => self.depth_sequence(content, depth),
        }
    }

    fn sequence(&mut self, nodes: &[MathNode<'_>], depth: usize) -> Result<(), MtefWriteError> {
        for node in nodes {
            self.charge()?;
            self.node(node, depth)?;
        }
        Ok(())
    }

    fn child(&mut self, nodes: &[MathNode<'_>], depth: usize) -> Result<(), MtefWriteError> {
        self.add(4)?; // LINE start/end and the conservative template framing
        self.sequence(nodes, depth.saturating_add(1))
    }

    fn node(&mut self, node: &MathNode<'_>, depth: usize) -> Result<(), MtefWriteError> {
        const CHAR: usize = 5;
        const STRUCTURE: usize = 64;
        match node {
            MathNode::Text(value) | MathNode::Number(value) => {
                self.text(value)?;
                self.add(value.chars().count().saturating_mul(CHAR))?;
            },
            MathNode::Symbol(symbol) => {
                if let Some(ch) = symbol.unicode {
                    self.character(ch)?;
                    self.add(CHAR)?;
                } else {
                    self.text(&symbol.name)?;
                    self.add(symbol.name.chars().count().saturating_mul(CHAR))?;
                }
            },
            MathNode::Operator(_) | MathNode::PredefinedSymbol(_) | MathNode::Space(_) => {
                self.add(CHAR)?;
            },
            MathNode::Frac {
                numerator,
                denominator,
                ..
            } => {
                self.add(STRUCTURE)?;
                self.child(numerator, depth)?;
                self.child(denominator, depth)?;
            },
            MathNode::Root { base, index } => {
                self.add(STRUCTURE)?;
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
                self.add(STRUCTURE)?;
                self.sequence(base, depth.saturating_add(1))?;
                self.child(exponent, depth)?;
            },
            MathNode::SubSup {
                base,
                subscript,
                superscript,
            } => {
                self.add(STRUCTURE)?;
                self.sequence(base, depth.saturating_add(1))?;
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
                self.add(STRUCTURE)?;
                self.child(pre_subscript, depth)?;
                self.sequence(base, depth.saturating_add(1))?;
            },
            MathNode::PreSubSup {
                base,
                pre_subscript,
                pre_superscript,
            } => {
                self.add(STRUCTURE)?;
                self.child(pre_subscript, depth)?;
                self.child(pre_superscript, depth)?;
                self.sequence(base, depth.saturating_add(1))?;
            },
            MathNode::Under { base, under, .. }
            | MathNode::Over {
                base, over: under, ..
            } => {
                self.add(STRUCTURE)?;
                self.child(base, depth)?;
                self.child(under, depth)?;
            },
            MathNode::UnderOver {
                base, under, over, ..
            } => {
                self.add(STRUCTURE)?;
                self.child(base, depth)?;
                self.child(under, depth)?;
                self.child(over, depth)?;
            },
            MathNode::Fenced {
                content, separator, ..
            } => {
                self.add(STRUCTURE)?;
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
                self.add(STRUCTURE.saturating_mul(2))?;
                if let Some(nodes) = lower_limit {
                    self.child(nodes, depth)?;
                }
                if let Some(nodes) = upper_limit {
                    self.child(nodes, depth)?;
                }
                if let Some(nodes) = integrand {
                    self.sequence(nodes, depth.saturating_add(1))?;
                }
            },
            MathNode::Function { name, argument } => {
                self.text(name)?;
                self.add(name.chars().count().saturating_mul(CHAR))?;
                self.sequence(argument, depth.saturating_add(1))?;
            },
            MathNode::PredefinedFunction { argument, .. } => {
                self.add(STRUCTURE)?;
                self.sequence(argument, depth.saturating_add(1))?;
            },
            MathNode::Matrix { rows, .. } => {
                let row_count = rows.len();
                let col_count = rows.iter().map(Vec::len).max().unwrap_or(0);
                if row_count > 63 || col_count > 63 {
                    return Err(MtefWriteError::MatrixTooLarge {
                        rows: row_count,
                        cols: col_count,
                        limit: 63,
                    });
                }
                self.add(
                    STRUCTURE.saturating_add(row_count.saturating_mul(col_count).saturating_mul(4)),
                )?;
                for row in rows {
                    for col in 0..col_count {
                        if let Some(cell) = row.get(col) {
                            self.child(cell, depth)?;
                        } else {
                            self.add(4)?;
                        }
                    }
                }
            },
            MathNode::EqArray { rows, .. } => {
                self.add(STRUCTURE)?;
                for row in rows {
                    self.child(row, depth)?;
                }
            },
            MathNode::Accent { base, .. }
            | MathNode::Bar { base, .. }
            | MathNode::BorderBox { content: base, .. }
            | MathNode::GroupChar { base, .. } => {
                self.add(STRUCTURE)?;
                self.sequence(base, depth.saturating_add(1))?;
            },
            MathNode::Style { content, .. } => self.sequence(content, depth.saturating_add(1))?,
            MathNode::Run {
                content,
                font,
                underline,
                overline,
                ..
            } => {
                if let Some(font) = font {
                    self.font(font)?;
                    self.add(4usize.saturating_add(font.len()))?;
                }
                if underline.is_some() {
                    self.add(STRUCTURE)?;
                }
                if overline.is_some() {
                    self.add(STRUCTURE)?;
                }
                self.sequence(content, depth.saturating_add(1))?;
            },
            MathNode::Row(content) => {
                self.sequence(content, depth.saturating_add(1))?;
            },
            MathNode::Phantom(content)
            | MathNode::Degree(content)
            | MathNode::Base(content)
            | MathNode::Argument(content)
            | MathNode::Numerator(content)
            | MathNode::Denominator(content)
            | MathNode::Integrand(content)
            | MathNode::LowerLimit(content)
            | MathNode::UpperLimit(content)
            | MathNode::Limit { content, .. } => {
                self.sequence(content, depth.saturating_add(1))?;
            },
            // Only direct root breaks reach `write_body`'s split_lines path;
            // nested break nodes are ignored by NodeWriter as ordinary leaves.
            // A top-level break is lowered to a PILE with two empty LINE
            // records around it.  The PILE start (4), two LINE frames (3
            // each), and the PILE end (1) cost eight bytes in addition to the
            // outer stream framing; charging seven used to admit one byte too
            // few and let the emitter exceed max_output_bytes.
            MathNode::LineBreak if depth == 0 => self.add(8)?,
            MathNode::LineBreak => {},
            MathNode::Error(message) => self.string(message)?,
        }
        Ok(())
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
