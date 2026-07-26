//! `MathNode` to MTEF record lowering
//!
//! Records are appended to a single caller-owned buffer; nothing is buffered on
//! the side. Structured nodes become TMPL records whose slots are LINE records,
//! in the slot order the MTEF template table defines:
//!
//! | Construct        | Template            | Slots                       |
//! |------------------|---------------------|-----------------------------|
//! | fraction         | `TMPL_FRACT`        | numerator, denominator      |
//! | root             | `TMPL_ROOT`         | radicand, degree            |
//! | scripts          | `TMPL_SUB`/`SUP`/`SUBSUP` | subscript, superscript (base precedes the record) |
//! | limits           | `TMPL_LIM`          | base, lower, upper          |
//! | large operators  | `TMPL_SUM`, ...     | integrand, lower, upper     |
//! | fences           | `TMPL_PAREN`, ...   | content                     |
//!
//! Constructs the format cannot express (phantoms, border boxes, pre-scripts)
//! degrade to their content rather than failing, and `MathNode::Error` is
//! dropped.

use super::charmap::{
    CharCode, accent_embellishment, char_code, char_code_in, function_name, operator_code,
    predefined_symbol_code, space_code, style_typeface,
};
use super::error::MtefWriteError;
use super::records::{
    set_char_embellished, write_char, write_embell, write_end, write_font, write_line_start,
    write_matrix_start, write_pile_start, write_template_start,
};
use super::templates::{
    Template, bar_template, fence_template, fraction_template, group_char_template,
    large_op_template, large_op_word, root_template, script_template,
};
use crate::ast::{Fence, LargeOperator, LineStyle, MathNode, MatrixFence, StyleType, Symbol};
use crate::mtef::constants::*;

/// Deepest AST nesting the writer will follow before giving up
///
/// Bounded so that a pathological tree returns an error instead of exhausting
/// the stack. MathType's own editor stops far below this.
pub(super) const MAX_NESTING_DEPTH: usize = 64;

/// Font style byte written with a FONT record for an unstyled font
const FONT_STYLE_PLAIN: u8 = 0;

/// Lowers AST nodes into MTEF records
pub(super) struct NodeWriter<'a> {
    /// Buffer records are appended to
    out: &'a mut Vec<u8>,
    /// Current recursion depth
    depth: usize,
    /// Offset of the options byte of the most recently written CHAR record
    ///
    /// Embellishments are attached by setting a flag on the character they
    /// decorate, which has already been written by the time we see the accent.
    last_char_options: Option<usize>,
    /// Typeface imposed by an enclosing style run, if any
    typeface: Option<u8>,
}

impl<'a> NodeWriter<'a> {
    /// Create a writer that appends to `out`
    pub(super) fn new(out: &'a mut Vec<u8>) -> Self {
        Self {
            out,
            depth: 0,
            last_char_options: None,
            typeface: None,
        }
    }

    /// Write `nodes` as a LINE record
    ///
    /// A LINE is the unit that fills a template slot, a pile row or a matrix
    /// cell, and is always terminated by an END record.
    pub(super) fn write_line(&mut self, nodes: &[MathNode<'_>]) -> Result<(), MtefWriteError> {
        self.enter()?;
        write_line_start(self.out);
        self.last_char_options = None;
        self.write_sequence(nodes)?;
        write_end(self.out);
        self.last_char_options = None;
        self.leave();
        Ok(())
    }

    /// Write a PILE record holding one LINE per row
    pub(super) fn write_pile(&mut self, rows: &[&[MathNode<'_>]]) -> Result<(), MtefWriteError> {
        self.enter()?;
        write_pile_start(self.out, ALIGN_CENTER, VALIGN_BASELINE);
        self.last_char_options = None;
        for row in rows {
            self.write_line(row)?;
        }
        write_end(self.out);
        self.leave();
        Ok(())
    }

    /// Write a sequence of sibling nodes
    pub(super) fn write_sequence(&mut self, nodes: &[MathNode<'_>]) -> Result<(), MtefWriteError> {
        for node in nodes {
            self.write_node(node)?;
        }
        Ok(())
    }

    // ------------------------------------------------------------------
    // Recursion guard
    // ------------------------------------------------------------------

    /// Enter one nesting level, refusing to recurse past [`MAX_NESTING_DEPTH`]
    fn enter(&mut self) -> Result<(), MtefWriteError> {
        if self.depth >= MAX_NESTING_DEPTH {
            return Err(MtefWriteError::DepthExceeded {
                limit: MAX_NESTING_DEPTH,
            });
        }
        self.depth += 1;
        Ok(())
    }

    /// Leave a nesting level entered by [`Self::enter`]
    fn leave(&mut self) {
        self.depth = self.depth.saturating_sub(1);
    }

    // ------------------------------------------------------------------
    // Node dispatch
    // ------------------------------------------------------------------

    /// Write a single node
    fn write_node(&mut self, node: &MathNode<'_>) -> Result<(), MtefWriteError> {
        match node {
            // Leaves
            MathNode::Text(text) => self.write_text(text, TYPEFACE_VARIABLE),
            MathNode::Number(number) => self.write_text(number, TYPEFACE_NUMBER),
            MathNode::Operator(operator) => self.write_char_code(operator_code(*operator)),
            MathNode::Symbol(symbol) => self.write_symbol(symbol),
            MathNode::PredefinedSymbol(symbol) => {
                self.write_char_code(predefined_symbol_code(*symbol))
            },
            MathNode::Space(space) => self.write_char_code(space_code(*space)),

            // Structures
            MathNode::Frac {
                numerator,
                denominator,
                frac_type,
                ..
            } => self.write_template(
                fraction_template(*frac_type),
                &[numerator.as_slice(), denominator.as_slice()],
            ),
            MathNode::Root { base, index } => match index {
                Some(index) if !index.is_empty() => {
                    self.write_template(root_template(true), &[base.as_slice(), index.as_slice()])
                },
                _ => self.write_template(root_template(false), &[base.as_slice()]),
            },
            MathNode::Power { base, exponent } => self.write_scripts(base, &[], exponent),
            MathNode::Sub { base, subscript } => self.write_scripts(base, subscript, &[]),
            MathNode::SubSup {
                base,
                subscript,
                superscript,
            } => self.write_scripts(base, subscript, superscript),

            // Pre-scripts have no MTEF template: keep the material, drop the
            // placement.
            MathNode::PreSub {
                base,
                pre_subscript,
            } => self.write_degraded(&[pre_subscript.as_slice(), base.as_slice()]),
            MathNode::PreSup {
                base,
                pre_superscript,
            } => self.write_degraded(&[pre_superscript.as_slice(), base.as_slice()]),
            MathNode::PreSubSup {
                base,
                pre_subscript,
                pre_superscript,
            } => self.write_degraded(&[
                pre_subscript.as_slice(),
                pre_superscript.as_slice(),
                base.as_slice(),
            ]),

            MathNode::Under { base, under, .. } => self.write_template(
                Template::new(TMPL_LIM, TV_DEFAULT),
                &[base.as_slice(), under.as_slice(), &[]],
            ),
            MathNode::Over { base, over, .. } => self.write_template(
                Template::new(TMPL_LIM, TV_DEFAULT),
                &[base.as_slice(), &[], over.as_slice()],
            ),
            MathNode::UnderOver {
                base, under, over, ..
            } => self.write_template(
                Template::new(TMPL_LIM, TV_DEFAULT),
                &[base.as_slice(), under.as_slice(), over.as_slice()],
            ),

            MathNode::Fenced {
                open,
                content,
                close,
                ..
            } => self.write_fenced(*open, *close, content),

            MathNode::LargeOp {
                operator,
                lower_limit,
                upper_limit,
                integrand,
                hide_lower,
                hide_upper,
            } => {
                let lower = visible_limit(lower_limit.as_deref(), *hide_lower);
                let upper = visible_limit(upper_limit.as_deref(), *hide_upper);
                let body = integrand.as_deref().unwrap_or(&[]);
                match large_op_template(*operator, !lower.is_empty() || !upper.is_empty()) {
                    Some(template) => self.write_template(template, &[body, lower, upper]),
                    None => self.write_word_operator(*operator, lower, upper, body),
                }
            },

            MathNode::Function { name, argument } => self.write_function(name.as_ref(), argument),
            MathNode::PredefinedFunction { function, argument } => {
                self.write_function(function_name(*function), argument)
            },

            MathNode::Matrix {
                rows, fence_type, ..
            } => self.write_matrix(rows, *fence_type),
            MathNode::EqArray { rows, .. } => {
                let rows: Vec<&[MathNode<'_>]> = rows.iter().map(Vec::as_slice).collect();
                self.write_pile(&rows)
            },

            // Decorations
            MathNode::Accent { base, accent, .. } => {
                self.write_sequence(base)?;
                self.attach_embellishment(accent_embellishment(*accent));
                Ok(())
            },
            MathNode::Bar { base, .. } => self.write_template(
                bar_template(TMPL_OBAR, Some(LineStyle::Single)),
                &[base.as_slice()],
            ),
            MathNode::GroupChar { base, position, .. } => {
                self.write_template(group_char_template(*position), &[base.as_slice(), &[]])
            },

            // Runs and styles
            MathNode::Style { style, content } => self.write_styled(Some(*style), content),
            MathNode::Run {
                content,
                style,
                font,
                underline,
                overline,
                ..
            } => self.write_run(content, *style, font.as_deref(), *underline, *overline),

            // Transparent wrappers
            MathNode::Row(content) => self.write_sequence(content),
            MathNode::Phantom(content)
            | MathNode::Degree(content)
            | MathNode::Base(content)
            | MathNode::Argument(content)
            | MathNode::Numerator(content)
            | MathNode::Denominator(content)
            | MathNode::Integrand(content)
            | MathNode::LowerLimit(content)
            | MathNode::UpperLimit(content) => self.write_sequence(content),
            MathNode::Limit { content, .. } => self.write_sequence(content),
            MathNode::BorderBox { content, .. } => self.write_sequence(content),

            // A line break only exists between the lines of a pile, which the
            // top-level writer builds; nested breaks have nowhere to go.
            MathNode::LineBreak => Ok(()),

            // Invalid input carries no equation content.
            MathNode::Error(_) => Ok(()),
        }
    }

    // ------------------------------------------------------------------
    // Leaves
    // ------------------------------------------------------------------

    /// Write one CHAR record per character of `text`
    ///
    /// `default_typeface` claims the ASCII range, where the node kind knows
    /// better than the generic mapping whether a character is a digit or a
    /// variable name. Everything else goes through the shared table, and an
    /// enclosing style run overrides both.
    fn write_text(&mut self, text: &str, default_typeface: u8) -> Result<(), MtefWriteError> {
        for ch in text.chars() {
            let code = match self.typeface {
                Some(typeface) => char_code_in(ch, typeface)?,
                None if ch.is_ascii() => char_code_in(ch, default_typeface)?,
                None => char_code(ch)?,
            };
            self.write_char_code(code)?;
        }
        Ok(())
    }

    /// Write a symbol, preferring its Unicode form over its name
    fn write_symbol(&mut self, symbol: &Symbol<'_>) -> Result<(), MtefWriteError> {
        match symbol.unicode {
            Some(ch) => self.write_char_code(char_code(ch)?),
            None => self.write_text(symbol.name.as_ref(), TYPEFACE_VARIABLE),
        }
    }

    /// Write a single CHAR record
    fn write_char_code(&mut self, code: CharCode) -> Result<(), MtefWriteError> {
        let typeface = self.typeface.unwrap_or(code.typeface);
        self.last_char_options = Some(write_char(self.out, typeface, code.mtcode));
        Ok(())
    }

    /// Attach an embellishment list to the character written most recently
    ///
    /// MTEF embellishments decorate a character, so an accent over a structured
    /// expression has nothing to attach to and is dropped.
    fn attach_embellishment(&mut self, embellishment: u8) {
        if let Some(offset) = self.last_char_options.take() {
            set_char_embellished(self.out, offset);
            write_embell(self.out, embellishment);
            write_end(self.out);
        }
    }

    // ------------------------------------------------------------------
    // Structures
    // ------------------------------------------------------------------

    /// Write a TMPL record whose slots hold `slots`, one LINE each
    fn write_template(
        &mut self,
        template: Template,
        slots: &[&[MathNode<'_>]],
    ) -> Result<(), MtefWriteError> {
        self.enter()?;
        write_template_start(self.out, template.selector, template.variation)?;
        self.last_char_options = None;
        for slot in slots {
            self.write_line(slot)?;
        }
        write_end(self.out);
        self.last_char_options = None;
        self.leave();
        Ok(())
    }

    /// Write a base followed by the script template that decorates it
    ///
    /// MTEF scripts attach to the object that precedes them, so the base is
    /// emitted into the enclosing line rather than into a slot.
    fn write_scripts(
        &mut self,
        base: &[MathNode<'_>],
        subscript: &[MathNode<'_>],
        superscript: &[MathNode<'_>],
    ) -> Result<(), MtefWriteError> {
        self.write_sequence(base)?;
        let template = script_template(!subscript.is_empty(), !superscript.is_empty());
        self.write_template(template, &[subscript, superscript])
    }

    /// Write fenced content, or bare content when neither delimiter is drawn
    fn write_fenced(
        &mut self,
        open: Fence,
        close: Fence,
        content: &[MathNode<'_>],
    ) -> Result<(), MtefWriteError> {
        match fence_template(open, close) {
            Some(template) => self.write_template(template, &[content]),
            None => self.write_sequence(content),
        }
    }

    /// Write a large operator that MathType spells with letters
    ///
    /// `lim`, `max` and friends are function-typeface characters carrying the
    /// limit template, which is how MathType stores them.
    fn write_word_operator(
        &mut self,
        operator: LargeOperator,
        lower: &[MathNode<'_>],
        upper: &[MathNode<'_>],
        body: &[MathNode<'_>],
    ) -> Result<(), MtefWriteError> {
        let word = large_op_word(operator).unwrap_or_default();
        self.enter()?;
        write_template_start(self.out, TMPL_LIM, TV_DEFAULT)?;
        self.last_char_options = None;

        // Slot one holds the operator name itself.
        self.enter()?;
        write_line_start(self.out);
        self.write_function_name(word)?;
        write_end(self.out);
        self.leave();

        self.write_line(lower)?;
        self.write_line(upper)?;
        write_end(self.out);
        self.last_char_options = None;
        self.leave();

        self.write_sequence(body)
    }

    /// Write a named function followed by its argument
    fn write_function(
        &mut self,
        name: &str,
        argument: &[MathNode<'_>],
    ) -> Result<(), MtefWriteError> {
        self.write_function_name(name)?;
        self.write_sequence(argument)
    }

    /// Write a function name in the function typeface
    ///
    /// Only ASCII letters go into that typeface: readers recognise a function
    /// by scanning a run of alphabetic function-typeface characters, so digits
    /// or punctuation would truncate (or invalidate) the run.
    fn write_function_name(&mut self, name: &str) -> Result<(), MtefWriteError> {
        for ch in name.chars() {
            let typeface = if ch.is_ascii_alphabetic() {
                TYPEFACE_FUNCTION
            } else {
                TYPEFACE_VARIABLE
            };
            self.write_char_code(char_code_in(ch, typeface)?)?;
        }
        Ok(())
    }

    /// Write a MATRIX record, wrapped in a fence template when one is asked for
    ///
    /// The MATRIX record itself has no delimiters: MathType draws a bracketed
    /// matrix as a fence template containing a matrix.
    fn write_matrix(
        &mut self,
        rows: &[Vec<Vec<MathNode<'_>>>],
        fence_type: MatrixFence,
    ) -> Result<(), MtefWriteError> {
        let fence = match fence_type {
            MatrixFence::None => Fence::None,
            MatrixFence::Paren => Fence::Paren,
            MatrixFence::Bracket => Fence::Bracket,
            MatrixFence::Brace => Fence::Brace,
            MatrixFence::Pipe => Fence::Pipe,
            MatrixFence::DoublePipe => Fence::DoublePipe,
        };

        match fence_template(fence, fence) {
            Some(template) => {
                self.enter()?;
                write_template_start(self.out, template.selector, template.variation)?;
                self.last_char_options = None;
                self.enter()?;
                write_line_start(self.out);
                self.write_matrix_record(rows)?;
                write_end(self.out);
                self.leave();
                write_end(self.out);
                self.leave();
                Ok(())
            },
            None => self.write_matrix_record(rows),
        }
    }

    /// Write the MATRIX record and its cells in row-major order
    fn write_matrix_record(
        &mut self,
        rows: &[Vec<Vec<MathNode<'_>>>],
    ) -> Result<(), MtefWriteError> {
        let row_count = rows.len();
        let col_count = rows.iter().map(Vec::len).max().unwrap_or(0);

        self.enter()?;
        write_matrix_start(self.out, row_count, col_count)?;
        self.last_char_options = None;
        for row in rows {
            for column in 0..col_count {
                match row.get(column) {
                    Some(cell) => self.write_line(cell)?,
                    // Ragged rows are padded so cells stay aligned.
                    None => self.write_line(&[])?,
                }
            }
        }
        write_end(self.out);
        self.leave();
        Ok(())
    }

    // ------------------------------------------------------------------
    // Runs and degradation
    // ------------------------------------------------------------------

    /// Write content under a style, restoring the previous typeface afterwards
    fn write_styled(
        &mut self,
        style: Option<StyleType>,
        content: &[MathNode<'_>],
    ) -> Result<(), MtefWriteError> {
        let previous = self.typeface;
        if let Some(typeface) = style.and_then(style_typeface) {
            self.typeface = Some(typeface);
        }
        let result = self.write_sequence(content);
        self.typeface = previous;
        result
    }

    /// Write a run, mapping its font, style and rules onto MTEF records
    fn write_run(
        &mut self,
        content: &[MathNode<'_>],
        style: Option<StyleType>,
        font: Option<&str>,
        underline: Option<LineStyle>,
        overline: Option<LineStyle>,
    ) -> Result<(), MtefWriteError> {
        if let Some(name) = font {
            write_font(self.out, TYPEFACE_TEXT, FONT_STYLE_PLAIN, name)?;
            self.last_char_options = None;
        }

        // Rules are templates wrapped around the run. A run carrying both keeps
        // the overline outermost and re-enters with the underline still to do.
        if let Some(line) = overline {
            return self.write_ruled(
                bar_template(TMPL_OBAR, Some(line)),
                content,
                style,
                underline,
            );
        }
        if let Some(line) = underline {
            return self.write_ruled(bar_template(TMPL_UBAR, Some(line)), content, style, None);
        }

        self.write_styled(style, content)
    }

    /// Write a styled run inside a rule template
    ///
    /// `remaining_underline` is the rule still to be drawn inside this one.
    fn write_ruled(
        &mut self,
        template: Template,
        content: &[MathNode<'_>],
        style: Option<StyleType>,
        remaining_underline: Option<LineStyle>,
    ) -> Result<(), MtefWriteError> {
        self.enter()?;
        write_template_start(self.out, template.selector, template.variation)?;
        self.last_char_options = None;
        self.enter()?;
        write_line_start(self.out);
        self.write_run(content, style, None, remaining_underline, None)?;
        write_end(self.out);
        self.leave();
        write_end(self.out);
        self.last_char_options = None;
        self.leave();
        Ok(())
    }

    /// Write the content of a construct MTEF cannot represent
    fn write_degraded(&mut self, parts: &[&[MathNode<'_>]]) -> Result<(), MtefWriteError> {
        for part in parts {
            self.write_sequence(part)?;
        }
        Ok(())
    }
}

/// A limit's content, or nothing when the limit is hidden
///
/// MTEF has no "present but hidden" state: a suppressed limit is simply an
/// empty slot.
fn visible_limit<'n, 'a>(limit: Option<&'n [MathNode<'a>]>, hidden: bool) -> &'n [MathNode<'a>] {
    match limit {
        Some(nodes) if !hidden => nodes,
        _ => &[],
    }
}
