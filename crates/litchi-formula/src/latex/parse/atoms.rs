// Control sequence dispatch
//
// `Parser::parse_command` turns a single control sequence, together with the
// arguments it consumes, into one AST node. Commands are tried in order:
// structural forms with fixed arity first, then the static lookup tables, then
// a graceful fallback that keeps unknown commands as symbols.

use super::commands::{
    ACCENTS, CMD_BEGIN, CMD_END, CMD_LEFT, CMD_LIMITS, CMD_LINE_BREAK, CMD_NOLIMITS, CMD_RIGHT,
    DELIMITER_GLYPHS, FUNCTION_WORDS, FUNCTIONS, IGNORED_COMMANDS, LARGE_OPERATORS, MATH_STYLES,
    OPERATORS, PREDEFINED_SYMBOLS, SPACES, TEXT_STYLES,
};
use super::error::LatexParseError;
use super::fences::fence_for_delimiter;
use super::parser::{Limit, NumberMode, Parser};
use super::scripts::Scripts;
use super::token::{Token, TokenKind};
use crate::ast::{
    Fence, FractionType, LargeOperator, MathNode, Position, StyleType, Symbol, VerticalAlignment,
};
use std::borrow::Cow;

/// `\frac` and its display/text-sized spellings.
const FRACTION_COMMANDS: [&str; 4] = ["frac", "dfrac", "tfrac", "cfrac"];
/// `\binom` and its display/text-sized spellings.
const BINOMIAL_COMMANDS: [&str; 3] = ["binom", "dbinom", "tbinom"];
/// Commands whose content is rendered invisibly but still occupies space.
const PHANTOM_COMMANDS: [&str; 3] = ["phantom", "hphantom", "vphantom"];
/// Commands that draw a frame around their content.
const BOX_COMMANDS: [&str; 2] = ["boxed", "fbox"];

/// `\sqrt`, which takes an optional index argument.
const CMD_SQRT: &str = "sqrt";
/// `\operatorname`, which names a function inline.
const CMD_OPERATORNAME: &str = "operatorname";
/// `\substack`, a one-column stack used inside limits.
const CMD_SUBSTACK: &str = "substack";
/// Pre-subscript form emitted by the AST-to-LaTeX converter.
const CMD_PRESUB: &str = "presub";
/// Pre-superscript form emitted by the AST-to-LaTeX converter.
const CMD_PRESUP: &str = "presup";
/// Combined pre-script form emitted by the AST-to-LaTeX converter.
const CMD_PRESUBSUP: &str = "presubsup";
/// `\underset`, the AST-to-LaTeX spelling of an under-script.
const CMD_UNDERSET: &str = "underset";
/// `\overset`, the AST-to-LaTeX spelling of an over-script.
const CMD_OVERSET: &str = "overset";
/// `\overline`.
const CMD_OVERLINE: &str = "overline";
/// `\underline`.
const CMD_UNDERLINE: &str = "underline";
/// `\overbrace`.
const CMD_OVERBRACE: &str = "overbrace";
/// `\underbrace`.
const CMD_UNDERBRACE: &str = "underbrace";

impl<'a> Parser<'a> {
    /// Parse the control sequence `name` under the cursor into a node.
    ///
    /// Returns `Ok(None)` for commands that carry no AST meaning, such as
    /// `\displaystyle` or the math-mode delimiters `\(` and `\)`.
    pub(crate) fn parse_command(
        &mut self,
        name: &'a str,
        token: Token<'a>,
    ) -> Result<Option<MathNode<'a>>, LatexParseError> {
        // These two are only ever consumed by the construct that opened them;
        // meeting one here means the source is unbalanced.
        match name {
            CMD_END => {
                return Err(LatexParseError::UnexpectedEnd {
                    position: token.start,
                });
            },
            CMD_RIGHT => {
                return Err(LatexParseError::UnmatchedRight {
                    position: token.start,
                });
            },
            _ => {},
        }

        self.advance();

        if let Some(node) = self.parse_structural_command(name, token)? {
            return Ok(Some(node));
        }
        self.parse_tabulated_command(name)
    }

    /// Handle commands whose shape is fixed rather than table-driven.
    fn parse_structural_command(
        &mut self,
        name: &'a str,
        token: Token<'a>,
    ) -> Result<Option<MathNode<'a>>, LatexParseError> {
        if FRACTION_COMMANDS.contains(&name) {
            return self
                .parse_fraction(name, FractionType::Bar, false)
                .map(Some);
        }
        if BINOMIAL_COMMANDS.contains(&name) {
            return self
                .parse_fraction(name, FractionType::NoBar, true)
                .map(Some);
        }
        if PHANTOM_COMMANDS.contains(&name) {
            let content = self.parse_required_argument(name)?;
            return Ok(Some(MathNode::Phantom(Box::new(content))));
        }
        if BOX_COMMANDS.contains(&name) {
            let content = self.parse_required_argument(name)?;
            return Ok(Some(MathNode::BorderBox {
                content: Box::new(content),
                style: None,
            }));
        }

        let node = match name {
            CMD_BEGIN => self.parse_environment(token)?,
            CMD_LEFT => self.parse_left_right(token)?,
            CMD_LINE_BREAK => {
                // `\\[2pt]` carries a spacing hint the AST does not model.
                self.parse_optional_argument()?;
                MathNode::LineBreak
            },
            CMD_SQRT => {
                let index = self.parse_optional_argument()?;
                let base = self.parse_required_argument(CMD_SQRT)?;
                MathNode::Root { base, index }
            },
            CMD_OPERATORNAME => {
                let function = self.read_command_name(CMD_OPERATORNAME)?;
                let argument = self.parse_function_argument()?;
                MathNode::Function {
                    name: Cow::Borrowed(function),
                    argument,
                }
            },
            CMD_SUBSTACK => self.parse_substack()?,
            CMD_PRESUB => {
                let base = self.parse_required_argument(CMD_PRESUB)?;
                let pre_subscript = self.parse_required_argument(CMD_PRESUB)?;
                MathNode::PreSub {
                    base,
                    pre_subscript,
                }
            },
            CMD_PRESUP => {
                let base = self.parse_required_argument(CMD_PRESUP)?;
                let pre_superscript = self.parse_required_argument(CMD_PRESUP)?;
                MathNode::PreSup {
                    base,
                    pre_superscript,
                }
            },
            CMD_PRESUBSUP => {
                let base = self.parse_required_argument(CMD_PRESUBSUP)?;
                let pre_subscript = self.parse_required_argument(CMD_PRESUBSUP)?;
                let pre_superscript = self.parse_required_argument(CMD_PRESUBSUP)?;
                MathNode::PreSubSup {
                    base,
                    pre_subscript,
                    pre_superscript,
                }
            },
            CMD_UNDERSET => {
                let under = self.parse_required_argument(CMD_UNDERSET)?;
                let base = self.parse_required_argument(CMD_UNDERSET)?;
                MathNode::Under {
                    base,
                    under,
                    position: None,
                }
            },
            CMD_OVERSET => {
                let over = self.parse_required_argument(CMD_OVERSET)?;
                let base = self.parse_required_argument(CMD_OVERSET)?;
                MathNode::Over {
                    base,
                    over,
                    position: None,
                }
            },
            CMD_OVERLINE => {
                let base = self.parse_required_argument(CMD_OVERLINE)?;
                MathNode::Bar {
                    base: Box::new(base),
                    position: Some(Position::Top),
                }
            },
            CMD_UNDERLINE => {
                let base = self.parse_required_argument(CMD_UNDERLINE)?;
                MathNode::Under {
                    base,
                    under: Vec::new(),
                    position: Some(Position::Bottom),
                }
            },
            CMD_OVERBRACE => self.parse_group_char(CMD_OVERBRACE, Position::Top)?,
            CMD_UNDERBRACE => self.parse_group_char(CMD_UNDERBRACE, Position::Bottom)?,
            _ => return Ok(None),
        };

        Ok(Some(node))
    }

    /// Handle commands resolved through the static lookup tables.
    fn parse_tabulated_command(
        &mut self,
        name: &'a str,
    ) -> Result<Option<MathNode<'a>>, LatexParseError> {
        if let Some(&accent) = ACCENTS.get(name) {
            let base = self.parse_required_argument(name)?;
            return Ok(Some(MathNode::Accent {
                base: Box::new(base),
                accent,
                position: Some(Position::Top),
            }));
        }
        if let Some(&operator) = LARGE_OPERATORS.get(name) {
            return self.parse_large_operator(operator).map(Some);
        }
        if let Some(&function) = FUNCTIONS.get(name) {
            let argument = self.parse_function_argument()?;
            return Ok(Some(MathNode::PredefinedFunction { function, argument }));
        }
        if FUNCTION_WORDS.contains(name) {
            let argument = self.parse_function_argument()?;
            return Ok(Some(MathNode::Function {
                name: Cow::Borrowed(name),
                argument,
            }));
        }
        if let Some(&style) = TEXT_STYLES.get(name) {
            return self.parse_text_style(name, style).map(Some);
        }
        if let Some(&style) = MATH_STYLES.get(name) {
            let content = self.parse_required_argument(name)?;
            return Ok(Some(MathNode::Style { style, content }));
        }
        if let Some(&space) = SPACES.get(name) {
            return Ok(Some(MathNode::Space(space)));
        }
        if let Some(&operator) = OPERATORS.get(name) {
            return Ok(Some(MathNode::Operator(operator)));
        }
        if let Some(&symbol) = PREDEFINED_SYMBOLS.get(name) {
            return Ok(Some(MathNode::PredefinedSymbol(symbol)));
        }
        if IGNORED_COMMANDS.contains(name) {
            return Ok(None);
        }
        // A delimiter without a partner still renders as its bracket glyph.
        if let Some(&glyph) = DELIMITER_GLYPHS.get(name) {
            return Ok(Some(MathNode::Symbol(Symbol {
                name: Cow::Borrowed(name),
                unicode: Some(glyph),
                variant: None,
            })));
        }

        // Unknown command: keep the name so the AST-to-LaTeX direction can
        // still render something recognisable instead of dropping content.
        Ok(Some(MathNode::Symbol(Symbol {
            name: Cow::Borrowed(name),
            unicode: None,
            variant: None,
        })))
    }

    /// Parse the two arguments of a fraction-shaped command.
    fn parse_fraction(
        &mut self,
        name: &str,
        frac_type: FractionType,
        parenthesised: bool,
    ) -> Result<MathNode<'a>, LatexParseError> {
        let numerator = self.parse_required_argument(name)?;
        let denominator = self.parse_required_argument(name)?;
        let fraction = MathNode::Frac {
            numerator,
            denominator,
            line_thickness: None,
            frac_type: Some(frac_type),
        };

        if !parenthesised {
            return Ok(fraction);
        }
        Ok(MathNode::Fenced {
            open: Fence::Paren,
            content: vec![fraction],
            close: Fence::Paren,
            separator: None,
        })
    }

    /// Parse a large operator together with its limits.
    ///
    /// `\limits` and `\nolimits` only move the limits relative to the operator,
    /// which the AST does not model, so they are accepted and skipped.
    fn parse_large_operator(
        &mut self,
        operator: LargeOperator,
    ) -> Result<MathNode<'a>, LatexParseError> {
        loop {
            self.skip_spaces();
            match self.peek().and_then(|token| token.command()) {
                Some(CMD_LIMITS) | Some(CMD_NOLIMITS) => self.advance(),
                _ => break,
            }
        }

        let Scripts {
            subscript: lower_limit,
            superscript: upper_limit,
        } = self.collect_scripts()?;
        Ok(MathNode::LargeOp {
            operator,
            hide_lower: lower_limit.is_none(),
            hide_upper: upper_limit.is_none(),
            lower_limit,
            upper_limit,
            integrand: None,
        })
    }

    /// Parse `\overbrace`/`\underbrace` into a group character node.
    fn parse_group_char(
        &mut self,
        name: &str,
        position: Position,
    ) -> Result<MathNode<'a>, LatexParseError> {
        let base = self.parse_required_argument(name)?;
        let vertical_alignment = match position {
            Position::Bottom => VerticalAlignment::Bottom,
            _ => VerticalAlignment::Top,
        };
        Ok(MathNode::GroupChar {
            base: Box::new(base),
            character: None,
            position: Some(position),
            vertical_alignment: Some(vertical_alignment),
        })
    }

    /// Parse a `\text`-family command, keeping its argument verbatim.
    ///
    /// The argument becomes a literal [`MathNode::Run`] rather than a
    /// [`MathNode::Style`], because the AST-to-LaTeX direction already renders
    /// a [`MathNode::Text`] carrying spaces as `\text{...}`; wrapping it in a
    /// style would add a redundant `\mathrm` layer on every round trip.
    fn parse_text_style(
        &mut self,
        name: &'a str,
        style: StyleType,
    ) -> Result<MathNode<'a>, LatexParseError> {
        self.skip_spaces();
        if !matches!(
            self.peek().map(|token| token.kind),
            Some(TokenKind::GroupOpen)
        ) {
            // `\text x` is unusual but harmless; fall back to math parsing.
            let content = self.parse_required_argument(name)?;
            return Ok(MathNode::Style { style, content });
        }

        let raw = self.read_raw_group()?;
        let content = if raw.is_empty() {
            Vec::new()
        } else {
            vec![MathNode::Text(unescape_text(raw))]
        };

        Ok(MathNode::Run {
            content,
            literal: Some(true),
            style: (style != StyleType::Normal).then_some(style),
            font: None,
            color: None,
            underline: None,
            overline: None,
            strike_through: None,
            double_strike_through: None,
        })
    }

    /// Read a `{name}` argument verbatim, used by `\operatorname`.
    fn read_command_name(&mut self, command: &str) -> Result<&'a str, LatexParseError> {
        self.skip_spaces();
        let position = self.offset();
        if !matches!(
            self.peek().map(|token| token.kind),
            Some(TokenKind::GroupOpen)
        ) {
            return Err(LatexParseError::MissingArgument {
                command: command.to_string(),
                position,
            });
        }
        Ok(self.read_raw_group()?.trim())
    }

    /// Parse the argument a named function applies to.
    ///
    /// A braced group is always taken; a bare operand is taken only when the
    /// next token can plausibly start one, so `\sin + 1` keeps the `+` outside.
    fn parse_function_argument(&mut self) -> Result<Vec<MathNode<'a>>, LatexParseError> {
        self.skip_spaces();
        let Some(token) = self.peek() else {
            return Ok(Vec::new());
        };
        if matches!(token.kind, TokenKind::GroupOpen) {
            return self.parse_group_contents();
        }
        if !starts_operand(&token) {
            return Ok(Vec::new());
        }

        Ok(match self.parse_atom(NumberMode::Run)? {
            Some(MathNode::Row(inner)) => inner,
            Some(node) => vec![node],
            None => Vec::new(),
        })
    }

    /// Parse a `\left ... \right` delimited expression.
    fn parse_left_right(&mut self, left: Token<'a>) -> Result<MathNode<'a>, LatexParseError> {
        let open = self.read_fence_delimiter()?;
        let content = self.parse_sequence(Limit::Right)?;

        let Some(right) = self.peek().filter(|token| token.is_command(CMD_RIGHT)) else {
            return Err(LatexParseError::UnmatchedLeft {
                position: left.start,
            });
        };
        let _ = right;
        self.advance();
        let close = self.read_fence_delimiter()?;

        Ok(MathNode::Fenced {
            open,
            content,
            close,
            separator: None,
        })
    }

    /// Consume the delimiter that follows `\left` or `\right`.
    fn read_fence_delimiter(&mut self) -> Result<Fence, LatexParseError> {
        self.skip_spaces();
        let position = self.offset();
        let Some(token) = self.peek() else {
            return Err(LatexParseError::MissingDelimiter { position });
        };
        let Some(fence) = fence_for_delimiter(&token) else {
            return Err(LatexParseError::MissingDelimiter {
                position: token.start,
            });
        };
        self.advance();
        Ok(fence)
    }
}

/// Characters that the AST-to-LaTeX direction escapes inside `\text{...}`.
const ESCAPED_TEXT_CHARS: [char; 11] = [' ', '#', '$', '%', '&', '_', '{', '}', '~', '^', '\\'];

/// Undo the backslash escaping applied to `\text{...}` content.
///
/// Borrows unless the argument actually contains an escape sequence.
fn unescape_text(raw: &str) -> Cow<'_, str> {
    if !raw.contains('\\') {
        return Cow::Borrowed(raw);
    }

    let mut out = String::with_capacity(raw.len());
    let mut chars = raw.chars();
    while let Some(ch) = chars.next() {
        if ch != '\\' {
            out.push(ch);
            continue;
        }
        match chars.next() {
            Some(next) if ESCAPED_TEXT_CHARS.contains(&next) => out.push(next),
            Some(next) => {
                out.push('\\');
                out.push(next);
            },
            None => out.push('\\'),
        }
    }
    Cow::Owned(out)
}

/// Report whether `token` can begin an operand of a named function.
fn starts_operand(token: &Token<'_>) -> bool {
    match token.kind {
        TokenKind::GroupOpen => true,
        TokenKind::Char(ch) => ch.is_alphanumeric() || ch == '(' || ch == '[' || !ch.is_ascii(),
        TokenKind::Command(name) => {
            if name == CMD_RIGHT || name == CMD_END || name == CMD_LINE_BREAK {
                return false;
            }
            !OPERATORS.contains_key(name)
                && !SPACES.contains_key(name)
                && !IGNORED_COMMANDS.contains(name)
        },
        _ => false,
    }
}
