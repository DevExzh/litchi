//! Typed, inert model of the legacy `EQ` field instruction syntax.
//!
//! ECMA-376 Part 1 §17.16.5.17 defines the `EQ` field, whose instruction text
//! encodes mathematical layout through backslash switches: fractions (`\f`),
//! radicals (`\r`), scripts (`\s`), integrals and sums (`\i`), arrays (`\a`),
//! brackets (`\b`), boxes (`\x`), overstrikes (`\o`), lists (`\l`), and
//! displacements (`\d`).
//!
//! This module parses that syntax into a typed tree of [`EquationSegment`]s.
//! The model is purely syntactic and inert: spacing values, characters, and
//! element text are exposed exactly as stored. Nothing here evaluates, lays
//! out, typesets, or renders an equation.

use crate::error::{RtfError, RtfResult};

/// Maximum number of switched groups in one `EQ` expression.
const MAX_EQUATION_GROUPS: usize = 256;
/// Maximum number of comma-separated elements in one switched group.
const MAX_GROUP_ELEMENTS: usize = 64;
/// Maximum number of sub-options on one switch.
const MAX_SWITCH_OPTIONS: usize = 16;
/// Upper bound for spacing arguments, in points.
const MAX_SPACING_POINTS: u32 = 16_384;
/// Upper bound for `\a\co` column counts.
const MAX_ARRAY_COLUMNS: u8 = 64;

fn malformed(message: impl Into<String>) -> RtfError {
    RtfError::MalformedDocument(message.into())
}

/// A spacing argument stored by an `EQ` switch, in points.
#[derive(Debug, Clone, Copy, PartialEq, Eq, PartialOrd, Ord)]
pub struct EquationSpacing(u32);

impl EquationSpacing {
    /// Return the spacing value in points.
    pub const fn points(self) -> u32 {
        self.0
    }

    fn parse(text: &str) -> RtfResult<Self> {
        let value: u32 = text
            .parse()
            .map_err(|_| malformed("RTF EQ switch has a non-numeric spacing argument"))?;
        if value > MAX_SPACING_POINTS {
            return Err(malformed(
                "RTF EQ spacing argument exceeds the safety limit",
            ));
        }
        Ok(Self(value))
    }
}

/// Column or element alignment shared by the `\a` and `\o` switches.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum EquationAlignment {
    /// Left alignment (`\al`).
    Left,
    /// Center alignment (`\ac`).
    Center,
    /// Right alignment (`\ar`).
    Right,
}

/// Options of the array switch (`\a`).
#[derive(Debug, Clone, Copy, Default, PartialEq, Eq)]
pub struct EquationArray {
    /// Element alignment (`\al`, `\ac`, or `\ar`).
    pub alignment: Option<EquationAlignment>,
    /// Column count (`\co`N).
    pub columns: Option<u8>,
    /// Vertical line spacing in points (`\vs`N).
    pub vertical_spacing: Option<EquationSpacing>,
    /// Horizontal column spacing in points (`\hs`N).
    pub horizontal_spacing: Option<EquationSpacing>,
}

/// Options of the bracket switch (`\b`).
#[derive(Debug, Clone, Copy, Default, PartialEq, Eq)]
pub struct EquationBracket {
    /// Left bracket character (`\lc`C, or both via `\bc`C).
    pub left: Option<char>,
    /// Right bracket character (`\rc`C, or both via `\bc`C).
    pub right: Option<char>,
}

/// Options of the displacement switch (`\d`).
#[derive(Debug, Clone, Copy, Default, PartialEq, Eq)]
pub struct EquationDisplace {
    /// Forward displacement in points (`\fo`N).
    pub forward: Option<EquationSpacing>,
    /// Backward displacement in points (`\ba`N).
    pub backward: Option<EquationSpacing>,
    /// Underline drawn through the following element (`\li`).
    pub underline: bool,
}

/// The symbol family selected by integral-switch sub-options.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum EquationIntegralSymbol {
    /// Integral sign (the plain `\i` default).
    Integral,
    /// Capital-sigma summation (`\su`).
    Summation,
    /// Capital-pi product (`\pr`).
    Product,
}

/// Options of the integral switch (`\i`).
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct EquationIntegral {
    /// Selected symbol family.
    pub symbol: EquationIntegralSymbol,
    /// Limits set inline instead of above/below (`\in`).
    pub inline_limits: bool,
    /// Fixed-size substitute character (`\fc`C).
    pub fixed_char: Option<char>,
    /// Variable-size substitute character (`\vc`C).
    pub variable_char: Option<char>,
}

impl Default for EquationIntegral {
    fn default() -> Self {
        Self {
            symbol: EquationIntegralSymbol::Integral,
            inline_limits: false,
            fixed_char: None,
            variable_char: None,
        }
    }
}

/// Options of the overstrike switch (`\o`).
#[derive(Debug, Clone, Copy, Default, PartialEq, Eq)]
pub struct EquationOverstrike {
    /// Element alignment (`\al`, `\ac`, or `\ar`).
    pub alignment: Option<EquationAlignment>,
}

/// Options of the script switch (`\s`).
#[derive(Debug, Clone, Copy, Default, PartialEq, Eq)]
pub struct EquationScript {
    /// Superscript raise in points (`\up`N).
    pub up: Option<EquationSpacing>,
    /// Subscript drop in points (`\do`N).
    pub down: Option<EquationSpacing>,
}

/// Options of the box switch (`\x`).
#[derive(Debug, Clone, Copy, Default, PartialEq, Eq)]
pub struct EquationBox {
    /// Border over the element (`\to`).
    pub top: bool,
    /// Border under the element (`\bo`).
    pub bottom: bool,
    /// Border left of the element (`\le`).
    pub left: bool,
    /// Border right of the element (`\ri`).
    pub right: bool,
}

/// A single `EQ` switch with its typed sub-options.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum EquationSwitch {
    /// Array (`\a`).
    Array(EquationArray),
    /// Bracketed element (`\b`).
    Bracket(EquationBracket),
    /// Displaced element (`\d`).
    Displace(EquationDisplace),
    /// Fraction (`\f`): elements are numerator and denominator.
    Fraction,
    /// Integral, sum, or product (`\i`).
    Integral(EquationIntegral),
    /// Comma- or semicolon-separated list (`\l`).
    List,
    /// Overstruck elements (`\o`).
    Overstrike(EquationOverstrike),
    /// Radical (`\r`): elements are the optional index and the radicand.
    Radical,
    /// Superscript or subscript (`\s`).
    Script(EquationScript),
    /// Boxed element (`\x`).
    Box(EquationBox),
}

/// One switched group: a switch plus its parenthesized elements.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct EquationGroup<'a> {
    /// The parsed switch and its sub-options.
    pub switch: EquationSwitch,
    /// Element text in stored order, each trimmed of surrounding whitespace.
    pub elements: Vec<&'a str>,
}

/// A segment of an `EQ` expression.
#[derive(Debug, Clone, PartialEq, Eq)]
pub enum EquationSegment<'a> {
    /// A switched group such as `\f(a,b)`.
    Switched(EquationGroup<'a>),
    /// Literal text stored between or around switched groups.
    Literal(&'a str),
}

/// A parsed `EQ` field expression.
#[derive(Debug, Clone, Default, PartialEq, Eq)]
pub struct EquationModel<'a> {
    segments: Vec<EquationSegment<'a>>,
}

impl<'a> EquationModel<'a> {
    /// Parse the instruction text following the `EQ` keyword.
    ///
    /// The result borrows from `expression`. Unknown switches, unterminated
    /// element lists, missing numeric or character arguments, and inputs
    /// exceeding the safety limits are reported as errors; the model is never
    /// partially returned for malformed input.
    pub fn parse(expression: &'a str) -> RtfResult<Self> {
        Parser::new(expression).run()
    }

    /// Return the expression segments in stored order.
    pub fn segments(&self) -> &[EquationSegment<'a>] {
        &self.segments
    }

    /// Iterate over only the switched groups, skipping literal text.
    pub fn groups(&self) -> impl Iterator<Item = &EquationGroup<'a>> {
        self.segments.iter().filter_map(|segment| match segment {
            EquationSegment::Switched(group) => Some(group),
            EquationSegment::Literal(_) => None,
        })
    }

    /// Whether the expression contains no segments at all.
    pub fn is_empty(&self) -> bool {
        self.segments.is_empty()
    }
}

struct Parser<'a> {
    input: &'a str,
    position: usize,
    segments: Vec<EquationSegment<'a>>,
    literal_start: Option<usize>,
}

impl<'a> Parser<'a> {
    fn new(input: &'a str) -> Self {
        Self {
            input,
            position: 0,
            segments: Vec::new(),
            literal_start: None,
        }
    }

    fn run(mut self) -> RtfResult<EquationModel<'a>> {
        while self.position < self.input.len() {
            if self.input.as_bytes()[self.position] == b'\\' {
                self.flush_literal();
                self.parse_group()?;
            } else {
                if self.literal_start.is_none() {
                    self.literal_start = Some(self.position);
                }
                self.advance_char();
            }
        }
        self.flush_literal();
        Ok(EquationModel {
            segments: self.segments,
        })
    }

    fn flush_literal(&mut self) {
        if let Some(start) = self.literal_start.take() {
            let text = self.input[start..self.position].trim();
            if !text.is_empty() {
                self.segments.push(EquationSegment::Literal(text));
            }
        }
    }

    fn advance_char(&mut self) {
        let width = self.input[self.position..]
            .chars()
            .next()
            .map(char::len_utf8)
            .unwrap_or(1);
        self.position += width;
    }

    fn peek(&self) -> Option<char> {
        self.input[self.position..].chars().next()
    }

    fn expect_backslash(&mut self) -> RtfResult<()> {
        if self.peek() == Some('\\') {
            self.position += 1;
            Ok(())
        } else {
            Err(malformed("RTF EQ switch option lacks a backslash"))
        }
    }

    /// Read exactly two ASCII letters forming a switch or sub-option name.
    fn read_name(&mut self) -> RtfResult<[u8; 2]> {
        let bytes = self.input.as_bytes();
        let end = self.position + 2;
        if end > bytes.len()
            || !bytes[self.position].is_ascii_alphabetic()
            || !bytes[self.position + 1].is_ascii_alphabetic()
        {
            return Err(malformed("RTF EQ switch has a truncated name"));
        }
        let name = [bytes[self.position], bytes[self.position + 1]];
        self.position = end;
        Ok(name)
    }

    fn read_number(&mut self) -> RtfResult<EquationSpacing> {
        let start = self.position;
        while self
            .peek()
            .is_some_and(|character| character.is_ascii_digit())
        {
            self.position += 1;
        }
        if self.position == start {
            return Err(malformed("RTF EQ switch option lacks a numeric argument"));
        }
        EquationSpacing::parse(&self.input[start..self.position])
    }

    /// Read the escaped single character argument of `\lc`, `\rc`, `\bc`,
    /// `\fc`, `\vc`: the character follows the option name in `\C` form, so
    /// `\bc\{` selects `{`.
    fn read_char(&mut self) -> RtfResult<char> {
        self.expect_backslash()
            .map_err(|_| malformed("RTF EQ switch option lacks a character argument"))?;
        let Some(character) = self.peek() else {
            return Err(malformed("RTF EQ switch option lacks a character argument"));
        };
        self.advance_char();
        Ok(character)
    }

    fn parse_group(&mut self) -> RtfResult<()> {
        if self.segments.len() >= MAX_EQUATION_GROUPS {
            return Err(malformed("RTF EQ expression exceeds the group limit"));
        }
        self.expect_backslash()?;
        let Some(&letter) = self.input.as_bytes().get(self.position) else {
            return Err(malformed("RTF EQ switch has a truncated name"));
        };
        if !letter.is_ascii_alphabetic() {
            return Err(malformed("RTF EQ switch has a non-alphabetic name"));
        }
        self.position += 1;
        let mut options = 0usize;
        let mut switch = match letter {
            b'a' => EquationSwitch::Array(EquationArray::default()),
            b'b' => EquationSwitch::Bracket(EquationBracket::default()),
            b'd' => EquationSwitch::Displace(EquationDisplace::default()),
            b'f' => EquationSwitch::Fraction,
            b'i' => EquationSwitch::Integral(EquationIntegral::default()),
            b'l' => EquationSwitch::List,
            b'o' => EquationSwitch::Overstrike(EquationOverstrike::default()),
            b'r' => EquationSwitch::Radical,
            b's' => EquationSwitch::Script(EquationScript::default()),
            b'x' => EquationSwitch::Box(EquationBox::default()),
            other => {
                return Err(malformed(format!(
                    "RTF EQ switch '\\{}' is not a known equation switch",
                    other as char
                )));
            },
        };
        while self.peek() == Some('\\') {
            options += 1;
            if options > MAX_SWITCH_OPTIONS {
                return Err(malformed("RTF EQ switch exceeds the option limit"));
            }
            self.expect_backslash()?;
            let name = self.read_name()?;
            Self::apply_option(&mut switch, name, self)?;
        }
        let elements = self.parse_elements()?;
        self.segments.push(EquationSegment::Switched(EquationGroup {
            switch,
            elements,
        }));
        Ok(())
    }

    fn apply_option(
        switch: &mut EquationSwitch,
        name: [u8; 2],
        parser: &mut Self,
    ) -> RtfResult<()> {
        let alignment = || match name {
            [b'a', b'l'] => Some(EquationAlignment::Left),
            [b'a', b'c'] => Some(EquationAlignment::Center),
            [b'a', b'r'] => Some(EquationAlignment::Right),
            _ => None,
        };
        match switch {
            EquationSwitch::Array(options) => match name {
                [b'a', b'l'] | [b'a', b'c'] | [b'a', b'r'] => {
                    options.alignment = alignment();
                    Ok(())
                },
                [b'c', b'o'] => {
                    let columns = parser.read_number()?.points();
                    let columns = u8::try_from(columns).map_err(|_| {
                        malformed("RTF EQ array column count exceeds the safety limit")
                    })?;
                    if columns > MAX_ARRAY_COLUMNS {
                        return Err(malformed(
                            "RTF EQ array column count exceeds the safety limit",
                        ));
                    }
                    options.columns = Some(columns);
                    Ok(())
                },
                [b'v', b's'] => {
                    options.vertical_spacing = Some(parser.read_number()?);
                    Ok(())
                },
                [b'h', b's'] => {
                    options.horizontal_spacing = Some(parser.read_number()?);
                    Ok(())
                },
                _ => Err(malformed("RTF EQ array switch has an unknown option")),
            },
            EquationSwitch::Bracket(options) => match name {
                [b'l', b'c'] => {
                    options.left = Some(parser.read_char()?);
                    Ok(())
                },
                [b'r', b'c'] => {
                    options.right = Some(parser.read_char()?);
                    Ok(())
                },
                [b'b', b'c'] => {
                    let bracket = parser.read_char()?;
                    options.left = Some(bracket);
                    options.right = Some(bracket);
                    Ok(())
                },
                _ => Err(malformed("RTF EQ bracket switch has an unknown option")),
            },
            EquationSwitch::Displace(options) => match name {
                [b'f', b'o'] => {
                    options.forward = Some(parser.read_number()?);
                    Ok(())
                },
                [b'b', b'a'] => {
                    options.backward = Some(parser.read_number()?);
                    Ok(())
                },
                [b'l', b'i'] => {
                    options.underline = true;
                    Ok(())
                },
                _ => Err(malformed(
                    "RTF EQ displacement switch has an unknown option",
                )),
            },
            EquationSwitch::Integral(options) => match name {
                [b's', b'u'] => {
                    options.symbol = EquationIntegralSymbol::Summation;
                    Ok(())
                },
                [b'p', b'r'] => {
                    options.symbol = EquationIntegralSymbol::Product;
                    Ok(())
                },
                [b'i', b'n'] => {
                    options.inline_limits = true;
                    Ok(())
                },
                [b'f', b'c'] => {
                    options.fixed_char = Some(parser.read_char()?);
                    Ok(())
                },
                [b'v', b'c'] => {
                    options.variable_char = Some(parser.read_char()?);
                    Ok(())
                },
                _ => Err(malformed("RTF EQ integral switch has an unknown option")),
            },
            EquationSwitch::Overstrike(options) => match name {
                [b'a', b'l'] | [b'a', b'c'] | [b'a', b'r'] => {
                    options.alignment = alignment();
                    Ok(())
                },
                _ => Err(malformed("RTF EQ overstrike switch has an unknown option")),
            },
            EquationSwitch::Script(options) => match name {
                [b'u', b'p'] => {
                    options.up = Some(parser.read_number()?);
                    Ok(())
                },
                [b'd', b'o'] => {
                    options.down = Some(parser.read_number()?);
                    Ok(())
                },
                _ => Err(malformed("RTF EQ script switch has an unknown option")),
            },
            EquationSwitch::Box(options) => match name {
                [b't', b'o'] => {
                    options.top = true;
                    Ok(())
                },
                [b'b', b'o'] => {
                    options.bottom = true;
                    Ok(())
                },
                [b'l', b'e'] => {
                    options.left = true;
                    Ok(())
                },
                [b'r', b'i'] => {
                    options.right = true;
                    Ok(())
                },
                _ => Err(malformed("RTF EQ box switch has an unknown option")),
            },
            EquationSwitch::Fraction | EquationSwitch::List | EquationSwitch::Radical => {
                Err(malformed("RTF EQ switch does not accept sub-options"))
            },
        }
    }

    /// Parse the optional parenthesized element list after a switch.
    fn parse_elements(&mut self) -> RtfResult<Vec<&'a str>> {
        if self.peek() != Some('(') {
            return Ok(Vec::new());
        }
        self.position += 1;
        let mut elements = Vec::new();
        let mut depth = 1usize;
        let mut element_start = self.position;
        loop {
            let Some(character) = self.peek() else {
                return Err(malformed("RTF EQ element list is unterminated"));
            };
            match character {
                '(' => depth += 1,
                ')' => {
                    depth -= 1;
                    if depth == 0 {
                        elements.push(self.input[element_start..self.position].trim());
                        self.position += 1;
                        if elements.len() > MAX_GROUP_ELEMENTS {
                            return Err(malformed("RTF EQ group exceeds the element limit"));
                        }
                        return Ok(elements);
                    }
                },
                ',' if depth == 1 => {
                    elements.push(self.input[element_start..self.position].trim());
                    if elements.len() > MAX_GROUP_ELEMENTS {
                        return Err(malformed("RTF EQ group exceeds the element limit"));
                    }
                    element_start = self.position + 1;
                },
                _ => {},
            }
            self.advance_char();
        }
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    fn only_group<'a>(model: &'a EquationModel<'a>) -> &'a EquationGroup<'a> {
        assert_eq!(model.segments().len(), 1);
        match &model.segments()[0] {
            EquationSegment::Switched(group) => group,
            EquationSegment::Literal(_) => panic!("expected a switched group"),
        }
    }

    #[test]
    fn parses_fraction_with_two_elements() {
        let model = EquationModel::parse(r"\f(x+1, y-2)").unwrap();
        let group = only_group(&model);
        assert_eq!(group.switch, EquationSwitch::Fraction);
        assert_eq!(group.elements, ["x+1", "y-2"]);
    }

    #[test]
    fn parses_array_with_all_options() {
        let model = EquationModel::parse(r"\a\ac\co2\vs4\hs6(a,b,c,d)").unwrap();
        let group = only_group(&model);
        assert_eq!(
            group.switch,
            EquationSwitch::Array(EquationArray {
                alignment: Some(EquationAlignment::Center),
                columns: Some(2),
                vertical_spacing: Some(EquationSpacing(4)),
                horizontal_spacing: Some(EquationSpacing(6)),
            })
        );
        assert_eq!(group.elements, ["a", "b", "c", "d"]);
    }

    #[test]
    fn parses_bracket_with_nested_parentheses_in_element() {
        let model = EquationModel::parse(r"\b\bc\{(\f(1,x)+1)").unwrap();
        let group = only_group(&model);
        assert_eq!(
            group.switch,
            EquationSwitch::Bracket(EquationBracket {
                left: Some('{'),
                right: Some('{'),
            })
        );
        assert_eq!(group.elements, [r"\f(1,x)+1"]);
    }

    #[test]
    fn parses_integral_symbol_and_limits() {
        let model = EquationModel::parse(r"\i\su(0,n,term)").unwrap();
        let group = only_group(&model);
        assert_eq!(
            group.switch,
            EquationSwitch::Integral(EquationIntegral {
                symbol: EquationIntegralSymbol::Summation,
                inline_limits: false,
                fixed_char: None,
                variable_char: None,
            })
        );
        assert_eq!(group.elements, ["0", "n", "term"]);

        let model = EquationModel::parse(r"\i\in\fc\|((,),(x))").unwrap();
        let group = only_group(&model);
        match group.switch {
            EquationSwitch::Integral(options) => {
                assert!(options.inline_limits);
                assert_eq!(options.fixed_char, Some('|'));
                assert_eq!(options.symbol, EquationIntegralSymbol::Integral);
            },
            _ => panic!("expected an integral switch"),
        }
    }

    #[test]
    fn parses_radical_with_and_without_index() {
        let model = EquationModel::parse(r"\r(3,x)").unwrap();
        assert_eq!(only_group(&model).elements, ["3", "x"]);
        let model = EquationModel::parse(r"\r(x)").unwrap();
        let group = only_group(&model);
        assert_eq!(group.switch, EquationSwitch::Radical);
        assert_eq!(group.elements, ["x"]);
    }

    #[test]
    fn parses_script_displacements() {
        let model = EquationModel::parse(r"\s\up9(sup)\s\do5(sub)").unwrap();
        assert_eq!(model.segments().len(), 2);
        let mut groups = model.groups();
        assert_eq!(
            groups.next().unwrap().switch,
            EquationSwitch::Script(EquationScript {
                up: Some(EquationSpacing(9)),
                down: None,
            })
        );
        assert_eq!(
            groups.next().unwrap().switch,
            EquationSwitch::Script(EquationScript {
                up: None,
                down: Some(EquationSpacing(5)),
            })
        );
        assert!(groups.next().is_none());
    }

    #[test]
    fn parses_box_displace_overstrike_and_list() {
        let model = EquationModel::parse(r"\x\to\bo(b)").unwrap();
        assert_eq!(
            only_group(&model).switch,
            EquationSwitch::Box(EquationBox {
                top: true,
                bottom: true,
                left: false,
                right: false,
            })
        );

        let model = EquationModel::parse(r"\d\fo10\li(x)").unwrap();
        assert_eq!(
            only_group(&model).switch,
            EquationSwitch::Displace(EquationDisplace {
                forward: Some(EquationSpacing(10)),
                backward: None,
                underline: true,
            })
        );

        let model = EquationModel::parse(r"\o\al(a,b,c)").unwrap();
        let group = only_group(&model);
        assert_eq!(
            group.switch,
            EquationSwitch::Overstrike(EquationOverstrike {
                alignment: Some(EquationAlignment::Left),
            })
        );
        assert_eq!(group.elements, ["a", "b", "c"]);

        let model = EquationModel::parse(r"\l(a;b,c)").unwrap();
        let group = only_group(&model);
        assert_eq!(group.switch, EquationSwitch::List);
        assert_eq!(group.elements, ["a;b", "c"]);
    }

    #[test]
    fn preserves_literal_text_between_groups() {
        let model = EquationModel::parse(r"2\f(1,2) + \r(x)").unwrap();
        assert_eq!(
            model.segments(),
            &[
                EquationSegment::Literal("2"),
                EquationSegment::Switched(EquationGroup {
                    switch: EquationSwitch::Fraction,
                    elements: vec!["1", "2"],
                }),
                EquationSegment::Literal("+"),
                EquationSegment::Switched(EquationGroup {
                    switch: EquationSwitch::Radical,
                    elements: vec!["x"],
                }),
            ]
        );
    }

    #[test]
    fn switch_without_parentheses_has_no_elements() {
        let model = EquationModel::parse(r"\f x").unwrap();
        assert_eq!(model.groups().next().unwrap().elements, Vec::<&str>::new());
        assert!(matches!(model.segments()[1], EquationSegment::Literal("x")));
    }

    #[test]
    fn rejects_unknown_switches_and_malformed_input() {
        assert!(EquationModel::parse(r"\q(1)").is_err());
        assert!(EquationModel::parse(r"\f(a").is_err());
        assert!(EquationModel::parse(r"\a\co(a,b)").is_err());
        assert!(EquationModel::parse(r"\s\up(x)").is_err());
        assert!(EquationModel::parse(r"\b\bc").is_err());
        assert!(EquationModel::parse(r"\f\zz(1)").is_err());
        assert!(EquationModel::parse("\\").is_err());
    }

    #[test]
    fn enforces_group_and_element_limits() {
        let too_many_elements = format!("\\l({})", "a,".repeat(MAX_GROUP_ELEMENTS + 1));
        assert!(EquationModel::parse(&too_many_elements).is_err());
        let too_many_groups = "\\f(a,b)".repeat(MAX_EQUATION_GROUPS + 1);
        assert!(EquationModel::parse(&too_many_groups).is_err());
    }
}
