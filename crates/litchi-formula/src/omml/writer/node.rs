// Per-node OMML serialization
//
// Each `MathNode` variant maps onto the corresponding OMML construct from
// ECMA-376 Part 1, §22.1. Structural components (numerator, base, limits, ...)
// are wrapped in the container elements the parser expects, so a
// serialize/parse round-trip preserves the AST.

use super::OmmlWriter;
use super::chars::*;
use super::names::*;
use crate::ast::{BorderBoxStyle, EqArrayProperties, Fence, MathNode, MatrixProperties, StyleType};
use crate::omml::error::OmmlError;
use std::borrow::Cow;

impl OmmlWriter {
    /// Serialize a single AST node
    pub(super) fn write_node(&mut self, node: &MathNode) -> Result<(), OmmlError> {
        match node {
            MathNode::Text(text) | MathNode::Number(text) => {
                self.text_run(text);
                Ok(())
            },
            MathNode::Operator(op) => {
                self.text_run(operator_char(*op));
                Ok(())
            },
            MathNode::Symbol(symbol) => {
                match symbol.unicode {
                    Some(ch) => {
                        let mut buf = [0u8; 4];
                        self.text_run(ch.encode_utf8(&mut buf));
                    },
                    None => self.text_run(&symbol.name),
                }
                Ok(())
            },
            MathNode::PredefinedSymbol(symbol) => {
                self.text_run(predefined_symbol_char(*symbol));
                Ok(())
            },
            MathNode::Frac {
                numerator,
                denominator,
                frac_type,
                ..
            } => self.write_fraction(numerator, denominator, *frac_type),
            MathNode::Root { base, index } => self.write_radical(base, index.as_deref()),
            MathNode::Power { base, exponent } => self.element(EL_SUPERSCRIPT, |w| {
                w.base_element(base)?;
                w.wrapped_nodes(EL_SUP, exponent)
            }),
            MathNode::Sub { base, subscript } => self.element(EL_SUBSCRIPT, |w| {
                w.base_element(base)?;
                w.wrapped_nodes(EL_SUB, subscript)
            }),
            MathNode::SubSup {
                base,
                subscript,
                superscript,
            } => self.element(EL_SUB_SUP, |w| {
                w.base_element(base)?;
                w.wrapped_nodes(EL_SUB, subscript)?;
                w.wrapped_nodes(EL_SUP, superscript)
            }),
            MathNode::PreSub {
                base,
                pre_subscript,
            } => self.write_pre_script(base, Some(pre_subscript.as_slice()), None),
            MathNode::PreSup {
                base,
                pre_superscript,
            } => self.write_pre_script(base, None, Some(pre_superscript.as_slice())),
            MathNode::PreSubSup {
                base,
                pre_subscript,
                pre_superscript,
            } => self.write_pre_script(
                base,
                Some(pre_subscript.as_slice()),
                Some(pre_superscript.as_slice()),
            ),
            MathNode::Under { base, under, .. } => self.element(EL_LIM_LOW, |w| {
                w.base_element(base)?;
                w.wrapped_nodes(EL_LIMIT, under)
            }),
            MathNode::Over { base, over, .. } => self.element(EL_LIM_UPP, |w| {
                w.base_element(base)?;
                w.wrapped_nodes(EL_LIMIT, over)
            }),
            MathNode::UnderOver {
                base, under, over, ..
            } => self.element(EL_LIM_LOW, |w| {
                w.element(EL_ELEMENT, |w| {
                    w.element(EL_LIM_UPP, |w| {
                        w.base_element(base)?;
                        w.wrapped_nodes(EL_LIMIT, over)
                    })
                })?;
                w.wrapped_nodes(EL_LIMIT, under)
            }),
            MathNode::Fenced {
                open,
                content,
                close,
                separator,
            } => self.write_fenced(*open, content, *close, separator.as_deref()),
            MathNode::LargeOp {
                operator,
                lower_limit,
                upper_limit,
                integrand,
                hide_lower,
                hide_upper,
            } => self.write_nary(
                large_operator_char(*operator),
                lower_limit.as_deref(),
                upper_limit.as_deref(),
                integrand.as_deref(),
                *hide_lower,
                *hide_upper,
            ),
            MathNode::Function { name, argument } => self.write_function(name, argument),
            MathNode::PredefinedFunction { function, argument } => {
                self.write_function(function_name_str(*function), argument)
            },
            MathNode::Matrix {
                rows,
                fence_type,
                properties,
            } => {
                let (open, close) = matrix_fence_pair(*fence_type);
                if open == Fence::None && close == Fence::None {
                    self.write_matrix(rows, properties.as_ref())
                } else {
                    // OMML expresses fenced matrices as a matrix inside a
                    // delimiter (ECMA-376 has no fence on m:m itself).
                    self.element(EL_DELIMITER, |w| {
                        w.element(EL_DELIMITER_PROPS, |w| {
                            w.val_element(EL_BEGIN_CHAR, fence_open_char(open));
                            w.val_element(EL_END_CHAR, fence_close_char(close));
                            Ok(())
                        })?;
                        w.element(EL_ELEMENT, |w| w.write_matrix(rows, properties.as_ref()))
                    })
                }
            },
            MathNode::EqArray { rows, properties } => {
                self.write_eq_array(rows, properties.as_ref())
            },
            MathNode::Accent {
                base,
                accent,
                position,
            } => self.element(EL_ACCENT, |w| {
                w.element(EL_ACCENT_PROPS, |w| {
                    w.val_element(EL_CHAR, accent_char(*accent));
                    if let Some(position) = position {
                        w.val_element(EL_POSITION, position_value(*position));
                    }
                    Ok(())
                })?;
                w.base_element(base)
            }),
            MathNode::Bar { base, position } => self.element(EL_BAR, |w| {
                if let Some(position) = position {
                    w.element(EL_BAR_PROPS, |w| {
                        w.val_element(EL_POSITION, position_value(*position));
                        Ok(())
                    })?;
                }
                w.base_element(base)
            }),
            MathNode::BorderBox { content, style } => self.element(EL_BORDER_BOX, |w| {
                if let Some(style) = style {
                    w.write_border_box_props(style)?;
                }
                w.base_element(content)
            }),
            MathNode::GroupChar {
                base,
                character,
                position,
                vertical_alignment,
            } => self.element(EL_GROUP_CHAR, |w| {
                let has_props =
                    character.is_some() || position.is_some() || vertical_alignment.is_some();
                if has_props {
                    w.element(EL_GROUP_CHAR_PROPS, |w| {
                        if let Some(character) = character {
                            w.val_element(EL_CHAR, character);
                        }
                        if let Some(position) = position {
                            w.val_element(EL_POSITION, position_value(*position));
                        }
                        if let Some(alignment) = vertical_alignment {
                            w.val_element(EL_VERT_JC, vertical_alignment_value(*alignment));
                        }
                        Ok(())
                    })?;
                }
                w.base_element(base)
            }),
            MathNode::Space(space) => {
                self.text_run(space_char(*space));
                Ok(())
            },
            MathNode::LineBreak => {
                self.text_run("\n");
                Ok(())
            },
            MathNode::Style { style, content } => {
                self.write_run(content, None, Some(*style), None, None)
            },
            MathNode::Run {
                content,
                literal,
                style,
                font,
                color,
                // Underline/overline/strike decorations have no faithful OMML
                // run-property mapping in this writer yet.
                underline: _,
                overline: _,
                strike_through: _,
                double_strike_through: _,
            } => self.write_run(content, *literal, *style, font.as_deref(), color.as_deref()),
            MathNode::Row(children) => self.write_all(children),
            MathNode::Phantom(content) => self.element(EL_PHANTOM, |w| w.base_element(content)),
            MathNode::Limit { content, .. } => self.write_all(content),
            MathNode::Degree(content)
            | MathNode::Base(content)
            | MathNode::Argument(content)
            | MathNode::Numerator(content)
            | MathNode::Denominator(content)
            | MathNode::Integrand(content)
            | MathNode::LowerLimit(content)
            | MathNode::UpperLimit(content) => self.write_all(content),
            MathNode::Error(message) => Err(OmmlError::SerializationError(format!(
                "cannot serialize error node: {}",
                message
            ))),
        }
    }

    /// `m:f` — fraction with optional `m:type` property
    fn write_fraction(
        &mut self,
        numerator: &[MathNode],
        denominator: &[MathNode],
        frac_type: Option<crate::ast::FractionType>,
    ) -> Result<(), OmmlError> {
        self.element(EL_FRACTION, |w| {
            if let Some(frac_type) = frac_type {
                w.element(EL_FRACTION_PROPS, |w| {
                    w.val_element(EL_FRACTION_TYPE, fraction_type_value(frac_type));
                    Ok(())
                })?;
            }
            w.wrapped_nodes(EL_NUMERATOR, numerator)?;
            w.wrapped_nodes(EL_DENOMINATOR, denominator)
        })
    }

    /// `m:rad` — radical; a missing degree serializes as a hidden degree
    fn write_radical(
        &mut self,
        base: &[MathNode],
        index: Option<&[MathNode]>,
    ) -> Result<(), OmmlError> {
        self.element(EL_RADICAL, |w| {
            match index {
                Some(index) => {
                    w.wrapped_nodes(EL_DEGREE, index)?;
                },
                None => {
                    w.element(EL_RADICAL_PROPS, |w| {
                        w.val_element(EL_DEGREE_HIDE, VAL_ON);
                        Ok(())
                    })?;
                    w.empty_element(EL_DEGREE);
                },
            }
            w.base_element(base)
        })
    }

    /// `m:sPre` — pre-scripts (sub, sup, base order per ECMA-376)
    fn write_pre_script(
        &mut self,
        base: &[MathNode],
        subscript: Option<&[MathNode]>,
        superscript: Option<&[MathNode]>,
    ) -> Result<(), OmmlError> {
        self.element(EL_PRE_SCRIPT, |w| {
            w.wrapped_nodes(EL_SUB, subscript.unwrap_or_default())?;
            w.wrapped_nodes(EL_SUP, superscript.unwrap_or_default())?;
            w.base_element(base)
        })
    }

    /// `m:d` (or `m:box` for a fenceless group) — delimiters
    fn write_fenced(
        &mut self,
        open: Fence,
        content: &[MathNode],
        close: Fence,
        separator: Option<&str>,
    ) -> Result<(), OmmlError> {
        if open == Fence::None && close == Fence::None && separator.is_none() {
            return self.element(EL_BOX, |w| w.base_element(content));
        }

        self.element(EL_DELIMITER, |w| {
            w.element(EL_DELIMITER_PROPS, |w| {
                w.val_element(EL_BEGIN_CHAR, fence_open_char(open));
                if let Some(separator) = separator {
                    w.val_element(EL_SEPARATOR_CHAR, separator);
                }
                w.val_element(EL_END_CHAR, fence_close_char(close));
                Ok(())
            })?;
            w.base_element(content)
        })
    }

    /// `m:nary` — n-ary operator with optional limits
    fn write_nary(
        &mut self,
        operator: &str,
        lower_limit: Option<&[MathNode]>,
        upper_limit: Option<&[MathNode]>,
        integrand: Option<&[MathNode]>,
        hide_lower: bool,
        hide_upper: bool,
    ) -> Result<(), OmmlError> {
        self.element(EL_NARY, |w| {
            w.element(EL_NARY_PROPS, |w| {
                w.val_element(EL_CHAR, operator);
                if hide_lower {
                    w.val_element(EL_SUB_HIDE, VAL_ON);
                }
                if hide_upper {
                    w.val_element(EL_SUP_HIDE, VAL_ON);
                }
                Ok(())
            })?;
            w.wrapped_nodes(EL_SUB, lower_limit.unwrap_or_default())?;
            w.wrapped_nodes(EL_SUP, upper_limit.unwrap_or_default())?;
            w.wrapped_nodes(EL_ELEMENT, integrand.unwrap_or_default())
        })
    }

    /// `m:func` — function application
    fn write_function(&mut self, name: &str, argument: &[MathNode]) -> Result<(), OmmlError> {
        self.element(EL_FUNCTION, |w| {
            w.element(EL_FUNCTION_NAME, |w| {
                w.text_run(name);
                Ok(())
            })?;
            w.base_element(argument)
        })
    }

    /// `m:m` — matrix body (without any surrounding delimiter)
    fn write_matrix(
        &mut self,
        rows: &[Vec<Vec<MathNode>>],
        properties: Option<&MatrixProperties>,
    ) -> Result<(), OmmlError> {
        self.element(EL_MATRIX, |w| {
            if let Some(properties) = properties {
                let base_jc = properties.base_alignment.and_then(base_alignment_value);
                let row_spacing = properties.row_spacing;
                if base_jc.is_some() || row_spacing.is_some() {
                    w.element(EL_MATRIX_PROPS, |w| {
                        if let Some(base_jc) = base_jc {
                            w.val_element(EL_BASE_JC, base_jc);
                        }
                        if let Some(row_spacing) = row_spacing {
                            w.val_element(EL_ROW_SPACING, &row_spacing.to_string());
                        }
                        Ok(())
                    })?;
                }
            }
            for row in rows {
                w.element(EL_MATRIX_ROW, |w| {
                    for cell in row {
                        w.wrapped_nodes(EL_ELEMENT, cell)?;
                    }
                    Ok(())
                })?;
            }
            Ok(())
        })
    }

    /// `m:eqArr` — equation array
    fn write_eq_array(
        &mut self,
        rows: &[Vec<MathNode>],
        properties: Option<&EqArrayProperties>,
    ) -> Result<(), OmmlError> {
        self.element(EL_EQ_ARRAY, |w| {
            if let Some(properties) = properties {
                w.write_eq_array_props(properties)?;
            }
            for row in rows {
                w.wrapped_nodes(EL_ELEMENT, row)?;
            }
            Ok(())
        })
    }

    /// `m:eqArrPr` — equation array properties
    fn write_eq_array_props(&mut self, properties: &EqArrayProperties) -> Result<(), OmmlError> {
        let base_jc = properties.base_alignment.and_then(base_alignment_value);
        let has_props = base_jc.is_some()
            || properties.max_distance.is_some()
            || properties.object_distance.is_some()
            || properties.row_spacing.is_some()
            || properties.row_spacing_rule.is_some();
        if !has_props {
            return Ok(());
        }

        self.element(EL_EQ_ARRAY_PROPS, |w| {
            if let Some(base_jc) = base_jc {
                w.val_element(EL_BASE_JC, base_jc);
            }
            if let Some(max_distance) = properties.max_distance {
                w.val_element(EL_MAX_DIST, &max_distance.to_string());
            }
            if let Some(object_distance) = properties.object_distance {
                w.val_element(EL_OBJ_DIST, &object_distance.to_string());
            }
            if let Some(row_spacing) = properties.row_spacing {
                w.val_element(EL_ROW_SPACING, &row_spacing.to_string());
            }
            if let Some(rule) = properties.row_spacing_rule.as_deref() {
                w.val_element(EL_ROW_SPACING_RULE, rule);
            }
            Ok(())
        })
    }

    /// `m:borderBoxPr` — border box hide/strike flags
    fn write_border_box_props(&mut self, style: &BorderBoxStyle) -> Result<(), OmmlError> {
        let flags: [(&str, bool); 8] = [
            (EL_HIDE_TOP, style.hide_top),
            (EL_HIDE_BOT, style.hide_bottom),
            (EL_HIDE_LEFT, style.hide_left),
            (EL_HIDE_RIGHT, style.hide_right),
            (EL_STRIKE_H, style.strike_horizontal),
            (EL_STRIKE_V, style.strike_vertical),
            (EL_STRIKE_BLTR, style.strike_bltr),
            (EL_STRIKE_TLBR, style.strike_tlbr),
        ];
        if flags.iter().all(|(_, enabled)| !enabled) {
            return Ok(());
        }
        self.element(EL_BORDER_BOX_PROPS, |w| {
            for (name, enabled) in flags {
                if enabled {
                    w.val_element(name, VAL_ON);
                }
            }
            Ok(())
        })
    }

    /// `m:r` — run with optional properties; leaf children are emitted as
    /// `m:t` inside the run, structural children after it.
    fn write_run(
        &mut self,
        content: &[MathNode],
        literal: Option<bool>,
        style: Option<StyleType>,
        font: Option<&str>,
        color: Option<&str>,
    ) -> Result<(), OmmlError> {
        self.open_element(EL_RUN);
        self.write_run_props(literal, style, font, color);

        let mut structural: Vec<&MathNode> = Vec::new();
        for child in content {
            match leaf_text(child) {
                Some(text) => {
                    self.open_element(EL_TEXT);
                    self.push_text(&text);
                    self.close_element(EL_TEXT);
                },
                None => structural.push(child),
            }
        }
        self.close_element(EL_RUN);

        for child in structural {
            self.write_node(child)?;
        }
        Ok(())
    }

    /// Emit `<m:rPr>` when any run property is present
    fn write_run_props(
        &mut self,
        literal: Option<bool>,
        style: Option<StyleType>,
        font: Option<&str>,
        color: Option<&str>,
    ) {
        if literal.is_none() && style.is_none() && font.is_none() && color.is_none() {
            return;
        }

        // Color is carried as an attribute on m:rPr (matching the parser)
        self.buffer.push('<');
        self.buffer.push_str(EL_RUN_PROPS);
        if let Some(color) = color {
            self.push_attr(ATTR_COLOR, color);
        }
        self.buffer.push('>');

        if let Some(literal) = literal {
            self.val_element(EL_LITERAL, if literal { VAL_ON } else { VAL_OFF });
        }
        if let Some(style) = style {
            let (element, value) = style_value(style);
            match element {
                StyleElement::Script => self.val_element(EL_SCRIPT_STYLE, value),
                StyleElement::Style => self.val_element(EL_STYLE, value),
            }
        }
        if let Some(font) = font {
            // The parser reads the normal-text font from m:nor text content
            self.open_element(EL_NORMAL_TEXT);
            self.push_text(font);
            self.close_element(EL_NORMAL_TEXT);
        }

        self.close_element(EL_RUN_PROPS);
    }
}

/// Text content for nodes that can live inside a run as `m:t`
fn leaf_text<'a>(node: &'a MathNode) -> Option<Cow<'a, str>> {
    match node {
        MathNode::Text(text) | MathNode::Number(text) => Some(Cow::Borrowed(text.as_ref())),
        MathNode::Operator(op) => Some(Cow::Borrowed(operator_char(*op))),
        MathNode::PredefinedSymbol(symbol) => Some(Cow::Borrowed(predefined_symbol_char(*symbol))),
        MathNode::Symbol(symbol) => match symbol.unicode {
            Some(ch) => Some(Cow::Owned(ch.to_string())),
            None => Some(Cow::Borrowed(symbol.name.as_ref())),
        },
        MathNode::Space(space) => Some(Cow::Borrowed(space_char(*space))),
        _ => None,
    }
}
