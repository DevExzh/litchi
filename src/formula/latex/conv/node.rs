// Node conversion logic for LaTeX conversion
//
// This module contains the core node conversion logic for converting
// MathNode AST elements to LaTeX format with optimized performance.

use super::converter::LatexConverter;
use super::error::LatexError;
use crate::formula::ast::{
    FunctionName, LimitType, MathNode, Position, PredefinedSymbol, VerticalAlignment,
};
use crate::formula::latex::operators::{
    accent_to_latex, fence_to_latex, is_standard_function, large_operator_to_latex,
    operator_to_latex, space_to_latex, style_to_latex,
};
use crate::formula::latex::symbols::convert_symbol;
use crate::formula::latex::templates::needs_grouping_for_scripts;
use crate::formula::latex::utils::{
    escape_latex_special_chars, is_valid_number_fast, needs_latex_protection,
};
use std::fmt::Write;

/// Convert predefined symbol to LaTeX
fn predefined_symbol_to_latex(symbol: PredefinedSymbol) -> &'static str {
    match symbol {
        PredefinedSymbol::Alpha => "\\alpha",
        PredefinedSymbol::Beta => "\\beta",
        PredefinedSymbol::Gamma => "\\gamma",
        PredefinedSymbol::Delta => "\\delta",
        PredefinedSymbol::Epsilon => "\\epsilon",
        PredefinedSymbol::Zeta => "\\zeta",
        PredefinedSymbol::Eta => "\\eta",
        PredefinedSymbol::Theta => "\\theta",
        PredefinedSymbol::Iota => "\\iota",
        PredefinedSymbol::Kappa => "\\kappa",
        PredefinedSymbol::Lambda => "\\lambda",
        PredefinedSymbol::Mu => "\\mu",
        PredefinedSymbol::Nu => "\\nu",
        PredefinedSymbol::Xi => "\\xi",
        PredefinedSymbol::Omicron => "o",
        PredefinedSymbol::Pi => "\\pi",
        PredefinedSymbol::Rho => "\\rho",
        PredefinedSymbol::Sigma => "\\sigma",
        PredefinedSymbol::Tau => "\\tau",
        PredefinedSymbol::Upsilon => "\\upsilon",
        PredefinedSymbol::Phi => "\\phi",
        PredefinedSymbol::Chi => "\\chi",
        PredefinedSymbol::Psi => "\\psi",
        PredefinedSymbol::Omega => "\\omega",
        PredefinedSymbol::AlphaCap => "A",
        PredefinedSymbol::BetaCap => "B",
        PredefinedSymbol::GammaCap => "\\Gamma",
        PredefinedSymbol::DeltaCap => "\\Delta",
        PredefinedSymbol::EpsilonCap => "E",
        PredefinedSymbol::ZetaCap => "Z",
        PredefinedSymbol::EtaCap => "H",
        PredefinedSymbol::ThetaCap => "\\Theta",
        PredefinedSymbol::IotaCap => "I",
        PredefinedSymbol::KappaCap => "K",
        PredefinedSymbol::LambdaCap => "\\Lambda",
        PredefinedSymbol::MuCap => "M",
        PredefinedSymbol::NuCap => "N",
        PredefinedSymbol::XiCap => "\\Xi",
        PredefinedSymbol::OmicronCap => "O",
        PredefinedSymbol::PiCap => "\\Pi",
        PredefinedSymbol::RhoCap => "P",
        PredefinedSymbol::SigmaCap => "\\Sigma",
        PredefinedSymbol::TauCap => "T",
        PredefinedSymbol::UpsilonCap => "\\Upsilon",
        PredefinedSymbol::PhiCap => "\\Phi",
        PredefinedSymbol::ChiCap => "X",
        PredefinedSymbol::PsiCap => "\\Psi",
        PredefinedSymbol::OmegaCap => "\\Omega",
        PredefinedSymbol::Aleph => "\\aleph",
        PredefinedSymbol::EulerGamma => "\\gamma",
        PredefinedSymbol::ExponentialE => "e",
        PredefinedSymbol::ImaginaryI => "i",
        PredefinedSymbol::Infinity => "\\infty",
    }
}

/// Convert function name to LaTeX
fn function_name_to_latex(function: FunctionName) -> &'static str {
    match function {
        FunctionName::Sin => "\\sin",
        FunctionName::Cos => "\\cos",
        FunctionName::Tan => "\\tan",
        FunctionName::Sec => "\\sec",
        FunctionName::Csc => "\\csc",
        FunctionName::Cot => "\\cot",
        FunctionName::ArcSin => "\\arcsin",
        FunctionName::ArcCos => "\\arccos",
        FunctionName::ArcTan => "\\arctan",
        FunctionName::ArcSec => "\\arcsec",
        FunctionName::ArcCsc => "\\arccsc",
        FunctionName::ArcCot => "\\arccot",
        FunctionName::Sinh => "\\sinh",
        FunctionName::Cosh => "\\cosh",
        FunctionName::Tanh => "\\tanh",
        FunctionName::Sech => "\\sech",
        FunctionName::Csch => "\\csch",
        FunctionName::Coth => "\\coth",
        FunctionName::Log => "\\log",
        FunctionName::Ln => "\\ln",
        FunctionName::Exp => "\\exp",
        FunctionName::Sqrt => "\\sqrt",
        FunctionName::Min => "\\min",
        FunctionName::Max => "\\max",
        FunctionName::Sup => "\\sup",
        FunctionName::Inf => "\\inf",
        FunctionName::Lim => "\\lim",
        FunctionName::Det => "\\det",
        FunctionName::Trace => "\\trace",
        FunctionName::Dim => "\\dim",
        FunctionName::Ker => "\\ker",
        FunctionName::Im => "\\Im",
        FunctionName::Re => "\\Re",
        FunctionName::Arg => "\\arg",
        FunctionName::Mod => "\\mod",
        FunctionName::Gcd => "\\gcd",
        FunctionName::Lcm => "\\lcm",
    }
}

impl LatexConverter {
    /// Convert a single MathNode to LaTeX format
    pub fn convert_node(&mut self, node: &MathNode) -> Result<(), LatexError> {
        convert_node_internal(self, node)
    }
}

/// Internal node conversion function
fn convert_node_internal(
    converter: &mut LatexConverter,
    node: &MathNode,
) -> Result<(), LatexError> {
    converter.stats.record_node();

    match node {
        MathNode::Text(text) => {
            if needs_latex_protection(text) {
                super::utils::extend_buffer_with_capacity(
                    &mut converter.buffer,
                    "\\text{",
                    text.len() + 2,
                );
                converter.buffer.push_str("\\text{");
                if escape_latex_special_chars(text, &mut converter.buffer) {
                    converter.stats.record_allocation(text.len());
                }
                converter.buffer.push('}');
            } else {
                super::utils::extend_buffer_with_capacity(&mut converter.buffer, text, 0);
            }
        },
        MathNode::Number(num) => {
            // Fast validation for numbers (helps with malformed input)
            debug_assert!(is_valid_number_fast(num), "Invalid number format: {num}");
            super::utils::extend_buffer_with_capacity(&mut converter.buffer, num, 0);
        },
        MathNode::Operator(op) => {
            let op_str = operator_to_latex(*op);
            converter.append_cached_command(op_str);
        },
        MathNode::Symbol(sym) => {
            convert_symbol(&mut converter.buffer, sym)?;
        },
        MathNode::PredefinedSymbol(symbol) => {
            let symbol_str = predefined_symbol_to_latex(*symbol);
            converter.append_cached_command(symbol_str);
        },
        MathNode::Frac {
            numerator,
            denominator,
            ..
        } => {
            converter.append_cached_command("\\frac{");
            for n in numerator.iter() {
                convert_node_internal(converter, n)?;
            }
            converter.buffer.push_str("}{");
            for n in denominator.iter() {
                convert_node_internal(converter, n)?;
            }
            converter.buffer.push('}');
        },
        MathNode::Root { base, index } => {
            if let Some(idx) = index {
                converter.buffer.push_str("\\sqrt[");
                for n in idx.iter() {
                    convert_node_internal(converter, n)?;
                }
                converter.buffer.push_str("]{");
            } else {
                converter.buffer.push_str("\\sqrt{");
            }
            for n in base.iter() {
                convert_node_internal(converter, n)?;
            }
            converter.buffer.push('}');
        },
        MathNode::Power { base, exponent } => {
            if needs_grouping_for_scripts(base) {
                converter.buffer.push('{');
                for n in base.iter() {
                    convert_node_internal(converter, n)?;
                }
                converter.buffer.push('}');
            } else {
                for n in base.iter() {
                    convert_node_internal(converter, n)?;
                }
            }
            converter.buffer.push_str("^{");
            for n in exponent.iter() {
                convert_node_internal(converter, n)?;
            }
            converter.buffer.push('}');
        },
        MathNode::Sub { base, subscript } => {
            if needs_grouping_for_scripts(base) {
                converter.buffer.push('{');
                for n in base.iter() {
                    convert_node_internal(converter, n)?;
                }
                converter.buffer.push('}');
            } else {
                for n in base.iter() {
                    convert_node_internal(converter, n)?;
                }
            }
            converter.buffer.push_str("_{");
            for n in subscript.iter() {
                convert_node_internal(converter, n)?;
            }
            converter.buffer.push('}');
        },
        MathNode::SubSup {
            base,
            subscript,
            superscript,
        } => {
            if needs_grouping_for_scripts(base) {
                converter.buffer.push('{');
                for n in base.iter() {
                    convert_node_internal(converter, n)?;
                }
                converter.buffer.push('}');
            } else {
                for n in base.iter() {
                    convert_node_internal(converter, n)?;
                }
            }
            converter.buffer.push_str("_{");
            for n in subscript.iter() {
                convert_node_internal(converter, n)?;
            }
            converter.buffer.push_str("}^{");
            for n in superscript.iter() {
                convert_node_internal(converter, n)?;
            }
            converter.buffer.push('}');
        },
        MathNode::PreSub {
            base,
            pre_subscript,
        } => {
            converter.buffer.push_str("\\presub{");
            for n in base.iter() {
                convert_node_internal(converter, n)?;
            }
            converter.buffer.push_str("}{");
            for n in pre_subscript.iter() {
                convert_node_internal(converter, n)?;
            }
            converter.buffer.push('}');
        },
        MathNode::PreSup {
            base,
            pre_superscript,
        } => {
            converter.buffer.push_str("\\presup{");
            for n in base.iter() {
                convert_node_internal(converter, n)?;
            }
            converter.buffer.push_str("}{");
            for n in pre_superscript.iter() {
                convert_node_internal(converter, n)?;
            }
            converter.buffer.push('}');
        },
        MathNode::PreSubSup {
            base,
            pre_subscript,
            pre_superscript,
        } => {
            converter.buffer.push_str("\\presubsup{");
            for n in base.iter() {
                convert_node_internal(converter, n)?;
            }
            converter.buffer.push_str("}{");
            for n in pre_subscript.iter() {
                convert_node_internal(converter, n)?;
            }
            converter.buffer.push_str("}{");
            for n in pre_superscript.iter() {
                convert_node_internal(converter, n)?;
            }
            converter.buffer.push('}');
        },
        MathNode::Under {
            base,
            under,
            position: _,
        } => {
            converter.append_cached_command("\\underset{");
            for n in under.iter() {
                convert_node_internal(converter, n)?;
            }
            converter.buffer.push_str("}{");
            for n in base.iter() {
                convert_node_internal(converter, n)?;
            }
            converter.buffer.push('}');
        },
        MathNode::Over {
            base,
            over,
            position: _,
        } => {
            converter.append_cached_command("\\overset{");
            for n in over.iter() {
                convert_node_internal(converter, n)?;
            }
            converter.buffer.push_str("}{");
            for n in base.iter() {
                convert_node_internal(converter, n)?;
            }
            converter.buffer.push('}');
        },
        MathNode::UnderOver {
            base,
            under,
            over,
            position: _,
        } => {
            converter.buffer.push_str("\\overset{");
            for n in over.iter() {
                convert_node_internal(converter, n)?;
            }
            converter.buffer.push_str("}{\\underset{");
            for n in under.iter() {
                convert_node_internal(converter, n)?;
            }
            converter.buffer.push_str("}{");
            for n in base.iter() {
                convert_node_internal(converter, n)?;
            }
            converter.buffer.push_str("}}");
        },
        MathNode::Fenced {
            open,
            content,
            close,
            separator: _,
        } => {
            converter.buffer.push_str(fence_to_latex(*open, true));
            for n in content.iter() {
                convert_node_internal(converter, n)?;
            }
            converter.buffer.push_str(fence_to_latex(*close, false));
        },
        MathNode::LargeOp {
            operator,
            lower_limit,
            upper_limit,
            integrand,
            hide_lower: _,
            hide_upper: _,
        } => {
            converter
                .buffer
                .push_str(large_operator_to_latex(*operator));

            if let Some(lower) = lower_limit {
                converter.buffer.push_str("_{");
                for n in lower.iter() {
                    convert_node_internal(converter, n)?;
                }
                converter.buffer.push('}');
            }

            if let Some(upper) = upper_limit {
                converter.buffer.push_str("^{");
                for n in upper.iter() {
                    convert_node_internal(converter, n)?;
                }
                converter.buffer.push('}');
            }

            if let Some(expr) = integrand {
                converter.buffer.push(' ');
                for n in expr.iter() {
                    convert_node_internal(converter, n)?;
                }
            }
        },
        MathNode::Function { name, argument } => {
            if is_standard_function(name) {
                write!(&mut converter.buffer, "\\{}", name)
                    .map_err(|e| LatexError::FormatError(e.to_string()))?;
            } else {
                write!(&mut converter.buffer, "\\operatorname{{{}}}", name)
                    .map_err(|e| LatexError::FormatError(e.to_string()))?;
            }
            converter.buffer.push('{');
            for n in argument.iter() {
                convert_node_internal(converter, n)?;
            }
            converter.buffer.push('}');
        },
        MathNode::PredefinedFunction { function, argument } => {
            converter.buffer.push_str(function_name_to_latex(*function));
            converter.buffer.push('{');
            for n in argument.iter() {
                convert_node_internal(converter, n)?;
            }
            converter.buffer.push('}');
        },
        MathNode::Matrix {
            rows,
            fence_type,
            properties,
        } => {
            super::matrix::convert_matrix_optimized_internal(
                converter,
                rows,
                *fence_type,
                properties.as_ref(),
            )?;
        },
        MathNode::EqArray {
            rows,
            properties: _,
        } => {
            converter.buffer.push_str("\\begin{align*}");
            for (i, row) in rows.iter().enumerate() {
                if i > 0 {
                    converter.buffer.push_str("\\\\");
                }
                for n in row.iter() {
                    convert_node_internal(converter, n)?;
                }
            }
            converter.buffer.push_str("\\end{align*}");
        },
        MathNode::Accent {
            base,
            accent,
            position: _,
        } => {
            converter.buffer.push_str(accent_to_latex(*accent));
            converter.buffer.push('{');
            for n in base.iter() {
                convert_node_internal(converter, n)?;
            }
            converter.buffer.push('}');
        },
        MathNode::Bar { base, position: _ } => {
            converter.buffer.push_str("\\bar{");
            for n in base.iter() {
                convert_node_internal(converter, n)?;
            }
            converter.buffer.push('}');
        },
        MathNode::BorderBox { content, style: _ } => {
            converter.buffer.push_str("\\boxed{");
            for n in content.iter() {
                convert_node_internal(converter, n)?;
            }
            converter.buffer.push('}');
        },
        MathNode::GroupChar {
            base,
            character,
            position,
            vertical_alignment,
        } => {
            let cmd = match (position, vertical_alignment) {
                (Some(Position::Top), _) => "\\overbrace",
                (Some(Position::Bottom), _) => "\\underbrace",
                (_, Some(VerticalAlignment::Top)) => "\\overbrace",
                (_, Some(VerticalAlignment::Bottom)) => "\\underbrace",
                _ => "\\overbrace",
            };
            converter.buffer.push_str(cmd);
            converter.buffer.push('{');
            for n in base.iter() {
                convert_node_internal(converter, n)?;
            }
            converter.buffer.push('}');
            if let Some(char) = character {
                converter.buffer.push_str("^{");
                converter.buffer.push_str(char);
                converter.buffer.push('}');
            }
        },
        MathNode::Space(space_type) => {
            converter.buffer.push_str(space_to_latex(*space_type));
        },
        MathNode::LineBreak => {
            converter.buffer.push_str("\\\\");
        },
        MathNode::Style { style, content } => {
            converter.buffer.push_str(style_to_latex(*style));
            converter.buffer.push('{');
            for n in content.iter() {
                convert_node_internal(converter, n)?;
            }
            converter.buffer.push('}');
        },
        MathNode::Row(nodes) => {
            for n in nodes {
                convert_node_internal(converter, n)?;
            }
        },
        MathNode::Phantom(content) => {
            converter.buffer.push_str("\\phantom{");
            for n in content.iter() {
                convert_node_internal(converter, n)?;
            }
            converter.buffer.push('}');
        },
        MathNode::Error(msg) => {
            write!(&mut converter.buffer, "\\text{{[Error: {}]}}", msg)
                .map_err(|e| LatexError::FormatError(e.to_string()))?;
        },
        MathNode::Run {
            content,
            literal: _,
            style,
            font,
            color,
            underline,
            overline,
            strike_through,
            double_strike_through,
        } => {
            if let Some(s) = style {
                converter.buffer.push_str(style_to_latex(*s));
                converter.buffer.push('{');
            }
            if let Some(f) = font {
                converter.buffer.push_str("\\fontfamily{");
                converter.buffer.push_str(f);
                converter.buffer.push_str("}\\selectfont{");
            }
            if let Some(c) = color {
                converter.buffer.push_str("\\color{");
                converter.buffer.push_str(c);
                converter.buffer.push_str("}{");
            }
            if underline.is_some() {
                converter.buffer.push_str("\\underline{");
            }
            if overline.is_some() {
                converter.buffer.push_str("\\overline{");
            }
            if strike_through.is_some() || double_strike_through.is_some() {
                converter.buffer.push_str("\\sout{");
            }

            for n in content.iter() {
                convert_node_internal(converter, n)?;
            }

            if strike_through.is_some() || double_strike_through.is_some() {
                converter.buffer.push('}');
            }
            if overline.is_some() {
                converter.buffer.push('}');
            }
            if underline.is_some() {
                converter.buffer.push('}');
            }
            if color.is_some() {
                converter.buffer.push('}');
            }
            if font.is_some() {
                converter.buffer.push('}');
            }
            if style.is_some() {
                converter.buffer.push('}');
            }
        },
        MathNode::Limit {
            content,
            limit_type,
        } => {
            let cmd = match limit_type {
                LimitType::Lower => "\\lim_{",
                LimitType::Upper => "\\lim^{",
            };
            super::utils::extend_buffer_with_capacity(&mut converter.buffer, cmd, 1);
            for n in content.iter() {
                convert_node_internal(converter, n)?;
            }
            converter.buffer.push('}');
        },
        MathNode::Degree(content)
        | MathNode::Base(content)
        | MathNode::Argument(content)
        | MathNode::Numerator(content)
        | MathNode::Denominator(content)
        | MathNode::Integrand(content)
        | MathNode::LowerLimit(content)
        | MathNode::UpperLimit(content) => {
            for n in content.iter() {
                convert_node_internal(converter, n)?;
            }
        },
    }

    Ok(())
}
