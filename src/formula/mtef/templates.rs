//! MTEF template parsing based on rtf2latex2e template system
//!
//! This module provides template definitions and parsing logic for MTEF templates.
//! Templates represent structured mathematical constructs like fractions, roots,
//! integrals, fences, etc.
//!
//! Based on rtf2latex2e Profile_TEMPLATES_5 template system.

use crate::formula::ast::{Fence, LargeOperator, MathNode, Operator};
use smallvec::SmallVec;

/// Template argument list type - a small vector of node vectors
pub type TemplateArgs<'a> = SmallVec<[SmallVec<[MathNode<'a>; 8]>; 4]>;

/// Template parser helper methods - based on rtf2latex2e Profile_TEMPLATES_5
pub struct TemplateParser;

/// Template definition structure
///
/// Defines a specific MTEF template with its selector, variation, and LaTeX template string.
/// The description field is used for documentation and debugging purposes.
#[derive(Debug)]
pub struct TemplateDef {
    /// Template selector (identifies template type)
    pub selector: u8,
    /// Template variation (specific form within type)
    pub variation: u16,
    /// Human-readable description (used for documentation and debugging)
    #[allow(dead_code)]
    pub description: &'static str,
    /// LaTeX template string with argument placeholders
    pub template: &'static str,
}

/// MTEF v5 Template definitions based on rtf2latex2e Profile_TEMPLATES_5
const MTEF_TEMPLATES: &[TemplateDef] = &[
    TemplateDef {
        selector: 0,
        variation: 1,
        description: "fence: angle-left only",
        template: "\\left\\langle #1[M]\\right.  ",
    },
    TemplateDef {
        selector: 0,
        variation: 2,
        description: "fence: angle-right only",
        template: "\\left. #1[M]\\right\\rangle  ",
    },
    TemplateDef {
        selector: 0,
        variation: 3,
        description: "fence: angle-both",
        template: "\\left\\langle #1[M]\\right\\rangle  ",
    },
    TemplateDef {
        selector: 1,
        variation: 1,
        description: "fence: paren-left only",
        template: "\\left( #1[M]\\right.  ",
    },
    TemplateDef {
        selector: 1,
        variation: 2,
        description: "fence: paren-right only",
        template: "\\left. #1[M]\\right)  ",
    },
    TemplateDef {
        selector: 1,
        variation: 3,
        description: "fence: paren-both",
        template: "\\left( #1[M]\\right)  ",
    },
    TemplateDef {
        selector: 2,
        variation: 1,
        description: "fence: brace-left only",
        template: "\\left\\{ #1[M]\\right.  ",
    },
    TemplateDef {
        selector: 2,
        variation: 2,
        description: "fence: brace-right only",
        template: "\\left. #1[M]\\right\\}  ",
    },
    TemplateDef {
        selector: 2,
        variation: 3,
        description: "fence: brace-both",
        template: "\\left\\{ #1[M]\\right\\}  ",
    },
    TemplateDef {
        selector: 3,
        variation: 1,
        description: "fence: brack-left only",
        template: "\\lef]t[ #1[M]\\right.  ",
    },
    TemplateDef {
        selector: 3,
        variation: 2,
        description: "fence: brack-right only",
        template: "\\left. #1[M]\\right]  ",
    },
    TemplateDef {
        selector: 3,
        variation: 3,
        description: "fence: brack-both",
        template: "\\left[ #1[M]\\right]  ",
    },
    TemplateDef {
        selector: 4,
        variation: 1,
        description: "fence: bar-left only",
        template: "\\left| #1[M]\\right.  ",
    },
    TemplateDef {
        selector: 4,
        variation: 2,
        description: "fence: bar-right only",
        template: "\\left. #1[M]\\right|  ",
    },
    TemplateDef {
        selector: 4,
        variation: 3,
        description: "fence: bar-both",
        template: "\\left| #1[M]\\right|  ",
    },
    TemplateDef {
        selector: 5,
        variation: 1,
        description: "fence: dbar-left only",
        template: "\\left\\| #1[M]\\right.  ",
    },
    TemplateDef {
        selector: 5,
        variation: 2,
        description: "fence: dbar-right only",
        template: "\\left. #1[M]\\right\\|  ",
    },
    TemplateDef {
        selector: 5,
        variation: 3,
        description: "fence: dbar-both",
        template: "\\left\\| #1[M]\\right\\|  ",
    },
    TemplateDef {
        selector: 6,
        variation: 1,
        description: "fence: floor",
        template: "\\left\\lfloor #1[M]\\right.  ",
    },
    TemplateDef {
        selector: 6,
        variation: 2,
        description: "fence: floor",
        template: "\\left. #1[M]\\right\\rfloor  ",
    },
    TemplateDef {
        selector: 6,
        variation: 3,
        description: "fence: floor",
        template: "\\left\\lfloor #1[M]\\right\\rfloor  ",
    },
    TemplateDef {
        selector: 7,
        variation: 1,
        description: "fence: ceiling",
        template: "\\left\\lceil #1[M]\\right.  ",
    },
    TemplateDef {
        selector: 7,
        variation: 2,
        description: "fence: ceiling",
        template: "\\left. #1[M]\\right\\rceil  ",
    },
    TemplateDef {
        selector: 7,
        variation: 3,
        description: "fence: ceiling",
        template: "\\left\\lceil #1[M]\\right\\rceil  ",
    },
    TemplateDef {
        selector: 8,
        variation: 0,
        description: "fence: LBLB",
        template: "\\left[ #1[M]\\right[  ",
    },
    TemplateDef {
        selector: 9,
        variation: 0,
        description: "fence: LPLP",
        template: "\\left( #1[M]\\right(  ",
    },
    TemplateDef {
        selector: 9,
        variation: 1,
        description: "fence: RPLP",
        template: "\\left) #1[M]\\right(  ",
    },
    TemplateDef {
        selector: 9,
        variation: 2,
        description: "fence: LBLP",
        template: "\\left[ #1[M]\\right(  ",
    },
    TemplateDef {
        selector: 9,
        variation: 3,
        description: "fence: RBLP",
        template: "\\left] #1[M]\\right(  ",
    },
    TemplateDef {
        selector: 9,
        variation: 16,
        description: "fence: LPRP",
        template: "\\left( #1[M]\\right)  ",
    },
    TemplateDef {
        selector: 9,
        variation: 17,
        description: "fence: RPRP",
        template: "\\left) #1[M]\\right)  ",
    },
    TemplateDef {
        selector: 9,
        variation: 18,
        description: "fence: LBRP",
        template: "\\left[ #1[M]\\right)  ",
    },
    TemplateDef {
        selector: 9,
        variation: 19,
        description: "fence: RBRP",
        template: "\\left] #1[M]\\right)  ",
    },
    TemplateDef {
        selector: 9,
        variation: 32,
        description: "fence: LPLB",
        template: "\\left( #1[M]\\right[  ",
    },
    TemplateDef {
        selector: 9,
        variation: 33,
        description: "fence: RPLB",
        template: "\\left) #1[M]\\right[  ",
    },
    TemplateDef {
        selector: 9,
        variation: 34,
        description: "fence: LBLB",
        template: "\\left[ #1[M]\\right[  ",
    },
    TemplateDef {
        selector: 9,
        variation: 35,
        description: "fence: RBLB",
        template: "\\left] #1[M]\\right[  ",
    },
    TemplateDef {
        selector: 9,
        variation: 48,
        description: "fence: LPRB",
        template: "\\left( #1[M]\\right]  ",
    },
    TemplateDef {
        selector: 9,
        variation: 49,
        description: "fence: RPRB",
        template: "\\left) #1[M]\\right]  ",
    },
    TemplateDef {
        selector: 9,
        variation: 50,
        description: "fence: LBRB",
        template: "\\left[ #1[M]\\right]  ",
    },
    TemplateDef {
        selector: 9,
        variation: 51,
        description: "fence: RBRB",
        template: "\\left] #1[M]\\right]  ",
    },
    TemplateDef {
        selector: 10,
        variation: 0,
        description: "root: sqroot",
        template: "\\sqrt{#1[M]}  ",
    },
    TemplateDef {
        selector: 10,
        variation: 1,
        description: "root: nthroot",
        template: "\\sqrt[#2[M]]{#1[M]}  ",
    },
    TemplateDef {
        selector: 11,
        variation: 0,
        description: "fract: tmfract",
        template: "\\frac{#1[M]}{#2[M]}  ",
    },
    TemplateDef {
        selector: 11,
        variation: 1,
        description: "fract: smfract",
        template: "\\frac{#1[M]}{#2[M]}  ",
    },
    TemplateDef {
        selector: 11,
        variation: 2,
        description: "fract: slfract",
        template: "{#1[M]}/{#2[M]}  ",
    },
    TemplateDef {
        selector: 11,
        variation: 3,
        description: "fract: slfract",
        template: "{#1[M]}/{#2[M]}  ",
    },
    TemplateDef {
        selector: 11,
        variation: 4,
        description: "fract: slfract",
        template: "{#1[M]}/{#2[M]}  ",
    },
    TemplateDef {
        selector: 11,
        variation: 5,
        description: "fract: smfract",
        template: "\\frac{#1[M]}{#2[M]}  ",
    },
    TemplateDef {
        selector: 11,
        variation: 6,
        description: "fract: slfract",
        template: "{#1[M]}/{#2[M]}  ",
    },
    TemplateDef {
        selector: 11,
        variation: 7,
        description: "fract: slfract",
        template: "{#1[M]}/{#2[M]}  ",
    },
    TemplateDef {
        selector: 12,
        variation: 0,
        description: "ubar: subar",
        template: "\\underline{#1[M]}  ",
    },
    TemplateDef {
        selector: 12,
        variation: 1,
        description: "ubar: dubar",
        template: "\\underline{\\underline{#1[M]}}  ",
    },
    TemplateDef {
        selector: 13,
        variation: 0,
        description: "obar: sobar",
        template: "\\overline{#1[M]}  ",
    },
    TemplateDef {
        selector: 13,
        variation: 1,
        description: "obar: dobar",
        template: "\\overline{\\overline{#1[M]}}  ",
    },
    TemplateDef {
        selector: 14,
        variation: 0,
        description: "larrow: box on top",
        template: "\\stackrel{#1[M]}{\\longleftarrow}  ",
    },
    TemplateDef {
        selector: 14,
        variation: 1,
        description: "larrow: box below ",
        template: "\\stackunder{#1[M]}{\\longleftarrow}  ",
    },
    TemplateDef {
        selector: 14,
        variation: 0,
        description: "rarrow: box on top",
        template: "\\stackrel{#1[M]}{\\longrightarrow}  ",
    },
    TemplateDef {
        selector: 14,
        variation: 1,
        description: "rarrow: box below ",
        template: "\\stackunder{#1[M]}{\\longrightarrow}  ",
    },
    TemplateDef {
        selector: 14,
        variation: 0,
        description: "barrow: box on top",
        template: "\\stackrel{#1[M]}{\\longleftrightarrow}  ",
    },
    TemplateDef {
        selector: 14,
        variation: 1,
        description: "barrow: box below ",
        template: "\\stackunder{#1[M]}{\\longleftrightarrow}  ",
    },
    TemplateDef {
        selector: 15,
        variation: 0,
        description: "integrals: single - no limits",
        template: "\\int #1[M]  ",
    },
    TemplateDef {
        selector: 15,
        variation: 1,
        description: "integrals: single - both",
        template: "\\int\\nolimits#2[L][STARTSUB][ENDSUB]#3[L][STARTSUP][ENDSUP]#1[M]  ",
    },
    TemplateDef {
        selector: 15,
        variation: 2,
        description: "integrals: double - both",
        template: "\\iint\\nolimits#2[L][STARTSUB][ENDSUB]#3[L][STARTSUP][ENDSUP]#1[M]  ",
    },
    TemplateDef {
        selector: 15,
        variation: 3,
        description: "integrals: triple - both",
        template: "\\iiint\\nolimits#2[L][STARTSUB][ENDSUB]#3[L][STARTSUP][ENDSUP]#1[M]  ",
    },
    TemplateDef {
        selector: 15,
        variation: 4,
        description: "integrals: contour - no limits",
        template: "\\oint #1[M]  ",
    },
    TemplateDef {
        selector: 15,
        variation: 8,
        description: "integrals: contour - no limits",
        template: "\\oint #1[M]  ",
    },
    TemplateDef {
        selector: 15,
        variation: 12,
        description: "integrals: contour - no limits",
        template: "\\oint #1[M]  ",
    },
    TemplateDef {
        selector: 16,
        variation: 0,
        description: "sum: limits top/bottom - both",
        template: "\\sum\\limits#2[L][STARTSUB][ENDSUB]#3[L][STARTSUP][ENDSUP]#1[M]  ",
    },
    TemplateDef {
        selector: 17,
        variation: 0,
        description: "product: limits top/bottom - both",
        template: "\\prod\\limits#2[L][STARTSUB][ENDSUB]#3[L][STARTSUP][ENDSUP]#1[M]  ",
    },
    TemplateDef {
        selector: 18,
        variation: 0,
        description: "coproduct: limits top/bottom - both",
        template: "\\dcoprod\\limits#2[L][STARTSUB][ENDSUB]#3[L][STARTSUP][ENDSUP]#1[M]  ",
    },
    TemplateDef {
        selector: 19,
        variation: 0,
        description: "union: limits top/bottom - both",
        template: "\\dbigcup\\limits#2[L][STARTSUB][ENDSUB]#3[L][STARTSUP][ENDSUP]#1[M]  ",
    },
    TemplateDef {
        selector: 20,
        variation: 0,
        description: "intersection: limits top/bottom - both",
        template: "\\dbigcap\\limits#2[L][STARTSUB][ENDSUB]#3[L][STARTSUP][ENDSUP]#1[M]  ",
    },
    TemplateDef {
        selector: 21,
        variation: 0,
        description: "integrals: single - both",
        template: "\\int#2[L][STARTSUB][ENDSUB]#3[L][STARTSUP][ENDSUP]#1[M]  ",
    },
    TemplateDef {
        selector: 22,
        variation: 0,
        description: "sum: single - both",
        template: "\\sum#2[L][STARTSUB][ENDSUB]#3[L][STARTSUP][ENDSUP]#1[M]  ",
    },
    TemplateDef {
        selector: 23,
        variation: 0,
        description: "limit: both",
        template: "#1 #2[L][STARTSUB][ENDSUB]#3[L][STARTSUP][ENDSUP]  ",
    },
    TemplateDef {
        selector: 24,
        variation: 0,
        description: "horizontal brace: lower",
        template: "\\stackunder{#2[M]}{\\underbrace{#1[M]}}  ",
    },
    TemplateDef {
        selector: 24,
        variation: 1,
        description: "horizontal brace: upper",
        template: "\\stackrel{#2[M]}{\\overbrace{#1[M]}}  ",
    },
    TemplateDef {
        selector: 25,
        variation: 0,
        description: "horizontal brace: lower",
        template: "\\stackunder{#2[M]}{\\underbrace{#1[M]}}  ",
    },
    TemplateDef {
        selector: 25,
        variation: 1,
        description: "horizontal brace: upper",
        template: "\\stackrel{#2[M]}{\\overbrace{#1[M]}}  ",
    },
    TemplateDef {
        selector: 25,
        variation: 0,
        description: "hbracket",
        template: " ",
    },
    TemplateDef {
        selector: 27,
        variation: 0,
        description: "script: sub",
        template: "#1[L][STARTSUB][ENDSUB]  ",
    },
    TemplateDef {
        selector: 27,
        variation: 1,
        description: "script: sub",
        template: "#1[L][STARTSUB][ENDSUB]  ",
    },
    TemplateDef {
        selector: 28,
        variation: 0,
        description: "script: super",
        template: "#2[L][STARTSUP][ENDSUP]  ",
    },
    TemplateDef {
        selector: 28,
        variation: 1,
        description: "script: super",
        template: "#2[L][STARTSUP][ENDSUP]  ",
    },
    TemplateDef {
        selector: 29,
        variation: 0,
        description: "script: subsup",
        template: "#1[L][STARTSUB][ENDSUB]#2[L][STARTSUP][ENDSUP]  ",
    },
];

impl TemplateParser {
    /// Find template by selector and variation
    pub fn find_template(selector: u8, variation: u16) -> Option<&'static TemplateDef> {
        MTEF_TEMPLATES
            .iter()
            .find(|t| t.selector == selector && t.variation == variation)
    }

    /// Parse template arguments and apply formatting
    ///
    /// Parses the template string using rtf2latex2e template format and substitutes
    /// the provided arguments. Template format uses LaTeX commands with placeholders
    /// like #1[M], #2[L], etc. and special markers [STARTSUB], [ENDSUB], etc.
    pub fn parse_template_arguments<'a>(template: &str, args: &TemplateArgs<'a>) -> MathNode<'a> {
        let mut result = String::new();
        let mut chars = template.chars().peekable();

        while let Some(ch) = chars.next() {
            if ch == '#' {
                // Parse argument placeholder like #1[M] or #2[L]
                if let Some(digit) = chars.next().and_then(|c| c.to_digit(10)) {
                    let arg_index = digit as usize - 1; // Convert to 0-based index

                    // Skip the mode specifier in brackets, e.g., [M] or [L]
                    if chars.next() == Some('[') {
                        for c in chars.by_ref() {
                            if c == ']' {
                                break;
                            }
                        }
                    }

                    // Substitute the argument
                    if arg_index < args.len() {
                        // Convert argument nodes to text for simple substitution
                        let mut arg_text = String::new();
                        for node in &args[arg_index] {
                            match node {
                                MathNode::Text(text) => arg_text.push_str(text),
                                MathNode::Number(num) => arg_text.push_str(num),
                                MathNode::Symbol(sym) => {
                                    if let Some(unicode) = sym.unicode {
                                        arg_text.push(unicode);
                                    } else {
                                        arg_text.push_str(&sym.name);
                                    }
                                },
                                _ => arg_text.push('?'), // Placeholder for complex nodes
                            }
                        }
                        result.push_str(&arg_text);
                    }
                } else {
                    result.push('#');
                }
            } else if ch == '[' {
                // Handle special markers like [STARTSUB], [ENDSUB], etc.
                let mut marker = String::new();
                for c in chars.by_ref() {
                    if c == ']' {
                        break;
                    }
                    marker.push(c);
                }

                match marker.as_str() {
                    "STARTSUB" => result.push_str("_{"),
                    "ENDSUB" => result.push('}'),
                    "STARTSUP" => result.push_str("^{"),
                    "ENDSUP" => result.push('}'),
                    _ => {
                        // Unknown marker, keep as is
                        result.push('[');
                        result.push_str(&marker);
                        result.push(']');
                    },
                }
            } else {
                result.push(ch);
            }
        }

        // Try to recognize common LaTeX patterns and convert to AST nodes
        Self::parse_latex_to_ast(&result, args)
    }

    /// Parse LaTeX string back to AST nodes for common patterns
    ///
    /// Recognizes common LaTeX constructs (fractions, roots, operators, etc.)
    /// and converts them back to proper AST nodes instead of plain text.
    fn parse_latex_to_ast<'a>(latex: &str, args: &TemplateArgs<'a>) -> MathNode<'a> {
        let latex = latex.trim();

        // Fraction: \frac{numerator}{denominator}
        if latex.starts_with("\\frac{")
            && latex.contains("}{")
            && latex[latex.find("}{").unwrap() + 2..].find('}').is_some()
        {
            // Try to find the actual nodes from args
            let mut numerator = Vec::new();
            let mut denominator = Vec::new();

            // Simple heuristic: first arg is numerator, second is denominator
            if args.len() >= 2 {
                numerator = args[0].iter().cloned().collect();
                denominator = args[1].iter().cloned().collect();
            }

            return MathNode::Frac {
                numerator,
                denominator,
                line_thickness: None,
                frac_type: None,
            };
        }

        // Root: \sqrt[index]{base} or \sqrt{base}
        if latex.starts_with("\\sqrt") {
            if latex.starts_with("\\sqrt[") {
                if let Some(rel_pos) = latex.strip_prefix("\\sqrt[").and_then(|s| s.find("]{")) {
                    let abs_pos = 6 + rel_pos;
                    if latex[abs_pos + 2..].find('}').is_some() {
                        let mut base = Vec::new();
                        let mut index = Vec::new();

                        if !args.is_empty() {
                            base = args[0].iter().cloned().collect();
                        }
                        if args.len() >= 2 {
                            index = args[1].iter().cloned().collect();
                        }

                        return MathNode::Root {
                            base,
                            index: Some(index),
                        };
                    }
                }
            } else if latex.starts_with("\\sqrt{") && latex[6..].find('}').is_some() {
                let mut base = Vec::new();
                if !args.is_empty() {
                    base = args[0].iter().cloned().collect();
                }

                return MathNode::Root { base, index: None };
            }
        }

        // Large operators with limits
        if latex.contains("\\sum") || latex.contains("\\prod") || latex.contains("\\int") {
            let operator = if latex.contains("\\sum") {
                LargeOperator::Sum
            } else if latex.contains("\\prod") {
                LargeOperator::Product
            } else {
                LargeOperator::Integral
            };

            let mut lower_limit = None;
            let mut upper_limit = None;
            let mut integrand = None;

            // Extract limits from _{...}^{...} patterns
            if let Some(sub_start) = latex.find("_{")
                && latex[sub_start + 2..].find('}').is_some()
                && args.len() >= 2
            {
                lower_limit = Some(args[1].iter().cloned().collect());
            }

            if let Some(sup_start) = latex.find("^{")
                && latex[sup_start + 2..].find('}').is_some()
                && args.len() >= 3
            {
                upper_limit = Some(args[2].iter().cloned().collect());
            }

            if !args.is_empty() {
                integrand = Some(args[0].iter().cloned().collect());
            }

            return MathNode::LargeOp {
                operator,
                lower_limit,
                upper_limit,
                integrand,
                hide_lower: false,
                hide_upper: false,
            };
        }

        // Subscripts and superscripts
        if latex.contains("_{") && latex.contains("^{") {
            // Both sub and superscript
            let mut base = Vec::new();
            let mut subscript = Vec::new();
            let mut superscript = Vec::new();

            if !args.is_empty() {
                base = args[0].iter().cloned().collect();
            }
            if args.len() >= 2 {
                subscript = args[1].iter().cloned().collect();
            }
            if args.len() >= 3 {
                superscript = args[2].iter().cloned().collect();
            }

            return MathNode::SubSup {
                base,
                subscript,
                superscript,
            };
        } else if latex.contains("_{") {
            // Subscript only
            let mut base = Vec::new();
            let mut subscript = Vec::new();

            if !args.is_empty() {
                base = args[0].iter().cloned().collect();
            }
            if args.len() >= 2 {
                subscript = args[1].iter().cloned().collect();
            }

            return MathNode::Sub { base, subscript };
        } else if latex.contains("^{") {
            // Superscript only
            let mut base = Vec::new();
            let mut exponent = Vec::new();

            if !args.is_empty() {
                base = args[0].iter().cloned().collect();
            }
            if args.len() >= 2 {
                exponent = args[1].iter().cloned().collect();
            }

            return MathNode::Power { base, exponent };
        }

        // Fences: \left...\right...
        if latex.contains("\\left") && latex.contains("\\right") {
            // Extract content between \left and \right
            if let Some(left_pos) = latex.find("\\left")
                && let Some(right_pos) = latex.find("\\right")
            {
                let content_start = latex[left_pos..]
                    .find('{')
                    .map(|p| left_pos + p + 1)
                    .unwrap_or(left_pos + 6);
                let content_end = right_pos;

                if content_start < content_end {
                    // Determine fence type
                    let open_fence = if latex.contains("\\left(") {
                        Fence::Paren
                    } else if latex.contains("\\left[") {
                        Fence::Bracket
                    } else if latex.contains("\\left{") {
                        Fence::Brace
                    } else if latex.contains("\\left|") {
                        Fence::Pipe
                    } else {
                        Fence::Paren // default
                    };

                    let close_fence = if latex.contains("\\right)") {
                        Fence::Paren
                    } else if latex.contains("\\right]") {
                        Fence::Bracket
                    } else if latex.contains("\\right}") {
                        Fence::Brace
                    } else if latex.contains("\\right|") {
                        Fence::Pipe
                    } else {
                        Fence::Paren // default
                    };

                    let mut content = Vec::new();
                    if !args.is_empty() {
                        content = args[0].iter().cloned().collect();
                    }

                    return MathNode::Fenced {
                        open: open_fence,
                        content,
                        close: close_fence,
                        separator: None,
                    };
                }
            }
        }

        // Default: return as text
        MathNode::Text(latex.to_string().into())
    }
    /// Parse a fraction template
    pub fn parse_fraction<'a>(
        numerator: Vec<MathNode<'a>>,
        denominator: Vec<MathNode<'a>>,
    ) -> MathNode<'a> {
        MathNode::Frac {
            numerator,
            denominator,
            line_thickness: None,
            frac_type: None,
        }
    }

    /// Parse a slash template (inline fraction)
    ///
    /// Public API for potential external use or future MTEF features
    #[allow(dead_code)] // Part of public template parsing API
    pub fn parse_slash<'a>(
        numerator: Vec<MathNode<'a>>,
        denominator: Vec<MathNode<'a>>,
    ) -> MathNode<'a> {
        MathNode::Row(vec![
            MathNode::Row(numerator),
            MathNode::Operator(Operator::Divide),
            MathNode::Row(denominator),
        ])
    }

    /// Parse a root template
    pub fn parse_root<'a>(
        base: Vec<MathNode<'a>>,
        index: Option<Vec<MathNode<'a>>>,
    ) -> MathNode<'a> {
        MathNode::Root { base, index }
    }

    /// Parse a subscript template
    pub fn parse_subscript<'a>(
        base: Vec<MathNode<'a>>,
        subscript: Vec<MathNode<'a>>,
    ) -> MathNode<'a> {
        MathNode::Sub { base, subscript }
    }

    /// Parse a superscript template
    pub fn parse_superscript<'a>(
        base: Vec<MathNode<'a>>,
        superscript: Vec<MathNode<'a>>,
    ) -> MathNode<'a> {
        MathNode::Power {
            base,
            exponent: superscript,
        }
    }

    /// Parse a subscript-superscript template
    pub fn parse_subsup<'a>(
        base: Vec<MathNode<'a>>,
        subscript: Vec<MathNode<'a>>,
        superscript: Vec<MathNode<'a>>,
    ) -> MathNode<'a> {
        MathNode::SubSup {
            base,
            subscript,
            superscript,
        }
    }

    /// Parse an underscript template
    ///
    /// Public API for potential external use or future MTEF features
    #[allow(dead_code)] // Part of public template parsing API
    pub fn parse_below<'a>(base: Vec<MathNode<'a>>, script: Vec<MathNode<'a>>) -> MathNode<'a> {
        MathNode::Under {
            base,
            under: script,
            position: None,
        }
    }

    /// Parse an overscript template
    ///
    /// Public API for potential external use or future MTEF features
    #[allow(dead_code)] // Part of public template parsing API
    pub fn parse_above<'a>(base: Vec<MathNode<'a>>, script: Vec<MathNode<'a>>) -> MathNode<'a> {
        MathNode::Over {
            base,
            over: script,
            position: None,
        }
    }

    /// Parse an underscript-overscript template
    ///
    /// Public API for potential external use or future MTEF features
    #[allow(dead_code)] // Part of public template parsing API
    pub fn parse_below_above<'a>(
        base: Vec<MathNode<'a>>,
        below: Vec<MathNode<'a>>,
        above: Vec<MathNode<'a>>,
    ) -> MathNode<'a> {
        MathNode::UnderOver {
            base,
            under: below,
            over: above,
            position: None,
        }
    }

    /// Parse a large operator template
    pub fn parse_large_op<'a>(
        operator: LargeOperator,
        lower_limit: Vec<MathNode<'a>>,
        upper_limit: Vec<MathNode<'a>>,
        integrand: Vec<MathNode<'a>>,
    ) -> MathNode<'a> {
        MathNode::LargeOp {
            operator,
            lower_limit: if lower_limit.is_empty() {
                None
            } else {
                Some(lower_limit)
            },
            upper_limit: if upper_limit.is_empty() {
                None
            } else {
                Some(upper_limit)
            },
            integrand: if integrand.is_empty() {
                None
            } else {
                Some(integrand)
            },
            hide_lower: false,
            hide_upper: false,
        }
    }

    /// Parse a fence template
    pub fn parse_fence<'a>(fence: Fence, content: Vec<MathNode<'a>>) -> MathNode<'a> {
        MathNode::Fenced {
            open: fence,
            content,
            close: fence,
            separator: None,
        }
    }

    /// Get large operator from template selector
    ///
    /// Maps MTEF template selectors to corresponding large operator types.
    /// Some selectors may map to the same operator type (e.g., multiple integral variants).
    ///
    /// Public API for template system, may be used by custom template handlers
    #[allow(dead_code)] // Part of public template mapping API
    pub fn large_op_from_selector(selector: u8) -> Option<LargeOperator> {
        match selector {
            15 => Some(LargeOperator::Integral), // TMPL_INTOP: integrals (single, double, triple, contour)
            16 => Some(LargeOperator::Sum),      // TMPL_SUM: summation
            17 => Some(LargeOperator::Product),  // TMPL_PROD: product
            18 => Some(LargeOperator::Coproduct), // TMPL_COPROD: coproduct
            19 => Some(LargeOperator::Union),    // TMPL_UNION: union
            20 => Some(LargeOperator::Intersection), // TMPL_INTER: intersection
            21 => Some(LargeOperator::Integral), // TMPL_IINTOP: single integral with limits
            22 => Some(LargeOperator::Sum),      // TMPL_IIINTOP: single sum with limits
            23 => Some(LargeOperator::Integral), // TMPL_OINTOP: contour integral / limit template
            _ => None,
        }
    }

    /// Get fence from template selector
    ///
    /// Maps MTEF template selectors to corresponding fence types.
    ///
    /// Public API for template system, may be used by custom template handlers
    #[allow(dead_code)] // Part of public template mapping API
    pub fn fence_from_selector(selector: u8) -> Option<Fence> {
        match selector {
            1 => Some(Fence::Paren),      // TMPL_PAREN: parentheses
            3 => Some(Fence::Bracket),    // TMPL_BRACKET: square brackets
            2 => Some(Fence::Brace),      // TMPL_BRACE: curly braces
            4 => Some(Fence::Pipe),       // TMPL_BAR: vertical bars
            5 => Some(Fence::DoublePipe), // TMPL_DBAR: double vertical bars
            _ => None,
        }
    }
}
