// OMML (Office Math Markup Language) Parser
//
// This module parses Microsoft Office Math Markup Language (OMML) into our AST.
// OMML is used in modern Office documents (.docx, .pptx, etc.) to represent
// mathematical formulas.
//
// This implementation provides comprehensive OMML parsing with:
// - High-performance streaming XML parsing
// - Modular element handlers for different OMML constructs
// - Comprehensive attribute parsing
// - Memory-efficient arena-based allocation
// - Support for all OMML elements and properties
//
// Reference: https://devblogs.microsoft.com/math-in-office/officemath/

mod elements;
mod attributes;
mod handlers;
mod properties;
mod utils;
mod lookup;
mod parser;
mod error;
mod context;

use crate::formula::ast::MathNode;


/// Re-export public API
pub use parser::OmmlParser;
pub use error::OmmlError;

#[cfg(test)]
mod tests {
    use super::*;
    use crate::formula::ast::{Formula, Fence, AccentType, LargeOperator};

    #[test]
    fn test_parse_simple_text() {
        let formula = Formula::new();
        let parser = OmmlParser::new(formula.arena());

        let xml = r#"<m:oMath><m:r><m:t>x</m:t></m:r></m:oMath>"#;
        let nodes = parser.parse(xml).unwrap();

        assert_eq!(nodes.len(), 1);
        match &nodes[0] {
            MathNode::Text(text) => assert_eq!(text.as_ref(), "x"),
            _ => panic!("Expected text node"),
        }
    }

    #[test]
    fn test_parse_multiple_text_runs() {
        let formula = Formula::new();
        let parser = OmmlParser::new(formula.arena());

        let xml = r#"<m:oMath>
            <m:r><m:t>a</m:t></m:r>
            <m:r><m:t>b</m:t></m:r>
            <m:r><m:t>c</m:t></m:r>
        </m:oMath>"#;
        let nodes = parser.parse(xml).unwrap();

        assert_eq!(nodes.len(), 3);
        for (i, node) in nodes.iter().enumerate() {
            match node {
                MathNode::Text(text) => {
                    let expected = match i {
                        0 => "a",
                        1 => "b",
                        2 => "c",
                        _ => unreachable!(),
                    };
                    assert_eq!(text.as_ref(), expected);
                }
                _ => panic!("Expected text node at position {}", i),
            }
        }
    }

    #[test]
    fn test_parse_fraction() {
        let formula = Formula::new();
        let parser = OmmlParser::new(formula.arena());

        let xml = r#"<m:oMath>
            <m:f>
                <m:num><m:r><m:t>1</m:t></m:r></m:num>
                <m:den><m:r><m:t>2</m:t></m:r></m:den>
            </m:f>
        </m:oMath>"#;

        let nodes = parser.parse(xml).unwrap();
        assert!(!nodes.is_empty());
        match &nodes[0] {
            MathNode::Frac { numerator, denominator, .. } => {
                assert!(!numerator.is_empty());
                assert!(!denominator.is_empty());
            }
            _ => panic!("Expected fraction node"),
        }
    }

    #[test]
    fn test_parse_fraction_with_properties() {
        let formula = Formula::new();
        let parser = OmmlParser::new(formula.arena());

        let xml = r#"<m:oMath>
            <m:f>
                <m:fPr><m:type>noBar</m:type></m:fPr>
                <m:num><m:r><m:t>a</m:t></m:r></m:num>
                <m:den><m:r><m:t>b</m:t></m:r></m:den>
            </m:f>
        </m:oMath>"#;

        let nodes = parser.parse(xml).unwrap();
        assert!(!nodes.is_empty());
        match &nodes[0] {
            MathNode::Frac { numerator, denominator, .. } => {
                assert!(!numerator.is_empty());
                assert!(!denominator.is_empty());
            }
            _ => panic!("Expected fraction node"),
        }
    }

    #[test]
    fn test_parse_delimiter() {
        let formula = Formula::new();
        let parser = OmmlParser::new(formula.arena());

        let xml = r#"<m:oMath>
            <m:d>
                <m:dPr>
                    <m:begChr>(</m:begChr>
                    <m:endChr>)</m:endChr>
                </m:dPr>
                <m:e><m:r><m:t>x</m:t></m:r></m:e>
            </m:d>
        </m:oMath>"#;

        let nodes = parser.parse(xml).unwrap();
        assert!(!nodes.is_empty());
        match &nodes[0] {
            MathNode::Fenced { open, close, .. } => {
                assert_eq!(*open, Fence::Paren);
                assert_eq!(*close, Fence::Paren);
            }
            _ => panic!("Expected fenced node"),
        }
    }

    #[test]
    fn test_parse_delimiter_brackets() {
        let formula = Formula::new();
        let parser = OmmlParser::new(formula.arena());

        let xml = r#"<m:oMath>
            <m:d>
                <m:dPr>
                    <m:begChr>[</m:begChr>
                    <m:endChr>]</m:endChr>
                </m:dPr>
                <m:e><m:r><m:t>x</m:t></m:r></m:e>
            </m:d>
        </m:oMath>"#;

        let nodes = parser.parse(xml).unwrap();
        assert!(!nodes.is_empty());
        match &nodes[0] {
            MathNode::Fenced { open, close, .. } => {
                assert_eq!(*open, Fence::Bracket);
                assert_eq!(*close, Fence::Bracket);
            }
            _ => panic!("Expected fenced node"),
        }
    }

    #[test]
    fn test_parse_function() {
        let formula = Formula::new();
        let parser = OmmlParser::new(formula.arena());

        let xml = r#"<m:oMath>
            <m:func>
                <m:fName><m:r><m:t>sin</m:t></m:r></m:fName>
                <m:e><m:r><m:t>x</m:t></m:r></m:e>
            </m:func>
        </m:oMath>"#;

        let nodes = parser.parse(xml).unwrap();
        assert!(!nodes.is_empty());
        match &nodes[0] {
            MathNode::Function { name, .. } => {
                assert_eq!(name.as_ref(), "sin");
            }
            _ => panic!("Expected function node"),
        }
    }

    #[test]
    fn test_parse_function_complex() {
        let formula = Formula::new();
        let parser = OmmlParser::new(formula.arena());

        let xml = r#"<m:oMath>
            <m:func>
                <m:fName><m:r><m:t>log</m:t></m:r></m:fName>
                <m:e>
                    <m:sSub>
                        <m:e><m:r><m:t>x</m:t></m:r></m:e>
                        <m:sub><m:r><m:t>2</m:t></m:r></m:sub>
                    </m:sSub>
                </m:e>
            </m:func>
        </m:oMath>"#;

        let nodes = parser.parse(xml).unwrap();
        assert!(!nodes.is_empty());
        match &nodes[0] {
            MathNode::Function { name, argument } => {
                assert_eq!(name.as_ref(), "log");
                assert!(!argument.is_empty());
            }
            _ => panic!("Expected function node"),
        }
    }

    #[test]
    fn test_parse_accent() {
        let formula = Formula::new();
        let parser = OmmlParser::new(formula.arena());

        let xml = r#"<m:oMath>
            <m:acc>
                <m:accPr>
                    <m:chr>^</m:chr>
                </m:accPr>
                <m:e><m:r><m:t>x</m:t></m:r></m:e>
            </m:acc>
        </m:oMath>"#;

        let nodes = parser.parse(xml).unwrap();
        assert!(!nodes.is_empty());
        match &nodes[0] {
            MathNode::Accent { accent, .. } => {
                assert_eq!(*accent, AccentType::Hat);
            }
            _ => panic!("Expected accent node"),
        }
    }

    #[test]
    fn test_parse_accent_bar() {
        let formula = Formula::new();
        let parser = OmmlParser::new(formula.arena());

        let xml = r#"<m:oMath>
            <m:acc>
                <m:accPr>
                    <m:chr>&#175;</m:chr>
                </m:accPr>
                <m:e><m:r><m:t>x</m:t></m:r></m:e>
            </m:acc>
        </m:oMath>"#;

        let nodes = parser.parse(xml).unwrap();
        assert!(!nodes.is_empty());
        match &nodes[0] {
            MathNode::Accent { accent, .. } => {
                assert_eq!(*accent, AccentType::Bar);
            }
            _ => panic!("Expected accent node"),
        }
    }

    #[test]
    fn test_parse_bar() {
        let formula = Formula::new();
        let parser = OmmlParser::new(formula.arena());

        let xml = r#"<m:oMath>
            <m:bar>
                <m:e><m:r><m:t>x</m:t></m:r></m:e>
            </m:bar>
        </m:oMath>"#;

        let nodes = parser.parse(xml).unwrap();
        assert!(!nodes.is_empty());
        match &nodes[0] {
            MathNode::Over { .. } => {
                // Bar is represented as Over node
            }
            _ => panic!("Expected over node"),
        }
    }

    #[test]
    fn test_parse_nary_with_limits() {
        let formula = Formula::new();
        let parser = OmmlParser::new(formula.arena());

        let xml = r#"<m:oMath>
            <m:nary>
                <m:naryPr>
                    <m:chr>∑</m:chr>
                </m:naryPr>
                <m:sub><m:r><m:t>i</m:t></m:r><m:t>=</m:t><m:r><m:t>1</m:t></m:r></m:sub>
                <m:sup><m:r><m:t>n</m:t></m:r></m:sup>
                <m:e><m:r><m:t>a</m:t></m:r><m:sub><m:r><m:t>i</m:t></m:r></m:sub></m:e>
            </m:nary>
        </m:oMath>"#;

        let nodes = match parser.parse(xml) {
            Ok(nodes) => nodes,
            Err(e) => {
                println!("Parse error: {:?}", e);
                panic!("Parse failed: {:?}", e);
            }
        };
        assert!(!nodes.is_empty());
        match &nodes[0] {
            MathNode::LargeOp { operator, lower_limit, upper_limit, .. } => {
                assert_eq!(*operator, LargeOperator::Sum);
                assert!(lower_limit.is_some());
                assert!(upper_limit.is_some());
            }
            _ => panic!("Expected large operator node"),
        }
    }

    #[test]
    fn test_parse_nary_integral() {
        let formula = Formula::new();
        let parser = OmmlParser::new(formula.arena());

        let xml = r#"<m:oMath>
            <m:nary>
                <m:naryPr>
                    <m:chr>∫</m:chr>
                </m:naryPr>
                <m:sub><m:r><m:t>0</m:t></m:r></m:sub>
                <m:sup><m:r><m:t>1</m:t></m:r></m:sup>
                <m:e><m:r><m:t>x</m:t></m:r><m:sup><m:r><m:t>2</m:t></m:r></m:sup></m:e>
            </m:nary>
        </m:oMath>"#;

        let nodes = parser.parse(xml).unwrap();
        assert!(!nodes.is_empty());
        match &nodes[0] {
            MathNode::LargeOp { operator, lower_limit, upper_limit, .. } => {
                assert_eq!(*operator, LargeOperator::Integral);
                assert!(lower_limit.is_some());
                assert!(upper_limit.is_some());
            }
            _ => panic!("Expected large operator node"),
        }
    }

    #[test]
    fn test_parse_superscript() {
        let formula = Formula::new();
        let parser = OmmlParser::new(formula.arena());

        let xml = r#"<m:oMath>
            <m:sSup>
                <m:e><m:r><m:t>x</m:t></m:r></m:e>
                <m:sup><m:r><m:t>2</m:t></m:r></m:sup>
            </m:sSup>
        </m:oMath>"#;

        let nodes = parser.parse(xml).unwrap();
        assert!(!nodes.is_empty());
        match &nodes[0] {
            MathNode::Power { base, exponent } => {
                assert!(!base.is_empty());
                assert!(!exponent.is_empty());
            }
            _ => panic!("Expected power node"),
        }
    }

    #[test]
    fn test_parse_subscript() {
        let formula = Formula::new();
        let parser = OmmlParser::new(formula.arena());

        let xml = r#"<m:oMath>
            <m:sSub>
                <m:e><m:r><m:t>x</m:t></m:r></m:e>
                <m:sub><m:r><m:t>i</m:t></m:r></m:sub>
            </m:sSub>
        </m:oMath>"#;

        let nodes = parser.parse(xml).unwrap();
        assert!(!nodes.is_empty());
        match &nodes[0] {
            MathNode::Sub { base, subscript } => {
                assert!(!base.is_empty());
                assert!(!subscript.is_empty());
            }
            _ => panic!("Expected sub node"),
        }
    }

    #[test]
    fn test_parse_subsup() {
        let formula = Formula::new();
        let parser = OmmlParser::new(formula.arena());

        let xml = r#"<m:oMath>
            <m:sSubSup>
                <m:e><m:r><m:t>x</m:t></m:r></m:e>
                <m:sub><m:r><m:t>i</m:t></m:r></m:sub>
                <m:sup><m:r><m:t>2</m:t></m:r></m:sup>
            </m:sSubSup>
        </m:oMath>"#;

        let nodes = parser.parse(xml).unwrap();
        assert!(!nodes.is_empty());
        match &nodes[0] {
            MathNode::SubSup { base, subscript, superscript } => {
                assert!(!base.is_empty());
                assert!(!subscript.is_empty());
                assert!(!superscript.is_empty());
            }
            _ => panic!("Expected subsup node"),
        }
    }

    #[test]
    fn test_parse_radical() {
        let formula = Formula::new();
        let parser = OmmlParser::new(formula.arena());

        let xml = r#"<m:oMath>
            <m:rad>
                <m:deg><m:r><m:t>2</m:t></m:r></m:deg>
                <m:e><m:r><m:t>x</m:t></m:r></m:e>
            </m:rad>
        </m:oMath>"#;

        let nodes = parser.parse(xml).unwrap();
        assert!(!nodes.is_empty());
        match &nodes[0] {
            MathNode::Root { base, index } => {
                assert!(!base.is_empty());
                assert!(index.is_some());
            }
            _ => panic!("Expected root node"),
        }
    }

    #[test]
    fn test_parse_radical_simple() {
        let formula = Formula::new();
        let parser = OmmlParser::new(formula.arena());

        let xml = r#"<m:oMath>
            <m:rad>
                <m:e><m:r><m:t>x</m:t></m:r></m:e>
            </m:rad>
        </m:oMath>"#;

        let nodes = parser.parse(xml).unwrap();
        assert!(!nodes.is_empty());
        match &nodes[0] {
            MathNode::Root { base, index } => {
                assert!(!base.is_empty());
                assert!(index.is_none());
            }
            _ => panic!("Expected root node"),
        }
    }

    #[test]
    fn test_parse_matrix() {
        let formula = Formula::new();
        let parser = OmmlParser::new(formula.arena());

        let xml = r#"<m:oMath>
            <m:m>
                <m:mPr>
                    <m:mcs>
                        <m:mc>
                            <m:mcPr>
                                <m:count>2</m:count>
                                <m:mcJc>center</m:mcJc>
                            </m:mcPr>
                        </m:mc>
                    </m:mcs>
                </m:mPr>
                <m:mr>
                    <m:e><m:r><m:t>a</m:t></m:r></m:e>
                    <m:e><m:r><m:t>b</m:t></m:r></m:e>
                </m:mr>
                <m:mr>
                    <m:e><m:r><m:t>c</m:t></m:r></m:e>
                    <m:e><m:r><m:t>d</m:t></m:r></m:e>
                </m:mr>
            </m:m>
        </m:oMath>"#;

        let nodes = parser.parse(xml).unwrap();
        assert!(!nodes.is_empty());
        match &nodes[0] {
            MathNode::Matrix { rows, .. } => {
                assert!(!rows.is_empty());
            }
            _ => panic!("Expected matrix node"),
        }
    }

    #[test]
    fn test_parse_box() {
        let formula = Formula::new();
        let parser = OmmlParser::new(formula.arena());

        let xml = r#"<m:oMath>
            <m:box>
                <m:e><m:r><m:t>x</m:t></m:r></m:e>
            </m:box>
        </m:oMath>"#;

        let nodes = parser.parse(xml).unwrap();
        assert!(!nodes.is_empty());
        match &nodes[0] {
            MathNode::Fenced { open, close, .. } => {
                assert_eq!(*open, Fence::None);
                assert_eq!(*close, Fence::None);
            }
            _ => panic!("Expected fenced node"),
        }
    }

    #[test]
    fn test_parse_phantom() {
        let formula = Formula::new();
        let parser = OmmlParser::new(formula.arena());

        let xml = r#"<m:oMath>
            <m:phant>
                <m:e><m:r><m:t>x</m:t></m:r></m:e>
            </m:phant>
        </m:oMath>"#;

        let nodes = parser.parse(xml).unwrap();
        assert!(!nodes.is_empty());
        match &nodes[0] {
            MathNode::Phantom(content) => {
                assert!(!content.is_empty());
            }
            _ => panic!("Expected phantom node"),
        }
    }

    #[test]
    fn test_parse_border_box() {
        let formula = Formula::new();
        let parser = OmmlParser::new(formula.arena());

        let xml = r#"<m:oMath>
            <m:borderBox>
                <m:e><m:r><m:t>x</m:t></m:r></m:e>
            </m:borderBox>
        </m:oMath>"#;

        let nodes = parser.parse(xml).unwrap();
        assert!(!nodes.is_empty());
        match &nodes[0] {
            MathNode::Fenced { open, close, .. } => {
                assert_eq!(*open, Fence::None);
                assert_eq!(*close, Fence::None);
            }
            _ => panic!("Expected fenced node"),
        }
    }

    #[test]
    fn test_parse_equation_array() {
        let formula = Formula::new();
        let parser = OmmlParser::new(formula.arena());

        let xml = r#"<m:oMath>
            <m:eqArr>
                <m:eqArrPr>
                    <m:baseJc>center</m:baseJc>
                </m:eqArrPr>
                <m:e><m:r><m:t>a</m:t></m:r><m:t>=</m:t><m:r><m:t>b</m:t></m:r></m:e>
                <m:e><m:r><m:t>c</m:t></m:r><m:t>=</m:t><m:r><m:t>d</m:t></m:r></m:e>
            </m:eqArr>
        </m:oMath>"#;

        let nodes = parser.parse(xml).unwrap();
        assert!(!nodes.is_empty());
        match &nodes[0] {
            MathNode::Matrix { rows, .. } => {
                assert!(!rows.is_empty());
            }
            _ => panic!("Expected matrix node"),
        }
    }

    #[test]
    fn test_parse_group_char() {
        let formula = Formula::new();
        let parser = OmmlParser::new(formula.arena());

        let xml = r#"<m:oMath>
            <m:groupChr>
                <m:groupChrPr>
                    <m:chr>{</m:chr>
                    <m:pos>top</m:pos>
                </m:groupChrPr>
                <m:e><m:r><m:t>x</m:t></m:r></m:e>
            </m:groupChr>
        </m:oMath>"#;

        let nodes = parser.parse(xml).unwrap();
        assert!(!nodes.is_empty());
        match &nodes[0] {
            MathNode::Over { .. } => {
                // Group char is represented as Over node
            }
            _ => panic!("Expected over node"),
        }
    }

    #[test]
    fn test_parse_spacing() {
        let formula = Formula::new();
        let parser = OmmlParser::new(formula.arena());

        let xml = r#"<m:oMath>
            <m:r><m:t>a</m:t></m:r>
            <m:sPre>
                <m:sPrePr><m:val>thickmathspace</m:val></m:sPrePr>
            </m:sPre>
            <m:r><m:t>b</m:t></m:r>
        </m:oMath>"#;

        let nodes = parser.parse(xml).unwrap();
        assert!(!nodes.is_empty());
        // Should contain text nodes and spacing
    }

    #[test]
    fn test_parse_complex_expression() {
        let formula = Formula::new();
        let parser = OmmlParser::new(formula.arena());

        let xml = r#"<m:oMath>
            <m:func>
                <m:fName><m:r><m:t>sin</m:t></m:r></m:fName>
                <m:e>
                    <m:f>
                        <m:num><m:r><m:t>x</m:t></m:r></m:num>
                        <m:den>
                            <m:sSup>
                                <m:e><m:r><m:t>y</m:t></m:r></m:e>
                                <m:sup><m:r><m:t>2</m:t></m:r></m:sup>
                            </m:sSup>
                        </m:den>
                    </m:f>
                </m:e>
            </m:func>
        </m:oMath>"#;

        let nodes = parser.parse(xml).unwrap();
        assert!(!nodes.is_empty());
        match &nodes[0] {
            MathNode::Function { name, argument } => {
                assert_eq!(name.as_ref(), "sin");
                assert!(!argument.is_empty());
            }
            _ => panic!("Expected function node"),
        }
    }

    #[test]
    fn test_parse_empty_math() {
        let formula = Formula::new();
        let parser = OmmlParser::new(formula.arena());

        let xml = r#"<m:oMath></m:oMath>"#;
        let nodes = parser.parse(xml).unwrap();
        assert!(nodes.is_empty());
    }

    #[test]
    fn test_parse_invalid_xml() {
        let formula = Formula::new();
        let parser = OmmlParser::new(formula.arena());

        let xml = r#"<m:oMath><invalid></m:oMath>"#;
        let result = parser.parse(xml);
        assert!(result.is_ok()); // Should handle unknown elements gracefully
    }

    #[test]
    fn test_parse_malformed_xml() {
        let formula = Formula::new();
        let parser = OmmlParser::new(formula.arena());

        let xml = r#"<m:oMath><m:r><m:t>unclosed"#;
        let result = parser.parse(xml);
        assert!(result.is_err()); // Should return error for malformed XML
    }

    #[test]
    fn test_parse_unicode_characters() {
        let formula = Formula::new();
        let parser = OmmlParser::new(formula.arena());

        let xml = r#"<m:oMath>
            <m:r><m:t>α</m:t></m:r>
            <m:r><m:t>β</m:t></m:r>
            <m:r><m:t>∑</m:t></m:r>
        </m:oMath>"#;
        let nodes = parser.parse(xml).unwrap();

        assert_eq!(nodes.len(), 3);
        match (&nodes[0], &nodes[1], &nodes[2]) {
            (MathNode::Text(a), MathNode::Text(b), MathNode::Text(c)) => {
                assert_eq!(a.as_ref(), "α");
                assert_eq!(b.as_ref(), "β");
                assert_eq!(c.as_ref(), "∑");
            }
            _ => panic!("Expected text nodes"),
        }
    }

    #[test]
    fn test_parse_complex_nested_expression() {
        let formula = Formula::new();
        let parser = OmmlParser::new(formula.arena());

        let xml = r#"<m:oMath>
            <m:sSup>
                <m:e>
                    <m:func>
                        <m:fName><m:r><m:t>sin</m:t></m:r></m:fName>
                        <m:e>
                            <m:f>
                                <m:num><m:r><m:t>x</m:t></m:r></m:num>
                                <m:den><m:r><m:t>y</m:t></m:r></m:den>
                            </m:f>
                        </m:e>
                    </m:func>
                </m:e>
                <m:sup><m:r><m:t>2</m:t></m:r></m:sup>
            </m:sSup>
        </m:oMath>"#;

        let nodes = parser.parse(xml).unwrap();
        assert!(!nodes.is_empty());
        // Verify the structure contains a power node
        match &nodes[0] {
            MathNode::Power { base, exponent } => {
                assert!(!base.is_empty());
                assert!(!exponent.is_empty());
            }
            _ => panic!("Expected power node"),
        }
    }

    #[test]
    fn test_parse_matrix_with_properties() {
        let formula = Formula::new();
        let parser = OmmlParser::new(formula.arena());

        let xml = r#"<m:oMath>
            <m:m>
                <m:mPr>
                    <m:mcs>
                        <m:mc>
                            <m:mcPr>
                                <m:count>2</m:count>
                                <m:mcJc>center</m:mcJc>
                            </m:mcPr>
                        </m:mc>
                    </m:mcs>
                </m:mPr>
                <m:mr>
                    <m:e><m:r><m:t>a</m:t></m:r></m:e>
                    <m:e><m:r><m:t>b</m:t></m:r></m:e>
                </m:mr>
                <m:mr>
                    <m:e><m:r><m:t>c</m:t></m:r></m:e>
                    <m:e><m:r><m:t>d</m:t></m:r></m:e>
                </m:mr>
            </m:m>
        </m:oMath>"#;

        let nodes = parser.parse(xml).unwrap();
        assert!(!nodes.is_empty());
        match &nodes[0] {
            MathNode::Matrix { rows, .. } => {
                assert_eq!(rows.len(), 2);
                assert_eq!(rows[0].len(), 2);
                assert_eq!(rows[1].len(), 2);
            }
            _ => panic!("Expected matrix node"),
        }
    }

    #[test]
    fn test_parse_nary_with_complex_limits() {
        let formula = Formula::new();
        let parser = OmmlParser::new(formula.arena());

        let xml = r#"<m:oMath>
            <m:nary>
                <m:naryPr>
                    <m:chr>∑</m:chr>
                </m:naryPr>
                <m:sub>
                    <m:sSub>
                        <m:e><m:r><m:t>i</m:t></m:r></m:e>
                        <m:sub><m:r><m:t>0</m:t></m:r></m:sub>
                    </m:sSub>
                </m:sub>
                <m:sup><m:r><m:t>n</m:t></m:r></m:sup>
                <m:e>
                    <m:sSup>
                        <m:e><m:r><m:t>x</m:t></m:r></m:e>
                        <m:sup><m:r><m:t>i</m:t></m:r></m:sup>
                    </m:sSup>
                </m:e>
            </m:nary>
        </m:oMath>"#;

        let nodes = parser.parse(xml).unwrap();
        assert!(!nodes.is_empty());
        match &nodes[0] {
            MathNode::LargeOp { operator, lower_limit, upper_limit, integrand, .. } => {
                assert_eq!(*operator, LargeOperator::Sum);
                assert!(lower_limit.is_some());
                assert!(upper_limit.is_some());
                assert!(integrand.is_some());
            }
            _ => panic!("Expected large operator node"),
        }
    }

    #[test]
    fn test_parse_accent_with_position() {
        let formula = Formula::new();
        let parser = OmmlParser::new(formula.arena());

        let xml = r#"<m:oMath>
            <m:acc>
                <m:accPr>
                    <m:chr>→</m:chr>
                </m:accPr>
                <m:e><m:r><m:t>v</m:t></m:r></m:e>
            </m:acc>
        </m:oMath>"#;

        let nodes = parser.parse(xml).unwrap();
        assert!(!nodes.is_empty());
        match &nodes[0] {
            MathNode::Accent { accent, base, .. } => {
                assert_eq!(*accent, AccentType::Vec);
                assert!(!base.is_empty());
            }
            _ => panic!("Expected accent node"),
        }
    }

    #[test]
    fn test_parse_group_character_with_position() {
        let formula = Formula::new();
        let parser = OmmlParser::new(formula.arena());

        let xml = r#"<m:oMath>
            <m:groupChr>
                <m:groupChrPr>
                    <m:chr>{</m:chr>
                    <m:pos>top</m:pos>
                    <m:vertJc>center</m:vertJc>
                </m:groupChrPr>
                <m:e>
                    <m:f>
                        <m:num><m:r><m:t>a</m:t></m:r></m:num>
                        <m:den><m:r><m:t>b</m:t></m:r></m:den>
                    </m:f>
                </m:e>
            </m:groupChr>
        </m:oMath>"#;

        let nodes = parser.parse(xml).unwrap();
        assert!(!nodes.is_empty());
        match &nodes[0] {
            MathNode::GroupChar { base, character, position, vertical_alignment } => {
                assert!(!base.is_empty());
                assert_eq!(character.as_deref(), Some("{"));
                assert_eq!(*position, Some(crate::formula::ast::Position::Top));
                assert_eq!(*vertical_alignment, Some(crate::formula::ast::VerticalAlignment::Center));
            }
            _ => panic!("Expected group character node"),
        }
    }

    #[test]
    fn test_parse_phantom_element() {
        let formula = Formula::new();
        let parser = OmmlParser::new(formula.arena());

        let xml = r#"<m:oMath>
            <m:phant>
                <m:e><m:r><m:t>x</m:t></m:r></m:e>
            </m:phant>
        </m:oMath>"#;

        let nodes = parser.parse(xml).unwrap();
        assert!(!nodes.is_empty());
        match &nodes[0] {
            MathNode::Phantom(content) => {
                assert!(!content.is_empty());
            }
            _ => panic!("Expected phantom node"),
        }
    }

    #[test]
    fn test_parse_radical_with_degree() {
        let formula = Formula::new();
        let parser = OmmlParser::new(formula.arena());

        let xml = r#"<m:oMath>
            <m:rad>
                <m:deg><m:r><m:t>3</m:t></m:r></m:deg>
                <m:e><m:r><m:t>x</m:t></m:r></m:e>
            </m:rad>
        </m:oMath>"#;

        let nodes = parser.parse(xml).unwrap();
        assert!(!nodes.is_empty());
        match &nodes[0] {
            MathNode::Root { base, index } => {
                assert!(!base.is_empty());
                assert!(index.is_some());
            }
            _ => panic!("Expected root node with index"),
        }
    }

    #[test]
    fn test_parse_spacing_element() {
        let formula = Formula::new();
        let parser = OmmlParser::new(formula.arena());

        let xml = r#"<m:oMath>
            <m:r><m:t>a</m:t></m:r>
            <m:sPre>
                <m:sPrePr><m:val>thickmathspace</m:val></m:sPrePr>
            </m:sPre>
            <m:r><m:t>b</m:t></m:r>
        </m:oMath>"#;

        let nodes = parser.parse(xml).unwrap();
        assert!(nodes.len() >= 2); // Should have at least text and spacing
    }

    #[test]
    fn test_validation_empty_math() {
        let formula = Formula::new();
        let parser = OmmlParser::new(formula.arena());

        let xml = "<m:oMath></m:oMath>";
        let result = parser.parse(xml);
        assert!(result.is_err());
    }

    #[test]
    fn test_validation_malformed_xml() {
        let formula = Formula::new();
        let parser = OmmlParser::new(formula.arena());

        let xml = "<m:oMath><unclosed>";
        let result = parser.parse(xml);
        assert!(result.is_err());
    }

    #[test]
    fn test_validation_invalid_nesting() {
        let formula = Formula::new();
        let parser = OmmlParser::new(formula.arena());

        // Numerator outside of fraction
        let xml = r#"<m:oMath>
            <m:num><m:r><m:t>1</m:t></m:r></m:num>
        </m:oMath>"#;
        let result = parser.parse(xml);
        assert!(result.is_err());
    }

    #[test]
    fn test_validation_missing_required_elements() {
        let formula = Formula::new();
        let parser = OmmlParser::new(formula.arena());

        // Fraction with empty numerator
        let xml = r#"<m:oMath>
            <m:f>
                <m:num></m:num>
                <m:den><m:r><m:t>2</m:t></m:r></m:den>
            </m:f>
        </m:oMath>"#;
        let result = parser.parse(xml);
        assert!(result.is_err());
    }

    #[test]
    fn test_predefined_symbols() {
        let formula = Formula::new();
        let parser = OmmlParser::new(formula.arena());

        let xml = r#"<m:oMath>
            <m:r><m:t>α</m:t></m:r>
            <m:r><m:t>β</m:t></m:r>
            <m:r><m:t>∞</m:t></m:r>
            <m:r><m:t>∑</m:t></m:r>
        </m:oMath>"#;

        let nodes = parser.parse(xml).unwrap();
        assert_eq!(nodes.len(), 4);
    }

    #[test]
    fn test_deep_nesting_limit() {
        let formula = Formula::new();
        let parser = OmmlParser::new(formula.arena());

        // Create deeply nested XML that exceeds the limit
        let mut xml = "<m:oMath>".to_string();
        for _ in 0..1010 {
            xml.push_str("<m:f><m:num>");
        }
        xml.push_str("<m:r><m:t>x</m:t></m:r>");
        for _ in 0..1010 {
            xml.push_str("</m:num><m:den><m:r><m:t>1</m:t></m:r></m:den></m:f>");
        }
        xml.push_str("</m:oMath>");

        let result = parser.parse(&xml);
        assert!(result.is_err());
    }

    #[test]
    fn test_run_properties() {
        let formula = Formula::new();
        let parser = OmmlParser::new(formula.arena());

        let xml = r#"<m:oMath>
            <m:r>
                <m:rPr>
                    <m:scr>bi</m:scr>
                    <m:sty>p</m:sty>
                    <m:nor>Times New Roman</m:nor>
                    <m:lit>1</m:lit>
                </m:rPr>
                <m:t>text</m:t>
            </m:r>
        </m:oMath>"#;

        let nodes = parser.parse(xml).unwrap();
        assert!(!nodes.is_empty());
        match &nodes[0] {
            MathNode::Run { content, literal, style, font, .. } => {
                assert!(!content.is_empty());
                assert!(literal.unwrap_or(false));
                assert_eq!(*style, Some(crate::formula::StyleType::BoldItalic));
                assert_eq!(font.as_deref(), Some("Times New Roman"));
            }
            _ => panic!("Expected run node"),
        }
    }
}

