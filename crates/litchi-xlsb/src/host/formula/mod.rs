//! XLSB formula parsing and generation
//!
//! Excel formulas in XLSB files are stored in a binary format using Reverse Polish Notation (RPN)
//! with Parse Tree Generators (Ptg tokens). This module provides parsing and generation of formulas.
//!
//! # Formula Token Types (Ptgs)
//!
//! Formulas are sequences of tokens that represent operands, operators, and functions:
//! - **Value tokens**: Numbers, strings, booleans, errors
//! - **Operand tokens**: Cell references, ranges, names
//! - **Operator tokens**: Add, subtract, multiply, divide, etc.
//! - **Function tokens**: SUM, IF, VLOOKUP, etc.
//!
//! # Binary Format
//!
//! Each token consists of:
//! 1. Token type byte (identifies the Ptg)
//! 2. Token data (variable length, depends on token type)
//!
//! # Reference
//!
//! - [MS-XLSB] Section 2.5.98 - Formulas
//! - [MS-XLS] Section 2.5.198 - Ptg (for token details, largely compatible)

mod error;
#[path = "../function_table.rs"]
mod function_table;
mod parser;
mod semantic;
mod validation;

pub mod model;
pub mod pivot;
pub mod resolution;
pub mod table;
pub mod text;

pub use model::ExternalSheet;
pub use pivot::{Aggregation, Item, Name, Scope, View};
pub use resolution::{Context, ExternalBook, SupportingLink};
pub use table::Definition;
pub(crate) use text::{CompilationContext, DefinedName, excel_name_eq};

use crate::package::error::{Error, Result};
#[cfg(test)]
use crate::package::external_link::Link;

pub use parser::{
    ArrayValue, BinaryOperator, Compiler, ExternalTableReference, Group, GroupKind,
    MAX_CELL_FORMULA_BYTES, MemoryKind, ParsedFormula, Parser, Range, Resolution, TableColumns,
    TableDataType, TableNamedColumns, TableReference, TableRowType, Token, UnaryOperator,
    ptg_types,
};

use semantic::{format_pivot_identifier, format_structured_reference};
use validation::{
    invalid, validate_named_table_columns, validate_pivot_identifier, validate_table_column_name,
    validate_table_name,
};

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_parse_operators() {
        let data = vec![0x03]; // PTG_ADD
        let mut parser = Parser::new(&data);
        let tokens = parser.parse().unwrap();
        assert_eq!(tokens.len(), 1);
        match &tokens[0] {
            Token::BinaryOp(BinaryOperator::Add) => {},
            _ => panic!("Expected Add operator"),
        }
    }

    #[test]
    fn test_parse_number() {
        let mut data = vec![0x1F]; // PTG_NUM
        data.extend_from_slice(&42.5f64.to_le_bytes());
        let mut parser = Parser::new(&data);
        let tokens = parser.parse().unwrap();
        assert_eq!(tokens.len(), 1);
        match &tokens[0] {
            Token::Number(n) if (*n - 42.5).abs() < 0.001 => {},
            _ => panic!("Expected number 42.5"),
        }
    }

    #[test]
    fn test_formula_converter() {
        let tokens = vec![
            Token::Number(1.0),
            Token::Number(2.0),
            Token::BinaryOp(BinaryOperator::Add),
        ];
        let formula = Compiler::tokens_to_string(&tokens);
        assert_eq!(formula, "(1+2)");
    }

    #[test]
    fn parses_ms_xlsb_brt_fmla_num_example_formula() {
        // [MS-XLSB] 3.7.37: PtgRef(C13), PtgInt(2), PtgMul.
        let rgce = vec![
            0x44, 0x0C, 0x00, 0x00, 0x00, 0x02, 0xC0, 0x1E, 0x02, 0x00, 0x05,
        ];
        let parsed = ParsedFormula {
            rgce: rgce.clone(),
            rgcb: Vec::new(),
        };
        let bytes = parsed.to_bytes().unwrap();
        let (roundtrip, consumed) = ParsedFormula::parse(&bytes).unwrap();
        assert_eq!(consumed, bytes.len());
        assert_eq!(roundtrip, parsed);

        let tokens = Parser::new(&rgce).parse().unwrap();
        assert_eq!(Compiler::try_tokens_to_string(&tokens).unwrap(), "(C13*2)");
    }

    #[test]
    fn compiler_matches_ms_xlsb_reference_and_multiply_tokens() {
        let formula = text::Compiler::compile("=C13*2").unwrap();
        assert_eq!(
            formula.rgce,
            vec![
                0x44, 0x0C, 0x00, 0x00, 0x00, 0x02, 0xC0, 0x1E, 0x02, 0x00, 0x05,
            ]
        );
    }

    #[test]
    fn compiler_emits_conditional_control_flow_attributes() {
        assert_eq!(
            text::Compiler::compile("IF(TRUE,1,2)").unwrap().rgce,
            vec![
                0x1D, 0x01, 0x19, 0x02, 0x07, 0x00, 0x1E, 0x01, 0x00, 0x19, 0x08, 0x0A, 0x00, 0x1E,
                0x02, 0x00, 0x19, 0x08, 0x03, 0x00, 0x42, 0x03, 0x01, 0x00,
            ]
        );
        assert_eq!(
            text::Compiler::compile("IFERROR(1,2)").unwrap().rgce,
            vec![
                0x1E, 0x01, 0x00, 0x19, 0x80, 0x07, 0x00, 0x1E, 0x02, 0x00, 0x19, 0x08, 0x02, 0x00,
                0x41, 0xE0, 0x01,
            ]
        );
        assert_eq!(
            text::Compiler::compile("CHOOSE(2,10,20)").unwrap().rgce,
            vec![
                0x1E, 0x02, 0x00, 0x19, 0x04, 0x02, 0x00, 0x06, 0x00, 0x07, 0x00, 0x0E, 0x00, 0x1E,
                0x0A, 0x00, 0x19, 0x08, 0x0A, 0x00, 0x1E, 0x14, 0x00, 0x19, 0x08, 0x03, 0x00, 0x42,
                0x03, 0x64, 0x00,
            ]
        );

        for source in ["IF(TRUE,1,2)", "IFERROR(1,2)", "CHOOSE(2,10,20)"] {
            let compiled = text::Compiler::compile(source).unwrap();
            let tokens = Parser::new(&compiled.rgce).parse().unwrap();
            assert_eq!(Compiler::try_tokens_to_string(&tokens).unwrap(), source);
        }

        let mut malformed_if = text::Compiler::compile("IF(TRUE,1,2)").unwrap().rgce;
        malformed_if[4] = 6;
        assert!(matches!(
            Parser::new(&malformed_if).parse(),
            Err(crate::formula::Error::InvalidFormula(_))
        ));
        let mut malformed_choose = text::Compiler::compile("CHOOSE(2,10,20)").unwrap().rgce;
        malformed_choose[7] = 5;
        assert!(matches!(
            Parser::new(&malformed_choose).parse(),
            Err(crate::formula::Error::InvalidFormula(_))
        ));
    }

    #[test]
    fn compiler_supports_ranges_functions_unicode_and_absolute_refs() {
        let formula = text::Compiler::compile("SUM($A$1:B3)+\"荔枝\"").unwrap();
        let tokens = Parser::new(&formula.rgce).parse().unwrap();
        let text = Compiler::try_tokens_to_string(&tokens).unwrap();
        assert_eq!(text, "(SUM($A$1:B3)+\"荔枝\")");
    }

    #[test]
    fn compiler_emits_contextual_names_and_sheet_references() {
        let worksheet_names = vec![
            "Data".to_string(),
            "O'Brien Data".to_string(),
            "Summary".to_string(),
        ];
        let defined_names = vec![
            DefinedName {
                name: "Rate".to_string(),
                sheet_id: None,
            },
            DefinedName {
                name: "Rate".to_string(),
                sheet_id: Some(1),
            },
        ];
        let sheet_ranges = std::cell::RefCell::new(Vec::new());
        let context = CompilationContext {
            worksheet_names: &worksheet_names,
            defined_names: &defined_names,
            tables: &[],
            supporting_links: &[],
            external_sheets: &[],
            external_books: &[],
            sheet_ranges: &sheet_ranges,
            current_sheet: 1,
        };
        let compiled =
            text::Compiler::compile_with_context("Rate+Data!A1:B2+'O''Brien Data'!$C$3", &context)
                .unwrap();

        assert!(compiled.rgce.starts_with(&[0x43, 2, 0, 0, 0]));
        assert!(
            compiled
                .rgce
                .windows(3)
                .any(|window| window == [0x5B, 2, 0])
        );
        assert!(
            compiled
                .rgce
                .windows(3)
                .any(|window| window == [0x5A, 3, 0])
        );
        assert!(text::Compiler::compile("Rate").is_err());
        assert!(text::Compiler::compile("Data!A1").is_err());

        let span = text::Compiler::compile_with_context(
            "SUM('Data:Summary'!A1)+Data:Summary!$B$2",
            &context,
        )
        .unwrap();
        assert_eq!(&*sheet_ranges.borrow(), &[(0, 2)]);
        assert_eq!(
            span.rgce
                .windows(3)
                .filter(|window| *window == [0x5A, 5, 0])
                .count(),
            2
        );
        assert!(text::Compiler::compile_with_context("Summary:Data!A1", &context).is_err());
    }

    #[test]
    fn builtin_function_table_is_sorted_unique_and_non_macro() {
        assert_eq!(function_table::BUILTIN_FUNCTIONS.len(), 363);
        assert!(
            function_table::BUILTIN_FUNCTIONS
                .windows(2)
                .all(|entries| entries[0].0 < entries[1].0)
        );

        let mut names = std::collections::HashSet::new();
        for &(index, name, min_args, max_args) in function_table::BUILTIN_FUNCTIONS {
            assert!(min_args <= max_args, "invalid arity for {name}");
            assert!(
                names.insert(name.to_ascii_uppercase()),
                "duplicate function name {name} at index {index}"
            );
        }

        assert!(!text::has_builtin_function(53)); // GOTO macro function
        assert!(!text::has_builtin_function(110)); // EXEC macro function
        assert!(!text::has_builtin_function(255)); // context-dependent UDF
        assert!(!text::has_builtin_function(468)); // future-function CONVERT
    }

    #[test]
    fn compiler_covers_legacy_analysis_and_ooxml_function_ranges() {
        let cases: &[(&str, &str, &[u8])] = &[
            ("ROUNDUP(1.2,0)", "ROUNDUP(1.2,0)", &[0x41, 0xD4, 0x00]),
            ("MEDIAN(1,2,3)", "MEDIAN(1,2,3)", &[0x42, 0x03, 0xE3, 0x00]),
            ("CUBESETCOUNT(1)", "CUBESETCOUNT(1)", &[0x41, 0xDF, 0x01]),
        ];
        for &(source, expected, token_suffix) in cases {
            let formula = text::Compiler::compile(source).unwrap();
            assert!(formula.rgce.ends_with(token_suffix));
            let tokens = Parser::new(&formula.rgce).parse().unwrap();
            assert_eq!(Compiler::try_tokens_to_string(&tokens).unwrap(), expected);
        }

        let accrint = text::Compiler::compile("ACCRINT(1,2,3,4,5,6,7,8)").unwrap();
        assert!(accrint.rgce.ends_with(&[0x42, 0x08, 0xD5, 0x01]));
        assert!(text::Compiler::compile("ACCRINT(1,2,3,4,5,6,7,8,9)").is_err());
    }

    #[test]
    fn compiler_and_parser_enforce_function_argument_grammars() {
        assert!(text::Compiler::compile("SUM()").is_err());
        assert!(text::Compiler::compile("COUNTIFS(A1,1)").is_ok());
        assert!(text::Compiler::compile("COUNTIFS(A1,1,B1)").is_err());
        assert!(text::Compiler::compile("SUMIFS(A1,B1,1)").is_ok());
        assert!(text::Compiler::compile("SUMIFS(A1,B1,1,C1)").is_err());
        assert!(text::Compiler::compile("GETPIVOTDATA(1,2,3)").is_ok());
        assert!(text::Compiler::compile("GETPIVOTDATA(1,2,3,4,5)").is_err());

        assert!(matches!(
            Parser::new(&[0x41, 0xE3, 0x00]).parse(),
            Err(crate::formula::Error::InvalidFormula(_))
        ));
        assert!(matches!(
            Parser::new(&[0x42, 0x02, 0xD4, 0x00]).parse(),
            Err(crate::formula::Error::InvalidFormula(_))
        ));
        assert!(matches!(
            Parser::new(&[0x42, 0x03, 0xE1, 0x01]).parse(),
            Err(crate::formula::Error::InvalidFormula(_))
        ));

        assert!(matches!(
            text::Compiler::compile("EXEC(\"calc\")"),
            Err(Error::UnsupportedFeature(_))
        ));
        assert!(matches!(
            text::Compiler::compile("CONVERT(1,\"m\",\"ft\")"),
            Err(Error::UnsupportedFeature(_))
        ));
    }

    #[test]
    fn variable_functions_support_the_full_u8_argument_count() {
        let formula_255 = format!("SUM({})", vec!["1"; 255].join(","));
        let compiled = text::Compiler::compile(&formula_255).unwrap();
        assert!(compiled.rgce.ends_with(&[0x42, 0xFF, 0x04, 0x00]));
        let tokens = Parser::new(&compiled.rgce).parse().unwrap();
        assert_eq!(
            Compiler::try_tokens_to_string(&tokens).unwrap(),
            formula_255
        );

        let formula_256 = format!("SUM({})", vec!["1"; 256].join(","));
        assert!(matches!(
            text::Compiler::compile(&formula_256),
            Err(Error::InvalidFormula(_))
        ));
    }

    #[test]
    fn compiler_and_converter_preserve_missing_arguments_and_parentheses() {
        let missing = text::Compiler::compile("IF(TRUE,,0)").unwrap();
        assert!(missing.rgce.contains(&ptg_types::PTG_MISSING_ARG));
        let tokens = Parser::new(&missing.rgce).parse().unwrap();
        assert_eq!(
            Compiler::try_tokens_to_string(&tokens).unwrap(),
            "IF(TRUE,,0)"
        );

        let parenthesized = text::Compiler::compile("(1+2)*3").unwrap();
        assert!(parenthesized.rgce.contains(&ptg_types::PTG_PAREN));
        let tokens = Parser::new(&parenthesized.rgce).parse().unwrap();
        assert_eq!(
            Compiler::try_tokens_to_string(&tokens).unwrap(),
            "(((1+2))*3)"
        );
    }

    #[test]
    fn parser_converts_binary_reference_operators() {
        let mut rgce = text::Compiler::compile("A1").unwrap().rgce;
        rgce.extend_from_slice(&text::Compiler::compile("B2").unwrap().rgce);
        rgce.push(ptg_types::PTG_UNION);
        let tokens = Parser::new(&rgce).parse().unwrap();
        assert_eq!(Compiler::try_tokens_to_string(&tokens).unwrap(), "(A1,B2)");
    }

    #[test]
    fn parser_decodes_all_reference_error_token_forms() {
        let cases = [
            (
                vec![0x4A, 1, 2, 3, 4, 5, 6],
                Token::ReferenceError {
                    is_area: false,
                    sheet_index: None,
                },
            ),
            (
                vec![0x4B, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0],
                Token::ReferenceError {
                    is_area: true,
                    sheet_index: None,
                },
            ),
            (
                vec![0x5C, 0x34, 0x12, 1, 2, 3, 4, 5, 6],
                Token::ReferenceError {
                    is_area: false,
                    sheet_index: Some(0x1234),
                },
            ),
            (
                vec![0x7D, 0x78, 0x56, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0],
                Token::ReferenceError {
                    is_area: true,
                    sheet_index: Some(0x5678),
                },
            ),
        ];

        for (bytes, expected) in cases {
            let tokens = Parser::new(&bytes).parse().unwrap();
            assert_eq!(tokens, vec![expected]);
            assert_eq!(Compiler::try_tokens_to_string(&tokens).unwrap(), "#REF!");
        }

        assert!(matches!(
            Parser::new(&[0x4B; 12]).parse(),
            Err(crate::formula::Error::InvalidFormula(_))
        ));
        assert!(matches!(
            Parser::new(&[0xAA, 0, 0, 0, 0, 0, 0]).parse(),
            Err(crate::formula::Error::InvalidFormula(_))
        ));
    }

    #[test]
    fn parser_resolves_internal_3d_references_and_defined_names() {
        let context = Context {
            worksheet_names: vec!["Data 1".to_string(), "Last".to_string()].into(),
            supporting_links: vec![SupportingLink::SelfWorkbook].into(),
            external_sheets: vec![
                ExternalSheet {
                    external_link: 0,
                    first_sheet: 0,
                    last_sheet: 0,
                },
                ExternalSheet {
                    external_link: 0,
                    first_sheet: 0,
                    last_sheet: 1,
                },
            ]
            .into(),
            external_books: Vec::new().into(),
            defined_names: vec!["Rate".to_string()].into(),
            tables: Vec::new().into(),
            pivot_views: Vec::new().into(),
            pivot_name_scopes: Vec::new().into(),
            active_pivot_scope: None,
            current_sheet: None,
        };

        let ref_3d = [0x5A, 0, 0, 1, 0, 0, 0, 0, 0xC0];
        let tokens = Parser::new(&ref_3d).parse().unwrap();
        assert_eq!(
            Compiler::try_tokens_to_string_with_resolution(&tokens, &context).unwrap(),
            "'Data 1'!A2"
        );
        assert!(Compiler::try_tokens_to_string(&tokens).is_err());

        let area_3d = [0x7B, 1, 0, 0, 0, 0, 0, 1, 0, 0, 0, 0, 0, 1, 0];
        let tokens = Parser::new(&area_3d).parse().unwrap();
        assert_eq!(
            Compiler::try_tokens_to_string_with_resolution(&tokens, &context).unwrap(),
            "'Data 1:Last'!$A$1:$B$2"
        );

        let name = [0x43, 1, 0, 0, 0];
        let tokens = Parser::new(&name).parse().unwrap();
        assert_eq!(
            Compiler::try_tokens_to_string_with_resolution(&tokens, &context).unwrap(),
            "Rate"
        );
    }

    #[test]
    fn contextual_reference_parser_rejects_invalid_indices_and_payloads() {
        let context = Context {
            worksheet_names: vec!["Sheet1".to_string()].into(),
            supporting_links: vec![SupportingLink::SelfWorkbook].into(),
            external_sheets: Vec::new().into(),
            external_books: Vec::new().into(),
            defined_names: Vec::new().into(),
            tables: Vec::new().into(),
            pivot_views: Vec::new().into(),
            pivot_name_scopes: Vec::new().into(),
            active_pivot_scope: None,
            current_sheet: None,
        };
        let invalid_xti = [0x5A, 0, 0, 0, 0, 0, 0, 0, 0];
        let tokens = Parser::new(&invalid_xti).parse().unwrap();
        assert!(matches!(
            Compiler::try_tokens_to_string_with_resolution(&tokens, &context),
            Err(crate::formula::Error::InvalidFormula(_))
        ));
        assert!(matches!(
            Parser::new(&[0x43, 0, 0, 0, 0]).parse(),
            Err(crate::formula::Error::InvalidFormula(_))
        ));
        assert!(matches!(
            Parser::new(&[0xDA, 0, 0, 0, 0, 0, 0, 0, 0]).parse(),
            Err(crate::formula::Error::InvalidFormula(_))
        ));
        assert!(matches!(
            Parser::new(&[0x5B; 14]).parse(),
            Err(crate::formula::Error::InvalidFormula(_))
        ));
    }

    #[test]
    fn parser_rejects_reserved_class_bits_and_invalid_absolute_ranges() {
        for bytes in [
            vec![0xA4, 0, 0, 0, 0, 0, 0],
            vec![0xA1, 0, 0],
            vec![0xC3, 1, 0, 0, 0],
            vec![0xB9, 0, 0, 1, 0, 0, 0],
        ] {
            assert!(matches!(
                Parser::new(&bytes).parse(),
                Err(crate::formula::Error::InvalidFormula(_))
            ));
        }

        let row_past_end = [0x44, 0x00, 0x00, 0x10, 0x00, 0, 0];
        assert!(matches!(
            Parser::new(&row_past_end).parse(),
            Err(crate::formula::Error::InvalidFormula(_))
        ));

        let reversed_area = [0x45, 1, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0];
        assert!(matches!(
            Parser::new(&reversed_area).parse(),
            Err(crate::formula::Error::InvalidCellReference(_))
        ));
    }

    #[test]
    fn resolves_same_sheet_supporting_links_in_the_consuming_sheet() {
        let context = Context {
            worksheet_names: vec!["First".to_string(), "Current Sheet".to_string()].into(),
            supporting_links: vec![SupportingLink::SameSheet].into(),
            external_sheets: vec![ExternalSheet {
                external_link: 0,
                first_sheet: -2,
                last_sheet: -2,
            }]
            .into(),
            external_books: Vec::new().into(),
            defined_names: Vec::new().into(),
            tables: Vec::new().into(),
            pivot_views: Vec::new().into(),
            pivot_name_scopes: Vec::new().into(),
            active_pivot_scope: None,
            current_sheet: None,
        }
        .for_sheet(1);
        let tokens = Parser::new(&[0x5A, 0, 0, 0, 0, 0, 0, 0, 0])
            .parse()
            .unwrap();
        assert_eq!(
            Compiler::try_tokens_to_string_with_resolution(&tokens, &context).unwrap(),
            "'Current Sheet'!$A$1"
        );
    }

    #[test]
    fn parser_resolves_external_workbook_references_and_names() {
        let context = Context {
            worksheet_names: Vec::new().into(),
            supporting_links: vec![SupportingLink::ExternalWorkbook(0)].into(),
            external_sheets: vec![ExternalSheet {
                external_link: 0,
                first_sheet: 0,
                last_sheet: 0,
            }]
            .into(),
            external_books: vec![ExternalBook {
                metadata: Link::workbook(
                    "Book.xlsx",
                    vec!["Data Sheet".to_string()],
                    vec!["Rate".to_string()],
                )
                .unwrap(),
            }]
            .into(),
            defined_names: Vec::new().into(),
            tables: Vec::new().into(),
            pivot_views: Vec::new().into(),
            pivot_name_scopes: Vec::new().into(),
            active_pivot_scope: None,
            current_sheet: None,
        };

        let reference = [0x5A, 0, 0, 0, 0, 0, 0, 0, 0];
        let tokens = Parser::new(&reference).parse().unwrap();
        assert_eq!(
            Compiler::try_tokens_to_string_with_resolution(&tokens, &context).unwrap(),
            "'[Book.xlsx]Data Sheet'!$A$1"
        );

        let name = [0x59, 0, 0, 1, 0, 0, 0];
        let tokens = Parser::new(&name).parse().unwrap();
        assert_eq!(
            Compiler::try_tokens_to_string_with_resolution(&tokens, &context).unwrap(),
            "'[Book.xlsx]'!Rate"
        );

        let invalid_name = [0x59, 0, 0, 2, 0, 0, 0];
        let tokens = Parser::new(&invalid_name).parse().unwrap();
        assert!(matches!(
            Compiler::try_tokens_to_string_with_resolution(&tokens, &context),
            Err(crate::formula::Error::InvalidFormula(_))
        ));
        assert!(matches!(
            Parser::new(&[0x59, 0, 0, 0, 0, 0, 0]).parse(),
            Err(crate::formula::Error::InvalidFormula(_))
        ));
    }

    #[test]
    fn scalar_errors_compile_and_roundtrip_canonically() {
        for &(literal, code) in text::FORMULA_ERRORS {
            let compiled = text::Compiler::compile(&literal.to_ascii_lowercase()).unwrap();
            assert_eq!(compiled.rgce, vec![ptg_types::PTG_ERR, code]);
            let tokens = Parser::new(&compiled.rgce).parse().unwrap();
            assert_eq!(Compiler::try_tokens_to_string(&tokens).unwrap(), literal);
        }

        let compiled = text::Compiler::compile("#DIV/0!+1").unwrap();
        let tokens = Parser::new(&compiled.rgce).parse().unwrap();
        assert_eq!(
            Compiler::try_tokens_to_string(&tokens).unwrap(),
            "(#DIV/0!+1)"
        );
        assert!(text::Compiler::compile("#SPILL!").is_err());
    }

    #[test]
    fn parser_rejects_invalid_scalar_boolean_and_error_values() {
        assert!(matches!(
            Parser::new(&[ptg_types::PTG_BOOL, 2]).parse(),
            Err(crate::formula::Error::InvalidFormula(_))
        ));
        assert!(matches!(
            Parser::new(&[ptg_types::PTG_ERR, 1]).parse(),
            Err(crate::formula::Error::InvalidFormula(_))
        ));
    }

    #[test]
    fn parser_consumes_attribute_payloads_and_converts_attr_sum() {
        let attr_sum = [ptg_types::PTG_INT, 1, 0, ptg_types::PTG_ATTR, 0x10, 0, 0];
        let tokens = Parser::new(&attr_sum).parse().unwrap();
        assert_eq!(Compiler::try_tokens_to_string(&tokens).unwrap(), "SUM(1)");

        let attr_choose = [ptg_types::PTG_ATTR, 0x04, 0x00, 0x00, 0x02, 0x00];
        assert_eq!(
            Parser::new(&attr_choose).parse().unwrap(),
            vec![Token::Attribute(0x04)]
        );

        assert!(matches!(
            Parser::new(&attr_choose[..5]).parse(),
            Err(crate::formula::Error::InvalidFormula(_))
        ));
    }

    #[test]
    fn parser_decodes_typed_array_ancillary_values() {
        let mut rgce = vec![0x40];
        rgce.extend_from_slice(&[0; 14]);
        let mut rgcb = Vec::new();
        rgcb.extend_from_slice(&2_u32.to_le_bytes());
        rgcb.extend_from_slice(&2_u32.to_le_bytes());
        rgcb.push(0x00);
        rgcb.extend_from_slice(&1_f64.to_le_bytes());
        rgcb.extend_from_slice(&[0x01, 0x01, 0x00, b'x', 0x00]);
        rgcb.extend_from_slice(&[0x02, 0x01]);
        rgcb.extend_from_slice(&[0x04, 0x07, 0x00, 0x00, 0x00]);

        let tokens = Parser::with_extra(&rgce, &rgcb).parse().unwrap();
        assert_eq!(
            tokens,
            vec![Token::Array {
                rows: 2,
                cols: 2,
                values: vec![
                    ArrayValue::Number(1.0),
                    ArrayValue::String("x".to_string()),
                    ArrayValue::Bool(true),
                    ArrayValue::Error(0x07),
                ],
            }]
        );
        assert_eq!(
            Compiler::try_tokens_to_string(&tokens).unwrap(),
            "{1,\"x\";TRUE,#DIV/0!}"
        );
    }

    #[test]
    fn parser_rejects_malformed_array_ancillary_data_without_large_allocation() {
        let mut rgce = vec![0x40];
        rgce.extend_from_slice(&[0; 14]);
        let mut impossible = Vec::new();
        impossible.extend_from_slice(&1_048_576_u32.to_le_bytes());
        impossible.extend_from_slice(&16_384_u32.to_le_bytes());
        assert!(matches!(
            Parser::with_extra(&rgce, &impossible).parse(),
            Err(crate::formula::Error::InvalidFormula(_))
        ));

        let mut invalid_bool = Vec::new();
        invalid_bool.extend_from_slice(&1_u32.to_le_bytes());
        invalid_bool.extend_from_slice(&1_u32.to_le_bytes());
        invalid_bool.extend_from_slice(&[0x02, 0x02]);
        assert!(matches!(
            Parser::with_extra(&rgce, &invalid_bool).parse(),
            Err(crate::formula::Error::InvalidFormula(_))
        ));

        let mut invalid_number = Vec::new();
        invalid_number.extend_from_slice(&1_u32.to_le_bytes());
        invalid_number.extend_from_slice(&1_u32.to_le_bytes());
        invalid_number.push(0x00);
        invalid_number.extend_from_slice(&f64::NEG_INFINITY.to_le_bytes());
        assert!(matches!(
            Parser::with_extra(&rgce, &invalid_number).parse(),
            Err(crate::formula::Error::InvalidFormula(_))
        ));
    }

    #[test]
    fn compiler_emits_and_roundtrips_array_constants() {
        let formula = text::Compiler::compile("SUM({1,\"x\";TRUE,#N/A})").unwrap();
        assert_eq!(
            &formula.rgce[..15],
            &[0x40, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0]
        );
        assert_eq!(&formula.rgcb[..8], &[2, 0, 0, 0, 2, 0, 0, 0]);
        let tokens = Parser::with_extra(&formula.rgce, &formula.rgcb)
            .parse()
            .unwrap();
        assert_eq!(
            Compiler::try_tokens_to_string(&tokens).unwrap(),
            "SUM({1,\"x\";TRUE,#N/A})"
        );

        assert!(matches!(
            text::Compiler::compile_shared("SUM({1,2})", 0, 0),
            Err(Error::InvalidFormula(_))
        ));
    }

    #[test]
    fn parser_consumes_memory_area_and_cached_ranges() {
        let left = text::Compiler::compile("A1").unwrap().rgce;
        let right = text::Compiler::compile("B2").unwrap().rgce;
        let expression_len = left.len() + right.len() + 1;
        let mut rgce = vec![0x46, 0, 0, 0, 0];
        rgce.extend_from_slice(&(expression_len as u16).to_le_bytes());
        rgce.extend_from_slice(&left);
        rgce.extend_from_slice(&right);
        rgce.push(ptg_types::PTG_UNION);

        let mut rgcb = Vec::new();
        rgcb.extend_from_slice(&1_u32.to_le_bytes());
        rgcb.extend_from_slice(&0_u32.to_le_bytes());
        rgcb.extend_from_slice(&1_u32.to_le_bytes());
        rgcb.extend_from_slice(&0_u32.to_le_bytes());
        rgcb.extend_from_slice(&1_u32.to_le_bytes());
        let tokens = Parser::with_extra(&rgce, &rgcb).parse().unwrap();
        assert!(matches!(
            &tokens[0],
            Token::Memory {
                kind: MemoryKind::Area,
                expression_bytes: 15,
                cached_ranges,
            } if cached_ranges == &vec![[0, 1, 0, 1]]
        ));
        assert_eq!(Compiler::try_tokens_to_string(&tokens).unwrap(), "(A1,B2)");
    }

    #[test]
    fn parser_rejects_truncated_memory_metadata() {
        let rgce = [0x46, 0, 0, 0, 0, 0, 0];
        let mut rgcb = Vec::new();
        rgcb.extend_from_slice(&u32::MAX.to_le_bytes());
        assert!(matches!(
            Parser::with_extra(&rgce, &rgcb).parse(),
            Err(crate::formula::Error::InvalidFormula(_))
        ));

        let oversized_expression = [0x49, 0x01, 0x00];
        assert!(matches!(
            Parser::new(&oversized_expression).parse(),
            Err(crate::formula::Error::InvalidFormula(_))
        ));
    }

    #[test]
    fn shared_formula_uses_relative_tokens_and_expands_per_target_cell() {
        // Real shared-formula pattern from POI bug66682.xlsb: the C3:C10
        // formula group references the cell one column earlier.
        let formula = text::Compiler::compile_shared("B3", 2, 2).unwrap();
        assert_eq!(formula.rgce, vec![0x4C, 0x00, 0x00, 0x00, 0x00, 0xFF, 0xFF]);

        let anchor_tokens = Parser::with_base_cell(&formula.rgce, 2, 2).parse().unwrap();
        assert_eq!(
            Compiler::try_tokens_to_string(&anchor_tokens).unwrap(),
            "B3"
        );
        let follower_tokens = Parser::with_base_cell(&formula.rgce, 3, 2).parse().unwrap();
        assert_eq!(
            Compiler::try_tokens_to_string(&follower_tokens).unwrap(),
            "B4"
        );
    }

    #[test]
    fn parses_real_poi_shared_formula_definition_losslessly() {
        // BrtShrFmla from POI bug66682.xlsb: C3:C10 refers one column left.
        let bytes = [
            0x02, 0x00, 0x00, 0x00, 0x09, 0x00, 0x00, 0x00, 0x02, 0x00, 0x00, 0x00, 0x02, 0x00,
            0x00, 0x00, 0x07, 0x00, 0x00, 0x00, 0x4C, 0x00, 0x00, 0x00, 0x00, 0xFF, 0xFF, 0x00,
            0x00, 0x00, 0x00,
        ];
        let group = Group::parse_shared(&bytes).unwrap();
        assert_eq!(group.kind, GroupKind::Shared);
        assert_eq!(group.range.to_a1(), "C3:C10");
        assert_eq!(group.to_record_data().unwrap(), bytes);

        let tokens = Parser::with_base_cell(&group.formula.rgce, 9, 2)
            .parse()
            .unwrap();
        assert_eq!(Compiler::try_tokens_to_string(&tokens).unwrap(), "B10");
    }

    #[test]
    fn parses_real_poi_array_formula_definition_losslessly() {
        // BrtArrFmla from POI bug66682.xlsb. Its PtgName is retained even
        // when the standalone formula converter cannot resolve that name.
        let bytes = [
            0x08, 0x00, 0x00, 0x00, 0x08, 0x00, 0x00, 0x00, 0x02, 0x00, 0x00, 0x00, 0x02, 0x00,
            0x00, 0x00, 0x01, 0x09, 0x00, 0x00, 0x00, 0x23, 0x02, 0x00, 0x00, 0x00, 0x42, 0x01,
            0xFF, 0x00, 0x00, 0x00, 0x00, 0x00,
        ];
        let group = Group::parse_array(&bytes).unwrap();
        assert_eq!(group.kind, GroupKind::Array);
        assert_eq!(group.range.to_a1(), "C9:C9");
        assert!(group.always_calculate);
        assert_eq!(group.to_record_data().unwrap(), bytes);
    }

    #[test]
    fn rejects_malformed_ptg_exp_and_array_flags() {
        let malformed = ParsedFormula {
            rgce: vec![ptg_types::PTG_EXP, 0, 0],
            rgcb: vec![],
        };
        assert!(matches!(
            malformed.exp_cell(),
            Err(crate::formula::Error::InvalidFormula(_))
        ));

        let mut array = Group {
            kind: GroupKind::Array,
            range: Range::new(0, 0, 0, 0).unwrap(),
            formula: text::Compiler::compile("1+1").unwrap(),
            always_calculate: false,
        }
        .to_record_data()
        .unwrap();
        array[16] = 0x80;
        assert!(matches!(
            Group::parse_array(&array),
            Err(crate::formula::Error::InvalidFormula(_))
        ));
    }

    #[test]
    fn shared_formula_preserves_mixed_absolute_references() {
        let formula = text::Compiler::compile_shared("$A1+B$2", 4, 3).unwrap();
        let tokens = Parser::with_base_cell(&formula.rgce, 7, 5).parse().unwrap();
        assert_eq!(
            Compiler::try_tokens_to_string(&tokens).unwrap(),
            "($A4+D$2)"
        );
    }

    #[test]
    fn cell_parsed_formula_accepts_the_empty_token_streams_excel_writes() {
        // `cce == 0`, `cb == 0`: no tokens and no ancillary data.
        let empty = [0_u8; 8];
        let (formula, consumed) = ParsedFormula::parse(&empty).unwrap();
        assert!(formula.rgce.is_empty());
        assert!(formula.rgcb.is_empty());
        assert_eq!(consumed, empty.len());
        // The empty stream must survive a write so reading cannot lose it.
        assert_eq!(formula.to_bytes().unwrap(), empty);
    }

    #[test]
    fn cell_parsed_formula_rejects_oversized_token_streams() {
        let mut oversized = Vec::new();
        oversized.extend_from_slice(&((MAX_CELL_FORMULA_BYTES as u32) + 1).to_le_bytes());
        oversized.extend_from_slice(&[0; 4]);
        assert!(matches!(
            ParsedFormula::parse(&oversized),
            Err(crate::formula::Error::InvalidFormula(_))
        ));
    }

    #[test]
    fn truncated_token_is_an_error_instead_of_becoming_unknown_bytes() {
        let error = Parser::new(&[0x44, 0x01]).parse().unwrap_err();
        assert!(matches!(error, crate::formula::Error::InvalidFormula(_)));
    }

    fn resident_table_reference(row_type: TableRowType, columns: TableColumns) -> Token {
        Token::TableReference(TableReference {
            sheet_index: 0,
            row_type: Some(row_type),
            columns: Some(columns),
            square_bracket_space: false,
            comma_space: false,
            data_type: TableDataType::Reference,
            invalid: false,
            list_index: Some(7),
            external: None,
        })
    }

    fn table_context() -> Context {
        Context {
            worksheet_names: vec!["Data".to_string()].into(),
            supporting_links: vec![SupportingLink::SelfWorkbook].into(),
            external_sheets: vec![ExternalSheet {
                external_link: 0,
                first_sheet: 0,
                last_sheet: 0,
            }]
            .into(),
            external_books: Vec::new().into(),
            defined_names: Vec::new().into(),
            tables: vec![
                Definition::try_new(
                    7,
                    0,
                    "Sales",
                    vec![
                        "Item".to_string(),
                        "Price]Gross".to_string(),
                        "@Tag".to_string(),
                    ],
                )
                .unwrap(),
            ]
            .into(),
            pivot_views: Vec::new().into(),
            pivot_name_scopes: Vec::new().into(),
            active_pivot_scope: None,
            current_sheet: Some(0),
        }
    }

    #[test]
    fn resolves_resident_structured_references_faithfully() {
        let context = table_context();
        for (row_type, expected) in [
            (TableRowType::Data, "Sales"),
            (TableRowType::All, "Sales[#All]"),
            (TableRowType::Headers, "Sales[#Headers]"),
            (TableRowType::DataAlternate, "Sales[#Data]"),
            (TableRowType::DataAndHeaders, "Sales[[#Headers],[#Data]]"),
            (TableRowType::Totals, "Sales[#Totals]"),
            (TableRowType::DataAndTotals, "Sales[[#Data],[#Totals]]"),
            (TableRowType::Current, "Sales[#This Row]"),
        ] {
            let token = resident_table_reference(row_type, TableColumns::All);
            assert_eq!(
                Compiler::try_tokens_to_string_with_resolution(&[token], &context).unwrap(),
                expected
            );
        }

        let token = resident_table_reference(
            TableRowType::Current,
            TableColumns::Range { first: 1, last: 2 },
        );
        assert_eq!(
            Compiler::try_tokens_to_string_with_resolution(&[token], &context).unwrap(),
            "Sales[[#This Row],[Price']Gross]:['@Tag]]"
        );

        let mut spaced = resident_table_reference(TableRowType::Current, TableColumns::One(0));
        let Token::TableReference(reference) = &mut spaced else {
            unreachable!()
        };
        reference.square_bracket_space = true;
        reference.comma_space = true;
        assert_eq!(
            Compiler::try_tokens_to_string_with_resolution(&[spaced], &context).unwrap(),
            "Sales[ [#This Row], [Item] ]"
        );
    }

    #[test]
    fn resolves_nonresident_structured_references_with_external_prefix() {
        let context = Context {
            worksheet_names: Vec::new().into(),
            supporting_links: vec![SupportingLink::ExternalWorkbook(0)].into(),
            external_sheets: vec![ExternalSheet {
                external_link: 0,
                first_sheet: 0,
                last_sheet: 0,
            }]
            .into(),
            external_books: vec![ExternalBook {
                metadata: Link::workbook("Book.xlsx", vec!["Data Sheet".to_string()], Vec::new())
                    .unwrap(),
            }]
            .into(),
            defined_names: Vec::new().into(),
            tables: Vec::new().into(),
            pivot_views: Vec::new().into(),
            pivot_name_scopes: Vec::new().into(),
            active_pivot_scope: None,
            current_sheet: None,
        };
        let token = Token::TableReference(TableReference {
            sheet_index: 0,
            row_type: None,
            columns: None,
            square_bracket_space: false,
            comma_space: false,
            data_type: TableDataType::Reference,
            invalid: false,
            list_index: None,
            external: Some(ExternalTableReference {
                table: "Remote".to_string(),
                row_type: TableRowType::Totals,
                columns: TableNamedColumns::One("Amount".to_string()),
            }),
        });
        assert_eq!(
            Compiler::try_tokens_to_string_with_resolution(&[token], &context).unwrap(),
            "'[Book.xlsx]Data Sheet'!Remote[[#Totals],[Amount]]"
        );
    }

    #[test]
    fn structured_reference_resolution_rejects_ambiguous_and_invalid_metadata() {
        assert!(Definition::try_new(0, 0, "Sales", vec!["A".into()]).is_err());
        assert!(Definition::try_new(1, 0, "_xlBad", vec!["A".into()]).is_err());
        assert!(Definition::try_new(1, 0, "Sales", Vec::new()).is_err());
        assert!(Definition::try_new(1, 0, "Sales", vec!["A".into(), "a".into()]).is_err());

        let token = resident_table_reference(TableRowType::Data, TableColumns::One(3));
        assert!(Compiler::try_tokens_to_string(std::slice::from_ref(&token)).is_err());
        assert!(
            Compiler::try_tokens_to_string_with_resolution(&[token], &table_context()).is_err()
        );

        let mut missing = table_context();
        missing.tables = Vec::new().into();
        assert!(
            Compiler::try_tokens_to_string_with_resolution(
                &[resident_table_reference(
                    TableRowType::Data,
                    TableColumns::All,
                )],
                &missing,
            )
            .is_err()
        );

        let mut ambiguous = table_context();
        ambiguous.tables = vec![
            ambiguous.tables[0].clone(),
            Definition::try_new(7, 0, "Other", vec!["A".into()]).unwrap(),
        ]
        .into();
        assert!(
            Compiler::try_tokens_to_string_with_resolution(
                &[resident_table_reference(
                    TableRowType::Data,
                    TableColumns::All,
                )],
                &ambiguous,
            )
            .is_err()
        );

        let mut wrong_sheet = table_context();
        wrong_sheet.external_sheets = vec![ExternalSheet {
            external_link: 0,
            first_sheet: 0,
            last_sheet: 1,
        }]
        .into();
        assert!(
            Compiler::try_tokens_to_string_with_resolution(
                &[resident_table_reference(
                    TableRowType::Data,
                    TableColumns::All,
                )],
                &wrong_sheet,
            )
            .is_err()
        );
    }
}

#[cfg(test)]
mod structured_reference_compiler_tests {
    use super::*;

    fn tables() -> Vec<Definition> {
        vec![
            Definition::try_new(
                7,
                0,
                "Sales",
                vec![
                    "Item".to_string(),
                    "Price]Gross".to_string(),
                    "@Tag".to_string(),
                    "Amount".to_string(),
                ],
            )
            .unwrap(),
        ]
    }

    #[test]
    fn compiles_parses_and_stringifies_resident_and_nonresident_structured_references() {
        let worksheet_names = vec!["Data".to_string()];
        let tables = tables();
        let defined_names = Vec::new();
        let supporting_links = vec![
            SupportingLink::ExternalWorkbook(0),
            SupportingLink::SelfWorkbook,
        ];
        let external_sheets = vec![
            ExternalSheet {
                external_link: 0,
                first_sheet: 0,
                last_sheet: 0,
            },
            ExternalSheet {
                external_link: 1,
                first_sheet: 0,
                last_sheet: 0,
            },
            ExternalSheet {
                external_link: 1,
                first_sheet: 0,
                last_sheet: 0,
            },
        ];
        let external_books = vec![ExternalBook {
            metadata: Link::workbook("Book.xlsx", vec!["Data Sheet".to_string()], Vec::new())
                .unwrap(),
        }];
        let sheet_ranges = std::cell::RefCell::new(Vec::new());
        let compile_context = CompilationContext {
            worksheet_names: &worksheet_names,
            defined_names: &defined_names,
            tables: &tables,
            supporting_links: &supporting_links,
            external_sheets: &external_sheets,
            external_books: &external_books,
            sheet_ranges: &sheet_ranges,
            current_sheet: 0,
        };
        let resolution_context = Context {
            worksheet_names: worksheet_names.clone().into(),
            supporting_links: supporting_links.clone().into(),
            external_sheets: external_sheets.clone().into(),
            external_books: external_books.clone().into(),
            defined_names: Vec::new().into(),
            tables: tables.clone().into(),
            pivot_views: Vec::new().into(),
            pivot_name_scopes: Vec::new().into(),
            active_pivot_scope: None,
            current_sheet: Some(0),
        };

        for source in [
            "Sales",
            "Sales[Item]",
            "Sales[#All]",
            "Sales[[#Headers],[#Data]]",
            "Sales[[#Data],[#Totals]]",
            "Sales[[#This Row],[Price']Gross]:['@Tag]]",
            "Sales[ [#This Row], [Item] ]",
            "'[Book.xlsx]Data Sheet'!Remote[[#Totals],[Amount]]",
        ] {
            let compiled = text::Compiler::compile_with_context(source, &compile_context).unwrap();
            let tokens = Parser::with_extra(&compiled.rgce, &compiled.rgcb)
                .parse()
                .unwrap();
            assert_eq!(
                Compiler::try_tokens_to_string_with_resolution(&tokens, &resolution_context,)
                    .unwrap(),
                source
            );
            assert!(matches!(tokens.as_slice(), [Token::TableReference(_)]));
        }
    }

    #[test]
    fn structured_reference_compiler_rejects_ambiguous_missing_and_unrepresentable_inputs() {
        let worksheet_names = vec!["Data".to_string(), "Other".to_string()];
        let defined_names = Vec::new();
        let supporting_links = Vec::new();
        let external_sheets = Vec::new();
        let external_books = Vec::new();
        let sheet_ranges = std::cell::RefCell::new(Vec::new());
        let base_tables = tables();
        let context = CompilationContext {
            worksheet_names: &worksheet_names,
            defined_names: &defined_names,
            tables: &base_tables,
            supporting_links: &supporting_links,
            external_sheets: &external_sheets,
            external_books: &external_books,
            sheet_ranges: &sheet_ranges,
            current_sheet: 0,
        };
        for source in [
            "Missing[Item]",
            "Sales[Missing]",
            "Sales[[Amount]:[Item]]",
            "Sales[[Item],[Amount]]",
            "Sales[[#Headers],[#Totals]]",
            "Sales[ Item]",
            "Sales[Item ]",
            "Sales[Bad'x]",
            "'[Book.xlsx]Data Sheet'!Remote[Amount]",
            "Other!Sales[Item]",
        ] {
            assert!(
                text::Compiler::compile_with_context(source, &context).is_err(),
                "{source} unexpectedly compiled"
            );
        }

        let ambiguous = vec![
            base_tables[0].clone(),
            Definition::try_new(8, 0, "sales", vec!["Item".to_string()]).unwrap(),
        ];
        let ambiguous_context = CompilationContext {
            tables: &ambiguous,
            ..context
        };
        assert!(text::Compiler::compile_with_context("Sales[Item]", &ambiguous_context).is_err());

        let wrong_sheet =
            vec![Definition::try_new(7, 1, "Sales", vec!["Item".to_string()]).unwrap()];
        let wrong_sheet_context = CompilationContext {
            tables: &wrong_sheet,
            ..context
        };
        assert!(text::Compiler::compile_with_context("Sales[Item]", &wrong_sheet_context).is_err());
    }
}

#[cfg(test)]
mod pivot_name_resolution_tests {
    use super::*;

    fn references() -> Vec<Name> {
        vec![
            Name::Field {
                name: "Sales".to_string(),
                aggregation: None,
            },
            Name::Field {
                name: "Gross Profit".to_string(),
                aggregation: Some(Aggregation::Average),
            },
            Name::Item {
                field_name: "Region".to_string(),
                item: Item::Name("North".to_string()),
            },
            Name::Item {
                field_name: "Sales Region".to_string(),
                item: Item::Name("O'Brien".to_string()),
            },
            Name::Item {
                field_name: "Quarter".to_string(),
                item: Item::AbsolutePosition(2),
            },
            Name::Item {
                field_name: "Quarter".to_string(),
                item: Item::RelativePosition(1),
            },
            Name::Item {
                field_name: "Quarter".to_string(),
                item: Item::RelativePosition(-1),
            },
        ]
    }

    fn scope() -> Scope {
        Scope::try_new(7, 1, "Sales Pivot".to_string(), references()).unwrap()
    }

    fn pivot_context() -> Context {
        Context {
            worksheet_names: vec!["Data".to_string(), "Report".to_string()].into(),
            supporting_links: Vec::new().into(),
            external_sheets: Vec::new().into(),
            external_books: Vec::new().into(),
            defined_names: Vec::new().into(),
            tables: Vec::new().into(),
            pivot_views: vec![View::try_new(7, 1, "Sales Pivot".to_string()).unwrap()].into(),
            pivot_name_scopes: vec![scope()].into(),
            active_pivot_scope: None,
            current_sheet: Some(1),
        }
        .for_pivot_formula(scope())
        .unwrap()
    }

    fn render(index: u32, context: &Context) -> Result<String> {
        Ok(Compiler::try_tokens_to_string_with_resolution(
            &[Token::PivotName(index)],
            context,
        )?)
    }

    #[test]
    fn resolves_pivot_names_to_faithful_field_and_item_syntax() {
        let context = pivot_context();
        assert_eq!(render(0, &context).unwrap(), "Sales");
        assert_eq!(render(1, &context).unwrap(), "AVERAGE('Gross Profit')");
        assert_eq!(render(2, &context).unwrap(), "Region[North]");
        assert_eq!(render(3, &context).unwrap(), "'Sales Region'['O''Brien']");
        assert_eq!(render(4, &context).unwrap(), "Quarter[2]");
        assert_eq!(render(5, &context).unwrap(), "Quarter[+1]");
        assert_eq!(render(6, &context).unwrap(), "Quarter[-1]");
    }

    #[test]
    fn rejects_missing_ambiguous_cross_sheet_and_out_of_range_pivot_metadata() {
        assert!(Compiler::try_tokens_to_string(&[Token::PivotName(0)]).is_err());

        let mut context = pivot_context();
        assert!(render(7, &context).is_err());
        context.current_sheet = Some(0);
        assert!(render(0, &context).is_err());

        let mut context = pivot_context();
        context.pivot_views = vec![
            View::try_new(7, 1, "Sales Pivot".to_string()).unwrap(),
            View::try_new(7, 1, "sales pivot".to_string()).unwrap(),
        ]
        .into();
        assert!(render(0, &context).is_err());

        let mut context = pivot_context();
        context.pivot_name_scopes = vec![scope(), scope()].into();
        assert!(render(0, &context).is_err());

        let mut context = pivot_context();
        context.pivot_name_scopes =
            vec![Scope::try_new(8, 1, "Sales Pivot".to_string(), references()).unwrap()].into();
        assert!(render(0, &context).is_err());
    }

    #[test]
    fn validates_bounded_pivot_names_and_positions() {
        assert!(View::try_new(1, 0, String::new()).is_err());
        assert!(
            Scope::try_new(
                1,
                0,
                "Pivot".to_string(),
                vec![Name::Item {
                    field_name: "Quarter".to_string(),
                    item: Item::AbsolutePosition(0),
                }],
            )
            .is_err()
        );
        assert!(
            Scope::try_new(
                1,
                0,
                "Pivot".to_string(),
                vec![Name::Item {
                    field_name: "Quarter".to_string(),
                    item: Item::RelativePosition(0),
                }],
            )
            .is_err()
        );
        assert!(
            Scope::try_new(
                1,
                0,
                "Pivot".to_string(),
                vec![Name::Field {
                    name: "bad\0field".to_string(),
                    aggregation: None,
                }],
            )
            .is_err()
        );
    }
}
