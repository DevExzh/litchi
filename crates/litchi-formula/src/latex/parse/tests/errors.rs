// Structurally broken input and resource limits

use super::*;

// -- error handling ------------------------------------------------------

#[test]
fn reports_unbalanced_groups() {
    assert_eq!(
        parse_err("{a"),
        LatexParseError::UnmatchedGroupOpen { position: 0 }
    );
    assert_eq!(
        parse_err("a}"),
        LatexParseError::UnmatchedGroupClose { position: 1 }
    );
}

#[test]
fn reports_unbalanced_fences() {
    assert_eq!(
        parse_err("\\left( a"),
        LatexParseError::UnmatchedLeft { position: 0 }
    );
    assert_eq!(
        parse_err("a \\right)"),
        LatexParseError::UnmatchedRight { position: 2 }
    );
    assert!(matches!(
        parse_err("\\left x \\right x"),
        LatexParseError::MissingDelimiter { .. }
    ));
}

#[test]
fn reports_unterminated_environments() {
    assert_eq!(
        parse_err("\\begin{matrix}a"),
        LatexParseError::UnclosedEnvironment {
            name: "matrix".to_string(),
            position: 0,
        }
    );
    assert!(matches!(
        parse_err("\\begin{matrix}a\\end{cases}"),
        LatexParseError::MismatchedEnvironment { .. }
    ));
    assert!(matches!(
        parse_err("\\end{matrix}"),
        LatexParseError::UnexpectedEnd { .. }
    ));
    assert!(matches!(
        parse_err("\\begin x"),
        LatexParseError::MissingEnvironmentName { .. }
    ));
}

#[test]
fn reports_a_duplicated_script() {
    assert!(matches!(
        parse_err("x^1^2"),
        LatexParseError::DuplicateScript { .. }
    ));
    assert!(matches!(
        parse_err("x_1_2"),
        LatexParseError::DuplicateScript { .. }
    ));
}

#[test]
fn reports_a_missing_argument() {
    assert!(matches!(
        parse_err("\\frac{a}"),
        LatexParseError::MissingArgument { .. }
    ));
}

#[test]
fn reports_a_script_with_nothing_to_apply_to() {
    assert!(matches!(
        parse_err("x^"),
        LatexParseError::UnexpectedEndOfInput { .. }
    ));
}

#[test]
fn reports_a_dangling_backslash() {
    assert_eq!(
        parse_err("x \\"),
        LatexParseError::IncompleteCommand { position: 2 }
    );
}

#[test]
fn refuses_input_that_nests_too_deeply() {
    let deep = "{".repeat(DEFAULT_MAX_DEPTH + 10) + &"}".repeat(DEFAULT_MAX_DEPTH + 10);
    assert!(matches!(
        parse_err(&deep),
        LatexParseError::NestingTooDeep { .. }
    ));
}

#[test]
fn honours_a_custom_depth_limit() {
    let parser = LatexParser::with_max_depth(2);
    assert_eq!(parser.max_depth(), 2);
    assert!(parser.parse("{a}").is_ok());
    assert!(parser.parse("{{{{a}}}}").is_err());
    // A zero limit is raised to one so flat input still parses.
    assert_eq!(LatexParser::with_max_depth(0).max_depth(), 1);
}

#[test]
fn accepts_empty_and_whitespace_only_input() {
    assert!(parse("").is_empty());
    assert!(parse("   \n\t ").is_empty());
    assert!(parse("% only a comment").is_empty());
}

#[test]
fn the_default_parser_matches_a_freshly_constructed_one() {
    assert_eq!(LatexParser::default(), LatexParser::new());
    assert_eq!(LatexParser::new().max_depth(), DEFAULT_MAX_DEPTH);
}
