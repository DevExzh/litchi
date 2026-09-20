//! Integration coverage for inert OpenFormula 1.4 name recognition.
//!
//! The catalog below is a checked-in extraction of the standard function
//! headings in ODF 1.4 Part 4, chapter 6. These tests inspect recognition and
//! lexical tokens only; they deliberately make no evaluation or arity claim.

use litchi_ods::codec::formula::{
    Formula, FormulaParser, Token, extract_functions, is_valid_function,
};

/// Normative ODF 1.4 Part 4 chapter 6 function names, extracted independently from the specification headings.
const STANDARD_OPENFORMULA_FUNCTION_NAMES: &[&str] = &[
    "MDETERM",
    "MINVERSE",
    "MMULT",
    "MUNIT",
    "TRANSPOSE",
    "BITAND",
    "BITLSHIFT",
    "BITOR",
    "BITRSHIFT",
    "BITXOR",
    "FINDB",
    "LEFTB",
    "LENB",
    "MIDB",
    "REPLACEB",
    "RIGHTB",
    "SEARCHB",
    "COMPLEX",
    "IMABS",
    "IMAGINARY",
    "IMARGUMENT",
    "IMCONJUGATE",
    "IMCOS",
    "IMCOSH",
    "IMCOT",
    "IMCSC",
    "IMCSCH",
    "IMDIV",
    "IMEXP",
    "IMLN",
    "IMLOG10",
    "IMLOG2",
    "IMPOWER",
    "IMPRODUCT",
    "IMREAL",
    "IMSIN",
    "IMSINH",
    "IMSEC",
    "IMSECH",
    "IMSQRT",
    "IMSUB",
    "IMSUM",
    "IMTAN",
    "DAVERAGE",
    "DCOUNT",
    "DCOUNTA",
    "DGET",
    "DMAX",
    "DMIN",
    "DPRODUCT",
    "DSTDEV",
    "DSTDEVP",
    "DSUM",
    "DVAR",
    "DVARP",
    "DATE",
    "DATEDIF",
    "DATEVALUE",
    "DAY",
    "DAYS",
    "DAYS360",
    "EASTERSUNDAY",
    "EDATE",
    "EOMONTH",
    "HOUR",
    "ISOWEEKNUM",
    "MINUTE",
    "MONTH",
    "NETWORKDAYS",
    "NOW",
    "SECOND",
    "TIME",
    "TIMEVALUE",
    "TODAY",
    "WEEKDAY",
    "WEEKNUM",
    "WORKDAY",
    "YEAR",
    "YEARFRAC",
    "DDE",
    "HYPERLINK",
    "ACCRINT",
    "ACCRINTM",
    "AMORLINC",
    "COUPDAYBS",
    "COUPDAYS",
    "COUPDAYSNC",
    "COUPNCD",
    "COUPNUM",
    "COUPPCD",
    "CUMIPMT",
    "CUMPRINC",
    "DB",
    "DDB",
    "DISC",
    "DOLLARDE",
    "DOLLARFR",
    "DURATION",
    "EFFECT",
    "FV",
    "FVSCHEDULE",
    "INTRATE",
    "IPMT",
    "IRR",
    "ISPMT",
    "MDURATION",
    "MIRR",
    "NOMINAL",
    "NPER",
    "NPV",
    "ODDFPRICE",
    "ODDFYIELD",
    "ODDLPRICE",
    "ODDLYIELD",
    "PDURATION",
    "PMT",
    "PPMT",
    "PRICE",
    "PRICEDISC",
    "PRICEMAT",
    "PV",
    "RATE",
    "RECEIVED",
    "RRI",
    "SLN",
    "SYD",
    "TBILLEQ",
    "TBILLPRICE",
    "TBILLYIELD",
    "VDB",
    "XIRR",
    "XNPV",
    "YIELD",
    "YIELDDISC",
    "YIELDMAT",
    "AREAS",
    "CELL",
    "COLUMN",
    "COLUMNS",
    "COUNT",
    "COUNTA",
    "COUNTBLANK",
    "COUNTIF",
    "COUNTIFS",
    "ERROR.TYPE",
    "FORMULA",
    "INFO",
    "ISBLANK",
    "ISERR",
    "ISERROR",
    "ISEVEN",
    "ISFORMULA",
    "ISLOGICAL",
    "ISNA",
    "ISNONTEXT",
    "ISNUMBER",
    "ISODD",
    "ISREF",
    "ISTEXT",
    "N",
    "NA",
    "NUMBERVALUE",
    "ROW",
    "ROWS",
    "SHEET",
    "SHEETS",
    "TYPE",
    "VALUE",
    "ADDRESS",
    "CHOOSE",
    "GETPIVOTDATA",
    "HLOOKUP",
    "INDEX",
    "INDIRECT",
    "LOOKUP",
    "MATCH",
    "MULTIPLE.OPERATIONS",
    "OFFSET",
    "VLOOKUP",
    "AND",
    "FALSE",
    "IF",
    "IFERROR",
    "IFNA",
    "NOT",
    "OR",
    "TRUE",
    "XOR",
    "ABS",
    "ACOS",
    "ACOSH",
    "ACOT",
    "ACOTH",
    "ASIN",
    "ASINH",
    "ATAN",
    "ATAN2",
    "ATANH",
    "BESSELI",
    "BESSELJ",
    "BESSELK",
    "BESSELY",
    "COMBIN",
    "COMBINA",
    "CONVERT",
    "COS",
    "COSH",
    "COT",
    "COTH",
    "CSC",
    "CSCH",
    "DEGREES",
    "DELTA",
    "ERF",
    "ERFC",
    "EUROCONVERT",
    "EVEN",
    "EXP",
    "FACT",
    "FACTDOUBLE",
    "GAMMA",
    "GAMMALN",
    "GCD",
    "GESTEP",
    "LCM",
    "LN",
    "LOG",
    "LOG10",
    "MOD",
    "MULTINOMIAL",
    "ODD",
    "PI",
    "POWER",
    "PRODUCT",
    "QUOTIENT",
    "RADIANS",
    "RAND",
    "RANDBETWEEN",
    "SEC",
    "SERIESSUM",
    "SIGN",
    "SIN",
    "SINH",
    "SECH",
    "SQRT",
    "SQRTPI",
    "SUBTOTAL",
    "SUM",
    "SUMIF",
    "SUMIFS",
    "SUMPRODUCT",
    "SUMSQ",
    "SUMX2MY2",
    "SUMX2PY2",
    "SUMXMY2",
    "TAN",
    "TANH",
    "CEILING",
    "INT",
    "FLOOR",
    "MROUND",
    "ROUND",
    "ROUNDDOWN",
    "ROUNDUP",
    "TRUNC",
    "AVEDEV",
    "AVERAGE",
    "AVERAGEA",
    "AVERAGEIF",
    "AVERAGEIFS",
    "BETADIST",
    "BETAINV",
    "BINOM.DIST.RANGE",
    "BINOMDIST",
    "LEGACY.CHIDIST",
    "CHISQDIST",
    "LEGACY.CHIINV",
    "CHISQINV",
    "LEGACY.CHITEST",
    "CONFIDENCE",
    "CORREL",
    "COVAR",
    "CRITBINOM",
    "DEVSQ",
    "EXPONDIST",
    "FDIST",
    "LEGACY.FDIST",
    "FINV",
    "LEGACY.FINV",
    "FISHER",
    "FISHERINV",
    "FORECAST",
    "FREQUENCY",
    "FTEST",
    "GAMMADIST",
    "GAMMAINV",
    "GAUSS",
    "GEOMEAN",
    "GROWTH",
    "HARMEAN",
    "HYPGEOMDIST",
    "INTERCEPT",
    "KURT",
    "LARGE",
    "LINEST",
    "LOGEST",
    "LOGINV",
    "LOGNORMDIST",
    "MAX",
    "MAXA",
    "MEDIAN",
    "MIN",
    "MINA",
    "MODE",
    "NEGBINOMDIST",
    "NORMDIST",
    "NORMINV",
    "LEGACY.NORMSDIST",
    "LEGACY.NORMSINV",
    "PEARSON",
    "PERCENTILE",
    "PERCENTRANK",
    "PERMUT",
    "PERMUTATIONA",
    "PHI",
    "POISSON",
    "PROB",
    "QUARTILE",
    "RANK",
    "RSQ",
    "SKEW",
    "SKEWP",
    "SLOPE",
    "SMALL",
    "STANDARDIZE",
    "STDEV",
    "STDEVA",
    "STDEVP",
    "STDEVPA",
    "STEYX",
    "LEGACY.TDIST",
    "TINV",
    "TREND",
    "TRIMMEAN",
    "TTEST",
    "VAR",
    "VARA",
    "VARP",
    "VARPA",
    "WEIBULL",
    "ZTEST",
    "ARABIC",
    "BASE",
    "BIN2DEC",
    "BIN2HEX",
    "BIN2OCT",
    "DEC2BIN",
    "DEC2HEX",
    "DEC2OCT",
    "DECIMAL",
    "HEX2BIN",
    "HEX2DEC",
    "HEX2OCT",
    "OCT2BIN",
    "OCT2DEC",
    "OCT2HEX",
    "ROMAN",
    "ASC",
    "CHAR",
    "CLEAN",
    "CODE",
    "CONCATENATE",
    "DOLLAR",
    "EXACT",
    "FIND",
    "FIXED",
    "JIS",
    "LEFT",
    "LEN",
    "LOWER",
    "MID",
    "PROPER",
    "REPLACE",
    "REPT",
    "RIGHT",
    "SEARCH",
    "SUBSTITUTE",
    "T",
    "TEXT",
    "TRIM",
    "UNICHAR",
    "UNICODE",
    "UPPER",
];

/// Names accepted by the pre-catalog parser and retained as compatibility coverage.
const LEGACY_COMPATIBILITY_FUNCTION_NAMES: &[&str] = &[
    "ABS",
    "ACOS",
    "ACOSH",
    "ACOT",
    "ACOTH",
    "ASIN",
    "ASINH",
    "ATAN",
    "ATAN2",
    "ATANH",
    "CEILING",
    "COS",
    "COSH",
    "COT",
    "COTH",
    "DEGREES",
    "EXP",
    "FACT",
    "FLOOR",
    "INT",
    "LN",
    "LOG",
    "LOG10",
    "MOD",
    "PI",
    "POWER",
    "PRODUCT",
    "QUOTIENT",
    "RADIANS",
    "RAND",
    "ROUND",
    "ROUNDDOWN",
    "ROUNDUP",
    "SIGN",
    "SIN",
    "SINH",
    "SQRT",
    "SUM",
    "SUMIF",
    "SUMIFS",
    "SUMSQ",
    "TAN",
    "TANH",
    "TRUNC",
    "AVERAGE",
    "AVERAGEA",
    "AVERAGEIF",
    "AVERAGEIFS",
    "COUNT",
    "COUNTA",
    "COUNTBLANK",
    "COUNTIF",
    "COUNTIFS",
    "MAX",
    "MAXA",
    "MEDIAN",
    "MIN",
    "MINA",
    "MODE",
    "PERCENTILE",
    "PERCENTRANK",
    "QUARTILE",
    "RANK",
    "STDEV",
    "STDEVA",
    "STDEVP",
    "STDEVPA",
    "VAR",
    "VARA",
    "VARP",
    "VARPA",
    "AND",
    "FALSE",
    "IF",
    "IFERROR",
    "IFNA",
    "NOT",
    "OR",
    "TRUE",
    "XOR",
    "CHAR",
    "CODE",
    "CONCATENATE",
    "EXACT",
    "FIND",
    "FIXED",
    "LEFT",
    "LEN",
    "LOWER",
    "MID",
    "PROPER",
    "REPLACE",
    "REPT",
    "RIGHT",
    "SEARCH",
    "SUBSTITUTE",
    "T",
    "TEXT",
    "TRIM",
    "UPPER",
    "VALUE",
    "DATE",
    "DATEVALUE",
    "DAY",
    "DAYS",
    "DAYS360",
    "HOUR",
    "MINUTE",
    "MONTH",
    "NOW",
    "SECOND",
    "TIME",
    "TIMEVALUE",
    "TODAY",
    "WEEKDAY",
    "YEAR",
    "ADDRESS",
    "CHOOSE",
    "COLUMN",
    "COLUMNS",
    "HLOOKUP",
    "INDEX",
    "INDIRECT",
    "LOOKUP",
    "MATCH",
    "OFFSET",
    "ROW",
    "ROWS",
    "VLOOKUP",
    "CELL",
    "ERROR.TYPE",
    "INFO",
    "ISBLANK",
    "ISERR",
    "ISERROR",
    "ISEVEN",
    "ISLOGICAL",
    "ISNA",
    "ISNONTEXT",
    "ISNUMBER",
    "ISODD",
    "ISREF",
    "ISTEXT",
    "N",
    "NA",
    "TYPE",
    "DB",
    "DDB",
    "FV",
    "IPMT",
    "IRR",
    "MIRR",
    "NPER",
    "NPV",
    "PMT",
    "PPMT",
    "PV",
    "RATE",
    "SLN",
    "SYD",
    "VDB",
];

fn alternating_ascii_case(name: &str) -> String {
    name.chars()
        .enumerate()
        .map(|(index, character)| {
            if index % 2 == 0 {
                character.to_ascii_lowercase()
            } else {
                character.to_ascii_uppercase()
            }
        })
        .collect()
}

fn parse_call(name: &str, argument_list: &str) -> Formula {
    let source = format!("of:={name}({argument_list})");
    FormulaParser::new(&source)
        .parse()
        .unwrap_or_else(|error| panic!("standard function {name} should lex: {error}"))
}

fn assert_function_token(formula: &Formula, expected: &str) {
    match formula.tokens.first() {
        Some(Token::Function(actual)) => assert_eq!(actual, expected),
        other => panic!("expected {expected} function token, got {other:?}"),
    }
}

#[test]
fn the_complete_odf14_catalog_is_case_insensitive() {
    assert_eq!(STANDARD_OPENFORMULA_FUNCTION_NAMES.len(), 393);
    for name in STANDARD_OPENFORMULA_FUNCTION_NAMES {
        assert!(is_valid_function(name), "canonical name rejected: {name}");
        let lower = name.to_ascii_lowercase();
        assert!(
            is_valid_function(&lower),
            "lowercase standard name rejected: {name}"
        );
        let mixed = alternating_ascii_case(name);
        assert!(
            is_valid_function(&mixed),
            "mixed-case standard name rejected: {name}"
        );
    }
}

#[test]
fn every_standard_name_is_a_function_token_when_invoked() {
    for name in STANDARD_OPENFORMULA_FUNCTION_NAMES {
        let formula = parse_call(name, "");
        assert_eq!(formula.text, format!("of:={name}()"));
        assert_function_token(&formula, name);
        let functions = extract_functions(&formula);
        assert_eq!(functions.len(), 1, "unexpected function count for {name}");
        assert_eq!(functions[0], *name);
    }
}

#[test]
fn representative_standard_families_are_lexed_without_evaluation() {
    let cases = [
        ("matrix", "MDETERM", "A1"),
        ("bitwise", "BITAND", "A1"),
        ("complex", "IMLOG10", "A1"),
        ("database", "DAVERAGE", "A1"),
        ("external DDE", "DDE", "\"never-contacted\""),
        (
            "external hyperlink",
            "HYPERLINK",
            "\"https://invalid.example/\"",
        ),
        ("base conversion", "BIN2DEC", "\"101\""),
        ("base conversion", "BIN2HEX", "\"101\""),
        ("digit-bearing math", "LOG10", "A1"),
        ("existing compatibility", "SUM", "A1"),
    ];

    for (family, name, arguments) in cases {
        let formula = parse_call(name, arguments);
        assert_function_token(&formula, name);
        assert_eq!(
            extract_functions(&formula).as_slice(),
            [name],
            "{family} function was not lexed as a call"
        );
    }
}

#[test]
fn prior_compatibility_names_remain_recognized() {
    assert_eq!(LEGACY_COMPATIBILITY_FUNCTION_NAMES.len(), 161);
    for name in LEGACY_COMPATIBILITY_FUNCTION_NAMES {
        assert!(
            STANDARD_OPENFORMULA_FUNCTION_NAMES.contains(name),
            "legacy name is missing from the standard catalog: {name}"
        );
        assert!(is_valid_function(name), "legacy name rejected: {name}");
    }
}

#[test]
fn ordinary_cell_references_remain_cell_reference_tokens() {
    for source in ["=A1", "=AA10", "=Sheet1.A1", "=$A$1", "=.A1"] {
        let formula = FormulaParser::new(source)
            .parse()
            .unwrap_or_else(|error| panic!("cell reference should parse: {source}: {error}"));
        assert!(
            matches!(formula.tokens.as_slice(), [Token::CellRef(_)]),
            "{source} was not retained as one cell-reference token: {:?}",
            formula.tokens
        );
    }

    // LOG10 is also lexically shaped like column LOG, row 10. Without an
    // opening parenthesis it remains a cell reference; the call test above
    // exercises the required function-call precedence.
    let bare = FormulaParser::new("=LOG10")
        .parse()
        .expect("bare A1-shaped identifier should remain valid");
    assert!(matches!(
        bare.tokens.as_slice(),
        [Token::CellRef(reference)]
            if reference.column == "LOG" && reference.row == 10
    ));
}

#[test]
fn unknown_malformed_and_non_catalog_names_are_rejected() {
    // The pre-existing lookup helper uppercases Unicode before consulting the
    // ASCII catalog. Preserve its two established case-compatibility results;
    // the parser's identifier lexer remains ASCII-only.
    assert!(is_valid_function("ſUM"));
    assert!(is_valid_function("ıF"));

    for name in [
        "",
        "INVALID_FUNCTION",
        "LOG10()",
        "LOG10_",
        "LOG１０",
        "ℒOG10",
        "LOG10\u{200B}",
        "ＢＩＮ２ＤＥＣ",
    ] {
        assert!(
            !is_valid_function(name),
            "non-catalog name was accepted: {name:?}"
        );
    }

    for source in [
        "=INVALID_FUNCTION()",
        "=BIN2DECC()",
        "=LOG10_()",
        "=LOG10.()",
        "=ℒOG10()",
        "=BIN２DEC()",
        "=LOG10\u{200B}()",
    ] {
        assert!(
            FormulaParser::new(source).parse().is_err(),
            "malformed or non-ASCII function name unexpectedly parsed: {source:?}"
        );
    }
}

#[test]
fn text_arguments_retain_unicode_and_require_a_terminating_quote() {
    let unicode = parse_call("UNICODE", "\"α🌟\"");
    assert_function_token(&unicode, "UNICODE");
    assert!(matches!(
        unicode.tokens.as_slice(),
        [
            Token::Function(name),
            Token::LParen,
            Token::String(value),
            Token::RParen
        ] if name == "UNICODE" && value == "α🌟"
    ));

    let doubled_quote = parse_call("CONCATENATE", "\"α\"\"🌟\"");
    assert!(matches!(
        doubled_quote.tokens.as_slice(),
        [
            Token::Function(name),
            Token::LParen,
            Token::String(value),
            Token::RParen
        ] if name == "CONCATENATE" && value == "α\"🌟"
    ));

    assert!(
        FormulaParser::new("of:=SUM(\"unterminated)")
            .parse()
            .is_err(),
        "an unterminated formula string must be rejected"
    );
}
