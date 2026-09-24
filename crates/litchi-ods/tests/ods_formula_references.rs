//! Integration coverage for inert OpenFormula 1.4 references.
//!
//! The fixtures follow Part 4 section 5.8.  They inspect only the structured
//! reference token and its decoded metadata: no workbook lookup, URI
//! resolution, network access, or formula evaluation is involved.

use litchi_core::{Error, Resource};
use litchi_ods::codec::formula::reference::{
    Address, Endpoint, EndpointValue, Limits, Reference, SheetSelector, Subtable,
};
use litchi_ods::codec::formula::{
    FormulaLimits, FormulaParser, ReferenceView, Token, extract_cell_refs, extract_references,
    is_valid_function,
};

fn parse_reference(value: &str) -> Reference {
    Reference::parse(value)
        .unwrap_or_else(|error| panic!("expected valid OpenFormula reference {value:?}: {error}"))
}

fn parse_formula_reference(value: &str) -> Reference {
    let formula = FormulaParser::new(value)
        .parse()
        .unwrap_or_else(|error| panic!("expected valid formula {value:?}: {error}"));
    assert_eq!(
        formula.text, value,
        "formula text must remain byte-for-byte intact"
    );
    formula
        .tokens
        .into_iter()
        .find_map(|token| match token {
            Token::Reference(reference) => Some(*reference),
            _ => None,
        })
        .unwrap_or_else(|| panic!("formula did not produce a Reference token: {value:?}"))
}

fn assert_cell(
    endpoint: &Endpoint,
    column: &str,
    row: u32,
    column_absolute: bool,
    row_absolute: bool,
) {
    let EndpointValue::Cell(cell) = &endpoint.value else {
        panic!("expected a cell endpoint, got {:?}", endpoint.value);
    };
    assert_eq!(cell.column.label, column);
    assert_eq!(cell.column.absolute, column_absolute);
    assert_eq!(cell.row.number, row);
    assert_eq!(cell.row.absolute, row_absolute);
}

fn assert_column(endpoint: &Endpoint, label: &str, absolute: bool) {
    let EndpointValue::Column(column) = &endpoint.value else {
        panic!("expected a column endpoint, got {:?}", endpoint.value);
    };
    assert_eq!(column.label, label);
    assert_eq!(column.absolute, absolute);
}

fn assert_row(endpoint: &Endpoint, number: u32, absolute: bool) {
    let EndpointValue::Row(row) = &endpoint.value else {
        panic!("expected a row endpoint, got {:?}", endpoint.value);
    };
    assert_eq!(row.number, number);
    assert_eq!(row.absolute, absolute);
}

fn assert_explicit_sheet(endpoint: &Endpoint, name: &str, absolute: bool, quoted: bool) {
    let SheetSelector::Explicit(locator) = &endpoint.sheet else {
        panic!(
            "expected an explicit sheet selector, got {:?}",
            endpoint.sheet
        );
    };
    assert_eq!(locator.sheet.name, name);
    assert_eq!(locator.sheet.absolute, absolute);
    assert_eq!(locator.sheet.quoted, quoted);
}

fn reference_body_with_length(length: usize) -> String {
    let name_limit = Limits::default().max_name_bytes();
    let fixed_name_bytes = name_limit * 3;
    // `'source'#'sheet'.'subtable'.'subtable'.A1` has fourteen bytes of
    // syntax around its four decoded names and final cell coordinate.
    assert!(length >= fixed_name_bytes + 14);
    let last_name_bytes = length - fixed_name_bytes - 14;
    assert!(last_name_bytes <= name_limit);
    format!(
        "'{}'#'{}'.'{}'.'{}'.A1",
        "a".repeat(name_limit),
        "b".repeat(name_limit),
        "c".repeat(name_limit),
        "d".repeat(last_name_bytes),
    )
}

#[test]
fn bracketed_references_are_tokens_and_original_formula_text_is_retained() {
    let formula = FormulaParser::new("of:=[.A1]")
        .parse()
        .expect("legacy-compatible bracketed cell remains parseable");
    assert_eq!(formula.text, "of:=[.A1]");
    assert!(matches!(formula.tokens.first(), Some(Token::CellRef(_))));

    let reference = parse_formula_reference("of:=[.A1:.B2]");
    match &reference {
        Reference::Local(Address::Cells(start, end)) => {
            assert_eq!(start.sheet, SheetSelector::Current);
            assert_eq!(end.sheet, SheetSelector::Inherited);
            assert_cell(start, "A", 1, false, false);
            assert_cell(end, "B", 2, false, false);
        },
        other => panic!("expected rich inherited cell range, got {other:?}"),
    }

    let formula = FormulaParser::new("=A1+Sheet1.$B$2")
        .parse()
        .expect("legacy bare cell references remain parseable");
    assert_eq!(formula.text, "=A1+Sheet1.$B$2");
    assert!(matches!(formula.tokens[0], Token::CellRef(_)));
    assert!(matches!(formula.tokens[2], Token::CellRef(_)));

    let nested = FormulaParser::new("of:=SUM([.A1:.B2])")
        .parse()
        .expect("a bracketed reference is inert inside a function call");
    assert_eq!(nested.text, "of:=SUM([.A1:.B2])");
    assert!(
        nested
            .tokens
            .iter()
            .any(|token| matches!(token, Token::Reference(_)))
    );
    assert!(
        nested
            .tokens
            .iter()
            .filter_map(|token| match token {
                Token::Reference(reference) => Some(reference.as_ref()),
                _ => None,
            })
            .any(|reference| matches!(reference, Reference::Local(Address::Cells(_, _))))
    );
    assert!(
        extract_references(&nested)
            .iter()
            .any(|reference| matches!(reference, ReferenceView::Rich(_)))
    );
    assert!(matches!(
        parse_formula_reference("of:=[.A1:.B2]").address(),
        Some(Address::Cells(_, _))
    ));
    match parse_formula_reference("of:=[.A1:.B2]").address() {
        Some(Address::Cells(start, end)) => {
            assert_eq!(start.sheet, SheetSelector::Current);
            assert_eq!(end.sheet, SheetSelector::Inherited);
        },
        other => panic!("rich address accessor lost inherited endpoint: {other:?}"),
    }
}

#[test]
fn all_six_range_alternatives_preserve_coordinate_kind_and_locator_state() {
    match parse_reference("[.A1]") {
        Reference::Local(Address::Cell(endpoint)) => {
            assert_eq!(endpoint.sheet, SheetSelector::Current);
            assert_cell(&endpoint, "A", 1, false, false);
        },
        other => panic!("expected one current-sheet cell, got {other:?}"),
    }

    match parse_reference("[.A1:.B$2]") {
        Reference::Local(Address::Cells(start, end)) => {
            assert_eq!(start.sheet, SheetSelector::Current);
            assert_eq!(end.sheet, SheetSelector::Inherited);
            assert_cell(&start, "A", 1, false, false);
            assert_cell(&end, "B", 2, false, true);
        },
        other => panic!("expected inherited cell range, got {other:?}"),
    }

    match parse_reference("[.$A:.$C]") {
        Reference::Local(Address::Columns(start, end)) => {
            assert_eq!(start.sheet, SheetSelector::Current);
            assert_eq!(end.sheet, SheetSelector::Inherited);
            assert_column(&start, "A", true);
            assert_column(&end, "C", true);
        },
        other => panic!("expected inherited whole-column range, got {other:?}"),
    }

    match parse_reference("[.$1:.$3]") {
        Reference::Local(Address::Rows(start, end)) => {
            assert_eq!(start.sheet, SheetSelector::Current);
            assert_eq!(end.sheet, SheetSelector::Inherited);
            assert_row(&start, 1, true);
            assert_row(&end, 3, true);
        },
        other => panic!("expected inherited whole-row range, got {other:?}"),
    }

    match parse_reference("[Sheet1.A1:Sheet2.$B$2]") {
        Reference::Local(Address::Cells(start, end)) => {
            assert_explicit_sheet(&start, "Sheet1", false, false);
            assert_explicit_sheet(&end, "Sheet2", false, false);
            assert_cell(&start, "A", 1, false, false);
            assert_cell(&end, "B", 2, true, true);
        },
        other => panic!("expected cross-sheet cell cuboid, got {other:?}"),
    }

    match parse_reference("[$Sheet1.$A:$Sheet2.$C]") {
        Reference::Local(Address::Columns(start, end)) => {
            assert_explicit_sheet(&start, "Sheet1", true, false);
            assert_explicit_sheet(&end, "Sheet2", true, false);
            assert_column(&start, "A", true);
            assert_column(&end, "C", true);
        },
        other => panic!("expected cross-sheet whole-column cuboid, got {other:?}"),
    }

    match parse_reference("[Sheet1.$1:Sheet2.$3]") {
        Reference::Local(Address::Rows(start, end)) => {
            assert_explicit_sheet(&start, "Sheet1", false, false);
            assert_explicit_sheet(&end, "Sheet2", false, false);
            assert_row(&start, 1, true);
            assert_row(&end, 3, true);
        },
        other => panic!("expected cross-sheet whole-row cuboid, got {other:?}"),
    }
}

#[test]
fn quoted_unicode_delimited_and_nested_locators_keep_full_metadata() {
    match parse_reference("['Q1 Sales'.$C$3]") {
        Reference::Local(Address::Cell(endpoint)) => {
            assert_explicit_sheet(&endpoint, "Q1 Sales", false, true);
            assert_cell(&endpoint, "C", 3, true, true);
        },
        other => panic!("expected quoted sheet cell, got {other:?}"),
    }

    match parse_reference("['Bob''s'.A1]") {
        Reference::Local(Address::Cell(endpoint)) => {
            assert_explicit_sheet(&endpoint, "Bob's", false, true);
            assert_cell(&endpoint, "A", 1, false, false);
        },
        other => panic!("expected doubled-apostrophe sheet cell, got {other:?}"),
    }

    match parse_reference("['日本語'.A1]") {
        Reference::Local(Address::Cell(endpoint)) => {
            assert_explicit_sheet(&endpoint, "日本語", false, true);
            assert_cell(&endpoint, "A", 1, false, false);
        },
        other => panic!("expected Unicode sheet cell, got {other:?}"),
    }

    match parse_reference("['A] : B.C'.A1]") {
        Reference::Local(Address::Cell(endpoint)) => {
            assert_explicit_sheet(&endpoint, "A] : B.C", false, true);
            assert_cell(&endpoint, "A", 1, false, false);
        },
        other => panic!("expected delimiters inside quoted sheet, got {other:?}"),
    }

    match parse_reference("[$'Q1 Sales'.$A$1]") {
        Reference::Local(Address::Cell(endpoint)) => {
            assert_explicit_sheet(&endpoint, "Q1 Sales", true, true);
            assert_cell(&endpoint, "A", 1, true, true);
        },
        other => panic!("expected absolute quoted sheet cell, got {other:?}"),
    }

    for (spelling, expected_name) in [("[[.A1]", "["), ("[S[heet.A1]", "S[heet")] {
        match parse_reference(spelling) {
            Reference::Local(Address::Cell(endpoint)) => {
                assert_explicit_sheet(&endpoint, expected_name, false, false);
                assert_cell(&endpoint, "A", 1, false, false);
            },
            other => panic!("expected delimiter-bearing unquoted sheet, got {other:?}"),
        }
    }

    match parse_reference("[Sheet:Name.A1:.B2]") {
        Reference::Local(Address::Cells(start, end)) => {
            assert_explicit_sheet(&start, "Sheet:Name", false, false);
            assert_eq!(end.sheet, SheetSelector::Inherited);
            assert_cell(&start, "A", 1, false, false);
            assert_cell(&end, "B", 2, false, false);
        },
        other => panic!("expected colon-bearing sheet name, got {other:?}"),
    }

    match parse_reference("[Sheet.A1.'Sub]Table'.'Sub.Table'.B2]") {
        Reference::Local(Address::Cell(endpoint)) => {
            assert_explicit_sheet(&endpoint, "Sheet", false, false);
            let SheetSelector::Explicit(locator) = &endpoint.sheet else {
                unreachable!("checked above");
            };
            assert_eq!(locator.subtables.len(), 3);
            match &locator.subtables[0] {
                Subtable::Cell(cell) => {
                    assert_eq!(cell.column.label, "A");
                    assert_eq!(cell.row.number, 1);
                },
                other => panic!("expected cell subtable, got {other:?}"),
            }
            match &locator.subtables[1] {
                Subtable::Name(name) => {
                    assert_eq!(name.name, "Sub]Table");
                    assert!(name.quoted);
                },
                other => panic!("expected quoted subtable, got {other:?}"),
            }
            match &locator.subtables[2] {
                Subtable::Name(name) => {
                    assert_eq!(name.name, "Sub.Table");
                    assert!(name.quoted);
                },
                other => panic!("expected second quoted subtable, got {other:?}"),
            }
            assert_cell(&endpoint, "B", 2, false, false);
        },
        other => panic!("expected nested subtable cell, got {other:?}"),
    }
}

#[test]
fn source_iris_are_decoded_and_remain_external_inert_metadata() {
    for (spelling, decoded) in [
        ("[''#.A1]", ""),
        ("['../book.ods'#.A1]", "../book.ods"),
        ("['#fragment'#.A1]", "#fragment"),
        (
            "['https://例え.テスト/こんにちは?x=✓'#.A1]",
            "https://例え.テスト/こんにちは?x=✓",
        ),
        (
            "['https://example.org/a#fragment'#.A1]",
            "https://example.org/a#fragment",
        ),
        ("['file:///O''Brien.ods'#.A1]", "file:///O'Brien.ods"),
    ] {
        let reference = parse_reference(spelling);
        assert_eq!(
            reference.source().map(|source| source.as_str()),
            Some(decoded)
        );
        assert!(matches!(reference.address(), Some(Address::Cell(_))));
        assert!(!reference.is_error());
        match reference {
            Reference::Source {
                address: Address::Cell(endpoint),
                ..
            } => {
                assert_eq!(endpoint.sheet, SheetSelector::Current);
                assert_cell(&endpoint, "A", 1, false, false);
            },
            other => panic!("expected source-qualified cell, got {other:?}"),
        }
    }

    let formula_reference = parse_formula_reference("of:=['file:///O''Brien.ods'#.A1]");
    assert_eq!(
        formula_reference.source().unwrap().as_str(),
        "file:///O'Brien.ods"
    );
    assert!(matches!(formula_reference, Reference::Source { .. }));

    let nested_formula = FormulaParser::new("of:=SUM(['file:///book.ods'#.A1])")
        .parse()
        .expect("external references remain inert inside a formula");
    let rich_references = extract_references(&nested_formula);
    assert!(matches!(
        rich_references.as_slice(),
        [ReferenceView::Rich(Reference::Source { .. })]
    ));
    assert!(extract_cell_refs(&nested_formula).is_empty());

    match parse_reference("['../book.ods'#Sheet1.A1:.B2]") {
        Reference::Source {
            address: Address::Cells(start, end),
            ..
        } => {
            assert_explicit_sheet(&start, "Sheet1", false, false);
            assert_eq!(end.sheet, SheetSelector::Inherited);
            assert_cell(&start, "A", 1, false, false);
            assert_cell(&end, "B", 2, false, false);
        },
        other => panic!("expected source-qualified inherited range, got {other:?}"),
    }
}

#[test]
fn reference_error_is_distinct_and_cannot_be_made_source_qualified() {
    let reference = parse_reference("[#REF!]");
    assert!(reference.is_error());
    assert_eq!(reference.source(), None);
    assert_eq!(reference.address(), None);
    assert!(matches!(reference, Reference::Error));

    assert!(Reference::parse("[''#REF!]").is_err());
    assert!(Reference::parse("['file:///book.ods'# #REF!]").is_err());
}

#[test]
fn public_reference_parser_requires_the_normative_bracket_wrapper() {
    for input in [".A1", "Sheet1.A1", "#REF!"] {
        assert!(
            Reference::parse(input).is_err(),
            "unbracketed reference was accepted by the public parser: {input:?}"
        );
    }
    assert!(Reference::parse("[.A1]").is_ok());
}

#[test]
fn representative_functions_and_bare_cell_controls_remain_unchanged() {
    for name in [
        "SUM", "LOG10", "BIN2DEC", "BITAND", "COMPLEX", "DCOUNT", "DDE", "VLOOKUP",
    ] {
        assert!(
            is_valid_function(name),
            "representative function was lost: {name}"
        );
    }
    for (spelling, canonical) in [("sum", "SUM"), ("LoG10", "LOG10"), ("bin2dec", "BIN2DEC")] {
        assert!(
            is_valid_function(spelling),
            "case-insensitive function was rejected: {spelling}"
        );
        assert!(is_valid_function(canonical));
    }

    for input in ["=A1", "=AA10", "=Sheet1.A1", "=$A$1", "=.A1"] {
        let formula = FormulaParser::new(input)
            .parse()
            .unwrap_or_else(|error| panic!("legacy cell reference rejected {input:?}: {error}"));
        assert_eq!(formula.text, input);
        assert!(
            matches!(formula.tokens.first(), Some(Token::CellRef(_))),
            "{input:?}"
        );
        assert!(
            formula
                .tokens
                .iter()
                .all(|token| !matches!(token, Token::Reference(_))),
            "{input:?}"
        );
    }
}

#[test]
fn formula_limits_are_exact_and_forward_into_nested_references() {
    let input = "=A1";
    let exact = FormulaLimits::default().with_max_bytes(input.len());
    let formula = FormulaParser::new(input)
        .parse_with_limits(&exact)
        .expect("formula at the byte limit must be admitted");
    assert_eq!(formula.text, input);
    assert_eq!(formula.tokens.len(), 1);

    let short = FormulaLimits::default().with_max_bytes(input.len() - 1);
    let byte_error = FormulaParser::new(input)
        .parse_with_limits(&short)
        .expect_err("formula one byte over the limit must fail");
    let Error::ResourceLimit(byte_limit) = byte_error else {
        panic!("formula byte overflow lost its typed limit error: {byte_error}");
    };
    assert_eq!(byte_limit.resource, Resource::InputBytes);
    assert_eq!(byte_limit.observed, input.len() as u64);
    assert_eq!(byte_limit.limit, (input.len() - 1) as u64);

    let one_token = FormulaLimits::default().with_max_tokens(1);
    let one = FormulaParser::new("=1")
        .parse_with_limits(&one_token)
        .expect("one token at the token limit must be admitted");
    assert_eq!(one.tokens.len(), 1);

    let token_error = FormulaParser::new("=1+2")
        .parse_with_limits(&one_token)
        .expect_err("a second token over the limit must fail");
    let Error::ResourceLimit(token_limit) = token_error else {
        panic!("formula token overflow lost its typed limit error: {token_error}");
    };
    assert_eq!(token_limit.resource, Resource::Objects);
    assert_eq!(token_limit.observed, 2);
    assert_eq!(token_limit.limit, 1);

    // Admission happens before `next_token`, so a malformed first token still
    // reports the configured token budget instead of a syntax fallback.
    let no_tokens = FormulaLimits::default().with_max_tokens(0);
    let malformed_error = FormulaParser::new("=~")
        .parse_with_limits(&no_tokens)
        .expect_err("zero token budget must reject before parsing malformed input");
    let Error::ResourceLimit(malformed_limit) = malformed_error else {
        panic!(
            "zero-token admission parsed malformed input instead of returning a limit: {malformed_error}"
        );
    };
    assert_eq!(malformed_limit.resource, Resource::Objects);
    assert_eq!(malformed_limit.observed, 1);
    assert_eq!(malformed_limit.limit, 0);

    let nested_exact =
        FormulaLimits::default().with_reference_limits(Limits::default().with_max_bytes(3));
    FormulaParser::new("of:=[.A1]")
        .parse_with_limits(&nested_exact)
        .expect("nested reference at its body byte limit must be admitted");

    let nested_short =
        FormulaLimits::default().with_reference_limits(Limits::default().with_max_bytes(2));
    let nested_error = FormulaParser::new("of:=[.A1]")
        .parse_with_limits(&nested_short)
        .expect_err("nested reference byte overflow must be forwarded");
    let Error::ResourceLimit(nested_limit) = nested_error else {
        panic!("nested reference limit was converted to syntax error: {nested_error}");
    };
    assert_eq!(nested_limit.resource, Resource::InputBytes);
    assert_eq!(nested_limit.observed, 3);
    assert_eq!(nested_limit.limit, 2);

    let nested_name =
        FormulaLimits::default().with_reference_limits(Limits::default().with_max_name_bytes(3));
    let name_error = FormulaParser::new("of:=[Sheet.A1]")
        .parse_with_limits(&nested_name)
        .expect_err("nested reference name overflow must be forwarded");
    assert!(matches!(name_error, Error::ResourceLimit(_)));
}

#[test]
fn malformed_mixed_and_overflow_references_are_refused() {
    for input in [
        "[]",
        "[.a1]",
        "[.A01]",
        "[.A0]",
        "[.A]",
        "[.1]",
        "[.A1:.B]",
        "[.A1:.1]",
        "[.A:.1]",
        "[.A1:Sheet.B2]",
        "[.A1:B2]",
        "[.A1:.B2:.C3]",
        "[Sheet.A1:Sheet.B2:Sheet.C3]",
        "[Sheet..A1]",
        "[Sheet.A1.]",
        "[Sheet.A1.'Sub].B2]",
        "[''.A1]",
        "['Q1 Sales.A1]",
        "['Bob's'.A1]",
        "[.A4294967296]",
        "[.A18446744073709551616]",
        "['file:///bad%ZZ'#.A1]",
        "['file:///bad path'#.A1]",
        "['file:///bad'#A1]",
        "['file:///bad'#.A1",
        "[#ref!]",
    ] {
        assert!(
            Reference::parse(input).is_err(),
            "malformed reference was accepted: {input:?}"
        );
    }
}

#[test]
fn source_and_locator_boundaries_are_bounded_without_losing_atomicity() {
    // `max_bytes` is the body budget: the two outer brackets do not consume
    // it.  Check both sides of the boundary through the public parser.
    let exact_body = Limits::default().with_max_bytes(".A1".len());
    assert!(Reference::parse_with_limits("[.A1]", &exact_body).is_ok());
    let one_byte_short = Limits::default().with_max_bytes(".A1".len() - 1);
    assert!(Reference::parse_with_limits("[.A1]", &one_byte_short).is_err());

    // FormulaParser uses the same body budget after it has removed the
    // bracket wrapper.  Keep the constructed names within the independent
    // decoded-name budget while making the body exactly 64 KiB and then one
    // byte larger.
    let default_body_limit = Limits::default().max_bytes();
    let exact_formula = format!("of:=[{}]", reference_body_with_length(default_body_limit));
    let exact = FormulaParser::new(&exact_formula)
        .parse()
        .expect("formula reference at the body limit must be admitted");
    assert_eq!(exact.text, exact_formula);
    assert!(exact.tokens.iter().any(|token| matches!(
        token,
        Token::Reference(reference) if matches!(reference.as_ref(), Reference::Source { .. })
    )));

    let over_formula = format!(
        "of:=[{}]",
        reference_body_with_length(default_body_limit + 1)
    );
    let over_error = FormulaParser::new(&over_formula)
        .parse()
        .expect_err("formula reference one byte over the body limit must fail");
    assert!(
        matches!(&over_error, Error::ResourceLimit(_)),
        "formula body overflow must retain its typed resource-limit cause: {over_error}"
    );

    // Name limits are measured after UTF-8 decoding and doubled-apostrophe
    // unescaping, rather than by their lexical spelling.
    let exact_unicode = Limits::default().with_max_name_bytes("界".len());
    assert!(Reference::parse_with_limits("['界'.A1]", &exact_unicode).is_ok());
    let unicode_short = Limits::default().with_max_name_bytes("界".len() - 1);
    assert!(Reference::parse_with_limits("['界'.A1]", &unicode_short).is_err());

    let exact_decoded = Limits::default().with_max_name_bytes(3);
    assert!(Reference::parse_with_limits("['a''b'.A1]", &exact_decoded).is_ok());
    let decoded_short = Limits::default().with_max_name_bytes(2);
    assert!(Reference::parse_with_limits("['a''b'.A1]", &decoded_short).is_err());
    assert!(Reference::parse_with_limits("['a''b'#.A1]", &exact_decoded).is_ok());
    assert!(Reference::parse_with_limits("['a''b'#.A1]", &decoded_short).is_err());

    let body_limit = Limits::default().with_max_bytes(2);
    assert!(Reference::parse_with_limits("[.A1]", &body_limit).is_err());

    let name_limit = Limits::default().with_max_name_bytes(3);
    assert!(Reference::parse_with_limits("[Sheet.A1]", &name_limit).is_err());
    assert!(Reference::parse_with_limits("['../book.ods'#.A1]", &name_limit).is_err());

    let component_limit = Limits::default().with_max_components(1);
    assert!(Reference::parse_with_limits("[.A1]", &component_limit).is_err());

    let oversized_source = "a".repeat(Limits::default().max_name_bytes() + 1);
    let oversized_reference = format!("['{oversized_source}'#.A1]");
    let source_error = Reference::parse(&oversized_reference)
        .expect_err("an oversized source must fail before it can be retained");
    assert!(
        matches!(&source_error, Error::ResourceLimit(_)),
        "source refusal must retain its typed resource-limit cause: {source_error}"
    );
    assert!(
        source_error
            .to_string()
            .to_ascii_lowercase()
            .contains("limit"),
        "source refusal lost its resource-limit cause: {source_error}"
    );
    let formula = format!("of:={oversized_reference}");
    let formula_error = FormulaParser::new(&formula)
        .parse()
        .expect_err("the formula path must propagate source admission refusal");
    assert!(
        matches!(&formula_error, Error::ResourceLimit(_)),
        "formula path must retain its typed resource-limit cause: {formula_error}"
    );
    assert!(
        formula_error
            .to_string()
            .to_ascii_lowercase()
            .contains("limit"),
        "formula path converted source limit refusal into a generic fallback: {formula_error}"
    );

    let valid = Reference::parse_with_limits(
        "[''#.A1:.B2]",
        &Limits::default()
            .with_max_bytes(64)
            .with_max_components(8)
            .with_max_name_bytes(8),
    )
    .expect("a bounded valid source reference remains readable");
    assert_eq!(valid.source().unwrap().as_str(), "");
    assert!(matches!(valid, Reference::Source { .. }));
}
