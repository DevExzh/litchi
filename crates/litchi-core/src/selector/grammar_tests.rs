//! Unit tests per grammar feature.

use std::borrow::Cow;

use super::*;

fn path(input: &str) -> Path<'_> {
    Path::parse(input).unwrap_or_else(|error| panic!("{input:?}: {error}"))
}

fn query(input: &str) -> Query<'_> {
    Query::parse(input).unwrap_or_else(|error| panic!("{input:?}: {error}"))
}

fn path_error(input: &str) -> ParseError {
    match Path::parse(input) {
        Ok(path) => panic!("{input:?} parsed as {path:?}"),
        Err(error) => error,
    }
}

fn query_error(input: &str) -> ParseError {
    match Query::parse(input) {
        Ok(query) => panic!("{input:?} parsed as {query:?}"),
        Err(error) => error,
    }
}

/// Asserts `display(parse(input)) == canonical` and that it parses back.
fn path_shows(input: &str, canonical: &str) {
    let parsed = path(input);
    assert_eq!(parsed.to_string(), canonical, "{input:?}");
    assert_eq!(path(canonical), parsed, "{canonical:?}");
}

fn query_shows(input: &str, canonical: &str) {
    let parsed = query(input);
    assert_eq!(parsed.to_string(), canonical, "{input:?}");
    assert_eq!(query(canonical), parsed, "{canonical:?}");
}

fn eq(key: &'static str, value: &'static str) -> Predicate<'static> {
    Predicate::Filter(Expr::Condition(Condition::Compare {
        key: Key::new(key),
        op: CompareOp::Eq,
        value: Value::text(value),
    }))
}

fn cmp(key: &'static str, op: CompareOp, value: &'static str) -> Expr<'static> {
    Expr::Condition(Condition::Compare {
        key: Key::new(key),
        op,
        value: Value::text(value),
    })
}

fn identity(key: &'static str, value: &'static str) -> Predicate<'static> {
    Predicate::Filter(Expr::Condition(Condition::Compare {
        key: Key::attribute(key),
        op: CompareOp::Eq,
        value: Value::text(value),
    }))
}

// ---------------------------------------------------------------- paths

#[test]
fn element_name_paths_are_one_based_strings_and_zero_based_values() {
    let parsed = path("/body/p[3]/r[1]");
    assert!(parsed.is_absolute());
    let names: Vec<_> = parsed.steps().iter().map(Step::name).collect();
    assert_eq!(names, ["body", "p", "r"]);
    assert_eq!(parsed.steps()[1].index(), Some(Position::new(2)));
    assert_eq!(parsed.steps()[2].index(), Some(Position::new(0)));
    assert_eq!(parsed.steps()[0].index(), None);
    path_shows("/slide[1]/shape[2]", "/slide[1]/shape[2]");
    path_shows("/body/p[007]", "/body/p[7]");
    path_shows("/body/p[ 2 ]", "/body/p[2]");
}

#[test]
fn root_and_relative_paths() {
    let root = path("/");
    assert!(root.is_root());
    assert_eq!(root.to_string(), "/");
    assert_eq!(root, Path::root());

    let anchor = Path::parse_relative("p[@paraId=1A2B3C4D]").expect("relative");
    assert!(!anchor.is_absolute());
    assert_eq!(
        anchor.steps()[0].predicates(),
        &[identity("paraId", "1A2B3C4D")]
    );
    assert_eq!(anchor.to_string(), "p[@paraId=1A2B3C4D]");
    assert_eq!(
        Path::parse_relative("tr[2]/tc[1]")
            .expect("relative")
            .to_string(),
        "tr[2]/tc[1]"
    );
}

#[test]
fn last_predicate() {
    let parsed = path("/body/p[last()]/r[LAST()]");
    assert_eq!(parsed.steps()[1].predicates(), &[Predicate::Last]);
    assert_eq!(parsed.steps()[2].predicates(), &[Predicate::Last]);
    assert_eq!(parsed.to_string(), "/body/p[last()]/r[last()]");
}

#[test]
fn stable_identity_predicates() {
    let parsed = path("/slide[1]/shape[@name=#s1-title]");
    let predicate = &parsed.steps()[1].predicates()[0];
    assert_eq!(predicate.identity(), Some(("name", "#s1-title")));
    assert!(predicate.is_structural());

    let parsed = path("/revision[@id=500][@type=moveTo]");
    assert_eq!(
        parsed.steps()[0].predicates(),
        &[identity("id", "500"), identity("type", "moveTo")]
    );
    path_shows(
        "/body/p[@paraId=\"1A2B 3C4D\"]",
        "/body/p[@paraId=\"1A2B 3C4D\"]",
    );
    path_shows("/body/p[@paraId='X']", "/body/p[@paraId=X]");
    path_shows("/bookmark[ @name = Intro ]", "/bookmark[@name=Intro]");
}

#[test]
fn generic_attribute_and_content_filter_predicates() {
    let parsed = path("/Sheet1/cell[value>5000 or value<300]");
    assert!(!parsed.is_structural());
    let parsed = path("/body/p[1]/r[bold=true]");
    assert_eq!(parsed.steps()[2].predicates(), &[eq("bold", "true")]);
    assert!(!parsed.is_structural());
    assert!(path("/slide[1]/shape[@name=Title]/p[2]").is_structural());
}

#[test]
fn bare_key_predicates() {
    let parsed = path("/Sheet1/col[B]");
    assert_eq!(parsed.steps()[1].predicates()[0].bare_key(), Some("B"));
    assert!(parsed.is_structural());
    assert_eq!(
        path("/formfield[agree_terms]").steps()[0].predicates()[0].bare_key(),
        Some("agree_terms")
    );
    assert_eq!(
        path("/namedrange[\"My Range\"]").to_string(),
        "/namedrange[\"My Range\"]"
    );
    // An all-digit key must stay distinguishable from an index.
    let digits = Path::from_steps([Step::new("x").with(Predicate::Filter(Expr::Condition(
        Condition::Exists(Key::new("12")),
    )))]);
    assert_eq!(digits.to_string(), "/x[\"12\"]");
    assert_eq!(path("/x[\"12\"]"), digits);
}

#[test]
fn spreadsheet_paths() {
    let parsed = path("/Sheet1/A1");
    let cell = parsed.steps()[1].cell_ref().expect("cell");
    assert_eq!((cell.column(), cell.row()), (0, 0));

    let parsed = path("/Sheet1/A1:C5");
    let range = parsed.steps()[1].cell_range().expect("range");
    assert_eq!(
        (
            range.start().column(),
            range.end().column(),
            range.end().row()
        ),
        (0, 2, 4)
    );
    assert_eq!(parsed.steps()[1].cell_ref(), None);

    assert_eq!(
        path("/Sheet1/row[5]").steps()[1].index(),
        Some(Position::new(4))
    );
    path_shows("/Balance Sheet/B18:E18", "/Balance Sheet/B18:E18");
    path_shows(
        "/1-Basics & Quartile/chart[1]",
        "/1-Basics & Quartile/chart[1]",
    );
    path_shows("/P&L", "/P&L");
    path_shows("/O'Brien/A1", "/O'Brien/A1");
    path_shows("/My Data (2024)/row[2]", "/My Data (2024)/row[2]");
    path_shows("/sheet[2]/B3", "/sheet[2]/B3");
    path_shows("/'a/b'/A1", "/\"a/b\"/A1");
    path_shows("/\"x[1]\"/A1", "/\"x[1]\"/A1");
    path_shows("/\"say \\\"hi\\\"\"/A1", "/say \"hi\"/A1");
    assert_eq!(
        path("/\"say \\\"hi\\\"\"/A1").steps()[0].name(),
        "say \"hi\""
    );
    path_shows("/styles/Heading1", "/styles/Heading1");
    assert_eq!(column_number("AB"), Some(28));
    assert_eq!(CellRef::parse("A0"), None);
    assert_eq!(CellRef::parse("A1048577"), None);
    assert_eq!(CellRange::parse("A1:B"), None);
}

#[test]
fn story_axes() {
    let cases = [
        ("/header[1]/p[1]", Axis::Header),
        ("/footer[2]/p[1]/r[1]", Axis::Footer),
        ("/footnote[2]", Axis::Footnote),
        ("/footnotes/footnote[2]", Axis::Footnote),
        ("/endnote[1]", Axis::Endnote),
        ("/comments/comment[3]", Axis::Comment),
        ("/slide[2]/notes", Axis::Slide),
        ("/slideMaster[1]/layout[2]", Axis::SlideMaster),
        ("/slidemaster[1]/slidelayout[2]/shape[1]", Axis::SlideMaster),
        ("/slideLayout[3]", Axis::SlideLayout),
        ("/chart[1]/series[2]", Axis::Chart),
        ("/notesMaster", Axis::NotesMaster),
        ("/noteSlide[1]", Axis::Notes),
        ("/body/p[1]", Axis::Body),
        ("/styles", Axis::Styles),
        ("/numbering/abstractNum[@id=1]", Axis::Numbering),
        ("/theme", Axis::Theme),
        ("/watermark", Axis::Watermark),
    ];
    for (input, axis) in cases {
        assert_eq!(
            path(input).story().map(|(_, axis)| axis),
            Some(axis),
            "{input}"
        );
    }
    let notes: Vec<_> = path("/slide[2]/notes").axes().collect();
    assert_eq!(notes, [(0, Axis::Slide), (1, Axis::Notes)]);
    let chart: Vec<_> = path("/Sheet1/chart[1]/series[2]").axes().collect();
    assert_eq!(chart, [(1, Axis::Chart)]);
    assert!(Axis::table_is_sorted());
    assert!(Axis::Header.is_story());
    assert!(!Axis::Styles.is_story());
    assert_eq!(Axis::classify("a-very-long-name-indeed"), None);
}

#[test]
fn path_syntax_errors_carry_offsets_and_codes() {
    let cases: &[(&str, ErrorKind, usize)] = &[
        ("", ErrorKind::Empty, 0),
        ("//body", ErrorKind::EmptySegment, 1),
        ("/body//p[1]", ErrorKind::EmptySegment, 6),
        ("/body/p[1]/", ErrorKind::EmptySegment, 10),
        ("/body/p[0]", ErrorKind::IndexZero, 8),
        ("/body/p[-1]", ErrorKind::IndexNegative, 8),
        (
            "/body/p[99999999999999999999999]",
            ErrorKind::IndexOverflow,
            8,
        ),
        ("/a\u{1}b", ErrorKind::ControlCharacter, 2),
        ("/a/\"unterminated", ErrorKind::UnterminatedQuote, 3),
        ("/x[name=O'Brien]", ErrorKind::QuoteInBareValue, 9),
        ("/x[progId=Excel*]", ErrorKind::Wildcard, 15),
        ("/x[a==b]", ErrorKind::MisparsedOperator, 5),
        ("/x[a=r\"b\"]", ErrorKind::RegexRequiresContains, 5),
    ];
    for &(input, kind, offset) in cases {
        let error = path_error(input);
        assert_eq!(
            (error.kind(), error.offset()),
            (kind, offset),
            "{input:?}: {error}"
        );
        assert_eq!(error.code(), ErrorCode::InvalidPath, "{input:?}");
        assert_eq!(error.grammar(), Grammar::Path);
    }
    let unexpected: &[(&str, Expected, Found, usize)] = &[
        ("body", Expected::Root, Found::Char('b'), 0),
        ("/body/p[1", Expected::CloseBracket, Found::End, 9),
        ("/body/p[1]x", Expected::StepEnd, Found::Char('x'), 10),
        ("/body/p[]", Expected::Predicate, Found::Char(']'), 8),
        ("/body/[1]", Expected::SegmentName, Found::Char('['), 6),
        ("/body]", Expected::StepEnd, Found::Char(']'), 5),
        (
            "/x[Full Name=1]",
            Expected::CloseBracket,
            Found::Char('N'),
            8,
        ),
        ("/x[a$b]", Expected::Operator, Found::Char('$'), 4),
        ("/x[a=]", Expected::Value, Found::Char(']'), 5),
        ("/x[a=1 and]", Expected::CloseBracket, Found::Char('a'), 7),
    ];
    for &(input, expected, found, offset) in unexpected {
        let error = path_error(input);
        assert_eq!(
            (error.kind(), error.offset()),
            (ErrorKind::Unexpected { expected, found }, offset),
            "{input:?}: {error}"
        );
    }
    let error = path_error("/body/p[x]]");
    assert_eq!(
        error.to_string(),
        "invalid path at byte 10: expected '/', '[', or end of path, found a character"
    );
    assert_eq!(path_error("/x[a*]").suggestion(), None);
    assert!(path_error("/x[a=b*]").suggestion().is_some());
}

#[test]
fn offsets_move_to_the_enclosing_text() {
    let error = path_error("/body/p[0]");
    let moved = error.offset_by(7);
    assert_eq!(
        (moved.kind(), moved.offset(), moved.code(), moved.grammar()),
        (error.kind(), 15, error.code(), error.grammar())
    );
    assert_eq!(
        moved.to_string(),
        "invalid path at byte 15: index 0 is invalid; indices are 1-based"
    );
    assert_eq!(error.offset_by(usize::MAX).offset(), usize::MAX);
}

#[test]
fn path_limits_are_typed() {
    let deep = "/a".repeat(limits::MAX_STEPS + 1);
    let error = path_error(&deep);
    assert_eq!(error.kind(), ErrorKind::LimitExceeded(Limit::Steps));
    assert_eq!(error.code(), ErrorCode::MaxDepthExceeded);
    assert!(Path::parse(&"/a".repeat(limits::MAX_STEPS)).is_ok());

    let many = format!("/a{}", "[1]".repeat(limits::MAX_PREDICATES + 1));
    assert_eq!(
        path_error(&many).kind(),
        ErrorKind::LimitExceeded(Limit::Predicates)
    );

    let long = format!("/{}", "a".repeat(limits::MAX_INPUT_BYTES));
    let error = path_error(&long);
    assert_eq!(error.kind(), ErrorKind::LimitExceeded(Limit::InputBytes));
    assert_eq!(error.code(), ErrorCode::InvalidPath);

    let nested = format!(
        "/a[{}b{}]",
        "(".repeat(limits::MAX_NESTING + 1),
        ")".repeat(limits::MAX_NESTING + 1)
    );
    let error = path_error(&nested);
    assert_eq!(error.kind(), ErrorKind::LimitExceeded(Limit::Nesting));
    assert_eq!(error.code(), ErrorCode::MaxDepthExceeded);
    let ok = format!(
        "/a[{}b{}]",
        "(".repeat(limits::MAX_NESTING),
        ")".repeat(limits::MAX_NESTING)
    );
    assert!(Path::parse(&ok).is_ok());

    let terms = format!("/a[{}]", vec!["k"; limits::MAX_TERMS + 1].join(" and "));
    assert_eq!(
        path_error(&terms).kind(),
        ErrorKind::LimitExceeded(Limit::Terms)
    );

    let mut built = Path::root();
    for _ in 0..limits::MAX_STEPS {
        built.push(Step::new("a")).expect("within bound");
    }
    assert_eq!(
        built.push(Step::new("a")).unwrap_err().kind(),
        ErrorKind::LimitExceeded(Limit::Steps)
    );
}

#[test]
fn paths_borrow_from_input_and_convert_to_owned() {
    let input = String::from("/slide[1]/shape[@name=Title]");
    let parsed = Path::parse(&input).expect("parses");
    let Predicate::Filter(Expr::Condition(Condition::Compare {
        value: Value::Text(text),
        ..
    })) = &parsed.steps()[1].predicates()[0]
    else {
        panic!("identity predicate");
    };
    assert!(matches!(text, Cow::Borrowed(_)));
    let owned: Path<'static> = parsed.clone().into_owned();
    drop(input);
    assert_eq!(owned.to_string(), "/slide[1]/shape[@name=Title]");
}

#[test]
fn deep_paths_spill_and_still_compare_structurally() {
    let text = "/body/tbl[1]/tr[2]/tc[3]/p[1]/r[2]/t[1]/x[1]";
    let parsed = path(text);
    assert_eq!(parsed.steps().len(), 8);
    assert_eq!(parsed.to_string(), text);
    let rebuilt = Path::from_steps(parsed.steps().iter().cloned());
    assert_eq!(rebuilt, parsed);
}

#[test]
fn paths_are_covariant() {
    fn same<'a>(left: &Path<'a>, right: &Path<'a>) -> bool {
        left == right
    }
    let owned: Path<'static> = "/body/p[1]".parse().expect("parses");
    let text = String::from("/body/p[1]");
    let borrowed = Path::parse(&text).expect("parses");
    assert!(same(&owned, &borrowed));
}

#[test]
fn position_boundary_conversions() {
    assert_eq!(Position::from_one_based(1), Some(Position::new(0)));
    assert_eq!(Position::from_one_based(0), None);
    assert_eq!(Position::new(usize::MAX).one_based(), None);
    let max = format!("/p[{}]", usize::MAX);
    assert_eq!(
        path(&max).steps()[0].index(),
        Some(Position::new(usize::MAX - 1))
    );
    // The one position without a 1-based form still displays without panic.
    let edge = Path::from_steps([Step::new("p").with(Predicate::Index(Position::new(usize::MAX)))]);
    assert_eq!(edge.to_string(), format!("/p[{}]", usize::MAX as u128 + 1));
}

// ---------------------------------------------------------------- queries

#[test]
fn element_type_selectors() {
    let parsed = query("paragraph");
    let compound = parsed.alternatives()[0].subject();
    assert_eq!(compound.element().name(), Some("paragraph"));
    assert!(parsed.scope().is_none());
    query_shows("*", "*");
    query_shows("  shape  ", "shape");
    query_shows("3dmodel", "3dmodel");
    query_shows("[style=Title]", "[style=Title]");
    assert_eq!(
        query("[style=Title]").alternatives()[0].first().element(),
        &Element::Implicit
    );
}

#[test]
fn attribute_operators() {
    let cases = [
        ("shape[fill=FF0000]", CompareOp::Eq),
        ("field[fieldType!=page]", CompareOp::Ne),
        ("cell[value\\!=0]", CompareOp::Ne),
        ("cell[formula~=SUM]", CompareOp::Contains),
        ("shape[width>=10cm]", CompareOp::Ge),
        ("shape[x<=34cm]", CompareOp::Le),
        ("paragraph[numId>0]", CompareOp::Gt),
        ("row[Salary<5000]", CompareOp::Lt),
    ];
    for (input, op) in cases {
        let parsed = query(input);
        let Predicate::Filter(Expr::Condition(Condition::Compare { op: found, .. })) =
            &parsed.alternatives()[0].subject().predicates()[0]
        else {
            panic!("{input}");
        };
        assert_eq!(*found, op, "{input}");
    }
    query_shows("cell[value\\!=0]", "cell[value!=0]");
    query_shows("paragraph[hangingIndent]", "paragraph[hangingIndent]");
    query_shows(
        "paragraph[style=Heading 1]",
        "paragraph[style=\"Heading 1\"]",
    );
    query_shows(
        "shape[fill=#FF0000][size>=24pt][text~=报告]",
        "shape[fill=#FF0000][size>=24pt][text~=报告]",
    );
    query_shows(
        "namedrange[name=\"_xlnm.Print_Area\"]",
        "namedrange[name=_xlnm.Print_Area]",
    );
    query_shows("row[name=\"A,B\"]", "row[name=A,B]");
    query_shows("row[url~=a/b]", "row[url~=a/b]");
    query_shows("x[a='it\\'s']", "x[a=\"it's\"]");
    query_shows("x[a=\"a\\\\b\"]", "x[a=\"a\\\\b\"]");
    query_shows("x[a=\"\"]", "x[a=\"\"]");
    query_shows("x[a=\"x*\"]", "x[a=\"x*\"]");
    query_shows("x[a=\"==\"]", "x[a=\"==\"]");
    assert_eq!(query("x[a='it\\'s']").to_string(), "x[a=\"it's\"]");
}

#[test]
fn regex_values() {
    let parsed = query("shape[text~=r\"^Q[1-4]\"]");
    let Predicate::Filter(Expr::Condition(Condition::Compare { value, .. })) =
        &parsed.alternatives()[0].subject().predicates()[0]
    else {
        panic!("regex predicate");
    };
    assert_eq!(value, &Value::Regex(Cow::Borrowed("^Q[1-4]")));
    query_shows("shape[text~=r'a\"b']", "shape[text~=r'a\"b']");
    query_shows("shape[text~=r\"a\\\"b\"]", "shape[text~=r\"a\\\"b\"]");
    assert!(Value::regex("ok").is_some());
    assert!(Value::regex("both \" and '").is_none());
    assert!(Value::regex("trailing\\").is_none());
    assert_eq!(
        query_error("x[a=r\"b\"]").kind(),
        ErrorKind::RegexRequiresContains
    );
}

#[test]
fn boolean_expressions_and_header_predicates() {
    let parsed = query("row[Salary>5000 and Region=EMEA]");
    assert_eq!(
        parsed.alternatives()[0].subject().predicates(),
        &[Predicate::Filter(Expr::And(vec![
            cmp("Salary", CompareOp::Gt, "5000"),
            cmp("Region", CompareOp::Eq, "EMEA"),
        ]))]
    );
    // `and` binds tighter than `or`.
    let parsed = query("row[a=1 or b=2 and c=3]");
    assert_eq!(
        parsed.alternatives()[0].subject().predicates(),
        &[Predicate::Filter(Expr::Or(vec![
            cmp("a", CompareOp::Eq, "1"),
            Expr::And(vec![
                cmp("b", CompareOp::Eq, "2"),
                cmp("c", CompareOp::Eq, "3")
            ]),
        ]))]
    );
    query_shows("row[(a=1 or b=2) and c=3]", "row[(a=1 or b=2) and c=3]");
    query_shows("row[(a=1 or b=2) or c=3]", "row[(a=1 or b=2) or c=3]");
    query_shows("row[(a=1)]", "row[a=1]");
    query_shows("row[not(Region=West)]", "row[not(Region=West)]");
    query_shows("row[NOT (a=1) AND b=2]", "row[not(a=1) and b=2]");
    query_shows(
        "row[Region=West and Sales>1000]",
        "row[Region=West and Sales>1000]",
    );
    query_shows(
        "row[Status=open or Priority=high]",
        "row[Status=open or Priority=high]",
    );
    query_shows("row[金额>5 or 金额<1]", "row[金额>5 or 金额<1]");
    query_shows("row[\"Full Name\"~=doe]", "row[\"Full Name\"~=doe]");
    query_shows("row[\"Amount, USD\">0]", "row[\"Amount, USD\">0]");
    query_shows("row[col.Salary>5000]", "row[col.Salary>5000]");
    query_shows(
        "paragraph[text~=\"salt and pepper\"]",
        "paragraph[text~=\"salt and pepper\"]",
    );
    query_shows("cell[a and b]", "cell[a and b]");
    // Keywords are only keywords in keyword position.
    query_shows("row[not=5]", "row[not=5]");
    query_shows("row[and=1 or or=2]", "row[and=1 or or=2]");
    query_shows("row[x=and]", "row[x=and]");
    query_shows(
        "row[style=Heading 1 and x=2]",
        "row[style=\"Heading 1\" and x=2]",
    );
    let mut leaves = 0;
    if let Predicate::Filter(expr) = &query("row[a=1 or not(b=2 and c)]").alternatives()[0]
        .subject()
        .predicates()[0]
    {
        expr.for_each_condition(&mut |_| leaves += 1);
    }
    assert_eq!(leaves, 3);
}

#[test]
fn pseudo_classes() {
    let parsed = query("cell:contains(\"#DIV/0!\")");
    assert_eq!(
        parsed.alternatives()[0].subject().pseudos(),
        &[Pseudo::Contains(Cow::Borrowed("#DIV/0!"))]
    );
    query_shows("run:contains(TODO)", "run:contains(TODO)");
    query_shows(
        "shape:contains( Use of Funds )",
        "shape:contains(Use of Funds)",
    );
    query_shows(
        "shape:contains(\" padded \")",
        "shape:contains(\" padded \")",
    );
    query_shows("cell:TBD", "cell:contains(TBD)");
    query_shows("cell:#REF!", "cell:contains(#REF!)");
    query_shows("p:empty", "p:empty");
    query_shows("image:no-alt", "image:no-alt");
    query_shows("cell:has(formula)", "cell:has(formula)");
    query_shows("cell:not(:has(formula))", "cell:not(:has(formula))");
    query_shows("cell:not(:empty)", "cell:not(:empty)");
    query_shows(
        "paragraph:has(run[bold=true], hyperlink)",
        "paragraph:has(run[bold=true], hyperlink)",
    );
    query_shows("Dashboard!:has(formula)", "Dashboard!:has(formula)");
    query_shows(
        "cell[type=Number]:contains(5)[bold=true]",
        "cell[type=Number][bold=true]:contains(5)",
    );
}

#[test]
fn combinators_and_unions() {
    let parsed = query("paragraph[style=Normal] > run[font!=Arial]");
    let complex = &parsed.alternatives()[0];
    assert_eq!(complex.first().element().name(), Some("paragraph"));
    assert_eq!(complex.rest()[0].0, Combinator::Child);
    assert_eq!(complex.subject().element().name(), Some("run"));
    query_shows("slide[1]>ole", "slide[1] > ole");
    query_shows("slide chart", "slide chart");
    query_shows("slide > table > tr", "slide > table > tr");
    query_shows("table   tr > tc", "table tr > tc");
    let parsed = query("paragraph, table");
    assert_eq!(parsed.alternatives().len(), 2);
    query_shows("paragraph,table ,  shape", "paragraph, table, shape");
}

#[test]
fn scopes() {
    let parsed = query("Sheet1!B[value!=0]");
    assert_eq!(parsed.scope(), Some(&Scope::Sheet(Cow::Borrowed("Sheet1"))));
    assert_eq!(
        parsed.alternatives()[0].subject().element().name(),
        Some("B")
    );
    query_shows("Sheet1!A1", "Sheet1!A1");
    let parsed = query("Sheet1!A1:B5");
    let range = parsed.alternatives()[0]
        .subject()
        .element()
        .cell_range()
        .expect("range");
    assert_eq!(range.end().row(), 4);
    query_shows("Sheet1!A1:B5", "Sheet1!A1:B5");
    query_shows(
        "'My Data (2024)'!row[Amount>5000]",
        "'My Data (2024)'!row[Amount>5000]",
    );
    query_shows("'O''Brien'!A1", "'O''Brien'!A1");
    assert_eq!(
        query("'O''Brien'!A1").scope(),
        Some(&Scope::Sheet(Cow::Borrowed("O'Brien")))
    );
    query_shows("My Sheet!A1", "'My Sheet'!A1");
    query_shows("1-Hero!A1", "'1-Hero'!A1");

    let parsed = query("/Assumptions/cell[type=Number]");
    let Some(Scope::Path(scope)) = parsed.scope() else {
        panic!("path scope");
    };
    assert_eq!(scope.to_string(), "/Assumptions");
    query_shows(
        "/Assumptions/cell[type=Number]",
        "/Assumptions/cell[type=Number]",
    );
    query_shows("/body", "/body");
    let Some(Scope::Path(scope)) = query("/body").scope().cloned() else {
        panic!("root scope");
    };
    assert!(scope.is_root());
    query_shows(
        "/slide[1]/shape:contains(\"a/b\")",
        "/slide[1]/shape:contains(a/b)",
    );
    query_shows("/Sheet1/row[url~=\"x/y\"]", "/Sheet1/row[url~=x/y]");
    query_shows("/My Data (2024)/row[2]", "/\"My Data (2024)\"/row[2]");
    query_shows("/'a/b'/A1:C5", "/\"a/b\"/A1:C5");
    query_shows("/Sheet1/A1:C5", "/Sheet1/A1:C5");
    query_shows("/body/p[1]/r[bold=true]", "/body/p[1]/r[bold=true]");
}

#[test]
fn selector_syntax_errors_carry_offsets_and_codes() {
    let cases: &[(&str, ErrorKind, usize)] = &[
        ("   ", ErrorKind::Empty, 3),
        ("slide + shape", ErrorKind::UnsupportedCombinator('+'), 6),
        ("slide ~ shape", ErrorKind::UnsupportedCombinator('~'), 6),
        ("cell:foo(x)", ErrorKind::UnknownPseudoFunction, 5),
        ("ole[progId=Excel*]", ErrorKind::Wildcard, 16),
        ("p[a===b]", ErrorKind::MisparsedOperator, 4),
        ("x[name=r\"x\"]", ErrorKind::RegexRequiresContains, 7),
        ("cell:contains(a\"b)", ErrorKind::QuoteInBareValue, 15),
        ("'unterminated!A1", ErrorKind::UnterminatedQuote, 0),
        ("p[0]", ErrorKind::IndexZero, 2),
    ];
    for &(input, kind, offset) in cases {
        let error = query_error(input);
        assert_eq!(
            (error.kind(), error.offset()),
            (kind, offset),
            "{input:?}: {error}"
        );
        assert_eq!(error.code(), ErrorCode::InvalidSelector, "{input:?}");
        assert_eq!(error.code().as_str(), "invalid_selector");
    }
    let unexpected: &[(&str, Expected, Found, usize)] = &[
        ("p[style=Normal", Expected::CloseBracket, Found::End, 14),
        ("p[]", Expected::Predicate, Found::Char(']'), 2),
        ("> shape", Expected::Compound, Found::Char('>'), 0),
        ("a > > b", Expected::Compound, Found::Char('>'), 4),
        ("a,", Expected::Compound, Found::End, 2),
        ("a#b", Expected::CombinatorOrEnd, Found::Char('#'), 1),
        ("cell:contains", Expected::OpenParen, Found::End, 13),
        ("cell:contains()", Expected::Value, Found::Char(')'), 14),
        ("cell:has(a", Expected::CloseParen, Found::End, 10),
        ("cell:", Expected::PseudoName, Found::End, 5),
        ("'Sheet1'A1", Expected::SheetBang, Found::Char('A'), 8),
        ("Sheet1!", Expected::Compound, Found::End, 7),
        ("/", Expected::Compound, Found::End, 1),
        ("a)", Expected::CombinatorOrEnd, Found::Char(')'), 1),
    ];
    for &(input, expected, found, offset) in unexpected {
        let error = query_error(input);
        assert_eq!(
            (error.kind(), error.offset()),
            (ErrorKind::Unexpected { expected, found }, offset),
            "{input:?}: {error}"
        );
    }
    // A malformed scope path is reported as a selector error at its offset.
    let error = query_error("/a//b");
    assert_eq!(
        (error.kind(), error.offset(), error.code()),
        (ErrorKind::EmptySegment, 2, ErrorCode::InvalidSelector)
    );
    assert_eq!(
        query_error("slide + shape").to_string(),
        "invalid selector at byte 6: combinator '+' is not supported; use '>' or whitespace"
    );
}

#[test]
fn selector_limits_are_typed() {
    let nested = format!(
        "a{}b{}",
        ":not(".repeat(limits::MAX_NESTING + 1),
        ")".repeat(limits::MAX_NESTING + 1)
    );
    let error = query_error(&nested);
    assert_eq!(error.kind(), ErrorKind::LimitExceeded(Limit::Nesting));
    assert_eq!(error.code(), ErrorCode::MaxDepthExceeded);
    let ok = format!(
        "a{}b{}",
        ":not(".repeat(limits::MAX_NESTING),
        ")".repeat(limits::MAX_NESTING)
    );
    assert_eq!(query(&ok).to_string(), ok);

    let union = vec!["a"; limits::MAX_ALTERNATIVES + 1].join(",");
    assert_eq!(
        query_error(&union).kind(),
        ErrorKind::LimitExceeded(Limit::Alternatives)
    );
    let chain = vec!["a"; limits::MAX_COMPOUNDS + 1].join(" > ");
    assert_eq!(
        query_error(&chain).kind(),
        ErrorKind::LimitExceeded(Limit::Compounds)
    );
    assert!(Query::parse(&vec!["a"; limits::MAX_COMPOUNDS].join(" > ")).is_ok());
    let pseudos = format!("a{}", ":empty".repeat(limits::MAX_PREDICATES + 1));
    assert_eq!(
        query_error(&pseudos).kind(),
        ErrorKind::LimitExceeded(Limit::Predicates)
    );
}

#[test]
fn builders_render_like_parsed_values() {
    let built = Query::new([
        Complex::new(
            Compound::new(Element::Named("row".into())).with(Predicate::Filter(Expr::And(vec![
                cmp("Salary", CompareOp::Gt, "5000"),
                cmp("Region", CompareOp::Eq, "EMEA"),
            ]))),
        )
        .then(
            Combinator::Child,
            Compound::new(Element::Implicit).with_pseudo(Pseudo::Empty),
        ),
    ])
    .with_scope(Scope::Sheet("Q1 Data".into()));
    assert_eq!(
        built.to_string(),
        "'Q1 Data'!row[Salary>5000 and Region=EMEA] > :empty"
    );
    assert_eq!(
        query("'Q1 Data'!row[Salary>5000 and Region=EMEA] > :empty"),
        built
    );
}

#[test]
fn existing_selector_vocabulary_is_unchanged() {
    #[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
    struct SheetId(u64);
    let by_name: Selector<'_, SheetId> = "Summary".into();
    assert!(matches!(by_name, Selector::Name(name) if name == "Summary"));
}

#[test]
fn ast_values_are_thread_safe_and_sized_for_the_stack() {
    fn send_sync<T: Send + Sync>() {}
    send_sync::<Path<'static>>();
    send_sync::<Query<'static>>();
    send_sync::<ParseError>();
    // The inline step array keeps short paths off the heap; keep the value
    // small enough to move freely. Bounds are review triggers, not targets.
    assert!(size_of::<Path<'static>>() <= 1024);
    assert!(size_of::<Query<'static>>() <= 512);
    assert!(size_of::<ParseError>() <= 32);
}

#[test]
fn diagnostic_rendering_omits_authored_selector_content() {
    let path = Path::parse("/PRIVATE[@name=PRIVATE]").unwrap();
    let query = Query::parse("PRIVATE:contains(PRIVATE)").unwrap();
    for debug in [
        format!("{path:?}"),
        format!("{query:?}"),
        format!("{:?}", path.steps()[0]),
        format!("{:?}", path.steps()[0].predicates()[0]),
    ] {
        assert!(!debug.contains("PRIVATE"));
    }
    let error = Path::parse("PRIVATE").unwrap_err();
    assert!(!error.to_string().contains("'P'"));
    assert!(!format!("{error:?}").contains("'P'"));
    assert_eq!(path.to_string(), "/PRIVATE[@name=PRIVATE]");
}
