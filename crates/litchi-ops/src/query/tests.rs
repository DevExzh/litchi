//! Unit tests of query compilation and the walker over the fake formats
//! (record 0010 §5).

use super::compile::compile;
use super::fake::{Doc, FakeKind, Item, Nav, Node, SheetKind, SlideKind, Table, WordKind};
use super::plan::{Positional, Structural};
use super::walk::{PlanRef, Walk};
use super::{Diagnostic, DiagnosticKind, IntoQuery, Plan, Query, QueryArg, run};
use crate::Code;
use crate::error::{Error, Result, Source};
use crate::path::PathArg;
use crate::path::resolve;
use crate::value::{Color, Length, Scheme, Transform, TransformKind};

// ---------------------------------------------------------------------------
// Fixtures
// ---------------------------------------------------------------------------

fn p(text: &str) -> Item<WordKind> {
    Item::new(WordKind::Paragraph).text(text)
}

fn r(text: &str) -> Item<WordKind> {
    Item::new(WordKind::Run).text(text)
}

fn tc(text: &str) -> Item<WordKind> {
    Item::new(WordKind::Cell).child(p(text).child(r(text)))
}

/// A report with headings, runs, a hyperlink, a table holding a nested
/// table, an empty paragraph, a content control, pictures, a header and a
/// footer.
fn report() -> Doc<WordKind> {
    let nested = Item::new(WordKind::Table).child(Item::new(WordKind::Row).child(tc("nested")));
    let table = Item::new(WordKind::Table)
        .prop("style", "Grid")
        .child(
            Item::new(WordKind::Row)
                .child(
                    Item::new(WordKind::Cell)
                        .child(p("Region").child(r("Region").prop("bold", true))),
                )
                .child(tc("Sales")),
        )
        .child(
            Item::new(WordKind::Row)
                .child(tc("West"))
                .child(tc("1200").child(nested)),
        );
    let body = Item::new(WordKind::Body)
        .child(
            p("Quarterly Report")
                .style("Heading1")
                .prop("style", "Heading1")
                .prop("paraId", "1A2B3C4D")
                .child(
                    r("Quarterly ")
                        .prop("bold", true)
                        .prop("size", Length::from_points(14)),
                )
                .child(r("Report").prop("italic", true)),
        )
        .child(
            p("Revenue grew 25% in APAC")
                .style("Normal")
                .prop("style", "Normal")
                .child(r("Revenue grew "))
                .child(
                    Item::new(WordKind::Hyperlink)
                        .prop("url", "https://example.com")
                        .child(
                            r("25%")
                                .prop("bold", true)
                                .prop("color", Color::rgb(0xFF, 0, 0)),
                        ),
                )
                .child(r(" in APAC").prop("size", Length::from_half_points(21))),
        )
        .child(table)
        .child(p("").style("Normal").prop("style", "Normal"))
        .child(
            Item::new(WordKind::Sdt)
                .prop("tag", "summary")
                .child(p("In summary").prop("style", "Heading2")),
        )
        .child(
            p("Outlook")
                .style("Heading1")
                .prop("style", "Heading1")
                .prop("numId", 3)
                .prop("keepNext", true)
                .prop("spaceBefore", Length::from_twips(240))
                .child(r("Outlook").prop("bold", true))
                .child(Item::new(WordKind::Picture).prop("name", "Chart image")),
        )
        .child(
            p("Logo").child(
                Item::new(WordKind::Picture)
                    .prop("name", "Logo")
                    .prop("alt", "Company logo"),
            ),
        );
    Doc::new(
        Item::new(WordKind::Document)
            .prop("title", "Q4 Report")
            .child(body)
            .child(
                Item::new(WordKind::Header)
                    .prop("type", crate::node::Value::Token("default"))
                    .child(p("Confidential")),
            )
            .child(Item::new(WordKind::Footer).child(p("Page 1"))),
    )
}

/// Two slides: titles, text boxes, pictures with and without alt text, a
/// table, a group and notes.
fn deck() -> Doc<SlideKind> {
    let accent = Color::scheme(Scheme::Accent1)
        .with_transforms(&[
            Transform::new(TransformKind::LumMod, 75_000).expect("in range"),
            Transform::new(TransformKind::LumOff, 25_000).expect("in range"),
        ])
        .expect("fits");
    let shape = |flavour: &'static str, title: bool, name: &str, text: &str| {
        Item::new(SlideKind::Shape)
            .flavour(flavour)
            .text(text)
            .prop("name", name.to_owned())
            .prop("isTitle", title)
    };
    let slide1 = Item::new(SlideKind::Slide)
        .prop("layout", "Title Slide")
        .child(shape("title", true, "Title 1", "Q4 Report").prop("id", 2))
        .child(
            shape("textbox", false, "TextBox 2", "Revenue grew 25%")
                .prop("id", 3)
                .prop("x", Length::from_emu(720_000))
                .prop("width", Length::from_emu(3_810_000))
                .prop("fill", Color::rgb(0xFF, 0, 0)),
        )
        .child(
            Item::new(SlideKind::Picture)
                .prop("name", "Logo")
                .prop("alt", ""),
        )
        .child(
            Item::new(SlideKind::Picture)
                .prop("name", "Photo")
                .prop("alt", "A photo"),
        )
        .child(
            Item::new(SlideKind::Table).child(Item::new(SlideKind::Row).child(
                Item::new(SlideKind::Cell).child(Item::new(SlideKind::Paragraph).text("Region")),
            )),
        );
    let slide2 = Item::new(SlideKind::Slide)
        .child(shape("title", true, "Title 1", "Outlook"))
        .child(
            Item::new(SlideKind::Group)
                .child(
                    shape("textbox", false, "Grouped", "Grouped text")
                        .prop("fill", accent)
                        .prop("width", Length::from_emu(360_000 * 12)),
                )
                .child(Item::new(SlideKind::Picture).prop("name", "Icon")),
        )
        .child(
            Item::new(SlideKind::Notes)
                .child(Item::new(SlideKind::Paragraph).text("Speaker notes")),
        );
    Doc::new(
        Item::new(SlideKind::Presentation)
            .child(slide1)
            .child(slide2),
    )
}

fn cell(a1: &str, text: &str) -> Item<SheetKind> {
    Item::new(SheetKind::Cell).cell(a1).text(text)
}

fn row(number: usize) -> Item<SheetKind> {
    Item::new(SheetKind::Row).number(number)
}

/// A workbook: `Data` holds a `Sales` table over A1:C5 whose row 4 has no
/// record and whose B5 is a percentage, plus a formula row; `Notes` holds
/// one cell.
fn book() -> Doc<SheetKind> {
    let data = Item::new(SheetKind::Sheet)
        .name("Data")
        .child(
            Item::new(SheetKind::Column)
                .number(2)
                .prop("width", crate::node::Value::Number(20.0)),
        )
        .child(row(1))
        .child(cell("A1", "Region"))
        .child(cell("B1", "Sales"))
        .child(cell("C1", "Status"))
        .child(row(2).prop("height", crate::node::Value::Number(30.0)))
        .child(cell("A2", "West"))
        .child(
            cell("B2", "1200")
                .raw(crate::node::Value::Number(1200.0))
                .prop("type", crate::node::Value::Token("Number")),
        )
        .child(cell("C2", "open"))
        .child(row(3).prop("hidden", true))
        .child(cell("A3", "East"))
        .child(
            cell("B3", "800")
                .raw(crate::node::Value::Number(800.0))
                .prop("type", crate::node::Value::Token("Number")),
        )
        .child(cell("C3", "closed"))
        .child(row(5))
        .child(cell("A5", "West"))
        .child(cell("B5", "50%").raw(crate::node::Value::Number(0.5)))
        .child(cell("C5", "open"))
        .child(row(6))
        .child(
            cell("B6", "2000.5")
                .raw(crate::node::Value::Number(2000.5))
                .prop("formula", "SUM(B2:B5)")
                .prop("type", crate::node::Value::Token("Number"))
                .prop("font.bold", true),
        );
    let notes = Item::new(SheetKind::Sheet)
        .name("Notes")
        .child(row(1))
        .child(cell("A1", "TODO: check West"));
    Doc::new(
        Item::new(SheetKind::Workbook)
            .child(data)
            .child(notes)
            .child(Item::new(SheetKind::NamedRange).prop("name", "Total")),
    )
    .table(Table {
        sheet: 0,

        columns: vec!["Region".into(), "Sales".into(), "Status".into()],
        first_column: 1,
        header_row: 1,
        last_row: 5,
        totals: false,
    })
}

// ---------------------------------------------------------------------------
// Helpers
// ---------------------------------------------------------------------------

fn compile_text<K: FakeKind>(text: &str) -> Result<Plan<'_>> {
    let query = Query::parse(text)?;
    compile::<&Nav<'static, K>>(&query)
}

fn compiled<K: FakeKind>(text: &str) -> Plan<'_> {
    compile_text::<K>(text).unwrap_or_else(|error| panic!("{text:?}: {error}"))
}

/// The paths of every match of `text`, in order.
fn paths<K: FakeKind>(doc: &Doc<K>, text: &str) -> Vec<String> {
    let nav = Nav::new(doc);
    let mut walk = run(&nav, text).unwrap_or_else(|error| panic!("{text:?}: {error}"));
    let mut out = Vec::new();
    while let Some(hit) = walk.next_hit() {
        let hit = hit.unwrap_or_else(|error| panic!("{text:?}: {error}"));
        out.push(hit.path().to_string());
    }
    out
}

/// The error of compiling or running `text`.
fn fail<K: FakeKind>(doc: &Doc<K>, text: &str) -> Error {
    let nav = Nav::new(doc);
    match run(&nav, text) {
        Err(error) => error,
        Ok(walk) => {
            for item in walk {
                if let Err(error) = item {
                    return error;
                }
            }
            panic!("{text:?} succeeded")
        },
    }
}

fn count<K: FakeKind>(doc: &Doc<K>, text: &str) -> usize {
    paths(doc, text).len()
}

fn assert_paths<K: FakeKind>(doc: &Doc<K>, text: &str, expected: &[&str]) {
    assert_eq!(paths(doc, text), expected, "{text}");
}

// ---------------------------------------------------------------------------
// Element tests and names
// ---------------------------------------------------------------------------

#[test]
fn element_names_go_through_the_alias_table_in_any_case() {
    let doc = report();
    let paragraphs = paths(&doc, "p");
    assert_eq!(paragraphs.len(), 13, "{paragraphs:?}");
    for alias in ["paragraph", "P", "Paragraph", "PARAGRAPH"] {
        assert_eq!(paths(&doc, alias), paragraphs, "{alias}");
    }
    assert_eq!(paths(&doc, "table"), paths(&doc, "tbl"));
    assert_eq!(paths(&doc, "cell"), paths(&doc, "TC"));
    assert_eq!(paths(&doc, "image"), paths(&doc, "picture"));
    assert_eq!(paths(&doc, "contentcontrol"), ["/body/sdt[1]"]);
}

#[test]
fn matches_come_in_document_order_with_incremental_paths() {
    let doc = report();
    assert_paths(
        &doc,
        "p",
        &[
            "/body/p[1]",
            "/body/p[2]",
            "/body/tbl[1]/tr[1]/tc[1]/p[1]",
            "/body/tbl[1]/tr[1]/tc[2]/p[1]",
            "/body/tbl[1]/tr[2]/tc[1]/p[1]",
            "/body/tbl[1]/tr[2]/tc[2]/p[1]",
            "/body/tbl[1]/tr[2]/tc[2]/tbl[1]/tr[1]/tc[1]/p[1]",
            "/body/p[3]",
            "/body/sdt[1]/p[1]",
            "/body/p[4]",
            "/body/p[5]",
            "/header[1]/p[1]",
            "/footer[1]/p[1]",
        ],
    );
}

#[test]
fn every_match_path_resolves_back_to_the_match() {
    let doc = report();
    let nav = Nav::new(&doc);
    for selector in ["p", "r", "tc", "picture", "hyperlink > r", "*"] {
        let mut walk = run(&nav, selector).expect(selector);
        while let Some(hit) = walk.next_hit() {
            let hit = hit.expect("hit");
            let path = hit.path().to_string();
            let arg = PathArg::parse(&path).unwrap_or_else(|error| panic!("{path}: {error}"));
            let found = resolve::get(&&nav, &arg).unwrap_or_else(|error| panic!("{path}: {error}"));
            let (Node::Item { item: a, .. }, Node::Item { item: b, .. }) = (found, *hit.node())
            else {
                panic!("{path} is not an element");
            };
            assert!(std::ptr::eq(a, b), "{selector}: {path} resolves elsewhere");
            assert_eq!(hit.kind(), a.kind());
        }
    }
}

#[test]
fn unknown_and_foreign_element_names_are_typed_errors() {
    let doc = report();
    let error = fail(&doc, "paragrph");
    assert_eq!(error.code(), Code::InvalidSelector);
    assert_eq!(error.suggestion(), Some("did you mean \"paragraph\"?"));
    let error = fail(&doc, "slide");
    assert_eq!(error.code(), Code::UnsupportedPath);
    assert!(error.message().contains("presentations"), "{error}");
    assert_eq!(error.format(), Some(litchi_core::FileFormat::Docx));
    // Known to the vocabulary, absent from the format.
    assert_eq!(fail(&doc, "customXml").code(), Code::UnsupportedPath);
    // Known, but the adapter does not model it.
    let error = fail(&doc, "toc");
    assert_eq!(error.code(), Code::UnsupportedCapability);
    // Modeled, but listed as a capability gap.
    let error = fail(&doc, "chart");
    assert_eq!(error.code(), Code::UnsupportedCapability);
    assert!(
        error
            .suggestion()
            .is_some_and(|text| text.contains("L-FAKE-R1"))
    );
    assert_eq!(fail(&book(), "slicer").code(), Code::UnsupportedCapability);
    assert_eq!(fail(&book(), "pivot").code(), Code::UnsupportedCapability);
}

#[test]
fn star_selects_top_level_elements_or_the_scopes_children() {
    let doc = report();
    assert_paths(
        &doc,
        "*",
        &[
            "/body/p[1]",
            "/body/p[2]",
            "/body/tbl[1]",
            "/body/p[3]",
            "/body/sdt[1]",
            "/body/p[4]",
            "/body/p[5]",
        ],
    );
    assert_paths(
        &doc,
        "/body/tbl[1]/*",
        &["/body/tbl[1]/tr[1]", "/body/tbl[1]/tr[2]"],
    );
    // A lone `*` enters nothing below the elements it lists.
    let nav = Nav::new(&doc);
    assert_eq!(run(&nav, "*").expect("runs").count(), 7);
    assert_eq!(
        nav.produced(),
        3 + 7,
        "the root's stories and the body's blocks"
    );
    let nav = Nav::new(&doc);
    assert_eq!(run(&nav, "/body/tbl[1]/*").expect("runs").count(), 2);
    assert_eq!(nav.produced(), 2);
    assert_eq!(count(&deck(), "*"), 8);
    assert_eq!(paths(&deck(), "*")[0], "/slide[1]/shape[1]");
}

#[test]
fn implicit_elements_test_every_kind() {
    let doc = report();
    // `[style=Heading1]` matches any element whose `style` is Heading1.
    assert_paths(&doc, "[style=Heading1]", &["/body/p[1]", "/body/p[4]"]);
    // `[2]` alone is the second element of the walk.
    assert_paths(&doc, "[2]", &["/body/p[1]"]);
}

#[test]
fn title_and_textbox_narrow_presentation_shapes() {
    let doc = deck();
    assert_paths(&doc, "title", &["/slide[1]/shape[1]", "/slide[2]/shape[1]"]);
    assert_paths(
        &doc,
        "textbox",
        &["/slide[1]/shape[2]", "/slide[2]/group[1]/shape[1]"],
    );
    assert_eq!(count(&doc, "shape"), 4);
    assert_eq!(count(&doc, "Title"), 2);
    // `title` and `textbox` keep their path meaning: plain shapes.
    let nav = Nav::new(&doc);
    let found = resolve::get(
        &&nav,
        &PathArg::parse("/slide[1]/textbox[1]").expect("parses"),
    )
    .expect("resolves");
    assert_eq!(found.item().and_then(Item::text_str), Some("Q4 Report"));
    // OfficeCLI's `title` key.
    assert_eq!(count(&doc, "shape[title=true]"), 2);
    assert_eq!(count(&doc, "shape[isTitle=false]"), 2);
}

#[test]
fn workbook_selectors_take_columns_and_ranges() {
    let doc = book();
    assert_paths(
        &doc,
        "B",
        &["/Data/B1", "/Data/B2", "/Data/B3", "/Data/B5", "/Data/B6"],
    );
    assert_paths(
        &doc,
        "A2:B3",
        &["/Data/A2", "/Data/B2", "/Data/A3", "/Data/B3"],
    );
    assert_paths(
        &doc,
        "B3:A2",
        &["/Data/A2", "/Data/B2", "/Data/A3", "/Data/B3"],
    );
    assert_paths(&doc, "Data!A1", &["/Data/A1"]);
    assert_paths(&doc, "Notes!A1", &["/Notes/A1"]);
    // Element names win over column letters (`row` and `col` are columns).
    assert_paths(&doc, "col", &["/Data/col[2]"]);
    assert_eq!(count(&doc, "row"), 6);
}

// ---------------------------------------------------------------------------
// Comparisons
// ---------------------------------------------------------------------------

#[test]
fn operators_follow_officecli_semantics() {
    let doc = report();
    assert_eq!(count(&doc, "p[style=heading1]"), 2, "= ignores case");
    assert_eq!(
        count(&doc, "p[style!=Heading1]"),
        11,
        "!= keeps absent keys"
    );
    assert_eq!(count(&doc, "p[style~=ead]"), 3, "~= is a substring");
    assert_eq!(count(&doc, "p[numId>2]"), 1);
    assert_eq!(count(&doc, "p[numId>=3]"), 1);
    assert_eq!(count(&doc, "p[numId<3]"), 0);
    assert_eq!(count(&doc, "p[numId<=3]"), 1);
    assert_eq!(count(&doc, "p[numId]"), 1, "[key] is present");
    assert_eq!(count(&doc, "p[keepNext=true]"), 1);
    assert_eq!(
        count(&doc, "p[keepNext=yes]"),
        1,
        "typed equality coerces the literal"
    );
    assert_eq!(
        run(
            &Nav::new(&doc),
            super::raw::query("p[@paraId=1a2b3c4d]").unwrap()
        )
        .unwrap()
        .count(),
        1,
        "@key compares like key"
    );
    assert_eq!(count(&doc, "p[style=Heading1 and keepNext=true]"), 1);
    assert_eq!(count(&doc, "p[style=Heading2 or keepNext=true]"), 2);
    assert_eq!(count(&doc, "p[not(style=Normal)]"), 11);
    assert_eq!(
        count(&doc, "p[style=Heading1][keepNext]"),
        1,
        "stacked brackets AND"
    );
}

#[test]
fn lengths_and_colors_compare_typed() {
    let doc = report();
    assert_eq!(count(&doc, "r[size>=14pt]"), 1);
    assert_eq!(count(&doc, "r[size>10pt]"), 2);
    assert_eq!(count(&doc, "r[size=10.5pt]"), 1);
    assert_eq!(
        count(&doc, "r[size=21]"),
        0,
        "a font size reads points, not half-points"
    );
    assert_eq!(
        count(&doc, "r[color=red]"),
        1,
        "a named colour equals its RGB"
    );
    assert_eq!(count(&doc, "r[color=#ff0000]"), 1);
    assert_eq!(count(&doc, "r[color=FF0000]"), 1);
    assert_eq!(count(&doc, "p[spaceBefore=12pt]"), 1);
    assert_eq!(
        count(&doc, "p[spaceBefore=0.4233cm]"),
        1,
        "lengths compare in EMU within 500"
    );
    assert_eq!(count(&doc, "p[spaceBefore=0.42cm]"), 0);
    let deck = deck();
    assert_eq!(count(&deck, "shape[width>=11cm]"), 1);
    assert_eq!(count(&deck, "shape[width>=4cm]"), 2);
    assert_eq!(count(&deck, "shape[x=2cm]"), 1);
    assert_eq!(count(&deck, "shape[fill=red]"), 1);
    assert_eq!(count(&deck, "shape[fill=accent1+lumMod75+lumOff25]"), 1);
    assert_eq!(count(&deck, "shape[fill~=accent1]"), 1);
}

#[test]
fn pseudo_keys_read_text_type_and_style() {
    let doc = report();
    assert_eq!(count(&doc, "p[text~=revenue]"), 1);
    assert_eq!(count(&doc, "p[type=paragraph]"), 13);
    assert_eq!(count(&doc, "[type=hyperlink]"), 1);
    let deck = deck();
    assert_eq!(count(&deck, "shape[type=title]"), 2);
    assert_eq!(count(&deck, "shape[type=textbox]"), 2);
    // OfficeCLI's dual style key: the style property or the style member.
    let styled = Doc::new(
        Item::new(WordKind::Document).child(
            Item::new(WordKind::Body)
                .child(p("member only").style("Quote"))
                .child(p("property only").prop("style", "Quote"))
                .child(p("both").style("Quote").prop("style", "Quote"))
                .child(p("neither")),
        ),
    );
    assert_eq!(count(&styled, "p[style=quote]"), 3);
    assert_eq!(count(&styled, "p[style!=Quote]"), 1);
    assert_eq!(count(&styled, "p[style~=quo]"), 3);
}

#[test]
fn workbook_cell_keys_take_officecli_aliases_and_value() {
    let doc = book();
    assert_paths(&doc, "cell[bold=true]", &["/Data/B6"]);
    assert_paths(&doc, "cell[font.bold=true]", &["/Data/B6"]);
    assert_paths(&doc, "cell[formula~=SUM]", &["/Data/B6"]);
    assert_paths(&doc, "cell[formula]", &["/Data/B6"]);
    assert_paths(&doc, "cell[value=West]", &["/Data/A2", "/Data/A5"]);
    assert_paths(&doc, "cell[value>1000]", &["/Data/B2", "/Data/B6"]);
    assert_paths(&doc, "cell[value>1000 and value<2000]", &["/Data/B2"]);
    assert_eq!(count(&doc, "cell[type=Number]"), 3);
    assert_paths(
        &doc,
        "Data!B[value!=1200]",
        &["/Data/B1", "/Data/B3", "/Data/B5", "/Data/B6"],
    );
}

#[test]
fn regex_literals_are_refused() {
    let doc = report();
    let error = fail(&doc, r#"p[text~=r"^Q[1-4]"]"#);
    assert_eq!(error.code(), Code::UnsupportedCapability);
    assert!(error.message().contains("regular-expression"), "{error}");
    assert!(error.suggestion().is_some_and(|text| text.contains("~=")));
    // Refused at compile time, before the document is read.
    let error = compile_text::<WordKind>(r#"shape[text~=r"^Q"], p"#).unwrap_err();
    assert_eq!(error.code(), Code::UnsupportedCapability);
}

#[test]
fn unknown_keys_are_warnings_that_evaluate_as_absent() {
    let plan = compiled::<WordKind>("r[blod=true]");
    let [diagnostic] = plan.diagnostics() else {
        panic!("{:?}", plan.diagnostics());
    };
    assert_eq!(diagnostic.kind(), DiagnosticKind::UnknownKey);
    assert_eq!(diagnostic.suggestion(), Some("did you mean 'bold'?"));
    assert!(diagnostic.message().contains("available:"), "{diagnostic}");
    assert!(diagnostic.to_string().starts_with("unknown_key: "));
    let doc = report();
    assert_eq!(count(&doc, "r[blod=true]"), 0);
    assert_eq!(count(&doc, "r[blod!=true]"), count(&doc, "r"));
    // `!=` on an unknown key is no warning: an absent key satisfies it.
    assert!(
        compiled::<WordKind>("r[blod!=true]")
            .diagnostics()
            .is_empty()
    );
    // A dotted key is matched by its last segment too.
    let plan = compiled::<SheetKind>("cell[font.blod=true]");
    assert_eq!(
        plan.diagnostics()[0].suggestion(),
        Some("did you mean 'font.bold'?")
    );
    // Relational comparisons with a non-numeric literal.
    let plan = compiled::<WordKind>("p[numId>high]");
    assert_eq!(plan.diagnostics()[0].kind(), DiagnosticKind::NonNumeric);
    assert_eq!(count(&doc, "p[numId>high]"), 0);
    // A key the format lists as a gap is an error, not a warning.
    assert_eq!(
        compile_text::<WordKind>("r[highlight=yellow]")
            .unwrap_err()
            .code(),
        Code::UnsupportedCapability
    );
}

// ---------------------------------------------------------------------------
// Pseudo-classes
// ---------------------------------------------------------------------------

#[test]
fn contains_and_empty_test_the_text() {
    let doc = report();
    assert_paths(&doc, "p:contains(REVENUE)", &["/body/p[2]"]);
    assert_paths(&doc, r#"p:contains("grew 25%")"#, &["/body/p[2]"]);
    assert_paths(&doc, "p:revenue", &["/body/p[2]"]);
    assert_paths(&doc, "p:contains(grew):contains(APAC)", &["/body/p[2]"]);
    assert_paths(&doc, "p:empty", &["/body/p[3]"]);
    assert_eq!(count(&doc, "p:not(:empty)"), 12);
    assert_eq!(
        count(&doc, "cell:contains(west)"),
        0,
        "a Word cell has no text of its own"
    );
    assert_eq!(count(&book(), "cell:contains(west)"), 3);
    assert_eq!(count(&book(), "cell:TODO"), 1);
}

#[test]
fn no_alt_tests_alternative_text() {
    assert_paths(&report(), "picture:no-alt", &["/body/p[4]/picture[1]"]);
    assert_paths(&report(), "image:no-alt", &["/body/p[4]/picture[1]"]);
    let deck = deck();
    assert_paths(
        &deck,
        "picture:no-alt",
        &["/slide[1]/picture[1]", "/slide[2]/group[1]/picture[1]"],
    );
    // A kind without alternative text never matches, with a diagnostic.
    let plan = compiled::<WordKind>("p:no-alt");
    assert_eq!(plan.diagnostics()[0].kind(), DiagnosticKind::NeverMatches);
    assert_eq!(count(&report(), "p:no-alt"), 0);
    // A kind whose `alt` is a listed gap is unsupported.
    assert_eq!(
        compile_text::<SlideKind>("group:no-alt")
            .unwrap_err()
            .code(),
        Code::UnsupportedCapability
    );
}

#[test]
fn has_searches_the_subtree_and_stops_at_the_first_hit() {
    let doc = report();
    assert_paths(&doc, "p:has(hyperlink)", &["/body/p[2]"]);
    assert_paths(
        &doc,
        "p:has(r[bold=true])",
        &[
            "/body/p[1]",
            "/body/p[2]",
            "/body/tbl[1]/tr[1]/tc[1]/p[1]",
            "/body/p[4]",
        ],
    );
    assert_paths(&doc, "tbl:has(tbl)", &["/body/tbl[1]"]);
    assert_paths(&doc, "tc:has(tbl tc)", &["/body/tbl[1]/tr[2]/tc[2]"]);
    assert_paths(
        &doc,
        "tbl:has(tr)",
        &["/body/tbl[1]", "/body/tbl[1]/tr[2]/tc[2]/tbl[1]"],
    );
    // A relative `> tr` is not selector syntax: `:has` takes selectors.
    assert_eq!(fail(&doc, "tbl:has(> tr)").code(), Code::InvalidSelector);
    assert_eq!(count(&doc, "p:has(picture[alt])"), 1);
    assert_eq!(count(&doc, "r:has(p)"), 0, "runs contain no paragraphs");
    // A `:has` on an ancestor compound searches that ancestor's subtree, and
    // combinators inside it see the ancestor's true ancestors: every cell
    // paragraph is `tr > tc > p`, so no cell qualifies here...
    assert_eq!(count(&doc, "tc:has(p:not(tr > tc > p)) > p"), 0);
    // ...while only the cells holding the nested table's paragraph have one
    // that is not a direct cell paragraph of a body-level table.
    assert_paths(
        &doc,
        "tc:has(p:not(body > tbl > tr > tc > p)) > p",
        &[
            "/body/tbl[1]/tr[2]/tc[2]/p[1]",
            "/body/tbl[1]/tr[2]/tc[2]/tbl[1]/tr[1]/tc[1]/p[1]",
        ],
    );
    // OfficeCLI's literal `cell:has(formula)` is the key test `[formula]`.
    let book = book();
    assert_paths(&book, "cell:has(formula)", &["/Data/B6"]);
    assert_paths(&book, "Data!:has(formula)", &["/Data/B6"]);
    assert_eq!(count(&book, "cell:not(:has(formula))"), 13);
}

#[test]
fn not_matches_the_candidate_itself() {
    let doc = report();
    assert_eq!(count(&doc, "p:not([style=Normal])"), 11);
    assert_eq!(count(&doc, "p:not(tc p)"), 8);
    assert_eq!(count(&doc, "r:not(hyperlink > r, p > r)"), 0);
    assert_eq!(count(&doc, "r:not(hyperlink r)"), count(&doc, "r") - 1);
    assert_paths(
        &doc,
        "p:not(:contains(e))",
        &[
            "/body/tbl[1]/tr[2]/tc[2]/p[1]",
            "/body/p[3]",
            "/body/sdt[1]/p[1]",
            "/body/p[4]",
            "/body/p[5]",
        ],
    );
}

// ---------------------------------------------------------------------------
// Combinators and unions
// ---------------------------------------------------------------------------

#[test]
fn combinators_chain_freely() {
    let doc = report();
    assert_paths(&doc, "hyperlink > r", &["/body/p[2]/hyperlink[1]/r[1]"]);
    assert_eq!(count(&doc, "p > r"), count(&doc, "r") - 1);
    assert_eq!(count(&doc, "body > p"), 5);
    assert_eq!(count(&doc, "body p"), 11);
    assert_eq!(count(&doc, "tbl p"), 5);
    assert_eq!(count(&doc, "tbl tbl p"), 1);
    assert_eq!(count(&doc, "tbl > tr > tc > p"), 5);
    assert_eq!(count(&doc, "body > tbl > tr > tc > p"), 4);
    assert_eq!(count(&doc, "sdt p"), 1);
    assert_eq!(count(&doc, "header p"), 1);
    assert_paths(
        &doc,
        "paragraph[style=Heading1] > run[bold=true]",
        &["/body/p[1]/r[1]", "/body/p[4]/r[1]"],
    );
    // The nearest `tc` has no `tbl` parent above a second `tbl`, but a
    // farther one does: the chain backtracks.
    assert_eq!(count(&doc, "tbl > tr tc p"), 5);
    assert_eq!(count(&doc, "body > tbl tc > p"), 5);
    assert_eq!(count(&doc, "tr > tc > tbl tc p"), 1);
    let deck = deck();
    assert_eq!(count(&deck, "slide > shape"), 3);
    assert_eq!(count(&deck, "slide shape"), 4);
    assert_eq!(count(&deck, "slide > table > tr"), 1);
    assert_eq!(count(&deck, "slide picture"), 3);
}

#[test]
fn ancestor_positions_select_among_siblings() {
    let doc = report();
    assert_eq!(count(&doc, "tbl[1] > tr"), 3, "every first table's rows");
    assert_eq!(count(&doc, "body > tbl[1] > tr"), 2);
    assert_paths(
        &doc,
        "tr[2] > tc[1] > p",
        &["/body/tbl[1]/tr[2]/tc[1]/p[1]"],
    );
    assert_paths(
        &doc,
        "tr[last()] > tc[last()] > p",
        &[
            "/body/tbl[1]/tr[2]/tc[2]/p[1]",
            "/body/tbl[1]/tr[2]/tc[2]/tbl[1]/tr[1]/tc[1]/p[1]",
        ],
    );
    let deck = deck();
    assert_eq!(count(&deck, "slide[2] shape"), 2);
    assert_paths(&deck, "slide[1] > shape[2]", &["/slide[1]/shape[2]"]);
    assert_paths(
        &deck,
        "slide[1]>picture",
        &["/slide[1]/picture[1]", "/slide[1]/picture[2]"],
    );
}

#[test]
fn unions_yield_each_node_once_in_document_order() {
    let doc = report();
    assert_paths(
        &doc,
        "picture, p[style=Heading1], picture[name=Logo]",
        &[
            "/body/p[1]",
            "/body/p[4]",
            "/body/p[4]/picture[1]",
            "/body/p[5]/picture[1]",
        ],
    );
    assert_eq!(count(&doc, "paragraph, table"), count(&doc, "p") + 2);
    // Each alternative counts its own `[N]`.
    assert_paths(&doc, "p[2], tbl[1]", &["/body/p[2]", "/body/tbl[1]"]);
    let error = fail(&doc, "p[last()], tbl");
    assert_eq!(error.code(), Code::InvalidSelector);
}

// ---------------------------------------------------------------------------
// Positions, numbers and limits
// ---------------------------------------------------------------------------

#[test]
fn subject_positions_select_the_nth_or_last_match() {
    let doc = report();
    assert_paths(&doc, "p[2]", &["/body/p[2]"]);
    assert_paths(&doc, "p[3]", &["/body/tbl[1]/tr[1]/tc[1]/p[1]"]);
    assert_paths(&doc, "p[style=Heading1][2]", &["/body/p[4]"]);
    assert_paths(&doc, "p[2][style=Heading1]", &["/body/p[4]"]);
    assert_paths(&doc, "p[99]", &[]);
    assert_paths(&doc, "p[last()]", &["/footer[1]/p[1]"]);
    assert_paths(&doc, "p[style=Heading1][last()]", &["/body/p[4]"]);
    assert_paths(&doc, "tbl[last()]", &["/body/tbl[1]/tr[2]/tc[2]/tbl[1]"]);
    assert_paths(
        &doc,
        "tc p[last()]",
        &["/body/tbl[1]/tr[2]/tc[2]/tbl[1]/tr[1]/tc[1]/p[1]"],
    );
    let error = fail(&doc, "p[1][2]");
    assert_eq!(error.code(), Code::InvalidSelector);
    // `[N]` ends the walk once found.
    let nav = Nav::new(&doc);
    let mut walk = run(&nav, "p[1]").expect("runs");
    assert!(walk.next().is_some());
    let visited = walk.visited();
    assert!(walk.next().is_none());
    assert_eq!(
        walk.visited(),
        visited,
        "nothing is visited after the Nth match"
    );
}

#[test]
fn workbook_rows_and_columns_keep_their_numbers() {
    let doc = book();
    assert_paths(&doc, "row[3]", &["/Data/row[3]"]);
    assert_paths(&doc, "row[4]", &[]);
    assert_paths(&doc, "Data!row[5]", &["/Data/row[5]"]);
    assert_paths(&doc, "row[last()]", &["/Notes/row[1]"]);
    assert_paths(&doc, "col[B]", &["/Data/col[2]"]);
    assert_paths(&doc, "col[2]", &["/Data/col[2]"]);
    assert_paths(&doc, "col[C]", &[]);
    assert_eq!(fail(&doc, "col[XFE]").code(), Code::InvalidSelector);
    assert_paths(&doc, "row[hidden=true]", &["/Data/row[3]"]);
    assert_paths(&doc, "row[height>20]", &["/Data/row[2]"]);
}

#[test]
fn limit_stops_the_walk_at_the_first_match() {
    // The first match is the 7th paragraph whatever the document's size:
    // the walk visits the body and seven paragraphs, then stops.
    for paragraphs in [10, 1_000, 20_000] {
        let doc = generated(paragraphs, 7);
        let nav = Nav::new(&doc);
        let walk = run(&nav, "p[style=Target]").expect("runs").limit(1);
        let mut walk = walk;
        let first = walk.next().expect("a match").expect("readable");
        assert_eq!(first.item().and_then(Item::text_str), Some("paragraph 6"));
        assert!(walk.next().is_none(), "limit(1) yields one match");
        assert_eq!(walk.visited(), 1 + 7, "{paragraphs} paragraphs");
        assert_eq!(nav.produced(), 1 + 7);
    }
    // `first()` is `limit(1)`.
    let doc = generated(1_000, 3);
    let nav = Nav::new(&doc);
    let first = run(&nav, "p[style=Target]")
        .expect("runs")
        .first()
        .expect("reads");
    assert_eq!(
        first.and_then(|node| node.item().and_then(Item::text_str)),
        Some("paragraph 2")
    );
    assert_eq!(nav.produced(), 1 + 3);
    // Without a limit, the walk visits everything it cannot prune.
    let nav = Nav::new(&doc);
    let all = run(&nav, "p[style=Target]").expect("runs").count();
    assert_eq!(all, 1);
    assert_eq!(nav.produced(), 1 + 1_000);
}

/// A body of `paragraphs` paragraphs with runs; paragraph `target` (1-based)
/// has style `Target`.
fn generated(paragraphs: usize, target: usize) -> Doc<WordKind> {
    Doc::new(
        Item::new(WordKind::Document).child(Item::new(WordKind::Body).children(
            (0..paragraphs).map(|index| {
                let style = if index + 1 == target {
                    "Target"
                } else {
                    "Normal"
                };
                p(&format!("paragraph {index}"))
                    .prop("style", style)
                    .child(r("text"))
                    .child(r("more"))
            }),
        )),
    )
}

#[test]
fn pruning_skips_subtrees_that_cannot_hold_a_subject() {
    let doc = report();
    let nav = Nav::new(&doc);
    // Paragraphs: runs, hyperlinks and pictures are never entered, so the
    // walk asks only for paragraph containers.
    let count = run(&nav, "p").expect("runs").count();
    assert_eq!(count, 13);
    let visited_p = nav.produced();
    let nav = Nav::new(&doc);
    let _ = run(&nav, "r").expect("runs").count();
    assert!(nav.produced() > visited_p, "runs need paragraphs entered");
    // A table query enters only containers of tables.
    let nav = Nav::new(&doc);
    let tables = run(&nav, "tbl").expect("runs").count();
    assert_eq!(tables, 2);
    assert!(nav.produced() < visited_p);
    // Header queries never enter the body's paragraphs.
    let nav = Nav::new(&doc);
    assert_eq!(run(&nav, "header").expect("runs").count(), 1);
    assert_eq!(nav.produced(), 1, "only the header was produced");
}

// ---------------------------------------------------------------------------
// Scopes
// ---------------------------------------------------------------------------

#[test]
fn path_and_sheet_scopes_resolve_through_the_adapter() {
    let doc = report();
    assert_paths(
        &doc,
        "/body/tbl[1]/p",
        &[
            "/body/tbl[1]/tr[1]/tc[1]/p[1]",
            "/body/tbl[1]/tr[1]/tc[2]/p[1]",
            "/body/tbl[1]/tr[2]/tc[1]/p[1]",
            "/body/tbl[1]/tr[2]/tc[2]/p[1]",
            "/body/tbl[1]/tr[2]/tc[2]/tbl[1]/tr[1]/tc[1]/p[1]",
        ],
    );
    assert_paths(
        &doc,
        "/body/table[1]/tr",
        &[
            "/body/table[1]/tr[1]",
            "/body/table[1]/tr[2]",
            "/body/table[1]/tr[2]/tc[2]/tbl[1]/tr[1]",
        ],
    );
    assert_paths(&doc, "/header[1]/p", &["/header[1]/p[1]"]);
    assert_paths(
        &doc,
        "/body/p[1]/r",
        &["/body/p[1]/r[1]", "/body/p[1]/r[2]"],
    );
    let error = fail(&doc, "/body/p[99]/r");
    assert_eq!(error.code(), Code::PathNotFound);
    assert_eq!(fail(&doc, "Sheet1!p").code(), Code::UnsupportedPath);
    let book = book();
    assert_paths(&book, "Notes!cell", &["/Notes/A1"]);
    assert_paths(&book, "notes!cell", &["/notes/A1"]);
    assert_paths(&book, "/Notes/cell", &["/Notes/A1"]);
    assert_eq!(fail(&book, "Missing!cell").code(), Code::PathNotFound);
}

// ---------------------------------------------------------------------------
// Header-row predicates
// ---------------------------------------------------------------------------

#[test]
fn header_rows_bind_columns_by_name() {
    let doc = book();
    assert_paths(&doc, "row[Region=West]", &["/Data/row[2]", "/Data/row[5]"]);
    assert_paths(&doc, "row[region=west]", &["/Data/row[2]", "/Data/row[5]"]);
    assert_paths(&doc, "row[Sales>1000]", &["/Data/row[2]"]);
    assert_paths(&doc, "row[Region=West and Sales>1000]", &["/Data/row[2]"]);
    assert_paths(
        &doc,
        "row[Status=open or Sales<900]",
        &["/Data/row[2]", "/Data/row[3]", "/Data/row[5]"],
    );
    assert_paths(
        &doc,
        "row[not(Region=West)]",
        &["/Data/row[3]", "/Data/row[4]"],
    );
    assert_paths(&doc, "row[col.Status=closed]", &["/Data/row[3]"]);
    assert_paths(&doc, "row[column.Status=closed]", &["/Data/row[3]"]);
    assert_paths(
        &doc,
        r#"row["Region"~=es]"#,
        &["/Data/row[2]", "/Data/row[5]"],
    );
    assert_paths(&doc, "row[B>1000]", &["/Data/row[2]"]);
    assert_paths(
        &doc,
        "Data!row[Sales>500]",
        &["/Data/row[2]", "/Data/row[3]"],
    );
    assert_paths(
        &doc,
        "/Data/row[Sales>500]",
        &["/Data/row[2]", "/Data/row[3]"],
    );
    // In a union, header-row binding still holds; rows without a record are
    // not visited.
    assert_paths(
        &doc,
        "row[Status!=open], cell[formula]",
        &["/Data/row[3]", "/Data/B6"],
    );
    // `[N]` on a row keeps its path meaning: the row number.
    assert_paths(&doc, "row[Region=West][2]", &["/Data/row[2]"]);
    assert_paths(&doc, "row[Region=West][3]", &[]);
    // Hidden rows match, as in Excel.
    assert_paths(&doc, "row[Region=East]", &["/Data/row[3]"]);
}

#[test]
fn header_rows_compare_raw_values_when_numeric() {
    let doc = book();
    // B5 displays "50%" and stores 0.5: a numeric literal compares the
    // stored value, a text literal the displayed one.
    assert_paths(&doc, "row[Sales=0.5]", &["/Data/row[5]"]);
    assert_paths(&doc, "row[Sales=50%]", &["/Data/row[5]"]);
    assert_paths(&doc, "row[Sales<1]", &["/Data/row[5]"]);
    // A data row without a record still binds, with empty cells.
    assert_paths(&doc, "row[Status!=open]", &["/Data/row[3]", "/Data/row[4]"]);
}

#[test]
fn header_row_predicates_follow_the_classification_rules() {
    let doc = book();
    let error = fail(&doc, "row[Region=West and hidden=true]");
    assert_eq!(error.code(), Code::InvalidSelector);
    assert!(error.message().contains("mix"), "{error}");
    let error = fail(&doc, "row[2=West]");
    assert_eq!(error.code(), Code::InvalidSelector);
    assert!(error.suggestion().is_some_and(|text| text.contains("col.")));
    assert_eq!(fail(&doc, "row[@colour=red]").code(), Code::InvalidSelector);
    assert_paths(&doc, "row[@hidden=true]", &["/Data/row[3]"]);
    let error = fail(&doc, "row[Bonus>5]");
    assert_eq!(error.code(), Code::NotFound);
    assert!(
        error
            .suggestion()
            .is_some_and(|text| text.contains("requested columns"))
    );
    let mut ambiguous = book();
    ambiguous.tables.push(Table {
        sheet: 1,

        columns: vec!["Region".into()],
        first_column: 1,
        header_row: 1,
        last_row: 1,
        totals: false,
    });
    let error = fail(&ambiguous, "row[Region=West]");
    assert_eq!(error.code(), Code::InvalidSelector);
    assert!(error.message().contains("ambiguous"), "{error}");
    assert_paths(
        &ambiguous,
        "Data!row[Region=West]",
        &["/Data/row[2]", "/Data/row[5]"],
    );
    // A bare row property that names a column is ambiguous.
    let mut colliding = book();
    colliding.tables[0].columns[2] = "Hidden".to_owned();
    let error = fail(&colliding, "row[hidden=true]");
    assert_eq!(error.code(), Code::InvalidSelector);
    assert!(
        error
            .suggestion()
            .is_some_and(|text| text.contains("@ for a row property"))
    );
    assert_paths(&colliding, "row[@hidden=true]", &["/Data/row[3]"]);
}

// ---------------------------------------------------------------------------
// Errors mid-walk
// ---------------------------------------------------------------------------

#[test]
fn adapter_errors_end_the_walk() {
    let doc = Doc::new(
        Item::new(WordKind::Document).child(
            Item::new(WordKind::Body)
                .child(p("one"))
                .child(p("two"))
                .child(p("three"))
                .fail_child(2),
        ),
    );
    let nav = Nav::new(&doc);
    let mut walk = run(&nav, "p").expect("runs");
    assert!(matches!(walk.next(), Some(Ok(_))));
    assert!(matches!(walk.next(), Some(Ok(_))));
    assert!(matches!(walk.next(), Some(Err(error)) if error.code() == Code::CorruptFile));
    assert!(walk.next().is_none());
    assert!(walk.next().is_none());
    // A failing text read surfaces from the pseudo-class that needs it.
    let doc = Doc::new(
        Item::new(WordKind::Document).child(Item::new(WordKind::Body).child(p("bad").fail_text())),
    );
    assert_eq!(fail(&doc, "p:contains(x)").code(), Code::CorruptFile);
    assert_eq!(
        count(&doc, "p"),
        1,
        "a walk that reads no text does not fail"
    );
}

// ---------------------------------------------------------------------------
// Plans and arguments
// ---------------------------------------------------------------------------

#[test]
fn plans_are_reusable_and_bound_to_their_format() {
    let plan = compiled::<WordKind>("p[style=Normal]");
    assert_eq!(plan.format(), litchi_core::FileFormat::Docx);
    assert_eq!(plan.alternatives(), 1);
    let first = report();
    let second = generated(30, 5);
    for doc in [&first, &second] {
        let nav = Nav::new(doc);
        let walk = Walk::new(&nav, PlanRef::Borrowed(&plan)).expect("same format");
        assert!(walk.count() > 0);
        let walk = run(&nav, &plan).expect("a borrowed plan");
        assert_eq!(walk.plan().format(), litchi_core::FileFormat::Docx);
    }
    let book = book();
    let nav = Nav::new(&book);
    let Err(error) = Walk::new(&nav, PlanRef::Borrowed(&plan)) else {
        panic!("a Word plan ran on a workbook");
    };
    assert_eq!(error.code(), Code::InvalidArgument);
    // An owned plan moves into the walk.
    let owned = compiled::<WordKind>("tbl");
    let nav = Nav::new(&first);
    assert_eq!(run(&nav, owned).expect("an owned plan").count(), 2);
    const fn send_sync<T: Send + Sync>() {}
    send_sync::<Plan<'static>>();
    assert!(format!("{plan:?}").starts_with("Plan {"));
}

#[test]
fn into_query_accepts_text_queries_and_plans() {
    let doc = report();
    let nav = Nav::new(&doc);
    let text = String::from("p:contains(revenue)");
    assert_eq!(run(&nav, text.as_str()).expect("&str").count(), 1);
    assert_eq!(run(&nav, &text).expect("&String").count(), 1);
    let query = Query::parse("p:contains(revenue)").expect("parses");
    assert_eq!(run(&nav, &query).expect("&Query").count(), 1);
    assert_eq!(run(&nav, query.clone()).expect("Query").count(), 1);
    let arg = QueryArg::parse("p").expect("parses");
    assert_eq!(arg.text(), Some("p"));
    assert!(arg.plan().is_none());
    assert_eq!(run(&nav, arg).expect("QueryArg").count(), 13);
    let plan = compiled::<WordKind>("p");
    let arg = (&plan).into_query().expect("a plan");
    assert!(arg.plan().is_some() && arg.query().is_none() && arg.text().is_none());
    // Syntax errors keep the grammar's typed error.
    let error = "p[".into_query().unwrap_err();
    assert_eq!(error.code(), Code::InvalidSelector);
    assert!(matches!(error.leaf(), Source::Selector(_)));
    let error = "p:not(".repeat(40).as_str().into_query().unwrap_err();
    assert_eq!(error.code(), Code::MaxDepthExceeded);
}

#[test]
fn compiled_plans_bind_structure_once() {
    let plan = compiled::<SheetKind>("row[3], col[B], B, A1:C5");
    let structural: Vec<Structural> = plan
        .compounds
        .iter()
        .flat_map(|compound| compound.tests.iter().copied())
        .collect();
    assert_eq!(
        structural,
        [
            Structural::Number(3),
            Structural::Number(2),
            Structural::Column(1),
            Structural::Range {
                first: (0, 0),
                last: (2, 4)
            },
        ]
    );
    let plan = compiled::<WordKind>("tbl > tr[2] > tc p[last()]");
    let chain = plan.chains[plan.top[0] as usize];
    assert_eq!(chain.position, Some(Positional::Last));
    assert_eq!(chain.len, 4);
    assert_eq!(plan.memo_slots, 3);
    let plan = compiled::<WordKind>("p[3]");
    assert_eq!(
        plan.chains[plan.top[0] as usize].position,
        Some(Positional::Nth(3))
    );
    // A literal is coerced once, at compile time: a bad colour literal does
    // not fail, it simply never equals (OfficeCLI compares text).
    let plan = compiled::<WordKind>("r[color=notacolour]");
    assert!(plan.diagnostics().is_empty());
    assert_eq!(count(&report(), "r[color=notacolour]"), 0);
}

// ---------------------------------------------------------------------------
// OfficeCLI corpus and hostile input
// ---------------------------------------------------------------------------

/// Every selector harvested from OfficeCLI (record 0005) compiles for the
/// fake format of at least one family, or is refused with a typed code that
/// says why; none is an internal error. Compiled plans run on that family's
/// fixture, where a scope the fixture does not have is `path_not_found`.
#[test]
fn harvested_officecli_selectors_compile_or_fail_typed() {
    let text = include_str!("../../../litchi-core/tests/fixtures/officecli_selectors.txt");
    let word = report();
    let deck = deck();
    let book = book();
    let mut compiled = [0_usize; 3];
    let mut refused = Vec::new();
    for line in text
        .lines()
        .filter(|line| !line.trim().is_empty() && !line.starts_with('#'))
    {
        let query = Query::parse(line).unwrap_or_else(|error| panic!("{line:?}: {error}"));
        let outcomes = [
            check(&word, &query, line),
            check(&deck, &query, line),
            check(&book, &query, line),
        ];
        for (index, outcome) in outcomes.into_iter().enumerate() {
            compiled[index] += usize::from(outcome);
        }
        if !outcomes.contains(&true) {
            refused.push(line);
        }
    }
    assert_eq!(refused, REFUSED);
    // Compiled per family: Word, presentation, workbook (132 harvested).
    assert_eq!(compiled, [63, 88, 48]);
}

/// Harvested selectors no fake family compiles: elements the fakes do not
/// model (`unsupported_capability`), list as gaps (`chart`, `pivottable`) or
/// no vocabulary knows (`drawing`), and a regex literal (record 0010 Q3).
const REFUSED: &[&str] = &[
    "/numbering",
    "/revision[@author=Alice]",
    "animation",
    "chart",
    "comment",
    "conditionalformatting",
    "connector",
    "drawing",
    "equation",
    "field",
    "field[fieldType!=page]",
    "field[fieldType=page]",
    "field[fieldType=seq]",
    "footnote",
    "formfield",
    "merge",
    "ole",
    "pivottable",
    "revision",
    "section",
    "slicer",
    "sparkline",
    "toc",
    "validation",
    "slide[1]>ole",
    "slide chart",
    "revision[@type=ins]",
    "ole[progId~=Excel]",
    r#"shape[text~=r"^Q[1-4]"]"#,
    "/slide[1]/shape[@id=10000]",
];

/// Compiles `query` for `doc`'s family and runs it; `true` when it
/// compiled, `false` for a typed refusal.
fn check<K: FakeKind>(doc: &Doc<K>, query: &Query<'_>, line: &str) -> bool {
    let plan = match compile::<&Nav<'static, K>>(query) {
        Ok(plan) => plan,
        Err(error) => {
            assert_typed(&error, line);
            return false;
        },
    };
    let nav = Nav::new(doc);
    let walk = Walk::new(&nav, PlanRef::Borrowed(&plan)).expect("same format");
    for item in walk.limit(64) {
        if let Err(error) = item {
            assert_typed(&error, line);
            break;
        }
    }
    true
}

/// `SplitMix64`, for reproducible mutations without a new dependency.
struct Rng(u64);

impl Rng {
    fn next(&mut self) -> u64 {
        self.0 = self.0.wrapping_add(0x9E37_79B9_7F4A_7C15);
        let mut z = self.0;
        z = (z ^ (z >> 30)).wrapping_mul(0xBF58_476D_1CE4_E5B9);
        z = (z ^ (z >> 27)).wrapping_mul(0x94D0_49BB_1331_11EB);
        z ^ (z >> 31)
    }

    fn below(&mut self, bound: usize) -> usize {
        usize::try_from(self.next() % u64::try_from(bound).expect("small")).expect("small")
    }
}

/// The fuzz target's invariants, on stable: harvested selectors mutated by
/// inserting, deleting and replacing selector tokens either fail to parse,
/// or compile for every fake family to a plan or a typed refusal, and every
/// plan runs to matches or a typed error. Never a panic or an internal
/// error.
#[test]
fn mutated_selectors_compile_and_run_typed() {
    const TOKENS: &[&str] = &[
        "[",
        "]",
        "(",
        ")",
        ":",
        ">",
        " ",
        ",",
        "=",
        "!=",
        "~=",
        ">=",
        "<",
        "@",
        "\"",
        "'",
        "*",
        "/",
        ".",
        "1",
        "2",
        "0",
        "a",
        "B",
        "p",
        "r",
        "tc",
        "tbl",
        "row",
        "cell",
        "shape",
        "slide",
        "title",
        "textbox",
        "col",
        "sheet",
        "Data!",
        "picture",
        ":not(",
        ":has(",
        ":empty",
        ":no-alt",
        ":contains(",
        "last()",
        " and ",
        " or ",
        "not(",
        "col.",
        "@hidden",
        "style",
        "bold",
        "Sales",
        "Region",
        "value",
        "text",
        "type",
        "10cm",
        "12pt",
        "red",
        "true",
    ];
    let text = include_str!("../../../litchi-core/tests/fixtures/officecli_selectors.txt");
    let seeds: Vec<&str> = text
        .lines()
        .filter(|line| !line.trim().is_empty() && !line.starts_with('#'))
        .collect();
    let (word, deck, book) = (report(), deck(), book());
    let mut rng = Rng(0x0010_0003);
    let mut parsed = 0;
    for _ in 0..20_000 {
        let mut selector = seeds[rng.below(seeds.len())].to_owned();
        for _ in 0..=rng.below(3) {
            let at = (0..=selector.len())
                .filter(|&index| selector.is_char_boundary(index))
                .nth(rng.below(selector.chars().count() + 1))
                .unwrap_or(selector.len());
            let end = (at..=selector.len())
                .filter(|&index| selector.is_char_boundary(index))
                .nth(1)
                .unwrap_or(selector.len());
            let token = TOKENS[rng.below(TOKENS.len())];
            match rng.below(3) {
                0 => selector.insert_str(at, token),
                1 => selector.replace_range(at..end, ""),
                _ => selector.replace_range(at..end, token),
            }
        }
        let Ok(query) = Query::parse(&selector) else {
            continue;
        };
        parsed += 1;
        let _ = check(&word, &query, &selector);
        let _ = check(&deck, &query, &selector);
        let _ = check(&book, &query, &selector);
    }
    // 9 205 of the 20 000 mutations parse with this seed.
    assert!(parsed > 5_000, "only {parsed} mutations parsed");
}

fn assert_typed(error: &Error, line: &str) {
    assert!(
        matches!(
            error.code(),
            Code::InvalidSelector
                | Code::UnsupportedPath
                | Code::UnsupportedCapability
                | Code::PathNotFound
                | Code::NotFound
                | Code::InvalidPath
        ),
        "{line:?}: {error}"
    );
}

#[test]
fn deep_documents_do_not_grow_the_call_stack() {
    // 20 000 nested tables, walked on a thread with a 256 KiB stack.
    let worker = std::thread::Builder::new()
        .stack_size(256 * 1024)
        .spawn(|| {
            let mut item = tc("deepest");
            for _ in 0..20_000 {
                item = Item::new(WordKind::Cell)
                    .child(Item::new(WordKind::Table).child(Item::new(WordKind::Row).child(item)));
            }
            let doc =
                Doc::new(
                    Item::new(WordKind::Document).child(Item::new(WordKind::Body).child(
                        Item::new(WordKind::Table).child(Item::new(WordKind::Row).child(item)),
                    )),
                );
            let nav = Nav::new(&doc);
            for selector in [
                "tbl tbl tc > p:contains(deepest)",
                "body > tbl tc > p",
                "tbl:has(p:contains(deepest))",
            ] {
                let mut walk = run(&nav, selector).expect("runs");
                assert_eq!(
                    walk.next().unwrap().unwrap_err().code(),
                    Code::MaxDepthExceeded
                );
                assert!(walk.next().is_none());
            }
            // Dropping the document's 60 000 nested items recurses in the
            // fake's own `Drop`; keep it out of this small stack.
            std::mem::forget(doc);
        })
        .expect("spawn");
    worker.join().expect("the deep walk completes");
}

#[test]
fn nested_pseudo_classes_and_lists_compile_within_the_grammar_limits() {
    let doc = report();
    let selector = format!("p{}", ":not(:has(r:not(:contains(x))))".repeat(4));
    assert!(compile_text::<WordKind>(&selector).is_ok());
    assert_eq!(
        count(&doc, &selector),
        count(&doc, "p:not(:has(r:not(:contains(x))))")
    );
    // The parser bounds nesting; compilation recurses only that deep.
    let selector = format!("p{}{}", ":not(".repeat(31), ")".repeat(31));
    let _ = compile_text::<WordKind>(&selector);
}

#[test]
fn diagnostics_render_and_compare() {
    let a = Diagnostic::new(
        DiagnosticKind::NonNumeric,
        "size",
        "non-numeric".to_owned(),
        None,
    );
    assert_eq!(a.to_string(), "non_numeric: non-numeric");
    assert_eq!(a.clone(), a);
    assert_eq!(DiagnosticKind::NeverMatches.as_str(), "never_matches");
}

#[test]
fn typed_queries_are_bounded_before_recursive_compilation() {
    use litchi_core::selector::{
        Complex, Compound, Condition, Element, Expr, Key, Predicate, Pseudo,
    };
    let query = Query::new([]);
    assert_eq!(
        compile::<&Nav<'_, WordKind>>(&query).unwrap_err().code(),
        Code::LimitExceeded
    );
    let mut expr = Expr::Condition(Condition::Exists(Key::new("bold")));
    for _ in 0..64 {
        expr = Expr::Not(Box::new(expr));
    }
    let query = Query::new([Complex::new(
        Compound::new(Element::Named("r".into())).with(Predicate::Filter(expr)),
    )]);
    assert_eq!(
        compile::<&Nav<'_, WordKind>>(&query).unwrap_err().code(),
        Code::MaxDepthExceeded
    );
    let query = Query::new([Complex::new(
        Compound::new(Element::Named("p".into()))
            .with_pseudo(Pseudo::Contains("x".repeat(16 * 1024 + 1).into())),
    )]);
    assert_eq!(
        compile::<&Nav<'_, WordKind>>(&query).unwrap_err().code(),
        Code::LimitExceeded
    );
}

#[test]
fn native_query_keys_require_raw_admission() {
    let doc = report();
    let nav = Nav::new(&doc);
    for text in ["p[@paraId=1a2b3c4d]", "p[paraId=1a2b3c4d]", "sdt[sdtId=1]"] {
        assert_eq!(run(&nav, text).err().unwrap().code(), Code::InvalidSelector);
    }
    let arg = super::raw::query("p[@paraId=1a2b3c4d]").unwrap();
    assert_eq!(run(&nav, arg).unwrap().count(), 1);
}

#[test]
fn query_debug_and_diagnostics_omit_authored_content() {
    for text in [
        "SENSITIVE",
        "SENSITIVE!p",
        "col[SENSITIVE]",
        "row[@SENSITIVE=1]",
        "p[text~=r\"SENSITIVE\"]",
    ] {
        let query = Query::parse(text).unwrap();
        assert!(!format!("{query:?}").contains("SENSITIVE"));
        if let Err(error) = compile::<&Nav<'_, WordKind>>(&query) {
            assert!(!format!("{error:?} {error}").contains("SENSITIVE"));
        }
        assert!(!format!("{:?}", text.into_query().unwrap()).contains("SENSITIVE"));
    }
    for text in [
        "p[SENSITIVE=x]",
        "p[size>SENSITIVE]",
        "p:contains(SENSITIVE)",
    ] {
        let plan = compile_text::<WordKind>(text).unwrap();
        assert!(!format!("{plan:?}").contains("SENSITIVE"));
        for diagnostic in plan.diagnostics() {
            assert!(!format!("{diagnostic:?} {diagnostic}").contains("SENSITIVE"));
        }
    }
}

#[test]
fn scope_admission_precedes_document_reads_and_raw_plans_reuse_it() {
    use litchi_core::selector::{Path, Scope, Step};
    let query = Query::parse("/body/p[@paraId=1A2B3C4D]/r").unwrap();
    assert_eq!(
        compile::<&Nav<'_, WordKind>>(&query).unwrap_err().code(),
        Code::InvalidSelector
    );
    let arg = super::raw::query(&query).unwrap();
    assert!(arg.into_plan::<&Nav<'_, WordKind>>().unwrap().raw);
    let ordinary = compiled::<WordKind>("p");
    let promoted = super::raw::query(&ordinary)
        .unwrap()
        .into_plan::<&Nav<'_, WordKind>>()
        .unwrap();
    assert!(promoted.raw);
    assert!(!ordinary.raw);
    let relative = Query::parse("p")
        .unwrap()
        .with_scope(Scope::Path(Box::new(Path::relative([Step::new("body")]))));
    assert_eq!(
        compile::<&Nav<'_, WordKind>>(&relative).unwrap_err().code(),
        Code::InvalidSelector
    );
    let deep = Query::parse("p")
        .unwrap()
        .with_scope(Scope::Path(Box::new(Path::from_steps(
            (0..65).map(|_| Step::new("body")),
        ))));
    assert_eq!(
        compile::<&Nav<'_, WordKind>>(&deep).unwrap_err().code(),
        Code::MaxDepthExceeded
    );
    let unknown = compiled::<SheetKind>("cell[paraId=x]");
    assert_eq!(unknown.diagnostics()[0].kind(), DiagnosticKind::UnknownKey);
}

#[test]
fn sparse_row_probes_are_bounded_and_terminal_row_does_not_overflow() {
    let mut doc = book();
    doc.tables[0].last_row = 1_000_100;
    let nav = Nav::new(&doc);
    assert!(
        run(&nav, "row[Region=West]")
            .unwrap()
            .limit(1)
            .next()
            .unwrap()
            .is_ok()
    );
    let error = run(&nav, "row[Region=West]")
        .unwrap()
        .collect::<Result<Vec<_>>>()
        .unwrap_err();
    assert_eq!(error.code(), Code::LimitExceeded);
    doc.tables[0].header_row = u32::MAX - 1;
    doc.tables[0].last_row = u32::MAX;
    let nav = Nav::new(&doc);
    assert!(run(&nav, "row[Region=West]").unwrap().next().is_none());
}

/// Record 0280: the walker skips nodes the adapter reports childless
/// (`Inspect::childless`). On every ODT and flat-ODT fixture, a query for
/// each element kind returns exactly the nodes of that kind an unpruned
/// depth-first walk over `Navigate::children` (all kinds) finds, in the
/// same order; `limit` and `first()` return their leading matches.
#[cfg(feature = "odt")]
#[test]
fn odt_queries_match_an_unpruned_walk() {
    use crate::format::odt::Nav as OdtNav;
    use crate::path::navigate::{ElementKind, KindSet, Navigate};
    use crate::query::inspect::Inspect;

    fn walk<'f>(
        nav: &OdtNav<'f>,
        node: crate::format::odt::Node<'f>,
        out: &mut Vec<(String, String)>,
    ) {
        let children: Vec<_> = nav
            .children(&node, KindSet::ALL)
            .expect("children")
            .collect::<Result<_>>()
            .expect("child");
        out.push((
            nav.kind(&node).name().to_owned(),
            crate::node::Node::Odt(node).path().to_string(),
        ));
        for child in children {
            walk(nav, child, out);
        }
    }

    fn collect(directory: &std::path::Path, out: &mut Vec<std::path::PathBuf>) {
        let Ok(entries) = std::fs::read_dir(directory) else {
            return;
        };
        for entry in entries {
            let path = entry.expect("entry").path();
            if path.is_dir() {
                collect(&path, out);
            } else if path
                .extension()
                .is_some_and(|extension| extension == "odt" || extension == "fodt")
            {
                out.push(path);
            }
        }
    }

    let mut opened = 0;
    let mut entries = Vec::new();
    collect(
        std::path::Path::new(concat!(env!("CARGO_MANIFEST_DIR"), "/../../test-data")),
        &mut entries,
    );
    entries.sort();
    for path in entries {
        let Ok(file) = crate::open(&path) else {
            continue;
        };
        let crate::File::Odt(handle) = &file else {
            continue;
        };
        opened += 1;
        let nav = OdtNav::new(handle);
        let mut all = Vec::new();
        walk(&nav, nav.root().expect("root"), &mut all);
        for element in <OdtNav<'_> as Inspect<'_>>::SPEC.elements {
            let name = element.kind.name();
            if element.gap.is_some() || matches!(name, "document") {
                continue;
            }
            let expected: Vec<&str> = all
                .iter()
                .filter(|(kind, _)| kind == name)
                .map(|(_, path)| path.as_str())
                .collect();
            let Ok(matches) = file.query(name) else {
                continue;
            };
            let actual: Vec<String> = matches
                .map(|node| node.expect("match").path().to_string())
                .collect();
            assert_eq!(actual, expected, "{} `{name}`", path.display());
            for limit in [1, 2, 5] {
                let limited: Vec<String> = file
                    .query(name)
                    .expect("query")
                    .limit(limit)
                    .map(|node| node.expect("match").path().to_string())
                    .collect();
                assert_eq!(
                    limited,
                    actual[..actual.len().min(limit)],
                    "{name} limit {limit}"
                );
            }
            let first = file
                .query(name)
                .expect("query")
                .first()
                .expect("first")
                .map(|node| node.path().to_string());
            assert_eq!(first.as_deref(), actual.first().map(String::as_str));
        }
    }
    assert!(opened >= 5, "only {opened} ODT fixtures opened");
}

// ---------------------------------------------------------------------------
// The single-subject fast path (record 0315)
// ---------------------------------------------------------------------------

/// The identity of a fake node: its element (by address) and worksheet, or
/// its row.
fn identity<K>(node: &Node<'_, K>) -> String {
    match node {
        Node::Item { item, sheet } => {
            format!("item {:p} {:?}", *item, sheet.map(std::ptr::from_ref))
        },
        Node::Row { sheet, number } => format!("row {:p} {number}", *sheet),
    }
}

/// Every match the walk yields as an iterator (the route `File::query`
/// takes), the visit count, and the error that ended the walk, if any.
fn trace<'d, K: FakeKind>(mut walk: Walk<'_, 'd, &Nav<'d, K>>) -> (Vec<String>, u64, Option<Code>) {
    let mut out = Vec::new();
    let mut failed = None;
    // The route `Matches` takes: the fast step when the walk has one.
    while let Some(item) = if walk.is_fast() {
        walk.next_fast_map(|node| node)
    } else {
        walk.next()
    } {
        match item {
            Ok(node) => out.push(identity(&node)),
            Err(error) => {
                failed = Some(error.code());
                break;
            },
        }
    }
    (out, walk.visited(), failed)
}

/// `text` on `doc` yields the same nodes, visits and error on the fast path
/// as on the general step, also with a limit, and the general step's hits
/// are those nodes; returns whether the plan takes the fast path.
fn same_as_general<K: FakeKind>(doc: &Doc<K>, text: &str) -> bool {
    let nav = Nav::new(doc);
    let fast = run(&nav, text).unwrap_or_else(|error| panic!("{text:?}: {error}"));
    let is_fast = fast.is_fast();
    let general = run(&nav, text).unwrap().general();
    assert!(!general.is_fast());
    let expected = trace(general);
    assert_eq!(trace(fast), expected, "{text}");
    let mut hits = run(&nav, text).unwrap();
    let mut nodes = Vec::new();
    while let Some(Ok(hit)) = hits.next_hit() {
        nodes.push(identity(hit.node()));
    }
    assert_eq!(
        nodes, expected.0,
        "{text}: the hits are the iterator's nodes"
    );
    for limit in [1, 2] {
        let fast = run(&nav, text).unwrap().limit(limit);
        let general = run(&nav, text).unwrap().general().limit(limit);
        assert_eq!(trace(fast), trace(general), "{text} limit {limit}");
    }
    is_fast
}

#[test]
fn fast_path_matches_the_general_step() {
    let doc = report();
    for text in [
        "p",
        "r",
        "paragraph",
        "r:contains(report)",
        "p:empty",
        "tbl",
        "/header[1]/p",
        "p:has(r)",
        "p:contains(Nothing)",
        "r[bold=true]",
        "p[style=Heading1]",
        "p[style=Nothing]",
    ] {
        assert!(same_as_general(&doc, text), "{text} takes the fast path");
    }
    for text in ["p, r", "p[2]", "p[last()]", "tbl > tr", "*", "body p"] {
        assert!(
            !same_as_general(&doc, text),
            "{text} takes the general step"
        );
    }
    let deck = deck();
    for text in [
        "shape",
        "shape:contains(Revenue)",
        "picture",
        "shape[width>=10cm]",
    ] {
        assert!(same_as_general(&deck, text), "{text}");
    }
    let book = book();
    for text in ["cell", "row", "/Notes/cell", "cell[value>1000]"] {
        assert!(same_as_general(&book, text), "{text}");
    }
    for text in ["B", "row[Sales>1000]"] {
        assert!(!same_as_general(&book, text), "{text}");
    }
}
