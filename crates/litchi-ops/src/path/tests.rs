//! Unit tests of path resolution over the fake formats.

use super::fake::{Fake, Fixture, Kind, Node, Presentation, Word, Workbook};
use super::navigate::{
    Available, ElementKind, Identities, KindSet, Lookup, Navigate, Resolved, Select,
};
use super::report::distance;
use super::resolve;
use super::vocabulary::{
    self, Bare, Case, ElementRule, Identity, IndexRule, Named, PRESENTATION, Vocabulary, WORD,
    WORKBOOK,
};
use super::{IntoPath, ParseError, Path, PathArg, Position, Predicate, Query, Step, limits};
use crate::error::Source;
use crate::{Code, Error, Family};

fn arg(text: &str) -> PathArg<'_> {
    super::raw::path(text).unwrap_or_else(|error| panic!("{text:?}: {error}"))
}

fn get<F: Fixture>(text: &str) -> Node {
    resolve::get(&Fake::<F>::new(), &arg(text)).unwrap_or_else(|error| panic!("{text:?}: {error}"))
}

fn fail<F: Fixture>(text: &str) -> Error {
    match resolve::get(&Fake::<F>::new(), &arg(text)) {
        Ok(node) => panic!("{text:?} resolved to {node:?}"),
        Err(error) => error,
    }
}

fn find<F: Fixture>(text: &str) -> Option<Node> {
    resolve::find(&Fake::<F>::new(), &arg(text)).unwrap_or_else(|error| panic!("{text:?}: {error}"))
}

/// Asserts the error's code, message and suggestion, and that it names the
/// fake format.
fn assert_error(error: &Error, code: Code, message: &str, suggestion: Option<&str>) {
    assert_eq!(error.code(), code, "{error}");
    assert_eq!(error.message(), message, "{error}");
    assert_eq!(error.suggestion(), suggestion, "{error}");
}

fn children<F: Fixture>(node: Node, kinds: KindSet<Kind<F>>) -> usize {
    let nav = Fake::<F>::new();
    let mut count = 0;
    for child in nav.children(&node, kinds).expect("children") {
        assert!(matches!(child.expect("child"), Node::Tree(_)));
        count += 1;
    }
    count
}

fn kind<F: Fixture>(name: &str) -> Kind<F> {
    Kind::<F>::from_name(name).unwrap_or_else(|| panic!("{name} is not modeled"))
}

// ---------------------------------------------------------------- vocabulary

#[test]
fn the_plans_m1_aliases_are_present() {
    // Record 0010 §4.2, one row per alias.
    let word: &[(&str, &str)] = &[
        ("paragraph", "p"),
        ("P", "p"),
        ("run", "r"),
        ("r", "r"),
        ("table", "tbl"),
        ("tbl", "tbl"),
        ("row", "tr"),
        ("tr", "tr"),
        ("cell", "tc"),
        ("tc", "tc"),
        ("hyperlink", "hyperlink"),
        ("sdt", "sdt"),
        ("contentcontrol", "sdt"),
        ("picture", "picture"),
        ("image", "picture"),
        ("img", "picture"),
        ("equation", "equation"),
        ("math", "equation"),
        ("formula", "equation"),
        ("oMath", "equation"),
        ("oMathPara", "equation"),
        ("toc", "toc"),
        ("tableofcontents", "toc"),
        ("ole", "ole"),
        ("oleobject", "ole"),
        ("object", "ole"),
        ("embed", "ole"),
        ("level", "level"),
        ("lvl", "level"),
        ("break", "break"),
        ("br", "break"),
        ("pagebreak", "break"),
    ];
    let presentation: &[(&str, &str)] = &[
        ("shape", "shape"),
        ("textbox", "shape"),
        ("title", "shape"),
        ("picture", "picture"),
        ("pic", "picture"),
        ("connector", "connector"),
        ("connection", "connector"),
        ("hyperlink", "hyperlink"),
        ("hlink", "hyperlink"),
        ("paragraph", "paragraph"),
        ("p", "paragraph"),
        ("run", "run"),
        ("r", "run"),
        ("tr", "tr"),
        ("row", "tr"),
        ("tc", "tc"),
        ("cell", "tc"),
        ("col", "col"),
        ("column", "col"),
        ("moderncomment", "modernComment"),
        ("modern-comment", "modernComment"),
        ("thread", "modernComment"),
        ("threadedcomment", "modernComment"),
        ("video", "video"),
        ("audio", "audio"),
        ("media", "media"),
        ("slidemaster", "slideMaster"),
        ("slidelayout", "slideLayout"),
        ("layout", "slideLayout"),
    ];
    let workbook: &[(&str, &str)] = &[
        ("col", "col"),
        ("column", "col"),
        ("comment", "comment"),
        ("note", "comment"),
        ("table", "table"),
        ("listobject", "table"),
        ("namedrange", "namedrange"),
        ("definedname", "namedrange"),
        ("pivottable", "pivottable"),
        ("pivot", "pivottable"),
        ("validation", "dataValidation"),
        ("dataValidation", "dataValidation"),
        ("cf", "cf"),
        ("conditionalformatting", "cf"),
        ("merge", "mergeCell"),
        ("mergedrange", "mergeCell"),
        ("ole", "ole"),
        ("oleobject", "ole"),
        ("object", "ole"),
        ("embed", "ole"),
        ("media", "picture"),
        ("image", "picture"),
    ];
    for (vocabulary, table) in [
        (&WORD, word),
        (&PRESENTATION, presentation),
        (&WORKBOOK, workbook),
    ] {
        for &(alias, canonical) in table {
            assert_eq!(
                vocabulary.canonical(alias),
                Some(canonical),
                "{:?} {alias}",
                vocabulary.family
            );
        }
    }
}

#[test]
fn every_story_axis_is_a_known_element() {
    // The `Axis` table of record 0005: each name belongs to some family.
    let axes = [
        "body",
        "chart",
        "comment",
        "comments",
        "customXml",
        "endnote",
        "endnotes",
        "footer",
        "footnote",
        "footnotes",
        "header",
        "layout",
        "modernComment",
        "notes",
        "noteSlide",
        "notesMaster",
        "numbering",
        "settings",
        "sheet",
        "slide",
        "slideLayout",
        "slideMaster",
        "styles",
        "theme",
        "watermark",
    ];
    for axis in axes {
        assert!(
            vocabulary::FAMILIES
                .iter()
                .any(|family| family.canonical(axis).is_some()),
            "{axis}"
        );
        assert!(super::Axis::classify(axis).is_some(), "{axis}");
    }
}

#[test]
fn names_fold_ascii_case() {
    assert_eq!(WORD.canonical("PARAGRAPH"), Some("p"));
    assert_eq!(WORD.canonical("TableOfContents"), Some("toc"));
    assert_eq!(WORD.canonical(&"p".repeat(33)), None);
    assert_eq!(WORD.canonical("Ä"), None);
    assert_eq!(WORD.canonical(""), None);
    assert_eq!(
        WORD.rule("ABSTRACTNUM").map(|rule| rule.name),
        Some("abstractNum")
    );
    assert_eq!(WORD.rule("nope"), None);
    assert_eq!(WORD.index("abstractnum"), Some(0));
    assert_eq!(WORD.index("watermark"), Some(WORD.elements.len() - 1));
    assert_eq!(WORD.spellings().count(), WORD.aliases.len());
    assert!(
        WORD.spellings()
            .all(|(alias, canonical)| WORD.canonical(alias) == Some(canonical))
    );
    assert_eq!(WORD.root, "document");
    assert_eq!(WORD.family, Family::Document);
    assert_eq!(PRESENTATION.family, Family::Presentation);
    assert_eq!(WORKBOOK.family, Family::Workbook);
}

#[test]
fn other_families_know_foreign_names() {
    let found =
        |own: &Vocabulary, name: &str| vocabulary::elsewhere(own, name).map(|other| other.family);
    assert_eq!(found(&WORD, "slide"), Some(Family::Presentation));
    assert_eq!(found(&WORD, "SHEET"), Some(Family::Workbook));
    assert_eq!(found(&WORKBOOK, "body"), Some(Family::Document));
    assert_eq!(found(&PRESENTATION, "namedrange"), Some(Family::Workbook));
    assert_eq!(found(&WORD, "paragrph"), None);
}

#[test]
fn element_rules_carry_the_plans_predicate_meanings() {
    let rule = |vocabulary: &Vocabulary, name: &str| *vocabulary.rule(name).expect(name);
    let p = rule(&WORD, "p");
    assert_eq!(
        p.identities,
        &[
            Identity {
                key: "paraId",
                case: Case::Ascii
            },
            Identity {
                key: "textId",
                case: Case::Ascii
            }
        ]
    );
    assert_eq!(
        p.identity("PARAID").map(|identity| identity.key),
        Some("paraId")
    );
    assert_eq!(p.identity("name"), None);
    assert_eq!(rule(&WORD, "footnote").index, IndexRule::NumberUnderRoot);
    assert_eq!(rule(&WORD, "level").index, IndexRule::Number);
    assert_eq!(rule(&WORD, "formfield").bare, Some(Bare::Name(Case::Exact)));
    assert_eq!(
        rule(&WORD, "styles").named,
        Named::Only {
            child: "style",
            case: Case::Exact
        }
    );
    assert_eq!(
        rule(&WORD, "bookmark")
            .identity("name")
            .map(|identity| identity.case),
        Some(Case::Exact)
    );
    assert_eq!(rule(&PRESENTATION, "placeholder").bare, Some(Bare::Token));
    assert_eq!(
        rule(&PRESENTATION, "shape")
            .identity("name")
            .map(|identity| identity.case),
        Some(Case::Caseless)
    );
    assert_eq!(rule(&WORKBOOK, "row").index, IndexRule::Row);
    let column: ElementRule = rule(&WORKBOOK, "col");
    assert_eq!(
        (column.index, column.bare),
        (IndexRule::Column, Some(Bare::Column))
    );
    assert!(rule(&WORKBOOK, "sheet").cells);
    assert_eq!(
        rule(&WORKBOOK, "workbook").named,
        Named::First {
            child: "sheet",
            case: Case::Caseless
        }
    );
    assert_eq!(
        rule(&WORKBOOK, "namedrange").bare,
        Some(Bare::Name(Case::Caseless))
    );
}

#[test]
fn values_compare_by_their_case_rule() {
    assert!(Case::Exact.matches("Heading1", "Heading1"));
    assert!(!Case::Exact.matches("Heading1", "heading1"));
    assert!(Case::Ascii.matches("1A2B3C4D", "1a2b3c4d"));
    assert!(!Case::Ascii.matches("Ärger", "ärger"));
    assert!(Case::Caseless.matches("Ärger", "äRGER"));
    assert!(Case::Caseless.matches("Straße", "STRAßE"));
    assert!(!Case::Caseless.matches("Sheet1", "Sheet2"));
}

// ------------------------------------------------------------------ IntoPath

#[test]
fn text_parses_once_into_absolute_or_relative_paths() {
    let absolute = arg("/body/p[3]/r[1]");
    assert!(absolute.is_absolute());
    assert_eq!(absolute.text(), Some("/body/p[3]/r[1]"));
    assert_eq!(absolute.to_string(), "/body/p[3]/r[1]");
    assert!(matches!(absolute.origin(), super::arg::Origin::Text(_)));

    let relative = arg("tr[2]/tc[1]");
    assert!(!relative.is_absolute());
    assert_eq!(relative.path().steps().len(), 2);

    let owned = String::from("/slide[1]");
    assert_eq!(
        (&owned).into_path().expect("parses").to_string(),
        "/slide[1]"
    );

    let error = "/body/p[0]".into_path().unwrap_err();
    assert_eq!(error.code(), Code::InvalidPath);
    let Source::Selector(parse) = error.leaf() else {
        panic!("{error:?}")
    };
    assert_eq!(parse.offset(), 8);
}

#[test]
fn parsed_paths_are_used_as_is() {
    let path = Path::parse("/body/p[2]").expect("parses");
    let borrowed = (&path).into_path().expect("borrowed");
    assert_eq!(borrowed.path(), &path);
    assert_eq!(borrowed.text(), None);
    assert!(matches!(borrowed.origin(), super::arg::Origin::Typed));

    // A `'static` path lends itself for any shorter lifetime.
    let owned: Path<'static> = "/body/p[1]".parse().expect("parses");
    fn lend<'p>(path: &'p Path<'static>) -> PathArg<'p> {
        path.into_path().expect("borrowed")
    }
    assert_eq!(lend(&owned).path(), &owned);

    let by_value = path.clone().into_path().expect("owned");
    assert_eq!(by_value.path(), &path);
    let again = by_value.clone().into_path().expect("identity");
    assert_eq!(again.to_string(), "/body/p[2]");
    assert_eq!(PathArg::from(path.clone()).path(), &path);
}

#[test]
fn excel_notation_becomes_a_workbook_path() {
    let cases = [
        ("Sheet1!A1", "/Sheet1/A1"),
        ("Sheet1!A1:C5", "/Sheet1/A1:C5"),
        ("'My Data'!A1:C5", "/My Data/A1:C5"),
        ("'It''s'!B2", "/It's/B2"),
        ("Sheet1!3:3", "/Sheet1/row[3]"),
        ("Sheet1!03:03", "/Sheet1/row[3]"),
        ("Sheet1!B:B", "/Sheet1/col[B]"),
        ("Sheet1!b:B", "/Sheet1/col[b]"),
        ("Sheet1!AB:ab", "/Sheet1/col[AB]"),
        ("Sheet1!row[3]", "/Sheet1/row[3]"),
        ("Sheet1!chart[1]/series[2]", "/Sheet1/chart[1]/series[2]"),
        ("'Q1!Results'!A1", "/Q1!Results/A1"),
    ];
    for (text, canonical) in cases {
        let arg = arg(text);
        assert_eq!(arg.to_string(), canonical, "{text}");
        assert!(arg.is_absolute(), "{text}");
        assert_eq!(arg.text(), Some(text));
        assert!(
            matches!(arg.origin(), super::arg::Origin::Sheet { .. }),
            "{text}"
        );
    }
    // `''` un-doubles into an owned name; plain names stay borrowed.
    assert!(matches!(arg("'It''s'!B2").path().steps()[0].name(), "It's"));
    // Whole rows and columns need the same run on both sides, as OfficeCLI.
    assert_eq!(arg("Sheet1!A:C").path().steps()[1].name(), "A:C");
    assert_eq!(arg("Sheet1!1:5").path().steps()[1].name(), "1:5");
}

#[test]
fn text_that_is_not_excel_notation_keeps_its_meaning() {
    // A leading `/` is a path; `!` is an ordinary name character there.
    let slash = arg("/Sheet1!A1");
    assert_eq!(slash.path().steps().len(), 1);
    assert_eq!(slash.path().steps()[0].name(), "Sheet1!A1");
    // `!` inside a predicate, before `!=`, or at the start is not notation.
    for text in ["shape[@name=Hello!]", "a!=b", "!A1"] {
        let arg = arg(text);
        assert!(!arg.is_absolute(), "{text}");
        assert!(
            matches!(arg.origin(), super::arg::Origin::Text(_)),
            "{text}"
        );
    }
    // A quote that does not close before `!` leaves the text to the grammar.
    for text in ["'unterminated!A1", "'quoted'A1"] {
        assert_eq!(
            PathArg::parse(text).unwrap_err().code(),
            Code::InvalidPath,
            "{text}"
        );
    }
}

#[test]
fn malformed_excel_notation_is_an_invalid_path() {
    let error = PathArg::parse("Sheet1!").unwrap_err();
    assert_error(
        &error,
        Code::InvalidPath,
        "nothing follows '!' at byte 7",
        Some(
            "write a cell, range, row or column after '!', e.g. Sheet1!A1, Sheet1!3:3 or Sheet1!B:B",
        ),
    );
    let error = PathArg::parse("''!A1").unwrap_err();
    assert_error(
        &error,
        Code::InvalidPath,
        "the sheet name before '!' is empty at byte 0",
        Some("name the sheet, e.g. Sheet1!A1 or 'My Data'!A1"),
    );
    for text in [
        "Sheet1!0:0",
        "Sheet1!99999999999999999999999:99999999999999999999999",
    ] {
        let error = PathArg::parse(text).unwrap_err();
        assert_eq!(error.code(), Code::InvalidPath, "{text}");
        assert!(error.message().starts_with("the row number"), "{error}");
    }
    // The grammar's typed error keeps an offset into the text as written.
    let error = PathArg::parse("Sheet1!A1[0]").unwrap_err();
    assert_eq!(error.code(), Code::InvalidPath);
    assert_eq!(
        error.message(),
        "invalid path at byte 10: index 0 is invalid; indices are 1-based"
    );
    let Source::Selector(parse) = error.leaf() else {
        panic!("{error:?}")
    };
    assert_eq!(parse.offset(), 10);
    // The sheet step counts against the step limit.
    let deep = format!("Sheet1!{}", vec!["a"; limits::MAX_STEPS].join("/"));
    let error = PathArg::parse(&deep).unwrap_err();
    assert_eq!(error.code(), Code::MaxDepthExceeded);
    let fits = format!("Sheet1!{}", vec!["a"; limits::MAX_STEPS - 1].join("/"));
    assert_eq!(arg(&fits).path().steps().len(), limits::MAX_STEPS);
}

#[test]
fn grammar_errors_convert_with_their_code_and_offset() {
    let path = Path::parse("/body/p[-1]").unwrap_err();
    let error = Error::from(path);
    assert_eq!(error.code(), Code::InvalidPath);
    assert_eq!(error.message(), path.to_string());
    assert_eq!(error.suggestion(), path.suggestion());
    assert!(error.suggestion().is_some());
    assert_eq!(error.format(), None);
    assert!(std::error::Error::source(&error).is_some());

    let query: ParseError = Query::parse("p[style=Heading1").unwrap_err();
    let error = Error::from(query);
    assert_eq!(error.code(), Code::InvalidSelector);
    assert_eq!(error.suggestion(), None);

    let deep = "/a".repeat(limits::MAX_STEPS + 1);
    let error = PathArg::parse(&deep).unwrap_err();
    assert_eq!(error.code(), Code::MaxDepthExceeded);
    assert_eq!(error.code().as_str(), "max_depth_exceeded");
}

// --------------------------------------------------------------- resolution

#[test]
fn word_paths_resolve_with_aliases_and_case_folding() {
    assert_eq!(get::<Word>("/").kind(), "document");
    assert_eq!(get::<Word>("/body").kind(), "body");
    let second = get::<Word>("/body/p[2]");
    assert_eq!(second.attr("paraId"), Some("00000002"));
    assert_eq!(get::<Word>("/BODY/Paragraph[2]"), second);
    assert_eq!(get::<Word>("/body/P[2]/R[1]").kind(), "r");
    assert_eq!(
        get::<Word>("/body/p[last()]").attr("paraId"),
        Some("0000000A")
    );
    assert_eq!(get::<Word>("/body/p").attr("paraId"), Some("1A2B3C4D"));
    // `/body/p[N]` counts paragraphs only, past the table and the content control.
    assert_eq!(children::<Word>(get::<Word>("/body/p[3]"), KindSet::ALL), 3);

    let first = get::<Word>("/body/p[1]");
    assert_eq!(get::<Word>("/body/p[@paraId=1a2b3c4d]"), first);
    assert_eq!(
        get::<Word>("/body/p[@PARAID=1A2B3C4D][@textId=77aa0001]"),
        first
    );
    assert_eq!(
        get::<Word>("/body/paragraph[@textId=77AA0001]/hyperlink[1]").kind(),
        "hyperlink"
    );

    assert_eq!(
        get::<Word>("/body/tbl[1]/tr[2]/tc[1]/p[1]/r[2]").kind(),
        "r"
    );
    assert_eq!(get::<Word>("/body/table[1]/row[1]/cell[2]").kind(), "tc");
    assert_eq!(get::<Word>("/body/sdt[@sdtId=7]/p[1]/r[1]").kind(), "r");
    assert_eq!(
        get::<Word>("/body/contentcontrol[1]").attr("sdtId"),
        Some("7")
    );

    assert_eq!(get::<Word>("/header[2]").attr("type"), Some("first"));
    assert_eq!(
        get::<Word>("/header[@type=FIRST]").attr("type"),
        Some("first")
    );
    assert_eq!(get::<Word>("/footer[1]/p[1]").kind(), "p");
    // `/footnote[N]` is note id N; `/footnotes/footnote[N]` is the Nth note.
    assert_eq!(get::<Word>("/footnote[2]").attr("footnoteId"), Some("2"));
    assert_eq!(
        get::<Word>("/footnotes/footnote[2]").attr("footnoteId"),
        Some("2")
    );
    assert_eq!(
        get::<Word>("/footnotes/footnote[1]").attr("footnoteId"),
        Some("1")
    );
    assert_eq!(
        get::<Word>("/footnotes/footnote[@footnoteId=1]/p[1]/r[1]").kind(),
        "r"
    );
    assert_eq!(
        get::<Word>("/comments/comment[@commentId=1]").attr("commentId"),
        Some("1")
    );

    // Style ids are names, compared exactly, never aliased.
    assert_eq!(
        get::<Word>("/styles/Heading1").attr("name"),
        Some("Heading1")
    );
    assert_eq!(get::<Word>("/styles/Heading1/tab[1]").kind(), "tab");
    assert_eq!(find::<Word>("/styles/heading1"), None);

    // `level[L]` is `w:ilvl` = L.
    assert_eq!(
        get::<Word>("/numbering/abstractNum[@id=1]/level[1]").attr("ilvl"),
        Some("1")
    );
    assert_eq!(
        get::<Word>("/numbering/abstractnum[1]/lvl[@ilvl=0]").attr("ilvl"),
        Some("0")
    );
    assert_eq!(get::<Word>("/numbering/num[@id=1]").kind(), "num");

    assert_eq!(get::<Word>("/formfield[agree]").attr("name"), Some("agree"));
    assert_eq!(find::<Word>("/formfield[AGREE]"), None);
    assert_eq!(get::<Word>("/bookmark[@name=Intro]").kind(), "bookmark");
    assert_eq!(find::<Word>("/bookmark[@name=intro]"), None);
    assert_eq!(get::<Word>("/section").kind(), "section");
}

#[test]
fn workbook_paths_resolve_sheets_cells_rows_and_columns() {
    let sheet = get::<Workbook>("/Sheet1");
    assert_eq!(sheet.attr("name"), Some("Sheet1"));
    assert_eq!(get::<Workbook>("/sheet1"), sheet);
    let at = |text: &str| super::CellRef::parse(text).expect(text);
    assert_eq!(get::<Workbook>("/sheet1/a1"), Node::Cell(at("A1")));
    assert_eq!(get::<Workbook>("Sheet1!Z99"), Node::Cell(at("Z99")));
    assert_eq!(
        get::<Workbook>("/Sheet1/A1:C5"),
        Node::Range(super::CellRange::parse("A1:C5").expect("range"))
    );
    assert_eq!(get::<Workbook>("'Data Q1'!A1:C5").kind(), "range");

    // `/sheet[N]` is the Nth tab; a sheet named like an element is a sheet.
    assert_eq!(get::<Workbook>("/sheet[2]").attr("name"), Some("Data Q1"));
    assert_eq!(
        get::<Workbook>("/sheet[last()]").attr("name"),
        Some("Straße")
    );
    assert_eq!(get::<Workbook>("/Chart").attr("name"), Some("Chart"));
    assert_eq!(get::<Workbook>("/Chart/B2"), Node::Cell(at("B2")));
    assert_eq!(get::<Workbook>("/STRAßE/B2"), Node::Cell(at("B2")));

    // Rows and columns are numbers; a row without a record still resolves.
    assert_eq!(get::<Workbook>("/Sheet1/row[3]").attr("n"), Some("3"));
    assert_eq!(
        get::<Workbook>("Sheet1!3:3"),
        get::<Workbook>("/Sheet1/row[3]")
    );
    assert_eq!(get::<Workbook>("/Sheet1/row[5]"), Node::Row(5));
    assert_eq!(get::<Workbook>("'Data Q1'!3:3"), Node::Row(3));
    let column = get::<Workbook>("/Sheet1/col[B]");
    assert_eq!(column.attr("n"), Some("2"));
    assert_eq!(get::<Workbook>("/Sheet1/column[2]"), column);
    assert_eq!(get::<Workbook>("Sheet1!b:b"), column);

    assert_eq!(
        get::<Workbook>("/namedrange[total]").attr("name"),
        Some("Total")
    );
    assert_eq!(
        get::<Workbook>("/namedrange[@name=RATES]").attr("name"),
        Some("Rates")
    );
    assert_eq!(
        get::<Workbook>("/definedname[2]").attr("name"),
        Some("Rates")
    );
    assert_eq!(
        get::<Workbook>("/Sheet1/table[@name=sales]").kind(),
        "table"
    );
    assert_eq!(get::<Workbook>("/Sheet1/listobject[1]").kind(), "table");
    assert_eq!(get::<Workbook>("/Sheet1/note[1]").kind(), "comment");
}

#[test]
fn presentation_paths_resolve_identities_tokens_and_tables() {
    let title = get::<Presentation>("/slide[1]/shape[@name=title 1]");
    assert_eq!(title.attr("id"), Some("2"));
    assert_eq!(get::<Presentation>("/slide[1]/shape[@id=2]"), title);
    assert_eq!(get::<Presentation>("/Slide[1]/textbox[1]"), title);
    assert_eq!(
        get::<Presentation>("/slide[1]/title[2]").attr("id"),
        Some("3")
    );
    assert_eq!(
        get::<Presentation>("/slide[2]/shape[@name=äRGER]").attr("id"),
        Some("2")
    );

    assert_eq!(
        get::<Presentation>("/slide[1]/placeholder[@type=BODY][@idx=1]").attr("id"),
        Some("4")
    );
    assert_eq!(
        get::<Presentation>("/slide[1]/placeholder[ctrTitle]").attr("id"),
        Some("6")
    );
    assert_eq!(
        get::<Presentation>("/slide[1]/ph[CTRTITLE]").attr("id"),
        Some("6")
    );
    assert_eq!(
        get::<Presentation>("/slide[1]/placeholder[2]").attr("id"),
        Some("5")
    );

    assert_eq!(
        get::<Presentation>("/slide[1]/table[@name=grid]/tr[1]/tc[2]").kind(),
        "tc"
    );
    assert_eq!(
        get::<Presentation>("/slide[1]/table[1]/row[1]/cell[2]").kind(),
        "tc"
    );
    assert_eq!(get::<Presentation>("/slide[1]/notes").kind(), "notes");
    assert_eq!(
        get::<Presentation>("/slidemaster[1]/slidelayout[1]/shape[1]").kind(),
        "shape"
    );
    assert_eq!(
        get::<Presentation>("/slideMaster[1]/layout[1]/shape[1]").kind(),
        "shape"
    );
}

#[test]
fn relative_paths_resolve_from_a_start_node() {
    let nav = Fake::<Word>::new();
    let body = get::<Word>("/body");
    let run =
        resolve::get_from(&nav, body, kind::<Word>("body"), &arg("p[2]/r[1]")).expect("found");
    assert_eq!(run.kind(), "r");
    // An absolute path starts at the root whatever the start node.
    let header =
        resolve::get_from(&nav, body, kind::<Word>("body"), &arg("/header[1]")).expect("found");
    assert_eq!(header.kind(), "header");
    // The empty relative path is the start node.
    let same = resolve::get_from(
        &nav,
        body,
        kind::<Word>("body"),
        &PathArg::from(Path::relative([])),
    );
    assert_eq!(same.expect("found"), body);
    assert_eq!(
        resolve::find_from(&nav, body, kind::<Word>("body"), &arg("p[99]")).expect("no error"),
        None
    );
    let error = resolve::get_from(&nav, body, kind::<Word>("body"), &arg("p[99]")).unwrap_err();
    assert_error(
        &error,
        Code::PathNotFound,
        "requested p was not found at byte 0",
        Some("available at the parent element: p[1]..p[4], tbl[1], sdt[1]"),
    );
    let error = resolve::get_from(&nav, body, kind::<Word>("body"), &arg("p[1]/r[9]")).unwrap_err();
    assert_eq!(error.message(), "requested r was not found at byte 5");
}

#[test]
fn find_returns_none_for_every_miss_and_get_reports_it() {
    let misses = [
        "/body/p[99]",
        "/body/p[99]/r[1]",
        "/body/p[@paraId=DEADBEEF]",
        "/styles/Missing",
        "/footnote[9]",
        "/header[@type=even]",
    ];
    for text in misses {
        assert_eq!(find::<Word>(text), None, "{text}");
        assert_eq!(fail::<Word>(text).code(), Code::PathNotFound, "{text}");
    }
    for text in [
        "/Sheet9",
        "/Sheet9/A1",
        "/Q1!A1",
        "/namedrange[Nope]",
        "/Sheet1/mergedrange[1]",
    ] {
        assert_eq!(find::<Workbook>(text), None, "{text}");
        assert_eq!(fail::<Workbook>(text).code(), Code::PathNotFound, "{text}");
    }
    assert_eq!(find::<Presentation>("/slide[3]"), None);
    // Hits are hits for both.
    assert_eq!(find::<Word>("/body/p[1]").map(Node::kind), Some("p"));
}

// ------------------------------------------------------------------- errors

#[test]
fn misses_are_path_not_found_with_what_exists() {
    let error = fail::<Word>("/body/p[99]");
    assert_error(
        &error,
        Code::PathNotFound,
        "requested p was not found at byte 6",
        Some("available at the parent element: p[1]..p[4], tbl[1], sdt[1]"),
    );
    assert_eq!(error.format(), Some(litchi_core::FileFormat::Docx));
    assert_eq!(
        error.to_string(),
        "path_not_found: requested p was not found at byte 6 (suggestion: available at the parent element: p[1]..p[4], tbl[1], sdt[1])"
    );
    assert_error(
        &fail::<Word>("/body/p[@paraId=DEADBEEF]"),
        Code::PathNotFound,
        "requested p was not found at byte 6",
        Some("available at the parent element: p[1]..p[4], tbl[1], sdt[1]"),
    );
    assert_error(
        &fail::<Word>("/styles/Missing"),
        Code::PathNotFound,
        "requested style was not found at byte 8",
        Some("available at the parent element: named child, named child"),
    );
    assert_error(
        &fail::<Workbook>("/Sheet9/A1"),
        Code::PathNotFound,
        "requested sheet was not found at byte 1",
        Some("available at the parent element: named child, named child, named child, named child"),
    );
    assert_error(
        &fail::<Workbook>("/Q1!A1"),
        Code::PathNotFound,
        "requested sheet was not found at byte 1",
        Some("available at the parent element: named child, named child, named child, named child"),
    );
    assert_error(
        &fail::<Workbook>("/Sheet1/mergedrange[1]"),
        Code::PathNotFound,
        "requested mergeCell was not found at byte 8",
        Some("available at the parent element: row[1]..row[2], col[1], table[1], comment[1]"),
    );
    assert_error(
        &fail::<Word>("/numbering/abstractNum[@id=1]/level[7]"),
        Code::PathNotFound,
        "requested level was not found at byte 30",
        Some("available at the parent element: level[1]..level[2]"),
    );
}

#[test]
fn unknown_names_are_invalid_paths_with_did_you_mean() {
    assert_error(
        &fail::<Word>("/body/paragrph[1]"),
        Code::InvalidPath,
        "unknown element at byte 6",
        Some("did you mean \"paragraph\"?"),
    );
    // A transposition costs one, as in OfficeCLI's Damerau suggester.
    assert_eq!(
        fail::<Word>("/body/tabel[1]").suggestion(),
        Some("did you mean \"table\"?")
    );
    assert_eq!(
        fail::<Word>("/body/x[1]").suggestion(),
        Some("did you mean one of \"p\", \"r\"?")
    );
    let far = fail::<Word>("/body/zzzzzzzz[1]");
    assert_eq!(far.code(), Code::InvalidPath);
    let known = far.suggestion().expect("suggestion");
    assert!(
        known.starts_with("known elements: abstractNum, axis, body, bookmark"),
        "{known}"
    );
    assert!(!known.contains("document"), "{known}");
    assert_eq!(
        fail::<Presentation>("/slide[1]/shap[1]").suggestion(),
        Some("did you mean \"shape\"?")
    );
    // Coordinates outside the grid are not cells.
    assert_error(
        &fail::<Workbook>("/Sheet1/XFE1"),
        Code::InvalidPath,
        "cell reference is outside the worksheet grid A1:XFD1048576 at byte 8",
        Some("columns run from A to XFD and rows from 1 to 1048576"),
    );
    assert_eq!(fail::<Workbook>("/Sheet1/A0:B2").code(), Code::InvalidPath);
    assert_eq!(
        fail::<Workbook>("/Sheet1/frobnicate")
            .suggestion()
            .map(|text| text.starts_with("known elements")),
        Some(true)
    );
}

#[test]
fn concepts_of_other_families_are_unsupported_paths() {
    assert_error(
        &fail::<Word>("/slide[1]"),
        Code::UnsupportedPath,
        "Docx documents do not have this element, which belongs to presentations at byte 1",
        Some(
            "Docx paths start at /body, /header[N], /footer[N], /footnote[N], /endnote[N], /comments/comment[N], /styles/<style id>, /numbering or /section[N]",
        ),
    );
    assert_error(
        &fail::<Workbook>("/body"),
        Code::UnsupportedPath,
        "Xlsx documents do not have this element, which belongs to word-processing documents at byte 1",
        Some(
            "Xlsx paths start at /<sheet name>, /sheet[N] or /namedrange[NAME] (cells as /Sheet1/A1 or Sheet1!A1)",
        ),
    );
    assert_error(
        &fail::<Presentation>("/Sheet1/A1"),
        Code::UnsupportedPath,
        "Pptx documents have no worksheets or cells at byte 1",
        Some(
            "Pptx paths start at /slide[N], /slidemaster[N], /slidelayout[N], /notesMaster or /theme",
        ),
    );
    assert_eq!(
        fail::<Presentation>("/sheet[1]/A1").code(),
        Code::UnsupportedPath
    );
    assert_error(
        &fail::<Word>("Sheet1!A1"),
        Code::UnsupportedPath,
        "Docx documents have no worksheets or cells at byte 0",
        Some(
            "Docx paths start at /body, /header[N], /footer[N], /footnote[N], /endnote[N], /comments/comment[N], /styles/<style id>, /numbering or /section[N]",
        ),
    );
    // A concept of the family that this format lacks.
    assert_eq!(
        fail::<Word>("/customXml").message(),
        "Docx documents have no \"customXml\" elements at byte 1"
    );
    assert_eq!(fail::<Word>("/customXml").code(), Code::UnsupportedPath);
    // A kind that cannot occur under its parent.
    assert_error(
        &fail::<Word>("/body/tr[1]"),
        Code::UnsupportedPath,
        "tr elements cannot occur under the parent element at byte 6",
        Some("available at the parent element: p[1]..p[4], tbl[1], sdt[1]"),
    );
    assert_eq!(fail::<Workbook>("/row[1]").code(), Code::UnsupportedPath);
    assert_eq!(
        fail::<Presentation>("/shape[1]").code(),
        Code::UnsupportedPath
    );
}

#[test]
fn elements_without_a_model_are_unsupported_capabilities() {
    assert_error(
        &fail::<Word>("/body/chart[1]"),
        Code::UnsupportedCapability,
        "reading \"chart\" elements from Docx is not supported yet at byte 6",
        Some("follow-up L-TEST-R1; capabilities() lists this gap"),
    );
    assert_error(
        &fail::<Word>("/tableofcontents[1]"),
        Code::UnsupportedCapability,
        "reading \"toc\" elements from Docx is not supported yet at byte 1",
        Some("capabilities() lists the elements Litchi reads from this format"),
    );
    assert_eq!(
        fail::<Workbook>("/Sheet1/pivot[1]").code(),
        Code::UnsupportedCapability
    );
    assert_eq!(
        fail::<Workbook>("/Sheet1/slicer[1]").code(),
        Code::UnsupportedCapability
    );
    assert_eq!(
        fail::<Presentation>("/slide[1]/chart[1]").code(),
        Code::UnsupportedCapability
    );
    assert_eq!(
        fail::<Presentation>("/slide[1]/slidezoom[1]").code(),
        Code::UnsupportedCapability
    );
}

#[test]
fn predicates_that_cannot_address_one_element_are_invalid_paths() {
    assert_error(
        &fail::<Word>("/body/p[@foo=1]"),
        Code::InvalidPath,
        "attribute is not an identity of p at byte 6",
        Some("p accepts p[N], p[last()]"),
    );
    assert_error(
        &fail::<Word>("/body/tbl[@id=1]"),
        Code::InvalidPath,
        "attribute is not an identity of tbl at byte 6",
        Some("tbl accepts tbl[N], tbl[last()]"),
    );
    assert_error(
        &fail::<Word>("/body/p[@paraId]"),
        Code::InvalidPath,
        "identity needs a value at byte 6",
        Some("p accepts p[N], p[last()]"),
    );
    assert_error(
        &fail::<Word>("/body/p[style=Heading1]"),
        Code::InvalidPath,
        "the requested step filters by content at byte 6",
        Some(
            "get and find address one element by position or identity; select by content with query",
        ),
    );
    for text in [
        "/body/p[@paraId!=1]",
        "/body/p[@paraId=1 or @paraId=2]",
        "/body/p[@paraId=1][x>2]",
    ] {
        assert_eq!(
            fail::<Word>(text).message().split(" filters").count(),
            2,
            "{text}"
        );
    }
    assert_error(
        &fail::<Word>("/body/p[1][2]"),
        Code::InvalidPath,
        "the requested step stacks predicates that are not identities at byte 6",
        Some("use one [N], [last()] or bare token, or stack only [@key=value] identities"),
    );
    assert_eq!(
        fail::<Word>("/body/p[@paraId=1A2B3C4D][last()]").code(),
        Code::InvalidPath
    );
    assert_error(
        &fail::<Word>("/body/p[XYZ]"),
        Code::InvalidPath,
        "p does not take this bare token at byte 6",
        Some("p accepts p[N], p[last()]"),
    );
    assert_error(
        &fail::<Word>("/styles/Heading1[1]"),
        Code::InvalidPath,
        "a style is named by the step alone, without predicates at byte 8",
        Some("write the child name alone"),
    );
    // The Excel-notation hint is for worksheet names only.
    assert_error(
        &fail::<Word>("/styles/Q1!A1"),
        Code::PathNotFound,
        "requested style was not found at byte 8",
        Some("available at the parent element: named child, named child"),
    );
    assert_error(
        &fail::<Presentation>("/slide[1]/placeholder[@idx=1]"),
        Code::InvalidPath,
        "the requested step matches 2 elements under the parent element at byte 10",
        Some(
            "add another identity predicate to pick one; available at the parent element: shape[1]..shape[2], placeholder[1]..placeholder[3], table[1], notes[1]",
        ),
    );
}

#[test]
fn coordinates_outside_the_grid_are_invalid_paths() {
    assert_error(
        &fail::<Workbook>("/Sheet1/col[XFE]"),
        Code::InvalidPath,
        "invalid worksheet column at byte 8",
        Some("columns run from A to XFD, e.g. col[B] or col[2]"),
    );
    assert_error(
        &fail::<Workbook>("/Sheet1/row[1048577]"),
        Code::InvalidPath,
        "row 1048577 is outside the worksheet grid at byte 8",
        Some("rows run from 1 to 1048576"),
    );
    assert_eq!(
        fail::<Workbook>("/Sheet1/col[16385]").message(),
        "column 16385 is outside the worksheet grid at byte 8"
    );
    assert_eq!(
        get::<Workbook>("/Sheet1/row[1048576]"),
        Node::Row(1_048_576)
    );
    // A hand-built index with no 1-based spelling.
    let path =
        Path::from_steps([Step::new("footnote").with(Predicate::Index(Position::new(usize::MAX)))]);
    let error = resolve::get(&Fake::<Word>::new(), &PathArg::from(&path)).unwrap_err();
    assert_eq!(error.code(), Code::InvalidPath);
    assert!(
        error.message().starts_with("the index is too large"),
        "{error}"
    );
}

#[test]
fn document_paths_must_be_absolute() {
    assert_error(
        &fail::<Word>("body/p[1]"),
        Code::InvalidPath,
        "a document path must start with '/' at byte 0",
        Some("prefix the path with '/'"),
    );
    let relative = Path::relative([Step::new("p")]);
    let error = resolve::find(&Fake::<Word>::new(), &PathArg::from(relative)).unwrap_err();
    assert_eq!(error.suggestion(), Some("prefix the path with '/'"));
}

#[test]
fn offsets_point_at_the_failing_step_in_the_text_as_written() {
    let at = |text: &str| -> String {
        let message = fail::<Word>(text).message().to_owned();
        let start = message.find("at byte ").expect("offset") + "at byte ".len();
        message[start..]
            .split(' ')
            .next()
            .expect("number")
            .to_owned()
    };
    assert_eq!(at("/body/p[99]"), "6");
    assert_eq!(at("/body/p[ 99 ]"), "6");
    assert_eq!(at("/\"body\"/p[99]"), "8");
    let text = "/body/p[@paraId=\"1A2B3C4D\"]/r[9]";
    assert_eq!(
        at(text),
        (text.find("/r[9]").expect("step") + 1).to_string()
    );
    // Quoted values may hold `/` and `]`.
    let text = "/bookmark[@name='a/b]']/x[1]";
    assert_eq!(
        at(text),
        (text.rfind("/x[1]").expect("step") + 1).to_string()
    );

    // Parsed paths count in their canonical spelling.
    let path = Path::parse("/body/p[ 99 ]").expect("parses");
    let error = resolve::get(&Fake::<Word>::new(), &PathArg::from(&path)).unwrap_err();
    assert_eq!(error.message(), "requested p was not found at byte 6");

    // Excel notation: step 0 is the sheet, later steps start after `!`.
    let error = fail::<Workbook>("Sheet1!A1/run[1]");
    assert_error(
        &error,
        Code::UnsupportedPath,
        "run elements cannot occur under the parent element at byte 10",
        None,
    );
    // `r` is a Word name, not a workbook one.
    assert_eq!(
        fail::<Workbook>("Sheet1!A1/r[1]").code(),
        Code::UnsupportedPath
    );
    let error = fail::<Workbook>("'No Such'!A1");
    assert_eq!(error.message(), "requested sheet was not found at byte 0");

    // Long paths are shortened in messages at a character boundary.
    let long = format!("/body/p[99]/{}", "ä".repeat(150));
    let error = fail::<Word>(&long);
    assert!(!error.message().contains("ä"), "{error}");
    assert!(error.message().len() < 300, "{error}");
}

// ------------------------------------------------------------------ helpers

#[test]
fn kind_sets_hold_kinds() {
    let p = kind::<Word>("p");
    let tbl = kind::<Word>("tbl");
    let set = KindSet::of(p);
    assert!(set.contains(p) && !set.contains(tbl));
    assert!(set.with(tbl).contains(tbl));
    assert_eq!(set.union(KindSet::of(tbl)), set.with(tbl));
    assert!(KindSet::<Kind<Word>>::EMPTY.is_empty() && !set.is_empty());
    assert!(KindSet::ALL.contains(tbl));
    let copy = set;
    assert_eq!(copy, set);
    assert_eq!(
        format!("{set:?}"),
        format!("KindSet({:#x})", 1_u128 << p.ordinal())
    );
    assert_eq!(format!("{p:?}"), "p");
    assert_eq!(Kind::<Word>::ROOT.name(), "document");
    assert_eq!(Kind::<Word>::from_name("toc"), None);
    assert_eq!(Kind::<Word>::from_name("customXml"), None);
}

#[test]
fn children_are_listed_by_kind_in_document_order() {
    let body = get::<Word>("/body");
    assert_eq!(children::<Word>(body, KindSet::of(kind::<Word>("p"))), 4);
    assert_eq!(
        children::<Word>(
            body,
            KindSet::of(kind::<Word>("p")).with(kind::<Word>("tbl"))
        ),
        5
    );
    assert_eq!(children::<Word>(body, KindSet::ALL), 6);
    assert_eq!(children::<Word>(body, KindSet::EMPTY), 0);
    assert_eq!(
        children::<Workbook>(get::<Workbook>("/Sheet1/A1"), KindSet::ALL),
        0
    );
}

#[test]
fn identities_list_their_keys_and_values() {
    let path = Path::parse("/x[@PARAID=1A][@textId=2B]").expect("parses");
    let rule = WORD.rule("p").expect("p");
    let identities = Identities::new(path.steps()[0].predicates(), rule.identities);
    let listed: Vec<(&str, &str)> = identities
        .iter()
        .map(|(identity, value)| (identity.key, value))
        .collect();
    assert_eq!(listed, [("paraId", "1A"), ("textId", "2B")]);
    assert_eq!(format!("{identities:?}"), "Identities { .. }");
}

#[test]
fn lookups_and_resolved_steps_describe_themselves() {
    let resolved = Resolved {
        kind: kind::<Word>("p"),
        select: Select::Ordinal(Position::new(1)),
    };
    let copy = resolved;
    assert_eq!(copy, resolved);
    assert!(format!("{resolved:?}").contains("Ordinal"));
    let lookup: Lookup<Node, super::fake::Gap> = Lookup::Ambiguous(2);
    assert_eq!(format!("{lookup:?}"), "Ambiguous(2)");
}

#[test]
fn available_lists_kinds_and_names_within_a_bound() {
    let mut available = Available::default();
    assert!(available.is_empty());
    available.kind("p", 0);
    assert!(available.is_empty());
    available.kind("p", 1);
    available.kind("tbl", 3);
    available.name("a/b");
    assert_eq!(available.render(), "p[1], tbl[1]..tbl[3], named child");
    let mut many = Available::default();
    for index in 0..40 {
        many.name(&format!("Sheet{index}"));
    }
    let rendered = many.render();
    assert!(rendered.ends_with("named child and 8 more"), "{rendered}");
}

#[test]
fn edit_distance_is_optimal_string_alignment() {
    let chars = |text: &str| text.chars().collect::<Vec<_>>();
    assert_eq!(distance(&chars("blod"), &chars("bold")), 1);
    assert_eq!(distance(&chars("kitten"), &chars("sitting")), 3);
    assert_eq!(distance(&chars(""), &chars("abc")), 3);
    assert_eq!(distance(&chars("abc"), &chars("")), 3);
    assert_eq!(distance(&chars("table"), &chars("table")), 0);
    assert_eq!(distance(&chars("ca"), &chars("abc")), 3);
}

/// A minimal adapter: it keeps `Navigate::available`'s default and maps the
/// canonical `p` to a kind whose name is missing from its vocabulary, as a
/// buggy adapter might.
mod minimal {
    use litchi_core::FileFormat;

    use super::super::fake::Gap;
    use super::super::navigate::{ElementKind, KindSet, Lookup, Navigate, Resolved};
    use super::super::vocabulary::{self, Vocabulary};
    use crate::Result;

    #[derive(Debug, Clone, Copy, PartialEq, Eq)]
    pub(super) enum Kind {
        Root,
        Body,
        Broken,
    }

    impl ElementKind for Kind {
        const VOCABULARY: &'static Vocabulary = &vocabulary::WORD;
        const ROOT: Self = Self::Root;

        fn from_name(name: &str) -> Option<Self> {
            match name {
                "body" => Some(Self::Body),
                "p" => Some(Self::Broken),
                _ => None,
            }
        }

        fn name(self) -> &'static str {
            match self {
                Self::Root => "document",
                Self::Body => "body",
                Self::Broken => "not-in-the-vocabulary",
            }
        }

        fn ordinal(self) -> u8 {
            self as u8
        }
    }

    pub(super) struct Minimal;

    impl Navigate<'_> for Minimal {
        type Kind = Kind;
        type Node = Kind;
        type Gap = Gap;
        type Children = std::iter::Empty<Result<Kind>>;

        const FORMAT: FileFormat = FileFormat::Doc;

        fn root(&self) -> Result<Kind> {
            Ok(Kind::Root)
        }

        fn child(&self, _parent: &Kind, step: &Resolved<'_, Kind>) -> Result<Lookup<Kind, Gap>> {
            Ok(match step.kind {
                Kind::Body => Lookup::Missing,
                kind => Lookup::Found(kind),
            })
        }

        fn children(&self, _parent: &Kind, _kinds: KindSet<Kind>) -> Result<Self::Children> {
            Ok(std::iter::empty())
        }
    }
}

#[test]
fn adapter_defaults_and_bugs_never_panic() {
    use minimal::{Kind, Minimal};

    // The default `available` describes nothing, so a miss has no suggestion.
    let error = resolve::get(&Minimal, &arg("/body")).unwrap_err();
    assert_error(
        &error,
        Code::PathNotFound,
        "requested body was not found at byte 1",
        None,
    );
    assert_eq!(error.format(), Some(litchi_core::FileFormat::Doc));
    // Names the adapter does not map are capability gaps of the format.
    assert_eq!(
        resolve::get(&Minimal, &arg("/header[1]"))
            .unwrap_err()
            .code(),
        Code::UnsupportedCapability
    );
    // A kind missing from its vocabulary is an internal error, not a panic.
    let error = resolve::get(&Minimal, &arg("/p[1]/r[1]")).unwrap_err();
    assert_error(
        &error,
        Code::InternalError,
        "element kind \"not-in-the-vocabulary\" is missing from its vocabulary at byte 1",
        None,
    );
    assert_eq!(
        Minimal
            .children(&Kind::Root, KindSet::ALL)
            .expect("children")
            .count(),
        0
    );
}

#[test]
fn ordinary_paths_use_semantic_names_and_raw_paths_opt_into_native_ids() {
    let nav = Fake::<Word>::new();
    for text in [
        "/body/p[@paraId=1A2B3C4D]",
        "/numbering/abstractNum[1]/level[@ilvl=0]",
    ] {
        let error = resolve::get(&nav, &PathArg::parse(text).unwrap()).unwrap_err();
        assert_eq!(error.code(), Code::InvalidPath);
        assert!(error.message().contains("raw"));
    }
    let text = "/body/p[@paraId=1A2B3C4D]";
    assert!(resolve::get(&nav, &super::raw::path(text).unwrap()).is_ok());
    assert!(resolve::get(&nav, &PathArg::parse("/body/p[1]").unwrap()).is_ok());
    let unknown =
        resolve::get(&nav, &PathArg::parse("/body/p[@PRIVATE=PRIVATE]").unwrap()).unwrap_err();
    assert!(unknown.message().contains("not an identity"));
    assert!(!format!("{unknown:?}").contains("PRIVATE"));
    let parts = [
        crate::node::PathStep::new("body"),
        crate::node::raw::identity_step("p", "paraId", crate::node::raw::Ident::Text("1A2B3C4D")),
    ];
    let path = crate::node::NodePath::new(&parts).to_path();
    assert!(resolve::get(&nav, &super::raw::path(path).unwrap()).is_ok());
}

#[test]
fn all_path_forms_are_bounded_before_resolution() {
    for suffix in ["!3:3", "!B:B", "!A1"] {
        let text = format!("{}{}", "x".repeat(limits::MAX_INPUT_BYTES), suffix);
        let error = PathArg::parse(&text).unwrap_err();
        let Source::Selector(parse) = error.leaf() else {
            panic!("typed limit lost")
        };
        assert_eq!(
            parse.kind(),
            super::ErrorKind::LimitExceeded(super::Limit::InputBytes)
        );
    }
    for text in ["'bad\nname'!A1", "bad\0!3:3"] {
        assert_eq!(PathArg::parse(text).unwrap_err().code(), Code::InvalidPath);
    }
    let nav = Fake::<Word>::new();
    let deep = Path::from_steps((0..=limits::MAX_STEPS).map(|_| Step::new("p")));
    assert_eq!(
        resolve::get(&nav, &PathArg::from(&deep))
            .unwrap_err()
            .code(),
        Code::MaxDepthExceeded
    );
    let oversized = Path::from_steps([Step::new("x".repeat(limits::MAX_INPUT_BYTES + 1))]);
    assert_eq!(
        resolve::get(&nav, &PathArg::from(&oversized))
            .unwrap_err()
            .code(),
        Code::InvalidPath
    );
    let mut step = Step::new("p");
    for _ in 0..=limits::MAX_PREDICATES {
        step = step.with(Predicate::Last);
    }
    let many = Path::from_steps([step]);
    assert_eq!(
        resolve::get(&nav, &PathArg::from(&many))
            .unwrap_err()
            .code(),
        Code::InvalidPath
    );
}

#[test]
fn path_diagnostics_do_not_echo_authored_input_or_document_names() {
    for text in [
        "/PRIVATE",
        "/body/p[@name=PRIVATE]",
        "/body/p[99]/PRIVATE",
        "PRIVATE!",
        "'PRIVATE'!A1",
    ] {
        let input = PathArg::parse(text);
        if let Ok(input) = &input {
            assert!(!format!("{input:?}").contains("PRIVATE"));
        }
        let error = match input {
            Ok(input) => resolve::get(&Fake::<Workbook>::new(), &input).unwrap_err(),
            Err(error) => error,
        };
        assert!(!format!("{error:?}").contains("PRIVATE"));
        assert!(!error.to_string().contains("Data Q1"));
    }
    let parsed = PathArg::parse("PRIVATE").unwrap();
    let error = resolve::get(&Fake::<Word>::new(), &parsed).unwrap_err();
    assert!(!error.to_string().contains("PRIVATE"));
    let error = Error::from(Path::parse("PRIVATE").unwrap_err());
    assert!(!format!("{error:?}").contains("'P'"));
    assert!(
        !std::error::Error::source(&error)
            .unwrap()
            .to_string()
            .contains("'P'")
    );
}

#[test]
fn semantic_name_matching_uses_canonical_unicode_case_folding() {
    assert!(Case::Caseless.matches("Straße", "STRASSE"));
    assert!(Case::Caseless.matches("é", "e\u{301}"));
    assert_eq!(
        get::<Workbook>("/STRASSE/B2"),
        get::<Workbook>("/Straße/B2")
    );
}
