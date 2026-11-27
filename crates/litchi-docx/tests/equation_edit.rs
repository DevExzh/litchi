//! Office Math equations of an opened document: inserted, replaced and
//! removed through the main-document transaction (and, in headers and notes,
//! through the story edit), with every byte outside the edited equation
//! kept.

use std::collections::BTreeMap;

use litchi_core::Position;
use litchi_docx::document::{
    ElementAnchor, ElementPath, ElementStep, Error, NewNote, NewStory, NewTable, NoteKind,
    ParagraphEquation, Refusal, StoryTarget,
};
use litchi_docx::header_footer::{Kind, Role};
use litchi_docx::{Builder, OfficeMath, OfficeMathParagraph, Package};

/// A Word document with one inline equation (`x² + 1/2 + √9`, beside a run)
/// and one display equation (`∑ₖ k = 10`).
const MATH: &str = concat!(
    env!("CARGO_MANIFEST_DIR"),
    "/../../test-data/microsoft-office/word/word-math-navigation.docx"
);
/// A Strict-conformance document with a display equation.
const STRICT: &str = concat!(
    env!("CARGO_MANIFEST_DIR"),
    "/../../test-data/libreoffice-core/sw/qa/extras/ooxmlexport/data/strict.docx"
);
const MAIN: &str = "/word/document.xml";

fn open(path: &str) -> Package {
    Package::from_bytes(&std::fs::read(path).unwrap()).unwrap()
}

fn paragraph(position: usize) -> ElementPath {
    ElementPath::new([ElementStep::Paragraph(Position::new(position))])
}

/// Every part's payload, after a save and a reopen.
fn saved_parts(package: &Package) -> BTreeMap<String, Vec<u8>> {
    let reopened = Package::from_bytes(&package.to_bytes().unwrap()).unwrap();
    let names: Vec<String> = reopened
        .opc_package()
        .iter_parts()
        .map(|part| part.partname().to_string())
        .collect();
    names
        .into_iter()
        .map(|name| {
            let payload = reopened.part_xml(&name).unwrap().unwrap().to_vec();
            (name, payload)
        })
        .collect()
}

fn main_xml(package: &Package) -> String {
    String::from_utf8(package.part_xml(MAIN).unwrap().unwrap().to_vec()).unwrap()
}

/// The direct body paragraphs holding inline (`display == false`) or display
/// equations, by transaction position.
fn equation_paragraphs(package: &Package, display: bool) -> Vec<usize> {
    let edit = package.edit().unwrap();
    let snapshot = edit.projected().unwrap();
    (0..snapshot.paragraph_count())
        .filter(|position| {
            let paragraph = snapshot.paragraph(Position::new(*position)).unwrap();
            let equations = if display {
                paragraph.paragraph_level_equations().unwrap()
            } else {
                paragraph.omml_equations().unwrap()
            };
            !equations.is_empty()
        })
        .collect()
}

/// The exact markup of the first inline or display equation of the direct
/// body paragraph at `position`.
fn equation_xml(package: &Package, position: usize, display: bool) -> String {
    let edit = package.edit().unwrap();
    let paragraph = edit
        .projected()
        .unwrap()
        .paragraph(Position::new(position))
        .unwrap();
    let mut equations = if display {
        paragraph.paragraph_level_equations().unwrap()
    } else {
        paragraph.omml_equations().unwrap()
    };
    equations.remove(0)
}

/// Assert that `after` is `before` with exactly the main part changed, and
/// that part changed exactly by replacing `old` with `new` once.
fn assert_only_change(before: &Package, after: &Package, old: &str, new: &str) {
    let before_parts = saved_parts(before);
    let after_parts = saved_parts(after);
    assert_eq!(
        before_parts.keys().collect::<Vec<_>>(),
        after_parts.keys().collect::<Vec<_>>()
    );
    for (name, payload) in &before_parts {
        if name != MAIN {
            assert_eq!(payload, &after_parts[name], "{name} changed");
        }
    }
    let before_main = String::from_utf8(before_parts[MAIN].clone()).unwrap();
    let after_main = String::from_utf8(after_parts[MAIN].clone()).unwrap();
    assert_eq!(before_main.matches(old).count(), 1, "{old}");
    assert_eq!(after_main, before_main.replacen(old, new, 1));
}

#[test]
fn word_equations_are_read_with_their_exact_markup() {
    let package = open(MATH);
    let inline = equation_paragraphs(&package, false);
    let display = equation_paragraphs(&package, true);
    assert_eq!((inline.len(), display.len()), (1, 1));
    let omml = equation_xml(&package, inline[0], false);
    assert!(omml.starts_with("<m:oMath><m:sSup>"), "{omml}");
    assert!(main_xml(&package).contains(omml.as_str()));
    let omml = equation_xml(&package, display[0], true);
    assert!(omml.starts_with("<m:oMath><m:nary>"), "{omml}");
    assert!(main_xml(&package).contains(omml.as_str()));
}

#[test]
fn replacing_an_inline_equation_changes_only_its_markup() {
    let package = open(MATH);
    let position = equation_paragraphs(&package, false)[0];
    let old = equation_xml(&package, position, false);
    let math = OfficeMath::from_xml("<m:oMath><m:r><m:t>y</m:t></m:r></m:oMath>").unwrap();
    let mut edit = package.edit().unwrap();
    assert!(
        edit.replace_equation_at(
            &paragraph(position),
            ParagraphEquation::Inline(Position::new(0)),
            &math
        )
        .unwrap()
    );
    let (changed, patch) = edit.commit().unwrap().into_parts();
    // The root declares `m`, so the new equation repeats no declaration.
    assert_only_change(
        &package,
        &changed,
        &old,
        "<m:oMath><m:r><m:t>y</m:t></m:r></m:oMath>",
    );
    // The patch replays onto the source and inverts exactly.
    let replayed = package.apply(&patch).unwrap().into_package();
    assert_eq!(main_xml(&replayed), main_xml(&changed));
    let restored = changed.apply(&patch.inverse()).unwrap().into_package();
    assert_eq!(main_xml(&restored), main_xml(&package));
    // An identical replacement changes nothing.
    let mut edit = changed.edit().unwrap();
    assert!(
        !edit
            .replace_equation_at(
                &paragraph(position),
                ParagraphEquation::Inline(Position::new(0)),
                &math
            )
            .unwrap()
    );
    assert!(edit.commit().unwrap().into_parts().1.is_noop());
}

#[test]
fn replacing_a_display_equation_keeps_its_math_paragraph() {
    let package = open(MATH);
    let position = equation_paragraphs(&package, true)[0];
    let old = equation_xml(&package, position, true);
    let mut edit = package.edit().unwrap();
    edit.replace_equation_at(
        &paragraph(position),
        ParagraphEquation::Display(Position::new(0)),
        &OfficeMath::text("z"),
    )
    .unwrap();
    let changed = edit.commit().unwrap().into_package();
    assert_only_change(
        &package,
        &changed,
        &old,
        "<m:oMath><m:r><m:t>z</m:t></m:r></m:oMath>",
    );
    assert!(main_xml(&changed).contains("<m:oMathPara><m:oMath><m:r><m:t>z</m:t>"));
}

#[test]
fn removing_equations_takes_only_their_markup() {
    let package = open(MATH);
    let inline = equation_paragraphs(&package, false)[0];
    let old = equation_xml(&package, inline, false);
    let mut edit = package.edit().unwrap();
    edit.remove_equation_at(
        &paragraph(inline),
        ParagraphEquation::Inline(Position::new(0)),
    )
    .unwrap();
    let changed = edit.commit().unwrap().into_package();
    assert_only_change(&package, &changed, &old, "");

    // The only equation of a math paragraph takes the math paragraph.
    let display = equation_paragraphs(&package, true)[0];
    let xml = main_xml(&package);
    let start = xml.find("<m:oMathPara>").unwrap();
    let end = xml.find("</m:oMathPara>").unwrap() + "</m:oMathPara>".len();
    let mut edit = package.edit().unwrap();
    edit.remove_equation_at(
        &paragraph(display),
        ParagraphEquation::Display(Position::new(0)),
    )
    .unwrap();
    let changed = edit.commit().unwrap().into_package();
    assert_only_change(&package, &changed, &xml[start..end], "");

    // A display-equation paragraph is removed as a paragraph.
    let mut edit = package.edit().unwrap();
    edit.remove_element_at(&paragraph(display)).unwrap();
    let changed = edit.commit().unwrap().into_package();
    assert!(equation_paragraphs(&changed, true).is_empty());
    assert_eq!(equation_paragraphs(&changed, false).len(), 1);
}

#[test]
fn equations_are_inserted_in_the_body_cells_headers_and_notes() {
    let mut builder = Builder::new().unwrap();
    builder
        .document()
        .unwrap()
        .add_paragraph_with_text("Area: ");
    let package = builder.build().unwrap();
    let mut edit = package.edit().unwrap();
    edit.insert_table_at(
        &ElementPath::body(),
        ElementAnchor::End,
        &NewTable::new(1, 1).unwrap().with_cell_text(0, 0, "Cell"),
    )
    .unwrap();
    let header = edit
        .add_header_footer(Role::Header, Kind::Primary, &NewStory::new("Top"))
        .unwrap();
    let note = edit
        .add_note_at(
            NoteKind::Footnote,
            &paragraph(0),
            None,
            &NewNote::new("Source."),
        )
        .unwrap();
    let package = edit.commit().unwrap().into_package();

    let mut edit = package.edit().unwrap();
    // Body: one inline equation before the first run, one display equation
    // at the end of the body.
    assert_eq!(
        edit.insert_equation_at(
            &paragraph(0),
            Some(Position::new(0)),
            &OfficeMath::text("a")
        )
        .unwrap(),
        Position::new(0)
    );
    let display = OfficeMathParagraph::from_equation(OfficeMath::text("b"));
    let added = edit
        .insert_display_equation_at(&ElementPath::body(), ElementAnchor::End, &display)
        .unwrap();
    assert_eq!(added, Position::new(1));
    // A table cell: an inline equation, then a display equation.
    let cell = ElementPath::new([
        ElementStep::Table(Position::new(0)),
        ElementStep::Row(Position::new(0)),
        ElementStep::Cell(Position::new(0)),
    ]);
    edit.insert_equation_at(
        &cell.child(ElementStep::Paragraph(Position::new(0))),
        None,
        &OfficeMath::text("c"),
    )
    .unwrap();
    assert_eq!(
        edit.insert_display_equation_at(
            &cell,
            ElementAnchor::End,
            &OfficeMathParagraph::from_equation(OfficeMath::text("d")),
        )
        .unwrap(),
        Position::new(1)
    );
    // A header and a footnote, through the story edit.
    edit.edit_story(StoryTarget::Part(&header), |story| {
        story.insert_equation_at(&paragraph(0), None, &OfficeMath::text("e"))
    })
    .unwrap();
    edit.edit_story(StoryTarget::Note(NoteKind::Footnote, note), |story| {
        story.insert_display_equation_at(
            &ElementPath::body(),
            ElementAnchor::End,
            &OfficeMathParagraph::from_equation(OfficeMath::text("f")),
        )
    })
    .unwrap();
    let changed = edit.commit().unwrap().into_package();
    let reopened = Package::from_bytes(&changed.to_bytes().unwrap()).unwrap();

    let document = reopened.document().unwrap();
    let body = document.paragraphs().unwrap();
    assert_eq!(
        body[0].omml_equations().unwrap(),
        ["<m:oMath><m:r><m:t>a</m:t></m:r></m:oMath>"]
    );
    let xml = main_xml(&reopened);
    assert!(xml.contains("<m:oMath><m:r><m:t>a</m:t></m:r></m:oMath><w:r>"));
    assert!(xml.contains(
        "<w:p><m:oMathPara><m:oMath><m:r><m:t>b</m:t></m:r></m:oMath></m:oMathPara></w:p>"
    ));
    assert!(xml.contains("<m:oMath><m:r><m:t>c</m:t></m:r></m:oMath></w:p>"));
    assert!(xml.contains(
        "<w:p><m:oMathPara><m:oMath><m:r><m:t>d</m:t></m:r></m:oMath></m:oMathPara></w:p></w:tc>"
    ));
    let header_xml =
        String::from_utf8(reopened.part_xml(&header).unwrap().unwrap().to_vec()).unwrap();
    assert!(
        header_xml.contains("<m:r><m:t>e</m:t></m:r></m:oMath></w:p>"),
        "{header_xml}"
    );
    let notes = document.footnotes().unwrap();
    let note = notes.iter().find(|found| found.id() == note).unwrap();
    let equations: Vec<String> = note
        .paragraphs()
        .unwrap()
        .iter()
        .flat_map(|paragraph| paragraph.paragraph_level_equations().unwrap())
        .collect();
    assert_eq!(equations.len(), 1);
    assert!(equations[0].contains("<m:t>f</m:t>"));
}

#[test]
fn strict_documents_receive_strict_math() {
    let package = open(STRICT);
    let mut edit = package.edit().unwrap();
    // `OfficeMath::text` binds the Transitional math namespace; in a Strict
    // document it is written in the Strict namespace the root declares.
    edit.insert_equation_at(&paragraph(0), None, &OfficeMath::text("q"))
        .unwrap();
    let changed = edit.commit().unwrap().into_package();
    let xml = main_xml(&changed);
    assert!(xml.contains("<m:oMath><m:r><m:t>q</m:t></m:r></m:oMath></w:p>"));
    assert!(!xml.contains("schemas.openxmlformats.org/officeDocument/2006/math"));
    let reopened = Package::from_bytes(&changed.to_bytes().unwrap()).unwrap();
    let equations = reopened.document().unwrap().paragraphs().unwrap()[0]
        .omml_equations()
        .unwrap();
    assert_eq!(equations, ["<m:oMath><m:r><m:t>q</m:t></m:r></m:oMath>"]);
}

fn refusal(error: Error) -> Refusal {
    match error {
        Error::Refused { reason, .. } => reason,
        other => panic!("not a refusal: {other:?}"),
    }
}

fn raw(body: &str) -> litchi_docx::document::raw::Tree {
    litchi_docx::document::raw::Tree::from_xml(
        format!(
            r#"<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"><w:body>{body}</w:body></w:document>"#
        )
        .into_bytes(),
        litchi_docx::document::raw::Limits::default(),
    )
    .unwrap()
}

#[test]
fn unsafe_equation_edits_are_refused_with_a_typed_reason() {
    let first = ParagraphEquation::Inline(Position::new(0));
    let math = OfficeMath::text("x");
    // No such equation.
    let mut edit = raw("<w:p><w:r><w:t>x</w:t></w:r></w:p>");
    let error = edit
        .replace_equation_at(&paragraph(0), first, &math)
        .unwrap_err();
    assert_eq!(refusal(error), Refusal::EquationNotFound);
    // An equation holding one end of a bookmark.
    let mut edit = raw(
        r#"<w:p><m:oMath><w:bookmarkStart w:id="1" w:name="b"/><m:r><m:t>y</m:t></m:r></m:oMath><w:bookmarkEnd w:id="1"/></w:p>"#,
    );
    let error = edit.remove_equation_at(&paragraph(0), first).unwrap_err();
    assert_eq!(refusal(error), Refusal::OpenRange);
    // An equation in alternate content, whose branches readers project.
    let mut edit = raw(
        r#"<w:p><mc:AlternateContent><mc:Choice Requires="m"><m:oMath/></mc:Choice><mc:Fallback/></mc:AlternateContent></w:p>"#,
    );
    let error = edit
        .replace_equation_at(&paragraph(0), first, &math)
        .unwrap_err();
    assert_eq!(refusal(error), Refusal::ComplexContent);
    let error = edit
        .insert_equation_at(&paragraph(0), None, &math)
        .unwrap_err();
    assert_eq!(refusal(error), Refusal::ComplexContent);
    // A path that names no paragraph.
    let mut edit = raw("<w:tbl><w:tr><w:tc><w:p/></w:tc></w:tr></w:tbl>");
    let before = edit.projected().xml_bytes().to_vec();
    let error = edit
        .insert_equation_at(
            &ElementPath::new([ElementStep::Table(Position::new(0))]),
            None,
            &math,
        )
        .unwrap_err();
    assert_eq!(refusal(error), Refusal::BlockNotFound);
    // A refusal leaves the edit unchanged.
    assert_eq!(edit.projected().xml_bytes(), before.as_slice());
    // Caller OMML that brings what an equation does not own.
    let mut edit = raw("<w:p><m:oMath><m:r><m:t>y</m:t></m:r></m:oMath></w:p>");
    for (xml, wanted) in [
        (
            r#"<m:oMath xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><m:r><w:drawing><a r:embed="rId999"/></w:drawing></m:r></m:oMath>"#,
            Refusal::EquationContent,
        ),
        (
            r#"<m:oMath><w:bookmarkStart w:id="9" w:name="b"/><m:r><m:t>x</m:t></m:r></m:oMath>"#,
            Refusal::EquationContent,
        ),
        (
            r#"<m:oMath><m:r><w:footnoteReference w:id="99"/></m:r></m:oMath>"#,
            Refusal::EquationContent,
        ),
        (
            "<m:oMath><m:e><w:p><w:r><w:t>x</w:t></w:r></w:p></m:e></m:oMath>",
            Refusal::EquationContent,
        ),
    ] {
        let hostile = OfficeMath::from_xml(xml).unwrap();
        let error = edit
            .replace_equation_at(&paragraph(0), first, &hostile)
            .unwrap_err();
        assert_eq!(refusal(error), wanted, "{xml}");
        let error = edit
            .insert_equation_at(&paragraph(0), None, &hostile)
            .unwrap_err();
        assert_eq!(refusal(error), wanted, "{xml}");
        let error = edit
            .insert_display_equation_at(
                &ElementPath::body(),
                ElementAnchor::End,
                &OfficeMathParagraph::from_equation(hostile),
            )
            .unwrap_err();
        assert_eq!(refusal(error), wanted, "{xml}");
    }
}

#[cfg(feature = "equation")]
mod latex {
    use super::*;
    use litchi_equation::Limits;

    #[test]
    fn latex_replaces_a_display_equation_as_word_writes_it() {
        let package = open(MATH);
        let position = equation_paragraphs(&package, true)[0];
        let (math, losses) =
            OfficeMath::from_latex(r"\frac{a}{b}+x^{2}", &Limits::DEFAULT).unwrap();
        assert!(losses.is_empty());
        let mut edit = package.edit().unwrap();
        edit.replace_equation_at(
            &paragraph(position),
            ParagraphEquation::Display(Position::new(0)),
            &math,
        )
        .unwrap();
        let changed = edit.commit().unwrap().into_package();
        let reopened = Package::from_bytes(&changed.to_bytes().unwrap()).unwrap();
        let read = equation_xml(&reopened, position, true);
        // No namespace declaration is repeated inside the part.
        assert!(read.starts_with("<m:oMath><m:f>"), "{read}");
        assert!(!read.contains("xmlns"));
        // The converter's LaTeX writer spaces after operators.
        let latex = litchi_equation::omml_to_latex(&read, &Limits::DEFAULT).unwrap();
        assert_eq!(latex.replace(' ', ""), r"\frac{a}{b}+x^{2}");
    }

    #[test]
    fn hostile_latex_is_refused_by_the_limits() {
        let deep = format!("{}x{}", "{".repeat(100_000), "}".repeat(100_000));
        assert!(OfficeMath::from_latex(&deep, &Limits::DEFAULT).is_err());
        let tight = Limits::new(16, 1_000, 8).unwrap();
        assert!(matches!(
            OfficeMath::from_latex(r"\frac{a}{b}+\frac{c}{d}", &tight),
            Err(litchi_equation::Error::InputTooLarge { .. })
        ));
    }
}

/// A paragraph of `tags` nested smart tags, each declaring 200 namespaces,
/// around `runs` empty runs and one equation.
fn deep_paragraph(tags: usize, runs: usize) -> litchi_docx::document::raw::Tree {
    let mut body = String::from("<w:p>");
    for tag in 0..tags {
        body.push_str(r#"<w:smartTag w:uri="urn:x" w:element="e""#);
        for declared in 0..200 {
            body.push_str(&format!(
                r#" xmlns:n{tag}x{declared}="urn:{tag}-{declared}""#
            ));
        }
        body.push('>');
    }
    body.push_str(&"<w:r/>".repeat(runs));
    body.push_str("<m:oMath><m:r><m:t>x</m:t></m:r></m:oMath>");
    body.push_str(&"</w:smartTag>".repeat(tags));
    body.push_str("</w:p>");
    raw(&body)
}

#[test]
fn equation_edits_scale_with_the_paragraph_not_its_bindings() {
    let first = ParagraphEquation::Inline(Position::new(0));
    let math = OfficeMath::text("y");
    let timed = |runs: usize| {
        let mut edit = deep_paragraph(64, runs);
        let started = std::time::Instant::now();
        assert!(
            edit.replace_equation_at(&paragraph(0), first, &math)
                .unwrap()
        );
        edit.remove_equation_at(&paragraph(0), first).unwrap();
        started.elapsed()
    };
    let small = timed(12_500);
    let large = timed(50_000);
    // Four times the content takes about four times as long; a scan that
    // copied every binding per element took sixteen.
    assert!(
        large < small * 10 + std::time::Duration::from_millis(500),
        "{small:?} -> {large:?}"
    );
}
