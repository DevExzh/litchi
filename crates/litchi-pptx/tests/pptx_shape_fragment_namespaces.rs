#![allow(
    clippy::unwrap_used,
    clippy::expect_used,
    reason = "test assertions panic on failure by design"
)]

//! The namespace contract of a shape span sliced out of a processed scene.
//!
//! The markup-compatibility codec declares each namespace once, where XML
//! requires it, so a borrowed shape span inherits bindings its own root does
//! not repeat. `Common::xml` keeps returning that exact span and
//! `Common::self_contained_xml` restores the bindings at the slice boundary
//! for a caller that parses the fragment on its own.

use litchi_pptx::shape::text::extract;
use litchi_pptx::shape::{Scene, Shape};
use litchi_pptx::table::Table;
use quick_xml::events::Event;
use quick_xml::name::{Namespace, ResolveResult};
use quick_xml::reader::NsReader;

const PML: &str = "http://schemas.openxmlformats.org/presentationml/2006/main";
const DML: &str = "http://schemas.openxmlformats.org/drawingml/2006/main";
const REL: &str = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const MCE: &str = "http://schemas.openxmlformats.org/markup-compatibility/2006";
const P14: &str = "http://schemas.microsoft.com/office/powerpoint/2010/main";
const TABLE_URI: &str = "http://schemas.openxmlformats.org/drawingml/2006/table";

/// A producer-shaped slide: every namespace is declared on the slide root and
/// the shapes below it only use the prefixes. The aliases are parameters so
/// that a deck outside the conventional `p`/`a`/`r` spelling is covered too.
fn slide(p: &str, a: &str, r: &str) -> String {
    format!(
        r#"<{p}:sld xmlns:{a}="{DML}" xmlns:{r}="{REL}" xmlns:{p}="{PML}" xmlns:mc="{MCE}" xmlns:p14="{P14}" mc:Ignorable="p14">
  <{p}:cSld><{p}:spTree>
    <{p}:nvGrpSpPr/><{p}:grpSpPr/>
    <{p}:sp>
      <{p}:nvSpPr><{p}:cNvPr id="2" name="Title"/><{p}:nvPr><{p}:ph type="title"/></{p}:nvPr></{p}:nvSpPr>
      <{p}:spPr/>
      <{p}:txBody><{a}:bodyPr/><{a}:p><{a}:r><{a}:t>Quarterly</{a}:t></{a}:r><{a}:br/><{a}:r><{a}:t>Review</{a}:t></{a}:r></{a}:p><{a}:p><{a}:r><{a}:t>Second</{a}:t></{a}:r></{a}:p></{p}:txBody>
    </{p}:sp>
    <{p}:graphicFrame>
      <{p}:nvGraphicFramePr><{p}:cNvPr id="5" name="Grid"/></{p}:nvGraphicFramePr>
      <{a}:graphic><{a}:graphicData uri="{TABLE_URI}"><{a}:tbl><{a}:tblPr firstRow="1"/><{a}:tblGrid><{a}:gridCol w="100"/></{a}:tblGrid><{a}:tr h="20"><{a}:tc><{a}:txBody><{a}:p><{a}:r><{a}:t>Cell</{a}:t></{a}:r></{a}:p></{a}:txBody></{a}:tc></{a}:tr></{a}:tbl></{a}:graphicData></{a}:graphic>
    </{p}:graphicFrame>
  </{p}:spTree></{p}:cSld>
</{p}:sld>"#
    )
}

/// How a standalone parse resolves a fragment's own root element.
fn root_namespace(fragment: &[u8]) -> Option<Vec<u8>> {
    let mut reader = NsReader::from_reader(fragment);
    loop {
        let (namespace, event) = reader.read_resolved_event().unwrap();
        match event {
            Event::Start(_) | Event::Empty(_) => {
                return match namespace {
                    ResolveResult::Bound(Namespace(value)) => Some(value.to_vec()),
                    ResolveResult::Unknown(_) | ResolveResult::Unbound => None,
                };
            },
            Event::Eof => return None,
            _ => {},
        }
    }
}

#[test]
fn borrowed_spans_stay_exact_slices_of_the_processed_owner() {
    let xml = slide("p", "a", "r");
    let scene = Scene::read(xml.as_bytes()).unwrap();
    assert!(scene.is_rewritten(), "the MCE input must be rewritten");

    for index in 0..scene.len() {
        let shape = scene.at(index).unwrap();
        let span = shape.span().unwrap();
        let start = usize::try_from(span.start()).unwrap();
        let end = start + usize::try_from(span.len()).unwrap();
        let borrowed = shape.xml().unwrap();
        assert_eq!(borrowed, &scene.xml()[start..end]);
        assert_eq!(
            borrowed.as_ptr(),
            scene.xml()[start..].as_ptr(),
            "shape {index} must stay borrowed from the owner"
        );
    }
}

#[test]
fn self_contained_spans_add_only_the_inherited_declarations() {
    let xml = slide("p", "a", "r");
    let scene = Scene::read(xml.as_bytes()).unwrap();

    let title = scene.at(0).unwrap();
    let borrowed = title.xml().unwrap();
    assert!(
        root_namespace(borrowed).is_none(),
        "the borrowed span inherits its root binding instead of repeating it"
    );

    let restored = title.self_contained_xml().unwrap();
    assert_eq!(
        root_namespace(restored.as_ref()).as_deref(),
        Some(PML.as_bytes())
    );
    assert!(
        restored.len() > borrowed.len(),
        "declarations are added, never removed"
    );
    assert!(restored.starts_with(b"<p:sp "));
    assert!(restored.ends_with(b"</p:sp>"));
    for namespace in [PML, DML, REL, MCE, P14] {
        assert!(
            restored
                .windows(namespace.len())
                .any(|window| window == namespace.as_bytes()),
            "{namespace} must be re-declared on the fragment root"
        );
    }
    assert_eq!(
        title.common().self_contained_xml().unwrap().as_ref(),
        restored.as_ref(),
        "the Shape accessor delegates to the Common one"
    );
}

#[test]
fn conventional_prefixes_read_the_same_from_either_accessor() {
    let xml = slide("p", "a", "r");
    let scene = Scene::read(xml.as_bytes()).unwrap();

    let title = scene.at(0).unwrap();
    let borrowed = title.xml().unwrap();
    let restored = title.self_contained_xml().unwrap();
    assert_eq!(
        extract(borrowed, Some('\n')).unwrap(),
        "Quarterly\nReview\nSecond"
    );
    assert_eq!(
        extract(restored.as_ref(), Some('\n')).unwrap(),
        extract(borrowed, Some('\n')).unwrap()
    );

    let frame = scene.at(1).unwrap();
    assert!(matches!(frame, Shape::Table(_)));
    let borrowed = frame.xml().unwrap();
    let restored = frame.self_contained_xml().unwrap();
    let from_borrowed = Table::from_graphic_frame(borrowed).unwrap();
    let from_restored = Table::from_graphic_frame(restored.as_ref()).unwrap();
    assert_eq!(from_borrowed.row_count().unwrap(), 1);
    assert_eq!(from_borrowed.column_count().unwrap(), 1);
    assert_eq!(
        from_borrowed.cell(0, 0).unwrap().unwrap().text().unwrap(),
        "Cell"
    );
    assert_eq!(
        from_restored.row_count().unwrap(),
        from_borrowed.row_count().unwrap()
    );
    assert_eq!(
        from_restored.cell(0, 0).unwrap().unwrap().text().unwrap(),
        from_borrowed.cell(0, 0).unwrap().unwrap().text().unwrap()
    );
    assert_eq!(
        from_restored.properties().unwrap(),
        from_borrowed.properties().unwrap()
    );
}

#[test]
fn unconventional_prefixes_need_the_self_contained_accessor() {
    let xml = slide("q", "d", "rel");
    let scene = Scene::read(xml.as_bytes()).unwrap();

    // The `a` and `p` fallbacks in the fragment scanners only cover the
    // conventional aliases, so a borrowed span from this deck reads empty.
    let title = scene.at(0).unwrap();
    assert_eq!(extract(title.xml().unwrap(), Some('\n')).unwrap(), "");
    let restored = title.self_contained_xml().unwrap();
    assert_eq!(
        extract(restored.as_ref(), Some('\n')).unwrap(),
        "Quarterly\nReview\nSecond"
    );

    let frame = scene.at(1).unwrap();
    assert!(Table::from_graphic_frame(frame.xml().unwrap()).is_err());
    let restored = frame.self_contained_xml().unwrap();
    let table = Table::from_graphic_frame(restored.as_ref()).unwrap();
    assert_eq!(table.row_count().unwrap(), 1);
    assert_eq!(table.cell(0, 0).unwrap().unwrap().text().unwrap(), "Cell");
}
