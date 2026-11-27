//! A source-backed write that removes a part (a header, a comment, an image,
//! a slide, a sheet) must leave no trace of it in the saved file: the
//! removed member and its relationships member are absent, every remaining
//! member is a part, a part's relationships or `[Content_Types].xml`, and the
//! saved package equals the eager owner's for the same edits. Each removal is
//! saved directly and after a later write that promotes the handle to an
//! owning package.
#![cfg(all(feature = "docx", feature = "pptx", feature = "xlsx"))]

use std::collections::{BTreeMap, BTreeSet};
use std::io::{Cursor, Read as _};
use std::path::PathBuf;

use litchi_core::FileFormat;
use litchi_opc::{OpcPackage, SourceLimits};
use litchi_ops::{File, NewNode, OpenOptions, SourceMode, props};

fn fixture(relative: &str) -> Vec<u8> {
    let path = PathBuf::from(env!("CARGO_MANIFEST_DIR"))
        .join("../../test-data/ooxml")
        .join(relative);
    std::fs::read(&path).unwrap_or_else(|error| panic!("{}: {error}", path.display()))
}

fn members(bytes: &[u8]) -> BTreeMap<String, Vec<u8>> {
    let mut archive = zip::ZipArchive::new(Cursor::new(bytes)).unwrap();
    (0..archive.len())
        .map(|index| {
            let mut member = archive.by_index(index).unwrap();
            let mut data = Vec::new();
            member.read_to_end(&mut data).unwrap();
            (member.name().unwrap().into_owned(), data)
        })
        .collect()
}

/// The members of a package that no relationship chain from the package
/// reaches: neither a reachable part, nor the relationships member of the
/// package or of a reachable part, nor `[Content_Types].xml`.
fn unreferenced(bytes: &[u8]) -> Vec<String> {
    let package = OpcPackage::from_bytes(bytes, SourceLimits::default()).unwrap();
    let mut known = BTreeSet::from(["[Content_Types].xml".to_owned(), "_rels/.rels".to_owned()]);
    let mut pending: Vec<_> = package
        .rels()
        .iter()
        .filter(|relationship| !relationship.is_external())
        .filter_map(|relationship| relationship.target_partname().ok())
        .collect();
    while let Some(partname) = pending.pop() {
        if !known.insert(partname.membername().to_owned()) {
            continue;
        }
        known.insert(partname.rels_uri().unwrap().membername().to_owned());
        if let Ok(part) = package.get_part(&partname) {
            pending.extend(
                part.rels()
                    .iter()
                    .filter(|relationship| !relationship.is_external())
                    .filter_map(|relationship| relationship.target_partname().ok()),
            );
        }
    }
    members(bytes)
        .into_keys()
        .filter(|name| !known.contains(name))
        .collect()
}

/// The path of the first node `selector` matches.
fn first(file: &File, selector: &str) -> String {
    file.query(selector)
        .unwrap()
        .next()
        .unwrap_or_else(|| panic!("no {selector}"))
        .unwrap()
        .view()
        .path()
        .to_string()
}

/// A member's bytes, with a relationships member reduced to its sorted
/// relationship elements: the two owners may order relationships
/// differently.
fn normalized(name: &str, data: &[u8]) -> Vec<String> {
    let text = String::from_utf8_lossy(data);
    if !name.ends_with(".rels") {
        return vec![text.into_owned()];
    }
    let mut relationships: Vec<String> = text
        .split("<Relationship ")
        .skip(1)
        .map(|element| element.split("/>").next().unwrap_or(element).to_owned())
        .collect();
    relationships.sort();
    relationships
}

type Edit = fn(File) -> File;

/// `text` without the `element` elements (`<Relationship `, `<Override `)
/// that `kept` does not contain verbatim.
fn without_dropped(text: &str, element: &str, kept: &str) -> String {
    let mut output = String::new();
    let mut rest = text;
    while let Some(start) = rest.find(element) {
        let end = start + rest[start..].find("/>").expect("an empty element") + 2;
        output.push_str(&rest[..start]);
        if kept.contains(&rest[start..end]) {
            output.push_str(&rest[start..end]);
        }
        rest = &rest[end..];
    }
    output.push_str(rest);
    output
}

/// The source-backed removal is a minimal edit: every kept member is byte
/// for byte the input's, except the edited slide, relationships members that
/// lose whole `Relationship` elements, and `[Content_Types].xml`, which
/// loses whole `Override` elements.
fn assert_minimal(
    context: &str,
    input: &BTreeMap<String, Vec<u8>>,
    saved: &BTreeMap<String, Vec<u8>>,
) {
    for (name, data) in saved {
        let original = input
            .get(name)
            .unwrap_or_else(|| panic!("{context}: {name} is new"));
        if original == data {
            continue;
        }
        let (original, data) = (
            String::from_utf8_lossy(original),
            String::from_utf8_lossy(data),
        );
        let element = if name.ends_with(".rels") {
            "<Relationship "
        } else if name == "[Content_Types].xml" {
            "<Override "
        } else {
            // The edited slide, or the presentation a slide left.
            assert!(
                name.starts_with("ppt/slides/slide") || name == "ppt/presentation.xml",
                "{context}: {name} changed"
            );
            continue;
        };
        assert_eq!(
            without_dropped(&original, element, &data),
            data,
            "{context}: {name} changed beyond dropped elements"
        );
    }
}

/// Remove the first `selector` node, optionally run `then`, and compare the
/// source-backed save with the eager one. Returns whether the removal dropped
/// an archive member. With `minimal`, the source-backed save must keep
/// every other byte of the input ([`assert_minimal`]) and only its member
/// names are compared with the eager save.
fn check(input: &[u8], selector: &str, then: Option<Edit>, minimal: bool) -> bool {
    let source = litchi_ops::open_bytes(input.to_vec()).unwrap();
    let eager = litchi_ops::open_bytes_with(
        input.to_vec(),
        &OpenOptions::new().source_mode(SourceMode::Eager),
    )
    .unwrap();
    let path = first(&source, selector);
    assert_eq!(path, first(&eager, selector));
    let run = |file: File| {
        file.remove(path.as_str()).map(|removed| {
            then.map_or(removed.clone(), |then| then(removed))
                .to_bytes()
                .unwrap()
        })
    };
    let saved = run(source).unwrap_or_else(|error| panic!("remove {path}: {error}"));
    // The eager owner refuses some removals the source-backed one performs
    // (an OLE frame under markup compatibility); those compare with the
    // input only.
    let oracle = run(eager).ok();
    let before: BTreeSet<String> = members(input).into_keys().collect();
    let after: BTreeSet<String> = members(&saved).into_keys().collect();
    let context = format!("remove {path} then {}", then.is_some());
    let input_unreferenced = unreferenced(input);
    for name in unreferenced(&saved) {
        assert!(
            input_unreferenced.contains(&name),
            "{context}: {name} is left without a part"
        );
    }
    let saved = members(&saved);
    let oracle = match oracle {
        Some(oracle) => members(&oracle),
        None => {
            assert!(minimal, "{context}: the eager owner refused the removal");
            if then.is_none() {
                assert_minimal(&context, &members(input), &saved);
            }
            return before.difference(&after).next().is_some();
        },
    };
    assert_eq!(
        saved.keys().collect::<Vec<_>>(),
        oracle.keys().collect::<Vec<_>>(),
        "{context}: the saved package holds other members than the eager one"
    );
    if minimal {
        if then.is_none() {
            assert_minimal(&context, &members(input), &saved);
        }
        return before.difference(&after).next().is_some();
    }
    let differing: Vec<&String> = saved
        .keys()
        .chain(oracle.keys())
        .filter(|name| {
            saved.get(*name).map(|data| normalized(name, data))
                != oracle.get(*name).map(|data| normalized(name, data))
        })
        .collect();
    assert!(
        differing.is_empty(),
        "{context}: the saved package differs from the eager one in {differing:?}; source has {:?}, eager has {:?}",
        saved.keys().collect::<Vec<_>>(),
        oracle.keys().collect::<Vec<_>>()
    );
    before.difference(&after).next().is_some()
}

fn docx_promote(file: File) -> File {
    let paragraph = first(&file, "/body/p");
    file.add(
        paragraph.as_str(),
        NewNode::new("footnote").prop("text", "PROMOTED"),
    )
    .unwrap()
    .file
}

fn pptx_promote(file: File) -> File {
    file.add(
        "/slide[1]",
        NewNode::new("comment")
            .prop("author", "Ann")
            .prop("text", "PROMOTED"),
    )
    .unwrap_or_else(|error| panic!("comment add: {error}"))
    .file
}

fn xlsx_promote(file: File) -> File {
    file.add(
        "/sheet[1]",
        NewNode::new("comment")
            .prop("ref", "Z99")
            .prop("text", "PROMOTED")
            .prop("author", "Ann"),
    )
    .unwrap()
    .file
}

fn docx_body_edit(file: File) -> File {
    file.add("/body", NewNode::new("paragraph").prop("text", "BODY"))
        .unwrap()
        .file
}

const PNG: &str = "data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAQAAAACCAYAAAB/qH1jAAAAEklEQVR4nGP4z8DwHxkzoAsAAA8hD/EEN8afAAAAAElFTkSuQmCC";

/// A document with one body paragraph holding one picture.
fn docx_picture() -> Vec<u8> {
    litchi_ops::create(FileFormat::Docx)
        .unwrap()
        .add("/body", NewNode::new("paragraph").prop("text", "Body"))
        .unwrap()
        .file
        .add("/body/p[1]", NewNode::new("picture").prop("src", PNG))
        .unwrap()
        .file
        .to_bytes()
        .unwrap()
}

/// A document with one body paragraph holding one chart.
fn docx_chart() -> Vec<u8> {
    litchi_ops::create(FileFormat::Docx)
        .unwrap()
        .add("/body", NewNode::new("paragraph").prop("text", "Body"))
        .unwrap()
        .file
        .add(
            "/body/p[1]",
            NewNode::new("chart")
                .prop("chartType", "bar")
                .prop("categories", "Q1,Q2")
                .prop("data", "East:1,2"),
        )
        .unwrap()
        .file
        .to_bytes()
        .unwrap()
}

#[test]
fn docx_removed_stories_comments_and_images_leave_no_member() {
    for (bytes, selector) in [
        (fixture("docx/headerFooter.docx"), "header"),
        (fixture("docx/headerFooter.docx"), "footer"),
        (fixture("docx/headerPic.docx"), "header"),
        (fixture("docx/testComment.docx"), "comment"),
        (docx_picture(), "picture"),
        (docx_chart(), "chart"),
    ] {
        let dropped = check(&bytes, selector, None, false);
        assert_eq!(check(&bytes, selector, Some(docx_promote), false), dropped);
        // Removing a header, a footer or the only picture of an image
        // removes its part; one comment of several does not.
        assert!(dropped || selector == "comment", "{selector}");
        // A body edit first, so the removal lands on an edited source.
        let edited = docx_body_edit(litchi_ops::open_bytes(bytes).unwrap())
            .to_bytes()
            .unwrap();
        check(&edited, selector, Some(docx_promote), false);
    }
}

fn deck() -> Vec<u8> {
    litchi_ops::create(FileFormat::Pptx)
        .unwrap()
        .add("/", NewNode::new("slide").prop("layout", "Blank"))
        .unwrap()
        .file
        .add("/", NewNode::new("slide").prop("layout", "Blank"))
        .unwrap()
        .file
        .add(
            "/slide[2]",
            NewNode::new("notes").prop("text", "Second notes"),
        )
        .unwrap()
        .file
        .to_bytes()
        .unwrap()
}

#[test]
fn pptx_removed_slides_and_pictures_leave_no_member() {
    assert!(check(&deck(), "/slide[2]", None, true));
    assert!(check(&deck(), "/slide[2]", Some(pptx_promote), true));
    let pictured = litchi_ops::open_bytes(deck())
        .unwrap()
        .add("/slide[1]", NewNode::new("picture").prop("src", PNG))
        .unwrap()
        .file
        .to_bytes()
        .unwrap();
    assert!(check(&pictured, "picture", None, true));
    assert!(check(&pictured, "picture", Some(pptx_promote), true));
    let charts = fixture("pptx/line-chart.pptx");
    assert!(check(&charts, "chart", None, true));
    assert!(check(&charts, "chart", Some(pptx_promote), true));
}

fn data(relative: &str) -> Vec<u8> {
    let path = PathBuf::from(env!("CARGO_MANIFEST_DIR"))
        .join("../../test-data")
        .join(relative);
    std::fs::read(&path).unwrap_or_else(|error| panic!("{}: {error}", path.display()))
}

/// LibreOffice relates each slide's SmartArt drawing without writing the
/// drawing Part (its content-types override stays). Removing the SmartArt
/// frames, in either order, one at a time or in one edit, keeps that dangling
/// edge, succeeds, saves, undoes and redoes exactly.
#[test]
fn pptx_smartart_removals_keep_a_dangling_drawing_edge_consistent() {
    let input = data("ooxml/pptx/smartart-autoTxRot.pptx");
    for order in [
        ["/slide[1]/smartart[1]", "/slide[2]/smartart[1]"],
        ["/slide[2]/smartart[1]", "/slide[1]/smartart[1]"],
    ] {
        let file = litchi_ops::open_bytes(input.clone()).unwrap();
        let separate = file
            .remove(order[0])
            .and_then(|file| file.remove(order[1]))
            .unwrap_or_else(|error| panic!("{order:?}: {error}"));
        let commit = file
            .edit(|tx| {
                tx.remove(order[0])?;
                tx.remove(order[1])
            })
            .unwrap_or_else(|error| panic!("{order:?}: {error}"));
        let (together, patch, ()) = commit.into_parts();
        let saved = together.to_bytes().unwrap();
        assert_eq!(members(&separate.to_bytes().unwrap()), members(&saved));
        let reopened = litchi_ops::open_bytes(saved.clone()).unwrap();
        // The third slide keeps its SmartArt.
        assert_eq!(reopened.query("smartart").unwrap().count(), 1);
        for name in unreferenced(&saved) {
            assert!(unreferenced(&input).contains(&name), "{order:?}: {name}");
        }
        let undone = patch.inverse().apply(&together).unwrap();
        assert_eq!(undone.to_bytes().unwrap(), input, "{order:?}: undo");
        let redone = patch.apply(&undone).unwrap();
        assert_eq!(members(&redone.to_bytes().unwrap()), members(&saved));
    }
}

/// Remove the first `selector` node from a source-backed open, run
/// [`check`], and return the members the removal dropped, sorted.
fn removed_members(input: &[u8], selector: &str) -> Vec<String> {
    check(input, selector, None, true);
    let file = litchi_ops::open_bytes(input.to_vec()).unwrap();
    let path = first(&file, selector);
    let saved = file.remove(path.as_str()).unwrap().to_bytes().unwrap();
    let after = members(&saved);
    members(input)
        .into_keys()
        .filter(|name| !after.contains_key(name))
        .collect()
}

#[test]
fn pptx_producer_shape_removals_are_minimal_and_leave_no_member() {
    // OLE frames under markup compatibility: both branches name the
    // embedding, which leaves once neither remains.
    // The legacy VML drawing that held only the removed object's shape, and
    // the fallback preview image it and the fallback picture showed, leave
    // too.
    for (input, gone) in [
        (
            "libreoffice-core/sd/qa/unit/data/pptx/ole.pptx",
            &[
                "ppt/drawings/_rels/vmlDrawing1.vml.rels",
                "ppt/drawings/vmlDrawing1.vml",
                "ppt/embeddings/Microsoft_Word_Document.docx",
                "ppt/media/image1.emf",
            ][..],
        ),
        (
            "libreoffice-core/sd/qa/unit/data/pptx/tdf103477.pptx",
            &[
                "ppt/drawings/_rels/vmlDrawing3.vml.rels",
                "ppt/drawings/vmlDrawing3.vml",
                "ppt/embeddings/oleObject3.bin",
                "ppt/media/image3.emf",
                // The frame's own customer-data tag.
                "ppt/tags/tag4.xml",
            ][..],
        ),
    ] {
        assert_eq!(removed_members(&data(input), "ole"), gone, "{input}");
    }
    // A SmartArt's cached drawing, which only its data part names, leaves
    // with the frame, and with it the pictures it showed.
    for (input, images) in [
        (
            "ooxml/pptx/smartart-picture-strip.pptx",
            &[
                "ppt/media/image1.png",
                "ppt/media/image2.png",
                "ppt/media/image3.png",
            ][..],
        ),
        (
            "ooxml/pptx/smartart-simple.pptx",
            &["ppt/media/image1.jpeg"][..],
        ),
    ] {
        let gone = removed_members(&data(input), "smartart");
        assert!(
            gone.iter()
                .any(|name| name.starts_with("ppt/diagrams/drawing")),
            "{input}: {gone:?}"
        );
        for image in images {
            assert!(
                gone.iter().any(|name| name == image),
                "{input}: {image} stays: {gone:?}"
            );
        }
    }
    // This picture's image is shared, so it stays; the removal is still
    // minimal.
    check(
        &data("libreoffice-core/sd/qa/unit/data/pptx/tdf169496_hidden_graphic.pptx"),
        "picture",
        None,
        true,
    );
    // Groups without related parts change only their slide.
    for input in [
        "ooxml/pptx/shapes.pptx",
        "libreoffice-core/sd/qa/unit/data/pptx/slide-sections.pptx",
        "unified-ops/acceptance-fixtures/modern-comments.pptx",
    ] {
        let bytes = data(input);
        let file = litchi_ops::open_bytes(bytes.clone()).unwrap();
        if file.query("group").unwrap().next().is_some() {
            check(&bytes, "group", None, true);
        }
    }
}

fn workbook() -> Vec<u8> {
    litchi_ops::create(FileFormat::Xlsx)
        .unwrap()
        .set("/Sheet1/A1", props! { "value" => "first" })
        .unwrap()
        .add("/", NewNode::new("sheet").prop("name", "Second"))
        .unwrap()
        .file
        .set("/Second/A1", props! { "value" => "second" })
        .unwrap()
        .to_bytes()
        .unwrap()
}

#[test]
fn xlsx_removed_sheets_and_charts_leave_no_member() {
    assert!(check(&workbook(), "/sheet[2]", None, false));
    assert!(check(&workbook(), "/sheet[2]", Some(xlsx_promote), false));
    let charts = fixture("xlsx/SimpleScatterChart.xlsx");
    assert!(check(&charts, "chart", None, false));
    assert!(check(&charts, "chart", Some(xlsx_promote), false));
}
