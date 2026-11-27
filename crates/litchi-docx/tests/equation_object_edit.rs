//! OLE equation objects (MathType and Equation Editor 3.0): read from Word's
//! files, inserted, replaced and removed through the package editor, with
//! every part outside the edited object kept byte for byte.
#![cfg(feature = "equation")]

use std::collections::BTreeMap;

use litchi_core::Position;
use litchi_docx::Package;
use litchi_docx::document::{
    ElementPath, ElementStep, EquationContent, EquationSource, Error, Refusal,
};
use litchi_equation::{Limits, MtefTarget};
use soapberry_zip::office::ArchiveReader;

/// Word 2007: one `Equation.3` object with an EMF preview.
const EQUATION_EDITOR: &str = concat!(
    env!("CARGO_MANIFEST_DIR"),
    "/../../test-data/libreoffice-core/sw/qa/extras/ooxmlexport/data/mathtype.docx"
);
/// Word 2010: one `Equation.DSMT4` object with a WMF preview, its run
/// lowered by `w:position`.
const MATHTYPE: &str = concat!(
    env!("CARGO_MANIFEST_DIR"),
    "/../../test-data/libreoffice-core/sw/qa/core/data/ooxml/pass/fdo78906.docx"
);

fn members(archive: &[u8]) -> BTreeMap<String, Vec<u8>> {
    let reader =
        ArchiveReader::new(archive, soapberry_zip::office::ArchiveLimits::default()).unwrap();
    let names: Vec<String> = reader.file_names().map(str::to_owned).collect();
    names
        .into_iter()
        .map(|name| {
            let bytes = reader.read(&name).unwrap();
            (name, bytes)
        })
        .collect()
}

fn text(bytes: &[u8]) -> String {
    String::from_utf8(bytes.to_vec()).unwrap()
}

fn paragraph(index: usize) -> ElementPath {
    ElementPath::new([ElementStep::Paragraph(Position::new(index))])
}

fn open(path: &str) -> Package {
    Package::from_bytes(&std::fs::read(path).unwrap()).unwrap()
}

fn save(package: Package) -> Vec<u8> {
    package.to_bytes().unwrap()
}

/// The `Equation Native` stream of the paragraph's object at `index`.
fn native(package: &Package, paragraph: usize, index: usize) -> Vec<u8> {
    let objects = package.document().unwrap().paragraphs().unwrap()[paragraph]
        .equation_objects()
        .unwrap();
    let id = objects[index].object_relationship_id().unwrap();
    let payload = package.related_part_payload(None, id).unwrap().unwrap();
    litchi_equation::equation_object_native(&payload).unwrap()
}

fn latex(source: &str) -> EquationContent {
    EquationContent::new(EquationSource::Latex(source), None, 12.0, &Limits::DEFAULT).unwrap()
}

#[test]
fn word_objects_are_read_with_their_kind_parts_and_size() {
    let package = open(EQUATION_EDITOR);
    let objects = package.document().unwrap().paragraphs().unwrap()[0]
        .equation_objects()
        .unwrap();
    assert_eq!(objects.len(), 1);
    assert_eq!(objects[0].prog_id(), "Equation.3");
    assert_eq!(objects[0].object_relationship_id(), Some("rId5"));
    assert_eq!(objects[0].preview_relationship_id(), Some("rId4"));
    assert_eq!(objects[0].shape_id(), Some("_x0000_i1025"));
    assert_eq!(objects[0].size_pt(), Some((30.85, 28.5)));
    let read = litchi_equation::read_mtef(&native(&package, 0, 0), &Limits::DEFAULT).unwrap();
    assert_eq!(read.info().unwrap().version, 3);

    let package = open(MATHTYPE);
    let objects = package.document().unwrap().paragraphs().unwrap()[0]
        .equation_objects()
        .unwrap();
    assert_eq!(objects.len(), 1);
    assert_eq!(objects[0].prog_id(), "Equation.DSMT4");
    let read = litchi_equation::read_mtef(&native(&package, 0, 0), &Limits::DEFAULT).unwrap();
    assert_eq!(read.info().unwrap().version, 5);
    assert!(!read.latex().is_empty());
}

#[test]
fn a_new_object_is_word_markup_with_its_two_parts() {
    let mut builder = litchi_docx::Builder::new().unwrap();
    builder
        .document()
        .unwrap()
        .add_paragraph_with_text("Area: ");
    let package = builder.build().unwrap();
    let before = members(&save(package.clone()));
    let mut edit = package.edit().unwrap();
    let content = latex(r"\frac{a}{b}");
    assert_eq!(
        edit.insert_equation_object_at(None, &paragraph(0), None, &content)
            .unwrap(),
        Position::new(0)
    );
    let second = latex("x^2");
    assert_eq!(
        edit.insert_equation_object_at(None, &paragraph(0), Some(Position::new(0)), &second)
            .unwrap(),
        Position::new(0)
    );
    let bytes = save(edit.commit().unwrap().into_package());
    let after = members(&bytes);
    let document = text(&after["word/document.xml"]);
    // The part's picture shape type is written once, in the first object of
    // the part; both shapes use it.
    assert_eq!(document.matches("<v:shapetype ").count(), 1, "{document}");
    assert_eq!(document.matches(r##"type="#_x0000_t75""##).count(), 2);
    assert!(
        document.contains(r#"<v:shape id="_x0000_i1025""#),
        "{document}"
    );
    assert!(
        document.contains(r#"<v:shape id="_x0000_i1026""#),
        "{document}"
    );
    assert!(document.contains(r#"ProgID="Equation.DSMT4""#));
    assert!(document.contains("<w:position w:val=\"-"), "{document}");
    assert!(document.contains(r#"DrawAspect="Content" ObjectID="_1000000001""#));
    assert!(document.contains(r#"ObjectID="_1000000002""#));
    for name in [
        "word/embeddings/oleObject1.bin",
        "word/embeddings/oleObject2.bin",
        "word/media/image1.png",
        "word/media/image2.png",
    ] {
        assert!(after.contains_key(name), "{name}");
    }
    let types = text(&after["[Content_Types].xml"]);
    assert!(types.contains("application/vnd.openxmlformats-officedocument.oleObject"));
    assert!(types.contains("image/png"));
    let relationships = text(&after["word/_rels/document.xml.rels"]);
    assert_eq!(relationships.matches("relationships/oleObject").count(), 2);
    for (name, contents) in &before {
        if !matches!(
            name.as_str(),
            "word/document.xml" | "word/_rels/document.xml.rels" | "[Content_Types].xml"
        ) {
            assert_eq!(contents, &after[name], "{name} changed");
        }
    }
    // Read back through the reader and the MTEF reader.
    let package = Package::from_bytes(&bytes).unwrap();
    let objects = package.document().unwrap().paragraphs().unwrap()[0]
        .equation_objects()
        .unwrap();
    assert_eq!(objects.len(), 2);
    let first = litchi_equation::read_mtef(&native(&package, 0, 0), &Limits::DEFAULT).unwrap();
    assert_eq!(first.latex(), "x^{2}");
    let second = litchi_equation::read_mtef(&native(&package, 0, 1), &Limits::DEFAULT).unwrap();
    assert_eq!(second.latex(), r"\frac{a}{b}");
    assert_eq!(second.info().unwrap().version, 5);
}

#[test]
fn a_new_object_references_the_parts_existing_shape_type() {
    let package = open(EQUATION_EDITOR);
    let mut edit = package.edit().unwrap();
    edit.insert_equation_object_at(None, &paragraph(0), None, &latex("y"))
        .unwrap();
    let bytes = save(edit.commit().unwrap().into_package());
    let document = text(&members(&bytes)["word/document.xml"]);
    assert_eq!(document.matches("<v:shapetype ").count(), 1);
    // The fixture declares v, o and r on its root, so the object declares
    // nothing of its own.
    assert!(document.contains(r#"<w:object w:dxaOrig=""#));
    assert!(
        document.contains(r#""><v:shape id="_x0000_i1026""#),
        "{document}"
    );
    assert!(
        document.contains(r#"<v:shape id="_x0000_i1026""#),
        "{document}"
    );
    assert!(document.contains(r#"ObjectID="_1490538857""#), "{document}");
}

#[test]
fn a_replaced_object_keeps_its_kind_and_every_other_byte() {
    for (path, version) in [(MATHTYPE, 5), (EQUATION_EDITOR, 3)] {
        let package = open(path);
        let before = members(&save(package.clone()));
        let mut edit = package.edit().unwrap();
        let kind = edit
            .equation_object_kind(None, &paragraph(0), Position::new(0), &Limits::DEFAULT)
            .unwrap();
        assert!(matches!(kind, MtefTarget::SameAs(ref info) if info.version == version));
        let content = EquationContent::new(
            EquationSource::Latex("a+b"),
            Some(&kind),
            12.0,
            &Limits::DEFAULT,
        )
        .unwrap();
        assert!(
            edit.set_equation_object_at(None, &paragraph(0), Position::new(0), &content)
                .unwrap()
        );
        let bytes = save(edit.commit().unwrap().into_package());
        let after = members(&bytes);
        let package = Package::from_bytes(&bytes).unwrap();
        let read = litchi_equation::read_mtef(&native(&package, 0, 0), &Limits::DEFAULT).unwrap();
        assert_eq!(read.info().unwrap().version, version);
        assert_eq!(read.latex().replace(' ', ""), "a+b");
        // The EMF or WMF preview was replaced by a PNG one; the old part and
        // its relationship are gone.
        let old_preview = before
            .keys()
            .find(|name| name.starts_with("word/media/"))
            .unwrap();
        assert!(!after.contains_key(old_preview), "{old_preview} kept");
        assert!(after.keys().any(|name| name.ends_with(".png")));
        let relationships = text(&after["word/_rels/document.xml.rels"]);
        assert!(!relationships.contains(old_preview.trim_start_matches("word/")));
        // The compound file kept its name.
        assert!(after.contains_key("word/embeddings/oleObject1.bin"));
        assert_ne!(
            after["word/embeddings/oleObject1.bin"],
            before["word/embeddings/oleObject1.bin"]
        );
        for (name, contents) in &before {
            if !(name == "word/document.xml"
                || name == "word/_rels/document.xml.rels"
                || name == "[Content_Types].xml"
                || name.starts_with("word/embeddings/")
                || name.starts_with("word/media/"))
            {
                assert_eq!(contents, &after[name], "{name} changed");
            }
        }
        // Only the object's markup changed.
        let old = text(&before["word/document.xml"]);
        let new = text(&after["word/document.xml"]);
        let object = |xml: &str| {
            let start = xml.find("<w:object ").unwrap();
            let end = xml.find("</w:object>").unwrap() + "</w:object>".len();
            (xml[..start].to_owned(), xml[end..].to_owned())
        };
        let (old_head, old_tail) = object(&old);
        let (new_head, new_tail) = object(&new);
        let strip_position = |head: &str| {
            head.rsplit_once("<w:rPr>")
                .map_or_else(|| head.to_owned(), |(before, _)| before.to_owned())
        };
        assert_eq!(strip_position(&old_head), strip_position(&new_head));
        assert_eq!(old_tail, new_tail);
    }
}

#[test]
fn a_png_preview_only_one_object_shows_is_replaced_in_place() {
    let mut builder = litchi_docx::Builder::new().unwrap();
    builder.document().unwrap().add_paragraph_with_text("x");
    let mut edit = builder.build().unwrap().edit().unwrap();
    edit.insert_equation_object_at(None, &paragraph(0), None, &latex("a"))
        .unwrap();
    let package = edit.commit().unwrap().into_package();
    let before = members(&save(package.clone()));
    let mut edit = package.edit().unwrap();
    let kind = edit
        .equation_object_kind(None, &paragraph(0), Position::new(0), &Limits::DEFAULT)
        .unwrap();
    let content = EquationContent::new(
        EquationSource::Latex("a^{2}"),
        Some(&kind),
        12.0,
        &Limits::DEFAULT,
    )
    .unwrap();
    edit.set_equation_object_at(None, &paragraph(0), Position::new(0), &content)
        .unwrap();
    let after = members(&save(edit.commit().unwrap().into_package()));
    assert_eq!(
        before.keys().collect::<Vec<_>>(),
        after.keys().collect::<Vec<_>>()
    );
    assert_ne!(
        after["word/media/image1.png"],
        before["word/media/image1.png"]
    );
    assert_eq!(
        after["word/_rels/document.xml.rels"],
        before["word/_rels/document.xml.rels"]
    );
}

#[test]
fn a_removed_object_takes_its_parts_and_relationships() {
    let package = open(MATHTYPE);
    let before = members(&save(package.clone()));
    let mut edit = package.edit().unwrap();
    edit.remove_equation_object_at(None, &paragraph(0), Position::new(0))
        .unwrap();
    let bytes = save(edit.commit().unwrap().into_package());
    let after = members(&bytes);
    assert!(!after.contains_key("word/embeddings/oleObject1.bin"));
    assert!(!after.contains_key("word/media/image1.wmf"));
    let relationships = text(&after["word/_rels/document.xml.rels"]);
    assert!(!relationships.contains("oleObject"));
    assert!(!relationships.contains("image1.wmf"));
    let document = text(&after["word/document.xml"]);
    assert!(!document.contains("w:object"));
    // The run held only the object: it went with it, its neighbours stay.
    let old = text(&before["word/document.xml"]);
    let start = old.find(r#"<w:r w:rsidR="00B62753""#).unwrap();
    let end = old[start..].find("</w:r>").unwrap() + start + "</w:r>".len();
    assert_eq!(document, format!("{}{}", &old[..start], &old[end..]));
    let package = Package::from_bytes(&bytes).unwrap();
    assert!(
        package.document().unwrap().paragraphs().unwrap()[0]
            .equation_objects()
            .unwrap()
            .is_empty()
    );
}

#[test]
fn a_removed_object_leaves_the_shape_type_others_still_use() {
    let package = open(EQUATION_EDITOR);
    let mut edit = package.edit().unwrap();
    edit.insert_equation_object_at(None, &paragraph(0), None, &latex("y"))
        .unwrap();
    let package = edit.commit().unwrap().into_package();
    let mut edit = package.edit().unwrap();
    edit.remove_equation_object_at(None, &paragraph(0), Position::new(0))
        .unwrap();
    let bytes = save(edit.commit().unwrap().into_package());
    let document = text(&members(&bytes)["word/document.xml"]);
    assert_eq!(document.matches("<v:shapetype ").count(), 1, "{document}");
    assert!(
        document.contains(r#"<w:r><w:pict><v:shapetype id="_x0000_t75""#),
        "{document}"
    );
    let package = Package::from_bytes(&bytes).unwrap();
    let objects = package.document().unwrap().paragraphs().unwrap()[0]
        .equation_objects()
        .unwrap();
    assert_eq!(objects.len(), 1);
    assert_eq!(objects[0].shape_id(), Some("_x0000_i1026"));
}

#[test]
fn raw_mtef_keeps_its_version_and_hostile_input_is_refused() {
    // MTEF 3 makes an Equation Editor 3.0 object.
    let package = open(EQUATION_EDITOR);
    let mtef3 = native(&package, 0, 0);
    let content =
        EquationContent::new(EquationSource::Mtef(&mtef3), None, 12.0, &Limits::DEFAULT).unwrap();
    assert_eq!(content.prog_id(), "Equation.3");
    // A replacement of another version is refused.
    let mathtype = MtefTarget::MathType;
    assert!(
        EquationContent::new(
            EquationSource::Mtef(&mtef3),
            Some(&mathtype),
            12.0,
            &Limits::DEFAULT
        )
        .is_err()
    );
    // Garbage is not MTEF.
    assert!(
        EquationContent::new(
            EquationSource::Mtef(b"garbage"),
            None,
            12.0,
            &Limits::DEFAULT
        )
        .is_err()
    );
    // Huge LaTeX stays inside the limits.
    let huge = "x+".repeat(2_000_000);
    assert!(
        EquationContent::new(EquationSource::Latex(&huge), None, 12.0, &Limits::DEFAULT).is_err()
    );
    // A missing object, and a missing paragraph.
    let mut edit = package.edit().unwrap();
    assert!(matches!(
        edit.remove_equation_object_at(None, &paragraph(0), Position::new(1)),
        Err(Error::Refused {
            reason: Refusal::EquationNotFound,
            ..
        })
    ));
    assert!(
        edit.insert_equation_object_at(None, &paragraph(9), None, &latex("x"))
            .is_err()
    );
}

#[test]
fn an_unreadable_embedding_keeps_the_kind_its_prog_id_names() {
    let package = open(MATHTYPE);
    let mut edit = package.edit().unwrap();
    edit.replace_part(
        "/word/embeddings/oleObject1.bin",
        b"not a compound file".to_vec(),
    )
    .unwrap();
    let kind = edit
        .equation_object_kind(None, &paragraph(0), Position::new(0), &Limits::DEFAULT)
        .unwrap();
    assert!(matches!(kind, MtefTarget::MathType));
    let content = EquationContent::new(
        EquationSource::Latex("x"),
        Some(&kind),
        12.0,
        &Limits::DEFAULT,
    )
    .unwrap();
    // The object gets a new compound file of its ProgID's kind.
    assert!(
        edit.set_equation_object_at(None, &paragraph(0), Position::new(0), &content)
            .unwrap()
    );
    let package = edit.commit().unwrap().into_package();
    let package = Package::from_bytes(&save(package)).unwrap();
    let objects = package.document().unwrap().paragraphs().unwrap()[0]
        .equation_objects()
        .unwrap();
    assert_eq!(objects[0].prog_id(), "Equation.DSMT4");
    let read = litchi_equation::read_mtef(&native(&package, 0, 0), &Limits::DEFAULT).unwrap();
    assert_eq!(read.latex(), "x");
    assert_eq!(read.info().unwrap().version, 5);
}

/// Every saved member whose bytes name `needle`.
fn traces(bytes: &[u8], needle: &str) -> Vec<String> {
    members(bytes)
        .into_iter()
        .filter(|(name, contents)| {
            name.contains(needle)
                || contents
                    .windows(needle.len())
                    .any(|window| window == needle.as_bytes())
        })
        .map(|(name, _)| name)
        .collect()
}

#[test]
fn an_object_inserted_and_removed_in_one_edit_leaves_no_trace() {
    use litchi_docx::document::{NewPicture, StoryTarget};
    // The main document: an equation object and a picture.
    let mut builder = litchi_docx::Builder::new().unwrap();
    builder.document().unwrap().add_paragraph_with_text("x");
    let base = builder.build().unwrap();
    let base_bytes = save(base.clone());
    let mut edit = base.edit().unwrap();
    edit.insert_equation_object_at(None, &paragraph(0), None, &latex("a"))
        .unwrap();
    let png = latex("b").preview_png().to_vec();
    edit.insert_picture_at(&paragraph(0), None, &NewPicture::new(png))
        .unwrap();
    edit.remove_picture_at(&paragraph(0), Position::new(0))
        .unwrap();
    edit.remove_equation_object_at(None, &paragraph(0), Position::new(0))
        .unwrap();
    let bytes = save(edit.commit().unwrap().into_package());
    for needle in ["oleObject", "image1.png", "embeddings", "media/"] {
        assert!(
            traces(&bytes, needle).is_empty(),
            "{needle}: {:?}",
            traces(&bytes, needle)
        );
    }
    assert_eq!(bytes, base_bytes);

    // A header.
    let (base, header, ..) = storied();
    let base_bytes = save(base.clone());
    let mut edit = base.edit().unwrap();
    let story = Some(StoryTarget::Part(&header));
    edit.insert_equation_object_at(story, &paragraph(0), None, &latex("a"))
        .unwrap();
    edit.remove_equation_object_at(story, &paragraph(0), Position::new(0))
        .unwrap();
    let bytes = save(edit.commit().unwrap().into_package());
    for needle in [
        "oleObject",
        "image1.png",
        "embeddings",
        "media/",
        "header1.xml.rels",
    ] {
        assert!(
            traces(&bytes, needle).is_empty(),
            "{needle}: {:?}",
            traces(&bytes, needle)
        );
    }
    assert_eq!(bytes, base_bytes);
}

/// A built document with a header, a footer, a footnote and an endnote,
/// and the stories' targets: header and footer part names, note ids.
fn storied() -> (Package, String, String, u32, u32) {
    use litchi_docx::document::{NewNote, NewStory, NoteKind};
    use litchi_docx::header_footer::{Kind, Role};
    let mut builder = litchi_docx::Builder::new().unwrap();
    builder.document().unwrap().add_paragraph_with_text("Claim");
    let mut edit = builder.build().unwrap().edit().unwrap();
    let header = edit
        .add_header_footer(Role::Header, Kind::Primary, &NewStory::new("Top"))
        .unwrap();
    let footer = edit
        .add_header_footer(Role::Footer, Kind::Primary, &NewStory::new("Bottom"))
        .unwrap();
    let footnote = edit
        .add_note_at(
            NoteKind::Footnote,
            &paragraph(0),
            None,
            &NewNote::new("Foot."),
        )
        .unwrap();
    let endnote = edit
        .add_note_at(
            NoteKind::Endnote,
            &paragraph(0),
            None,
            &NewNote::new("End."),
        )
        .unwrap();
    (
        edit.commit().unwrap().into_package(),
        header,
        footer,
        footnote,
        endnote,
    )
}

#[test]
fn objects_of_every_story_are_inserted_replaced_and_removed() {
    use litchi_docx::document::{NoteKind, StoryTarget};
    let (base, header, footer, footnote, endnote) = storied();
    let base_bytes = save(base.clone());
    let notes = |kind: &str| {
        members(&base_bytes)
            .keys()
            .find(|name| name.starts_with(&format!("word/{kind}")) && name.ends_with(".xml"))
            .map(|name| format!("/{name}"))
            .unwrap()
    };
    let stories = [
        (StoryTarget::Part(&header), header.clone()),
        (StoryTarget::Part(&footer), footer.clone()),
        (
            StoryTarget::Note(NoteKind::Footnote, footnote),
            notes("footnotes"),
        ),
        (
            StoryTarget::Note(NoteKind::Endnote, endnote),
            notes("endnotes"),
        ),
    ];
    for (story, part) in stories {
        let member = part.trim_start_matches('/').to_owned();
        let rels = {
            let (folder, file) = member.rsplit_once('/').unwrap();
            format!("{folder}/_rels/{file}.rels")
        };
        // Insert.
        let mut edit = base.edit().unwrap();
        let size = edit
            .equation_object_font_size_pt(Some(story), &paragraph(0), None)
            .unwrap();
        let content =
            EquationContent::new(EquationSource::Latex("x^{2}"), None, size, &Limits::DEFAULT)
                .unwrap();
        assert_eq!(
            edit.insert_equation_object_at(Some(story), &paragraph(0), None, &content)
                .unwrap(),
            Position::new(0)
        );
        let (inserted, patch) = edit.commit().unwrap().into_parts();
        let inserted_bytes = save(inserted.clone());
        let after = members(&inserted_bytes);
        let story_xml = text(&after[&member]);
        assert!(
            story_xml.contains(r#"ProgID="Equation.DSMT4""#),
            "{part}: {story_xml}"
        );
        let relationships = text(&after[&rels]);
        assert!(
            relationships.contains("relationships/oleObject"),
            "{relationships}"
        );
        assert!(
            relationships.contains("relationships/image"),
            "{relationships}"
        );
        // The main part's relationships and every other part are untouched.
        let before = members(&base_bytes);
        for (name, contents) in &before {
            if *name != member && *name != rels && name != "[Content_Types].xml" {
                assert_eq!(Some(contents), after.get(name), "{part}: {name} changed");
            }
        }
        let id = story_xml
            .split(r#"r:id=""#)
            .last()
            .and_then(|tail| tail.split('"').next())
            .unwrap()
            .to_owned();
        let payload = inserted
            .related_part_payload(Some(&part), &id)
            .unwrap()
            .expect("embedding");
        let native = litchi_equation::equation_object_native(&payload).unwrap();
        let read = litchi_equation::read_mtef(&native, &Limits::DEFAULT).unwrap();
        assert_eq!(read.latex(), "x^{2}");
        // The patch replays on its source and its inverse restores it; its
        // durable form records the story's relationships (package-graph
        // changes publish through `Package::apply`).
        assert_eq!(
            save(base.apply(&patch).unwrap().into_package()),
            inserted_bytes
        );
        patch.to_durable(durable_limits()).unwrap();
        assert_eq!(
            save(inserted.apply(&patch.inverse()).unwrap().into_package()),
            base_bytes
        );

        // Replace.
        let mut edit = inserted.edit().unwrap();
        let kind = edit
            .equation_object_kind(
                Some(story),
                &paragraph(0),
                Position::new(0),
                &Limits::DEFAULT,
            )
            .unwrap();
        let content = EquationContent::new(
            EquationSource::Latex("y"),
            Some(&kind),
            size,
            &Limits::DEFAULT,
        )
        .unwrap();
        edit.set_equation_object_at(Some(story), &paragraph(0), Position::new(0), &content)
            .unwrap();
        let (replaced, patch) = edit.commit().unwrap().into_parts();
        let payload = replaced
            .related_part_payload(Some(&part), &id)
            .unwrap()
            .expect("embedding");
        let native = litchi_equation::equation_object_native(&payload).unwrap();
        assert_eq!(
            litchi_equation::read_mtef(&native, &Limits::DEFAULT)
                .unwrap()
                .latex(),
            "y"
        );
        assert_eq!(
            save(replaced.apply(&patch.inverse()).unwrap().into_package()),
            inserted_bytes
        );

        // Remove: the parts and the story's relationships go.
        let mut edit = replaced.edit().unwrap();
        edit.remove_equation_object_at(Some(story), &paragraph(0), Position::new(0))
            .unwrap();
        let (removed, patch) = edit.commit().unwrap().into_parts();
        let removed_bytes = save(removed.clone());
        let after = members(&removed_bytes);
        assert!(
            !after.keys().any(|name| {
                name.starts_with("word/embeddings/") || name.starts_with("word/media/")
            }),
            "{part}: {:?}",
            after.keys().collect::<Vec<_>>()
        );
        let relationships = after
            .get(&rels)
            .map(|bytes| text(bytes))
            .unwrap_or_default();
        assert!(!relationships.contains("oleObject"), "{relationships}");
        assert!(!text(&after[&member]).contains("w:object"));
        assert_eq!(
            save(removed.apply(&patch.inverse()).unwrap().into_package()),
            save(replaced.clone())
        );
        assert_eq!(
            save(replaced.apply(&patch).unwrap().into_package()),
            removed_bytes
        );
    }
}

#[test]
fn previews_are_drawn_at_the_runs_resolved_size() {
    // A run's own size wins over the styles.
    let package = open(MATHTYPE);
    let mut edit = package.edit().unwrap();
    let defaults = edit
        .equation_object_font_size_pt(None, &paragraph(0), None)
        .unwrap();
    assert!(defaults > 0.0);
    let mut builder = litchi_docx::Builder::new().unwrap();
    builder.document().unwrap().add_paragraph_with_text("x");
    let mut edit = builder.build().unwrap().edit().unwrap();
    edit.format_run_at(
        &ElementPath::new([
            ElementStep::Paragraph(Position::new(0)),
            ElementStep::Run(Position::new(0)),
        ]),
        |run| run.size = Some(litchi_docx::properties::HalfPoints(40)),
    )
    .unwrap();
    let size = edit
        .equation_object_font_size_pt(None, &paragraph(0), None)
        .unwrap();
    assert!((size - 20.0).abs() < f64::EPSILON, "{size}");
    // A larger size draws a larger preview and lowers the run further.
    let small =
        EquationContent::new(EquationSource::Latex("y"), None, 10.0, &Limits::DEFAULT).unwrap();
    let large =
        EquationContent::new(EquationSource::Latex("y"), None, 20.0, &Limits::DEFAULT).unwrap();
    assert!(large.size_pt().1 > small.size_pt().1);
}

fn durable_limits() -> litchi_core::patch::PatchLimits {
    litchi_core::patch::PatchLimits::new(
        litchi_core::patch::BlobLimits::new(64, 32 * 1024 * 1024, 32 * 1024 * 1024),
        1024 * 1024,
        32,
        8,
        256 * 1024,
        512 * 1024,
    )
}
