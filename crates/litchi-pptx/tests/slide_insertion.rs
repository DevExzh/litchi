#![allow(clippy::unwrap_used, clippy::expect_used, reason = "test assertions")]
use litchi_core::Position;
use litchi_pptx::{
    Error, Package,
    opened::{Limits, Patch},
};
use soapberry_zip::office::{ArchiveReader, StreamingArchiveWriter};

fn layout(package: &Package) -> String {
    package
        .presentation()
        .unwrap()
        .slide_layouts()
        .unwrap()
        .into_iter()
        .find(|layout| layout.kind().unwrap().as_deref() == Some("blank"))
        .unwrap()
        .name()
        .unwrap()
}
fn source(strict: bool) -> Package {
    let bytes = Package::new().unwrap().to_bytes().unwrap();
    let archive =
        ArchiveReader::new(&bytes, soapberry_zip::office::ArchiveLimits::default()).unwrap();
    let mut writer =
        StreamingArchiveWriter::new(soapberry_zip::office::StreamingArchiveLimits::default());
    for name in archive.file_names() {
        let bytes = archive.read(name).unwrap();
        let mut xml = String::from_utf8(bytes).unwrap();
        if strict {
            for (old, new) in [
                (
                    "http://schemas.openxmlformats.org/presentationml/2006/main",
                    "http://purl.oclc.org/ooxml/presentationml/main",
                ),
                (
                    "http://schemas.openxmlformats.org/drawingml/2006/main",
                    "http://purl.oclc.org/ooxml/drawingml/main",
                ),
                (
                    "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
                    "http://purl.oclc.org/ooxml/officeDocument/relationships",
                ),
            ] {
                xml = xml.replace(old, new);
            }
        }
        if [
            "[Content_Types].xml",
            "ppt/_rels/presentation.xml.rels",
            "ppt/presentation.xml",
        ]
        .contains(&name)
        {
            let at = xml.rfind("</").unwrap();
            xml.insert_str(at, "<!--insertion sentinel-->");
        }
        writer.write_stored(name, xml.as_bytes()).unwrap();
    }
    Package::from_reader(std::io::Cursor::new(writer.finish().unwrap().into_inner())).unwrap()
}
fn names(package: &Package) -> Vec<String> {
    package
        .presentation()
        .unwrap()
        .slides()
        .unwrap()
        .iter()
        .map(|s| s.name().unwrap())
        .collect()
}
#[test]
fn insertions_into_empty_and_populated_lists_compose_and_invert_losslessly() {
    for strict in [false, true] {
        let original = source(strict);
        let original_bytes = original.to_bytes().unwrap();
        let layout = layout(&original);
        let mut edit = original.edit().unwrap();
        let first = edit.insert_slide(Position::new(0), &layout).unwrap();
        edit.add_text_box(Position::new(0), "new slide text", (0, 0, 914400, 914400))
            .unwrap();
        let second = edit.insert_slide(Position::new(0), &layout).unwrap();
        let third = edit.insert_slide(Position::new(2), &layout).unwrap();
        assert_eq!(edit.slides().unwrap().len(), 3);
        assert_ne!(first.id(), second.id());
        assert_ne!(second.id(), third.id());
        let (changed, patch) = edit.commit().unwrap().into_parts();
        assert_eq!(names(&changed), [second.name(), first.name(), third.name()]);
        let wire = Patch::from_bytes(&patch.to_bytes().unwrap(), Limits::default()).unwrap();
        let replay = original.apply(&wire).unwrap().into_package();
        assert_eq!(replay.to_bytes().unwrap(), changed.to_bytes().unwrap());
        let reopened =
            Package::from_reader(std::io::Cursor::new(replay.to_bytes().unwrap())).unwrap();
        assert_eq!(names(&reopened), names(&changed));
        let restored = replay
            .apply(
                &Patch::from_bytes(&wire.inverse().to_bytes().unwrap(), Limits::default()).unwrap(),
            )
            .unwrap()
            .into_package();
        assert_eq!(restored.to_bytes().unwrap(), original_bytes);
        assert_eq!(original.to_bytes().unwrap(), original_bytes);
        let after = changed.to_bytes().unwrap();
        let archive =
            ArchiveReader::new(&after, soapberry_zip::office::ArchiveLimits::default()).unwrap();
        for name in [
            "[Content_Types].xml",
            "ppt/_rels/presentation.xml.rels",
            "ppt/presentation.xml",
        ] {
            assert!(
                archive
                    .read(name)
                    .unwrap()
                    .windows(18)
                    .any(|v| v == b"insertion sentinel")
            );
        }
        let slide = archive.read("ppt/slides/slide1.xml").unwrap();
        let text = String::from_utf8(slide).unwrap();
        assert!(text.contains(if strict {
            "http://purl.oclc.org/ooxml/presentationml/main"
        } else {
            "http://schemas.openxmlformats.org/presentationml/2006/main"
        }));
    }
}
#[test]
fn failed_insertion_preserves_prior_staged_edits() {
    let original = source(false);
    let layout = layout(&original);
    let mut edit = original.edit().unwrap();
    edit.insert_slide(Position::new(0), &layout).unwrap();
    assert!(matches!(
        edit.insert_slide(Position::new(2), &layout),
        Err(Error::SlideIndexOutOfBounds { .. })
    ));
    assert!(matches!(
        edit.insert_slide(Position::new(1), "missing"),
        Err(Error::SlideLayoutNameNotFound(_))
    ));
    edit.add_text_box(Position::new(0), "still usable", (0, 0, 914400, 914400))
        .unwrap();
    let changed = edit.commit().unwrap().into_package();
    assert_eq!(names(&changed).len(), 1);
    let limits = Limits::new(1024, 1024, 1024, 10, 1 << 20, 1 << 20).unwrap();
    let root = original
        .opened_presentation_with_options(
            litchi_pptx::opened::EditOptions::default().limits(limits),
        )
        .unwrap();
    let mut bounded = root.edit();
    assert!(matches!(
        bounded.insert_slide(Position::new(0), &layout),
        Err(Error::Limit { .. } | Error::Opc(_))
    ));
    assert!(bounded.commit().unwrap().patch().is_noop());
}

fn with_incoming_reference(bytes: &[u8]) -> Vec<u8> {
    let archive =
        ArchiveReader::new(bytes, soapberry_zip::office::ArchiveLimits::default()).unwrap();
    let mut writer =
        StreamingArchiveWriter::new(soapberry_zip::office::StreamingArchiveLimits::default());
    for name in archive.file_names() {
        writer
            .write_stored(name, &archive.read(name).unwrap())
            .unwrap();
    }
    writer
        .write_stored("custom/incoming.xml", b"<incoming/>")
        .unwrap();
    writer.write_stored("custom/_rels/incoming.xml.rels", br#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="incoming" Type="urn:incoming" Target="../PPT/SLIDES/SLIDE1.XML"/></Relationships>"#).unwrap();
    writer.finish().unwrap().into_inner()
}
#[test]
fn inverse_refuses_a_new_incoming_owner_atomically() {
    let original = source(false);
    let mut edit = original.edit().unwrap();
    edit.insert_slide(Position::new(0), &layout(&original))
        .unwrap();
    let (changed, patch) = edit.commit().unwrap().into_parts();
    let dependent = Package::from_reader(std::io::Cursor::new(with_incoming_reference(
        &changed.to_bytes().unwrap(),
    )))
    .unwrap();
    let before = dependent.to_bytes().unwrap();
    assert!(
        dependent.apply(&patch.inverse()).is_err(),
        "inverse must not orphan an unrelated owner"
    );
    assert_eq!(dependent.to_bytes().unwrap(), before);
}
#[test]
fn insertion_does_not_activate_an_existing_dangling_target() {
    let original = source(false);
    let original = Package::from_reader(std::io::Cursor::new(with_incoming_reference(
        &original.to_bytes().unwrap(),
    )))
    .unwrap();
    let mut edit = original.edit().unwrap();
    edit.insert_slide(Position::new(0), &layout(&original))
        .unwrap();
    let changed = edit.commit().unwrap().into_package();
    let bytes = changed.to_bytes().unwrap();
    let archive =
        ArchiveReader::new(&bytes, soapberry_zip::office::ArchiveLimits::default()).unwrap();
    assert!(
        archive.read("ppt/slides/slide1.xml").is_err(),
        "an existing dangling target reserves the name"
    );
    assert!(archive.read("ppt/slides/slide2.xml").is_ok());
}

#[test]
fn duplicate_layout_names_are_a_typed_atomic_refusal() {
    let original = source(false);
    let blank = layout(&original);
    let (member, other) = original
        .presentation()
        .unwrap()
        .slide_layouts()
        .unwrap()
        .into_iter()
        .find_map(|candidate| {
            let name = candidate.name().unwrap();
            (name != blank).then(|| {
                (
                    candidate.part().part().partname().membername().to_owned(),
                    name,
                )
            })
        })
        .unwrap();
    let bytes = original.to_bytes().unwrap();
    let archive =
        ArchiveReader::new(&bytes, soapberry_zip::office::ArchiveLimits::default()).unwrap();
    let mut writer =
        StreamingArchiveWriter::new(soapberry_zip::office::StreamingArchiveLimits::default());
    for name in archive.file_names() {
        let mut bytes = archive.read(name).unwrap();
        if name == member {
            let xml = String::from_utf8(bytes).unwrap();
            let needle = format!("name=\"{other}\"");
            assert!(xml.contains(&needle));
            bytes = xml
                .replace(&needle, &format!("name=\"{blank}\""))
                .into_bytes();
        }
        writer.write_stored(name, &bytes).unwrap();
    }
    let original =
        Package::from_reader(std::io::Cursor::new(writer.finish().unwrap().into_inner())).unwrap();
    let mut edit = original.edit().unwrap();
    assert!(matches!(
        edit.insert_slide(Position::new(0), &blank),
        Err(Error::AmbiguousSlideLayoutName { matches: 2, .. })
    ));
    assert!(edit.commit().unwrap().patch().is_noop());
}
#[test]
fn insertion_and_inverse_preserve_unrelated_destination_metadata() {
    let original = source(false);
    let mut edit = original.edit().unwrap();
    edit.insert_slide(Position::new(0), &layout(&original))
        .unwrap();
    let patch = edit.commit().unwrap().into_patch();
    let bytes = original.to_bytes().unwrap();
    let archive =
        ArchiveReader::new(&bytes, soapberry_zip::office::ArchiveLimits::default()).unwrap();
    let mut writer =
        StreamingArchiveWriter::new(soapberry_zip::office::StreamingArchiveLimits::default());
    let mut expected = Vec::new();
    for name in archive.file_names() {
        let mut bytes = archive.read(name).unwrap();
        if ["[Content_Types].xml", "ppt/_rels/presentation.xml.rels"].contains(&name) {
            let mut xml = String::from_utf8(bytes).unwrap();
            let at = xml.find("?>").unwrap() + 2;
            xml.insert_str(at, "<!--independent destination markup-->");
            bytes = xml.into_bytes();
            expected.push((name.to_owned(), bytes.clone()));
        }
        writer.write_stored(name, &bytes).unwrap();
    }
    let overlay =
        Package::from_reader(std::io::Cursor::new(writer.finish().unwrap().into_inner())).unwrap();
    let wire = Patch::from_bytes(&patch.to_bytes().unwrap(), Limits::default()).unwrap();
    let changed = overlay.apply(&wire).unwrap().into_package();
    let bytes = changed.to_bytes().unwrap();
    let archive =
        ArchiveReader::new(&bytes, soapberry_zip::office::ArchiveLimits::default()).unwrap();
    for (name, _) in &expected {
        assert!(
            String::from_utf8(archive.read(name).unwrap())
                .unwrap()
                .contains("independent destination markup")
        );
    }
    let restored = changed
        .apply(&Patch::from_bytes(&wire.inverse().to_bytes().unwrap(), Limits::default()).unwrap())
        .unwrap()
        .into_package();
    let bytes = restored.to_bytes().unwrap();
    let archive =
        ArchiveReader::new(&bytes, soapberry_zip::office::ArchiveLimits::default()).unwrap();
    for (name, bytes) in expected {
        assert_eq!(archive.read(&name).unwrap(), bytes);
    }
}

#[test]
fn inverse_checks_root_dependencies_but_allows_external_references() {
    let original = source(false);
    let mut edit = original.edit().unwrap();
    edit.insert_slide(Position::new(0), &layout(&original))
        .unwrap();
    let (changed, patch) = edit.commit().unwrap().into_parts();
    let bytes = changed.to_bytes().unwrap();
    for external in [false, true] {
        let archive =
            ArchiveReader::new(&bytes, soapberry_zip::office::ArchiveLimits::default()).unwrap();
        let mut writer =
            StreamingArchiveWriter::new(soapberry_zip::office::StreamingArchiveLimits::default());
        for name in archive.file_names() {
            let mut bytes = archive.read(name).unwrap();
            if name == "_rels/.rels" {
                let mut xml = String::from_utf8(bytes).unwrap();
                let mode = if external {
                    " TargetMode=\"External\""
                } else {
                    ""
                };
                let at = xml.rfind("</").unwrap();
                xml.insert_str(at, &format!("<Relationship Id=\"incoming\" Type=\"urn:incoming\" Target=\"ppt/slides/slide1.xml\"{mode}/>"));
                bytes = xml.into_bytes();
            }
            writer.write_stored(name, &bytes).unwrap();
        }
        let overlay =
            Package::from_reader(std::io::Cursor::new(writer.finish().unwrap().into_inner()))
                .unwrap();
        let before = overlay.to_bytes().unwrap();
        let result = overlay.apply(&patch.inverse());
        if external {
            assert!(result.is_ok());
        } else {
            assert!(result.is_err());
        }
        assert_eq!(overlay.to_bytes().unwrap(), before);
    }
}

#[test]
fn fresh_package_insertion_uses_the_same_immutable_transaction() {
    let original = Package::new().unwrap();
    let before = original.to_bytes().unwrap();
    let mut edit = original.edit().unwrap();
    let slide = edit
        .insert_slide(Position::new(0), &layout(&original))
        .unwrap();
    edit.add_text_box(slide.name(), "fresh", (0, 0, 914400, 914400))
        .unwrap();
    let (changed, patch) = edit.commit().unwrap().into_parts();
    let replay = original
        .apply(&Patch::from_bytes(&patch.to_bytes().unwrap(), Limits::default()).unwrap())
        .unwrap()
        .into_package();
    assert_eq!(changed.to_bytes().unwrap(), replay.to_bytes().unwrap());
    let restored = replay.apply(&patch.inverse()).unwrap().into_package();
    assert_eq!(restored.to_bytes().unwrap(), before);
    assert!(
        original
            .presentation()
            .unwrap()
            .slides()
            .unwrap()
            .is_empty()
    );
}

#[test]
fn forward_replay_refuses_to_activate_a_destination_dangling_reference() {
    let original = source(false);
    let mut edit = original.edit().unwrap();
    edit.insert_slide(Position::new(0), &layout(&original))
        .unwrap();
    let patch = Patch::from_bytes(
        &edit.commit().unwrap().into_patch().to_bytes().unwrap(),
        Limits::default(),
    )
    .unwrap();
    let destination = Package::from_reader(std::io::Cursor::new(with_incoming_reference(
        &original.to_bytes().unwrap(),
    )))
    .unwrap();
    let before = destination.to_bytes().unwrap();
    assert!(destination.apply(&patch).is_err());
    assert_eq!(destination.to_bytes().unwrap(), before);
}

#[test]
fn insertion_refuses_dangling_xml_bindings_ranges_and_unknown_extensions() {
    for strict in [false, true] {
        let base = source(strict);
        let bytes = base.to_bytes().unwrap();
        let archive =
            ArchiveReader::new(&bytes, soapberry_zip::office::ArchiveLimits::default()).unwrap();
        let rels =
            String::from_utf8(archive.read("ppt/_rels/presentation.xml.rels").unwrap()).unwrap();
        let free = (1..100)
            .find(|n| !rels.contains(&format!("Id=\"rId{n}\"")))
            .unwrap();
        for feature in [
            format!(
                "<p:custShowLst><p:custShow name=\"dangling\" id=\"1\"><p:sldLst><p:sld r:id=\"rId{free}\"/></p:sldLst></p:custShow></p:custShowLst>"
            ),
            "<p:showPr><p:sldRg st=\"1\" end=\"2\"/></p:showPr>".to_owned(),
            "<p:extLst><p:ext uri=\"unknown\"/></p:extLst>".to_owned(),
        ] {
            let mut writer = StreamingArchiveWriter::new(
                soapberry_zip::office::StreamingArchiveLimits::default(),
            );
            for name in archive.file_names() {
                let mut payload = archive.read(name).unwrap();
                if name == "ppt/presentation.xml" {
                    let mut xml = String::from_utf8(payload).unwrap();
                    let end = xml.rfind("</").unwrap();
                    xml.insert_str(end, &feature);
                    payload = xml.into_bytes();
                }
                writer.write_stored(name, &payload).unwrap();
            }
            let source =
                Package::from_reader(std::io::Cursor::new(writer.finish().unwrap().into_inner()))
                    .unwrap();
            let before = source.to_bytes().unwrap();
            let mut edit = source.edit().unwrap();
            assert!(matches!(
                edit.insert_slide(Position::new(0), &layout(&source)),
                Err(Error::Refused(refusal)) if refusal.operation() == "insert_slide"
            ));
            assert_eq!(
                edit.commit().unwrap().into_package().to_bytes().unwrap(),
                before
            );
        }
    }
}
