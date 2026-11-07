//! Corpus evidence only. Ink, comments and OLE do not qualify other owners.
#![cfg(feature = "encryption")]
use litchi_pptx::{
    Package,
    presentation::embedded::{controls, ink},
    presentation_properties::metadata::{changes, revision},
};
use quick_xml::{events::Event, name::ResolveResult, reader::NsReader};
use std::{io::Cursor, path::Path};

fn corpus(mut check: impl FnMut(&str, &Package)) {
    let root = Path::new(env!("CARGO_MANIFEST_DIR")).join("../../test-data/microsoft-office");
    let manifest: serde_json::Value =
        serde_json::from_slice(&std::fs::read(root.join("manifest.json")).unwrap()).unwrap();
    let mut files = 0;
    let mut encrypted = 0;
    for item in manifest["files"].as_array().unwrap() {
        let name = item["path"].as_str().unwrap();
        if !name.starts_with("powerpoint/")
            || !matches!(
                Path::new(name).extension().and_then(|e| e.to_str()),
                Some("pptx" | "pptm" | "potx" | "potm" | "ppsx" | "ppsm")
            )
        {
            continue;
        }
        let bytes = std::fs::read(root.join(name)).unwrap();
        let p = if name == "powerpoint/powerpoint-encrypted.pptx" {
            encrypted += 1;
            Package::from_reader_with_password(Cursor::new(&bytes), "OfficeFixture42!").unwrap()
        } else {
            Package::from_reader(Cursor::new(&bytes)).unwrap()
        };
        assert!(
            p.presentation().unwrap().slide_count().unwrap() > 0,
            "{name}"
        );
        check(name, &p);
        if p.encryption().is_some() {
            assert!(matches!(
                p.to_bytes(),
                Err(litchi_pptx::Error::EncryptionPolicy { .. })
            ));
        } else {
            assert!(p.to_bytes().unwrap() == bytes, "{name}: source changed");
        }
        assert!(std::fs::read(root.join(name)).unwrap() == bytes);
        files += 1;
    }
    assert_eq!((files, encrypted), (15, 1));
}

// Inspect all XML branches, including inactive MCE alternatives and orphan
// parts, in addition to the typed presentation/slide discovery below.
fn elements(p: &Package, mut check: impl FnMut(&[u8], &[u8])) {
    for info in p.opc().unwrap().iter_parts() {
        if !info.partname().as_str().ends_with(".xml")
            && !info.content_type().ends_with("+xml")
            && !matches!(info.content_type(), "text/xml" | "application/xml")
        {
            continue;
        }
        let part = p.opc().unwrap().get_part(info.partname()).unwrap();
        let mut r = NsReader::from_reader(part.blob());
        loop {
            let (ns, event) = r.read_resolved_event().unwrap();
            match event {
                Event::Start(e) | Event::Empty(e) => {
                    let uri = match &ns {
                        ResolveResult::Bound(u) => u.as_ref(),
                        ResolveResult::Unbound => b"",
                        ResolveResult::Unknown(_) => {
                            panic!("unbound namespace in {}", part.partname())
                        },
                    };
                    check(uri, e.local_name().as_ref());
                },
                Event::Eof => break,
                _ => {},
            }
        }
    }
}
fn presentation_namespace(ns: &[u8]) -> bool {
    matches!(
        ns,
        b"http://schemas.openxmlformats.org/presentationml/2006/main"
            | b"http://purl.oclc.org/ooxml/presentationml/main"
    )
}

#[test]
fn native_powerpoint_corpus_has_ole_but_no_activex_controls() {
    let mut ole = 0;
    corpus(|name, p| {
        let view = p.presentation().unwrap();
        for (index, slide) in view.slides().unwrap().iter().enumerate() {
            assert!(
                controls::load_slide(
                    p.opc().unwrap(),
                    index,
                    slide.part().part(),
                    &mut controls::Limits::default()
                )
                .unwrap()
                .is_empty(),
                "{name}"
            );
        }
        for part in p.opc().unwrap().iter_parts() {
            assert!(
                !part.content_type().to_ascii_lowercase().contains("activex"),
                "{name}"
            );
            for rel in part.rels().iter() {
                assert!(
                    !rel.reltype().to_ascii_lowercase().contains("activex")
                        && !rel.reltype().ends_with("/control"),
                    "{name}: {}",
                    rel.reltype()
                );
            }
        }
        elements(p, |ns, local| {
            assert_ne!(
                ns, b"http://schemas.microsoft.com/office/2006/activeX",
                "{name}"
            );
            assert!(
                !(presentation_namespace(ns) && matches!(local, b"control" | b"controls")),
                "{name}"
            );
            ole += usize::from(presentation_namespace(ns) && local == b"oleObj");
        });
    });
    assert_eq!(ole, 4);
}

#[test]
fn native_powerpoint_corpus_has_comments_but_no_revision_or_changes_information() {
    let mut comment_parts = 0;
    corpus(|name, p| {
        let opc = p.opc().unwrap();
        assert!(
            revision::load(opc, revision::Limits::DEFAULT)
                .unwrap()
                .is_none(),
            "{name}"
        );
        assert!(
            changes::load(opc, changes::Limits::DEFAULT)
                .unwrap()
                .is_none(),
            "{name}"
        );
        for part in opc.iter_parts() {
            assert!(
                !matches!(
                    part.content_type(),
                    revision::CONTENT_TYPE | changes::CONTENT_TYPE
                ),
                "{name}"
            );
            comment_parts +=
                usize::from(part.content_type() == "application/vnd.ms-powerpoint.comments+xml");
            for rel in part.rels().iter() {
                assert!(
                    !matches!(
                        rel.reltype(),
                        revision::RELATIONSHIP_TYPE | changes::RELATIONSHIP_TYPE
                    ),
                    "{name}"
                );
            }
        }
        elements(p, |_, local| {
            assert!(
                !matches!(local, b"revInfo" | b"chgInfo"),
                "{name}: new review owner"
            )
        });
    });
    assert_eq!(comment_parts, 2);
}

#[test]
fn native_powerpoint_corpus_content_parts_are_ink_only() {
    let mut anchors = 0;
    let mut ink_files = Vec::new();
    corpus(|name, p| {
        let parts = p.presentation().unwrap().content_parts().unwrap();
        if !parts.is_empty() {
            ink_files.push(name.to_string());
        }
        let mut physical_anchors = 0;
        elements(p, |ns, local| {
            physical_anchors += usize::from(presentation_namespace(ns) && local == b"contentPart");
        });
        assert_eq!(
            physical_anchors,
            parts.len(),
            "{name}: inspect inactive/orphan owners too"
        );
        for part in &parts {
            let payload = part.payload().unwrap();
            assert_eq!(
                payload.content_type(),
                ink::CONTENT_TYPE,
                "{name}: new non-Ink evidence"
            );
            assert_eq!(payload.part_name().as_str(), "/ppt/ink/ink1.xml");
            assert_eq!(payload.bytes().len(), 2424);
            assert_eq!(part.relationship_id(), "rId2");
        }
        anchors += parts.len();
    });
    assert_eq!(anchors, 2);
    assert_eq!(
        ink_files,
        [
            "powerpoint/powerpoint-ink-stroke.pptx",
            "powerpoint/powerpoint-ink-replay.pptx"
        ]
    );
}
