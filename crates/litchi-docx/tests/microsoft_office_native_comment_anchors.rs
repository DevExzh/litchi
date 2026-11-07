//! Native modern-comment metadata and inherited legacy anchors; derivatives stay inert.
use litchi_core::Position;
use litchi_docx::builder::{OleObject, VmlShape};
use litchi_docx::drawing::{AnchorId, LegacyAnchor, LegacyAnchorKind};
use litchi_docx::modern_comments::{
    self as modern, Conformance, Extension, ExtensionList, Reaction, ReactionInfo, ReactionUser,
};
use litchi_docx::{Builder, Package, RunContent};
use litchi_ooxml_common::mce::Limits;
use litchi_opc::{OpcPackage, PackURI};
use std::io::Cursor;
const VISUAL: &str = "word-visual-objects.docx";
const REVIEW: &str = "word-review-controls.docx";
fn native(name: &str) -> Vec<u8> {
    std::fs::read(
        std::path::Path::new(env!("CARGO_MANIFEST_DIR"))
            .join("../../test-data/microsoft-office/word")
            .join(name),
    )
    .unwrap()
}
fn package(name: &str) -> Package {
    Package::from_bytes(&native(name)).unwrap()
}
fn members(bytes: &[u8]) -> std::collections::BTreeMap<String, Vec<u8>> {
    let reader = litchi_opc::phys_pkg::OwnedPhysPkgReader::from_bytes(
        bytes.to_vec(),
        litchi_opc::SourceLimits::default(),
    )
    .unwrap();
    reader
        .member_names()
        .unwrap()
        .into_iter()
        .map(|name| {
            let data = reader.read_member(&name).unwrap();
            (name, data)
        })
        .collect()
}

fn assert_members_equal(left: &[u8], right: &[u8]) {
    let a = members(left);
    let b = members(right);
    assert_eq!(a.keys().collect::<Vec<_>>(), b.keys().collect::<Vec<_>>());
    for (name, bytes) in a {
        assert!(b[&name] == bytes, "changed {name}");
    }
}

fn metadata(p: &Package) -> modern::Metadata {
    modern::load_modern_comment_metadata(
        p.opc_package(),
        &PackURI::new("/word/document.xml").unwrap(),
    )
    .unwrap()
}
fn anchors(paragraphs: &[litchi_docx::Paragraph]) -> Vec<(usize, LegacyAnchor)> {
    paragraphs
        .iter()
        .enumerate()
        .flat_map(|(i, p)| p.legacy_anchors().unwrap().into_iter().map(move |a| (i, a)))
        .collect()
}

fn objects(p: &Package) -> Vec<litchi_docx::object::EmbeddedObject> {
    p.document()
        .unwrap()
        .paragraphs()
        .unwrap()
        .iter()
        .flat_map(|p| p.runs().unwrap())
        .flat_map(|r| r.contents().unwrap())
        .filter_map(|c| {
            if let RunContent::EmbeddedObject(v) = c {
                Some(*v)
            } else {
                None
            }
        })
        .collect()
}
fn save(p: &Package) -> Package {
    let f = tempfile::NamedTempFile::with_suffix(".docx").unwrap();
    p.save(f.path()).unwrap();
    Package::open(f.path()).unwrap()
}
#[test]
fn native_ole_anchor_reads_document_root_namespace_in_eager_and_source_views() {
    let bytes = native(VISUAL);
    let p = Package::from_bytes(&bytes).unwrap();
    let a = anchors(&p.document().unwrap().paragraphs().unwrap());
    assert_eq!(a.len(), 1);
    assert_eq!(a[0].0, 6);
    assert_eq!(a[0].1.kind(), LegacyAnchorKind::Object);
    assert_eq!(a[0].1.anchor_id(), AnchorId::new(0x61c01e2c));
    let source = litchi_docx::source_backed::Package::from_reader(Cursor::new(&bytes)).unwrap();
    assert_eq!(
        anchors(&source.document().unwrap().paragraphs().unwrap()),
        a
    );
    // The native VML compatibility branches are not selected by these paragraph views.
    let vector = package("word-vector-diagram.docx");
    assert!(anchors(&vector.document().unwrap().paragraphs().unwrap()).is_empty());
    assert_members_equal(&save(&p).to_bytes().unwrap(), &bytes);
    assert_members_equal(&p.to_bytes().unwrap(), &bytes);
}
#[test]
fn native_embedded_workbook_and_preview_survive_fresh_legacy_anchor_authoring() {
    let p = package(VISUAL);
    let before = p.to_bytes().unwrap();
    let list = objects(&p);
    assert_eq!(list.len(), 1);
    let original = &list[0];
    let main = p.opc_package().main_document_part().unwrap();
    let payload = original
        .ole_data(p.opc_package(), main.rels())
        .unwrap()
        .unwrap();
    let preview = original
        .preview_image()
        .unwrap()
        .data(p.opc_package(), main.rels())
        .unwrap();
    assert_eq!(original.prog_id(), Some("Excel.Sheet.12"));
    assert_eq!(payload.len(), 18372);
    for present in [true, false] {
        let mut object = OleObject::new(original.prog_id().unwrap(), payload.to_vec()).unwrap();
        object.set_preview(preview.to_vec()).unwrap();
        object.set_size_emu(2489200, 1028700).unwrap();
        object.set_anchor_id(AnchorId::new(0x61c01e2c).unwrap());
        if !present {
            object.clear_anchor_id();
        }
        let mut vml = VmlShape::rectangle();
        vml.set_anchor_id(AnchorId::new(1).unwrap());
        vml.set_anchor_id(AnchorId::new(0x7fff_ffff).unwrap());
        if !present {
            vml.clear_anchor_id();
        }
        let mut b = Builder::new().unwrap();
        b.document().unwrap().add_ole_object(object).unwrap();
        b.document().unwrap().add_vml_shape(vml).unwrap();
        let out = save(&b.build().unwrap());
        let found = anchors(&out.document().unwrap().paragraphs().unwrap());
        assert_eq!(found.len(), 2);
        assert_eq!(found[0].1.kind(), LegacyAnchorKind::Object);
        assert_eq!(found[1].1.kind(), LegacyAnchorKind::Picture);
        assert_eq!(
            found[0].1.anchor_id(),
            if present {
                AnchorId::new(0x61c01e2c)
            } else {
                None
            }
        );
        assert_eq!(
            found[1].1.anchor_id(),
            if present {
                AnchorId::new(0x7fff_ffff)
            } else {
                None
            }
        );
        let copied = objects(&out);
        assert_eq!(copied.len(), 1);
        let rels = out.opc_package().main_document_part().unwrap().rels();
        assert_eq!(
            copied[0]
                .ole_data(out.opc_package(), rels)
                .unwrap()
                .unwrap()
                .as_ref(),
            payload.as_ref()
        );
        assert_eq!(
            copied[0]
                .preview_image()
                .unwrap()
                .data(out.opc_package(), rels)
                .unwrap(),
            preview
        );
    }
    assert_eq!(AnchorId::new(0), None);
    assert_eq!(AnchorId::new(0x8000_0000), None);
    assert_members_equal(&p.to_bytes().unwrap(), &before);
    assert_members_equal(&before, &native(VISUAL));
}
fn fragment(xml: &[u8], name: &[u8]) -> Vec<u8> {
    use quick_xml::events::Event;
    let mut r = quick_xml::Reader::from_reader(xml);
    let mut start = None;
    let mut depth = 0;
    loop {
        let at = r.buffer_position() as usize;
        let e = r.read_event().unwrap();
        let end = r.buffer_position() as usize;
        match e {
            Event::Start(e) => {
                if start.is_none() && e.name().as_ref() == name {
                    start = Some((at, depth));
                }
                depth += 1;
            },
            Event::Empty(e) if start.is_none() && e.name().as_ref() == name => {
                return xml[at..end].to_vec();
            },
            Event::End(_) => {
                depth -= 1;
                if start.is_some_and(|(_, d)| d == depth) {
                    return xml[start.unwrap().0..end].to_vec();
                }
            },
            Event::Eof => panic!("missing fragment"),
            _ => {},
        }
    }
}
#[test]
fn native_anchor_and_comment_metadata_survive_unrelated_text_publication() {
    for name in [VISUAL, REVIEW] {
        let p = package(name);
        let before = p.to_bytes().unwrap();
        let parts = members(&before);
        let expected_anchors = anchors(&p.document().unwrap().paragraphs().unwrap());
        let expected_comments = metadata(&p);
        let mut e = p.edit().unwrap();
        e.replace_run_text(
            Position::new(0),
            Position::new(0),
            "NATIVE_METADATA_PRESERVED",
        )
        .unwrap();
        let c = e.commit().unwrap();
        let out = save(c.package());
        let after = out.to_bytes().unwrap();
        let changed = members(&after);
        assert_eq!(
            out.document().unwrap().paragraphs().unwrap()[0]
                .text()
                .unwrap(),
            "NATIVE_METADATA_PRESERVED"
        );
        assert_eq!(
            anchors(&out.document().unwrap().paragraphs().unwrap()),
            expected_anchors
        );
        assert_eq!(metadata(&out), expected_comments);
        assert_eq!(
            parts.keys().collect::<Vec<_>>(),
            changed.keys().collect::<Vec<_>>()
        );
        for (key, value) in &parts {
            if key != "word/document.xml" {
                assert!(changed[key] == *value, "{name}: changed {key}");
            }
        }
        if name == VISUAL {
            assert_eq!(
                fragment(&parts["word/document.xml"], b"w:object"),
                fragment(&changed["word/document.xml"], b"w:object")
            );
        }
        assert_members_equal(
            &out.apply(&c.patch().inverse())
                .unwrap()
                .into_package()
                .to_bytes()
                .unwrap(),
            &before,
        );
        assert_members_equal(&p.to_bytes().unwrap(), &before);
    }
}
#[test]
fn native_comment_extension_graphs_keep_exact_durable_ids_times_and_parent_links() {
    for (name, paragraphs, durable, times) in [
        (
            REVIEW,
            vec![0x44bdd156, 0x3fd135f2],
            vec![0x472edeae, 0x7a8903a0],
            vec!["2026-10-15T05:50:00Z", "2026-10-15T05:51:00Z"],
        ),
        (
            "word-text-structure.docx",
            vec![0x27aaa1ec],
            vec![0x54870228],
            vec!["2026-10-15T05:22:00Z"],
        ),
        (
            "word-document-protection.docx",
            vec![0x73a0d512],
            vec![0x19fbd274],
            vec!["2026-10-15T08:20:00Z"],
        ),
    ] {
        let p = package(name);
        let m = metadata(&p);
        let ext = m.comments_extended.as_ref().unwrap();
        assert_eq!(
            ext.iter().map(|v| v.paragraph_id).collect::<Vec<_>>(),
            paragraphs
        );
        assert!(ext.iter().all(|v| !v.done));
        assert_eq!(ext[0].parent_paragraph_id, None);
        if ext.len() == 2 {
            assert_eq!(ext[1].parent_paragraph_id, Some(paragraphs[0]));
        }
        let ids = m.comments_ids.as_ref().unwrap();
        assert_eq!(
            ids.iter().map(|v| v.durable_id).collect::<Vec<_>>(),
            durable
        );
        assert_eq!(
            ids.iter().map(|v| v.paragraph_id).collect::<Vec<_>>(),
            paragraphs
        );
        let comments = m.comments_extensible.as_ref().unwrap();
        assert_eq!(
            comments.iter().map(|v| v.durable_id).collect::<Vec<_>>(),
            durable
        );
        assert_eq!(
            comments
                .iter()
                .map(|v| v.date_utc.as_deref().unwrap())
                .collect::<Vec<_>>(),
            times
        );
        assert!(
            comments
                .iter()
                .all(|v| v.reactions.is_empty() && v.intelligent_placeholder.is_none())
        );
        let people = m.people.as_ref().unwrap();
        assert_eq!(people.len(), 1);
        assert_eq!(people[0].author, "Ryker Zhu");
        let presence = people[0].presence.as_ref().unwrap();
        assert_eq!(presence.provider_id, "None");
        assert_eq!(presence.user_id, "Ryker Zhu");
        for mode in [Conformance::Transitional, Conformance::Strict] {
            let limits = Limits::default();
            assert_eq!(
                modern::parse_comments_extended(
                    &modern::write_comments_extended(ext, mode, &limits).unwrap(),
                    &limits
                )
                .unwrap(),
                *ext
            );
            assert_eq!(
                modern::parse_comments_ids(
                    &modern::write_comments_ids(ids, mode, &limits).unwrap(),
                    &limits
                )
                .unwrap(),
                *ids
            );
            assert_eq!(
                modern::parse_comments_extensible(
                    &modern::write_comments_extensible(comments, mode, &limits).unwrap(),
                    &limits
                )
                .unwrap(),
                *comments
            );
            assert_eq!(
                modern::parse_people(
                    &modern::write_people(people, mode, &limits).unwrap(),
                    &limits
                )
                .unwrap(),
                *people
            );
        }
        assert_members_equal(&p.to_bytes().unwrap(), &native(name));
    }
}
fn extensions() -> ExtensionList {
    ExtensionList::new(vec![
        Extension::new(
            Some("  urn:litchi:extension-one \n".into()),
            r#"<f:flag xmlns:f="urn:litchi:inert">yes &amp; no</f:flag>"#,
        )
        .unwrap(),
        Extension::new(None, r#"<f:second xmlns:f="urn:litchi:inert" value="2"/>"#).unwrap(),
    ])
    .unwrap()
}
#[test]
fn native_comment_models_round_trip_inert_reaction_and_ordered_extension_derivatives() {
    let p = package(REVIEW);
    let before = p.to_bytes().unwrap();
    let native = metadata(&p).comments_extensible.unwrap();
    let mut values = native.clone();
    let ext = extensions();
    assert_eq!(ext.extensions()[0].uri(), Some("urn:litchi:extension-one"));
    let reaction = Reaction {
        reaction_type: 1,
        reactions: vec![ReactionInfo {
            date_utc: Some("2026-10-15T05:52:30.125Z".into()),
            user: Some(ReactionUser {
                user_id: "Ryker Zhu".into(),
                user_name: "Ryker Zhu".into(),
                user_provider: "None".into(),
                extensions: Some(ext.clone()),
            }),
            extensions: Some(ext.clone()),
        }],
        extensions: Some(ext.clone()),
    };
    values[0].reactions.push(reaction.clone());
    let mut unknown = reaction.clone();
    unknown.reaction_type = 2;
    values[1].reactions.push(unknown);
    for mode in [Conformance::Transitional, Conformance::Strict] {
        let limits = Limits::default();
        let xml = modern::write_comments_extensible(&values, mode, &limits).unwrap();
        assert!(
            std::str::from_utf8(&xml)
                .unwrap()
                .contains("{CE6994B0-6A32-4C9F-8C6B-6E91EDA988CE}")
        );
        assert_eq!(
            modern::parse_comments_extensible(&xml, &limits).unwrap(),
            values
        );
    }
    let mut invalid = values.clone();
    invalid[0].reactions[0].reactions[0].date_utc = Some("2026-10-15T05:52:30+05:00".into());
    assert!(matches!(
        modern::write_comments_extensible(&invalid, Conformance::Transitional, &Limits::default()),
        Err(litchi_docx::Error::Invalid(_))
    ));
    invalid = values.clone();
    invalid[0].reactions.push(reaction);
    assert!(
        modern::write_comments_extensible(&invalid, Conformance::Transitional, &Limits::default())
            .is_err()
    );
    let mut list = ext.clone();
    assert_eq!(list.remove(0).unwrap(), ext.extensions()[0]);
    assert_eq!(list.extensions(), &ext.extensions()[1..]);
    assert!(Extension::new(None, "<a/><b/>").is_err());
    for v in &mut values {
        v.reactions.clear();
    }
    assert_eq!(values, native);
    assert_members_equal(&p.to_bytes().unwrap(), &before);
}
#[test]
fn native_existing_modern_comment_graph_refuses_create_only_store_atomically() {
    let p = package(REVIEW);
    let before = p.to_bytes().unwrap();
    let value = metadata(&p);
    let mut opc = OpcPackage::from_bytes(&before, litchi_opc::SourceLimits::default()).unwrap();
    let ids = modern::RelationshipIds {
        comments_extended: Some("rIdNewExtended".into()),
        comments_ids: Some("rIdNewIds".into()),
        comments_extensible: Some("rIdNewExtensible".into()),
        people: Some("rIdNewPeople".into()),
    };
    assert!(
        matches!(modern::store_modern_comment_metadata(&mut opc,&PackURI::new("/word/document.xml").unwrap(),&value,&ids,Conformance::Transitional),Err(litchi_docx::Error::Invalid(message)) if message.contains("part or relationship already exists"))
    );
    let mut output = Vec::new();
    opc.write_to(&mut output).unwrap();
    assert_members_equal(&output, &before);
    assert_eq!(
        modern::load_modern_comment_metadata(&opc, &PackURI::new("/word/document.xml").unwrap())
            .unwrap(),
        value
    );
}
