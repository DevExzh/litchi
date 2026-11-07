//! Explicit conflict and ink metadata on native Word parents; no merge or recognition runtime.
use litchi_core::Position;
use litchi_docx::{Error, Package, ink, revision::conflict};
use litchi_opc::{BlobPart, OpcPackage, PackURI, SourceLimits, phys_pkg::OwnedPhysPkgReader};
use sha2::{Digest, Sha256};
use std::{collections::BTreeMap, path::Path};
const MAIN: &str = "word/document.xml";
const HEADER: &str = "word/header1.xml";
const W: &str = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
const W14: &str = "http://schemas.microsoft.com/office/word/2010/wordml";
const MC: &str = "http://schemas.openxmlformats.org/markup-compatibility/2006";
const REL: &str = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const PARENTS: [&str; 3] = [
    "word/word-tables-layout.docx",
    "word/word-recorded-macro.docm",
    "word/word-layout-extensions.dotx",
];
const DONOR: &str = "powerpoint/powerpoint-ink-stroke.pptx";
const NATIVE_INK: &str = "/word/ink/native-office-75.xml";
fn native(name: &str) -> Vec<u8> {
    std::fs::read(
        Path::new(env!("CARGO_MANIFEST_DIR"))
            .join("../../test-data/microsoft-office")
            .join(name),
    )
    .unwrap()
}
fn uri(name: &str) -> PackURI {
    PackURI::new(format!("/{name}")).unwrap()
}
fn members(p: &Package) -> BTreeMap<String, Vec<u8>> {
    let reader =
        OwnedPhysPkgReader::from_bytes(p.to_bytes().unwrap(), SourceLimits::default()).unwrap();
    reader
        .member_names()
        .unwrap()
        .into_iter()
        .map(|name| {
            let bytes = reader.read_member(&name).unwrap();
            (name, bytes)
        })
        .collect()
}
type Graph = BTreeMap<String, (String, Vec<(String, String, String, bool)>)>;
fn graph(p: &Package) -> Graph {
    let p = p.opc_package();
    let rels = |rels: &litchi_opc::Relationships| {
        let mut items = rels
            .iter()
            .map(|r| {
                (
                    r.r_id().into(),
                    r.reltype().into(),
                    r.target_ref().into(),
                    r.is_external(),
                )
            })
            .collect::<Vec<_>>();
        items.sort();
        items
    };
    let mut result = BTreeMap::from([("/".into(), (String::new(), rels(p.rels())))]);
    for p in p.iter_parts() {
        result.insert(
            p.partname().to_string(),
            (p.content_type().into(), rels(p.rels())),
        );
    }
    result
}
fn equivalent(a: &Package, b: &Package) {
    assert_eq!(graph(a), graph(b));
    let a = members(a);
    let b = members(b);
    assert_eq!(a.keys().collect::<Vec<_>>(), b.keys().collect::<Vec<_>>());
    for (name, bytes) in a {
        if ![
            "[Content_Types].xml",
            "word/_rels/document.xml.rels",
            "word/_rels/header1.xml.rels",
        ]
        .contains(&name.as_str())
        {
            assert!(b.get(&name) == Some(&bytes), "changed {name}");
        }
    }
}
fn only_changed(a: &Package, b: &Package, owned: &str) {
    assert_eq!(graph(a), graph(b));
    let a = members(a);
    let b = members(b);
    assert_eq!(a.keys().collect::<Vec<_>>(), b.keys().collect::<Vec<_>>());
    for (name, bytes) in a {
        if name != owned {
            assert!(
                b.get(&name) == Some(&bytes),
                "unowned member changed {name}"
            );
        }
    }
}
fn reopen(p: &Package) -> Package {
    let file = tempfile::NamedTempFile::with_suffix(".docx").unwrap();
    p.save(file.path()).unwrap();
    let q = Package::open(file.path()).unwrap();
    assert!(members(p) == members(&q));
    q
}
fn rewrite(p: &Package, name: &str, bytes: Vec<u8>) -> Package {
    let mut e = p.edit_opc();
    e.package_mut()
        .get_part_mut(&uri(name))
        .unwrap()
        .set_blob(bytes);
    e.commit().unwrap().into_package()
}
fn append(p: &Package, name: &str, fragment: &str) -> Package {
    let source = String::from_utf8(members(p)[name].clone()).unwrap();
    let marker = if name == MAIN {
        "<w:sectPr"
    } else {
        "</w:hdr>"
    };
    let at = source.rfind(marker).unwrap();
    let mut bytes = source.as_bytes()[..at].to_vec();
    bytes.extend_from_slice(fragment.as_bytes());
    bytes.extend_from_slice(&source.as_bytes()[at..]);
    rewrite(p, name, bytes)
}
fn conflict_fragment(id: i32) -> String {
    let declaration = format!(r#"xmlns:w="{W}" xmlns:c="{W14}" xmlns:mc="{MC}" mc:Ignorable="c""#);
    format!(
        r#"<w:p {declaration}><c:conflictIns w:id="{id}" w:author="Insert Author" w:date="2026-11-07T10:20:30Z"><w:r><w:t>added 日本語 &amp; kept</w:t></w:r></c:conflictIns><c:conflictDel w:id="{}" w:author="Delete Author"><w:r><w:delText>removed 🍎</w:delText></w:r></c:conflictDel></w:p><w:p {declaration}><c:customXmlConflictInsRangeStart w:id="{}" w:author="Range Insert"/><w:customXml w:uri="urn:inert" w:element="added"><w:r><w:t>range insertion</w:t></w:r></w:customXml><c:customXmlConflictInsRangeEnd w:id="{}"/><c:customXmlConflictDelRangeStart w:id="{}" w:author="Range Delete"/><w:customXml w:uri="urn:inert" w:element="deleted"><w:r><w:t>range deletion</w:t></w:r></w:customXml><c:customXmlConflictDelRangeEnd w:id="{}"/></w:p><w:tbl {declaration}><w:tr><w:trPr><c:conflictIns w:id="{}" w:author="Property Insert"/><c:conflictDel w:id="{}" w:author="Property Delete"/></w:trPr><w:tc><w:p/></w:tc></w:tr></w:tbl>"#,
        id + 1,
        id + 2,
        id + 2,
        id + 3,
        id + 3,
        id + 4,
        id + 5
    )
}
fn conflict_snapshot(p: &Package, part: &str) -> conflict::Snapshot {
    p.conflict_stories()
        .unwrap()
        .into_iter()
        .find(|s| s.part() == &uri(part))
        .unwrap()
        .snapshot()
        .clone()
}
fn conflicts_publish(p: &Package, c: &conflict::Commit, part: &str) -> Package {
    let q = reopen(c.package());
    only_changed(p, &q, part);
    assert!(members(&p.apply_conflicts(c.patch()).unwrap().into_package()) == members(&q));
    assert!(
        members(
            &q.apply_conflicts(&c.patch().inverse())
                .unwrap()
                .into_package()
        ) == members(p)
    );
    q
}
fn without_annotations(s: &conflict::Snapshot) -> Vec<u8> {
    let mut cuts = Vec::new();
    for c in &s.inventory().conflicts {
        if c.scope == conflict::Scope::Inline {
            cuts.push((c.start_tag.start(), c.start_tag.end()));
            cuts.push((c.content.end(), c.span.end()));
        } else {
            cuts.push((c.span.start(), c.span.end()));
        }
    }
    for r in &s.inventory().ranges {
        cuts.push((r.start_span.start(), r.start_span.end()));
        cuts.push((r.end_span.start(), r.end_span.end()));
    }
    cuts.sort();
    let mut from = 0;
    let mut result = Vec::new();
    for (start, end) in cuts {
        assert!(start >= from);
        result.extend_from_slice(&s.source()[from..start]);
        from = end;
    }
    result.extend_from_slice(&s.source()[from..]);
    result
}

#[test]
fn native_docx_conflict_metadata_edits_and_annotation_removal_preserve_story_payloads() {
    use conflict::{Kind, Scope};
    for file in PARENTS {
        let bytes = native(file);
        let base = Package::from_bytes(&bytes).unwrap();
        assert!(base.conflicts().unwrap().inventory().conflicts.is_empty());
        let p = append(&base, MAIN, &conflict_fragment(75));
        let p = if file == PARENTS[0] {
            append(&p, HEADER, &conflict_fragment(175))
        } else {
            p
        };
        let parts = if file == PARENTS[0] {
            vec![(MAIN, 75), (HEADER, 175)]
        } else {
            vec![(MAIN, 75)]
        };
        for (part, id) in parts {
            let snapshot = conflict_snapshot(&p, part);
            let inventory = snapshot.inventory();
            assert_eq!(
                inventory
                    .conflicts
                    .iter()
                    .map(|c| (c.kind, c.scope, c.metadata.id.get()))
                    .collect::<Vec<_>>(),
                [
                    (Kind::Insert, Scope::Inline, id),
                    (Kind::Delete, Scope::Inline, id + 1),
                    (Kind::Insert, Scope::Property, id + 4),
                    (Kind::Delete, Scope::Property, id + 5)
                ]
            );
            assert_eq!(
                inventory
                    .ranges
                    .iter()
                    .map(|r| (r.kind, r.metadata.id.get()))
                    .collect::<Vec<_>>(),
                [(Kind::Insert, id + 2), (Kind::Delete, id + 3)]
            );
            let span = inventory.conflicts[0].text_extent().unwrap();
            assert_eq!(
                &snapshot.source()[span.start()..span.end()],
                "added 日本語 &amp; kept".as_bytes()
            );
            let mut e = p.edit_story_conflicts(&uri(part));
            e.set_conflict_author(0, "Reviewed 日本語 & <saved>")
                .unwrap();
            e.set_conflict_date(0, Some("2026-11-07T12:34:56+08:00".into()))
                .unwrap();
            e.set_range_author(1, "Range reviewer 🍎").unwrap();
            let c = e.commit().unwrap();
            let changed = conflicts_publish(&p, &c, part);
            let after = conflict_snapshot(&changed, part);
            assert_eq!(
                after.inventory().conflicts[0].metadata.author,
                "Reviewed 日本語 & <saved>"
            );
            assert_eq!(
                after.inventory().conflicts[0].metadata.date.as_deref(),
                Some("2026-11-07T12:34:56+08:00")
            );
            assert_eq!(
                after.inventory().ranges[1].metadata.author,
                "Range reviewer 🍎"
            );
            assert_eq!(without_annotations(&snapshot), without_annotations(&after));
            assert_eq!(
                changed.document().unwrap().text().unwrap(),
                p.document().unwrap().text().unwrap()
            );
            assert!(matches!(
                changed.apply_conflicts(c.patch()),
                Err(Error::PatchConflict { .. })
            ));
            let mut e = changed.edit_story_conflicts(&uri(part));
            e.set_conflict_author(0, "Reviewed 日本語 & <saved>")
                .unwrap();
            assert!(e.commit().unwrap().patch().is_noop());
            let mut e = changed.edit_story_conflicts(&uri(part));
            for index in 0..after.inventory().conflicts.len() {
                e.remove_conflict(index).unwrap();
            }
            for index in 0..after.inventory().ranges.len() {
                e.remove_range(index).unwrap();
            }
            let c = e.commit().unwrap();
            let removed = conflicts_publish(&changed, &c, part);
            let remaining = conflict_snapshot(&removed, part);
            assert!(
                remaining.inventory().conflicts.is_empty()
                    && remaining.inventory().ranges.is_empty()
            );
            assert_eq!(remaining.source(), without_annotations(&after));
        }
        assert_eq!(base.to_bytes().unwrap(), bytes);
        assert_eq!(native(file), bytes);
    }
}

#[test]
fn native_docx_conflict_invalid_metadata_is_atomic_and_unpaired_ranges_are_rejected() {
    let bytes = native(PARENTS[0]);
    let base = Package::from_bytes(&bytes).unwrap();
    let p = append(&base, MAIN, &conflict_fragment(75));
    let mut e = p.edit_conflicts();
    e.set_conflict_author(0, "valid staged author").unwrap();
    assert!(
        matches!(e.set_conflict_date(0,Some("not-a-date".into())),Err(Error::Invalid(m)) if m=="conflict date is not a valid xsd:dateTime lexical value")
    );
    assert!(
        matches!(e.set_range_author(0,"bad\0author"),Err(Error::Invalid(m)) if m=="conflict author contains an XML-forbidden character")
    );
    let c = e.commit().unwrap();
    let q = conflicts_publish(&p, &c, MAIN);
    assert_eq!(
        q.conflicts().unwrap().inventory().conflicts[0]
            .metadata
            .author,
        "valid staged author"
    );
    let source = String::from_utf8(members(&p)[MAIN].clone()).unwrap();
    let malformed = source.replace("<c:customXmlConflictInsRangeEnd w:id=\"77\"/>", "");
    assert_ne!(malformed, source);
    let malformed = rewrite(&p, MAIN, malformed.into_bytes());
    assert!(
        matches!(malformed.conflicts(),Err(Error::Invalid(m)) if m=="orphaned conflict range start marker")
    );
    assert_eq!(base.to_bytes().unwrap(), bytes);
    assert_eq!(native(PARENTS[0]), bytes);
}

fn donor() -> (Vec<u8>, Vec<u8>) {
    let p = OpcPackage::from_bytes(&native(DONOR), SourceLimits::default()).unwrap();
    let ink = p
        .get_part(&PackURI::new("/ppt/ink/ink1.xml").unwrap())
        .unwrap()
        .blob()
        .to_vec();
    let image = p
        .get_part(&PackURI::new("/ppt/media/image1.png").unwrap())
        .unwrap()
        .blob()
        .to_vec();
    assert_eq!((ink.len(), image.len()), (2424, 1645));
    let sha = |bytes: &[u8]| {
        Sha256::digest(bytes)
            .iter()
            .map(|b| format!("{b:02x}"))
            .collect::<String>()
    };
    assert_eq!(
        sha(&ink),
        "153cd907463545144a3cdf362afbc52a6308f4db40a49c5590cfde5233dc883c"
    );
    assert_eq!(
        sha(&image),
        "f233c87a6418eb3f73de40c9a4ed3a93a45dbb5f3d0f0a5497dafb675d9fa183"
    );
    (ink, image)
}
fn native_ink(p: &Package) -> Package {
    let (ink, _) = donor();
    let mut e = p.edit_opc();
    let opc = e.package_mut();
    opc.add_part(Box::new(BlobPart::new(
        PackURI::new(NATIVE_INK).unwrap(),
        "application/inkml+xml".into(),
        ink,
    )));
    for name in [MAIN, HEADER] {
        let part = opc.get_part_mut(&uri(name)).unwrap();
        let text = String::from_utf8(part.blob().to_vec()).unwrap();
        let fragment = format!(
            r#"<w:r xmlns:w="{W}" xmlns:r="{REL}"><w:contentPart r:id="rIdNativeInk75"/></w:r>"#
        );
        assert!(text.contains("</w:p>"));
        part.set_blob(
            text.replacen("</w:p>", &format!("{fragment}</w:p>"), 1)
                .into_bytes(),
        );
        assert!(part.rels().get("rIdNativeInk75").is_none());
        part.rels_mut().add_relationship(
            format!("{REL}/customXml"),
            "ink/native-office-75.xml".into(),
            "rIdNativeInk75".into(),
            false,
        );
    }
    e.commit().unwrap().into_package()
}
fn prepared() -> ink::Prepared {
    ink::Draft::default()
        .context(
            ink::ContextDraft::new(ink::ContextKind::InkDrawing)
                .with_xml_id("ctx75")
                .unwrap(),
        )
        .unwrap()
        .brush(ink::BrushDraft::new("br75").unwrap())
        .unwrap()
        .trace(
            ink::TraceDraft::new("0 0, 10 20, -5 30")
                .unwrap()
                .with_context_ref("#ctx75")
                .unwrap()
                .with_brush_ref("#br75")
                .unwrap(),
        )
        .unwrap()
        .trace(
            ink::TraceDraft::new("5 6, 7 8")
                .unwrap()
                .with_context_ref("#ctx75")
                .unwrap()
                .with_brush_ref("#br75")
                .unwrap(),
        )
        .unwrap()
        .finish()
        .unwrap()
}
fn ink_publish(p: &Package, c: &ink::Commit) -> Package {
    let q = reopen(c.package());
    equivalent(&p.apply_ink(c.patch()).unwrap().into_package(), &q);
    equivalent(
        &q.apply_ink(&c.patch().inverse()).unwrap().into_package(),
        p,
    );
    assert_eq!(
        p.document().unwrap().text().unwrap(),
        q.document().unwrap().text().unwrap()
    );
    q
}

#[test]
fn native_office_ink_payload_shared_between_word_stories_is_read_cloned_and_collected() {
    use ink::BrushPropertyName as Brush;
    use litchi_docx::package::story::StoryKind;
    let bytes = native(PARENTS[0]);
    let base = Package::from_bytes(&bytes).unwrap();
    assert!(base.ink().unwrap().annotations().is_empty());
    let p = reopen(&native_ink(&base));
    let snapshot = p.ink().unwrap();
    assert_eq!(snapshot.annotations().len(), 2);
    assert_eq!(snapshot.distinct_payload_count(), 1);
    assert_eq!(
        snapshot
            .annotations()
            .iter()
            .map(|a| a.location().kind())
            .collect::<Vec<_>>(),
        [StoryKind::Main, StoryKind::Header]
    );
    for a in snapshot.annotations() {
        assert_eq!(a.trace_count(), 1);
        assert_eq!(
            a.brush_properties()
                .map(|p| (
                    p.name().clone(),
                    p.value().to_owned(),
                    p.units().map(str::to_owned)
                ))
                .collect::<Vec<_>>(),
            [
                (Brush::Width, "0.035".into(), Some("cm".into())),
                (Brush::Height, "0.035".into(), Some("cm".into())),
                (Brush::Color, "#E71224".into(), None)
            ]
        );
    }
    let payload = prepared();
    let mut e = p.edit_ink().unwrap();
    assert!(e.replace(Position::new(0), payload.clone(), None).unwrap());
    assert!(!e.replace(Position::new(0), payload, None).unwrap());
    assert!(
        matches!(e.replace(Position::new(99),prepared(),None),Err(Error::Invalid(m)) if m=="DOCX Ink annotation selector is out of range")
    );
    let c = e.commit().unwrap();
    let changed = ink_publish(&p, &c);
    assert_eq!(
        changed
            .ink()
            .unwrap()
            .annotations()
            .iter()
            .map(|a| a.trace_count())
            .collect::<Vec<_>>(),
        [2, 1]
    );
    assert_eq!(changed.ink().unwrap().distinct_payload_count(), 2);
    assert_eq!(
        changed
            .opc_package()
            .get_part(&PackURI::new(NATIVE_INK).unwrap())
            .unwrap()
            .blob(),
        donor().0
    );
    assert_eq!(members(&changed)[HEADER], members(&p)[HEADER]);
    assert!(matches!(
        changed.apply_ink(c.patch()),
        Err(Error::PatchConflict {
            operation: "apply_ink",
            reason: "the package does not hold the Ink patch source"
        })
    ));
    let mut e = changed.edit_ink().unwrap();
    assert!(e.remove(Position::new(0)).unwrap());
    let c = e.commit().unwrap();
    let one = ink_publish(&changed, &c);
    assert_eq!(one.ink().unwrap().annotations().len(), 1);
    assert_eq!(one.ink().unwrap().annotations()[0].trace_count(), 1);
    assert_eq!(
        one.opc_package()
            .get_part(&PackURI::new(NATIVE_INK).unwrap())
            .unwrap()
            .blob(),
        donor().0
    );
    let mut e = one.edit_ink().unwrap();
    assert!(e.remove(Position::new(0)).unwrap());
    let c = e.commit().unwrap();
    let removed = ink_publish(&one, &c);
    assert!(removed.ink().unwrap().annotations().is_empty());
    assert!(
        removed
            .opc_package()
            .get_part(&PackURI::new(NATIVE_INK).unwrap())
            .is_err()
    );
    assert_eq!(base.to_bytes().unwrap(), bytes);
    assert_eq!(native(PARENTS[0]), bytes);
}

#[test]
fn native_docx_ink_authoring_preserves_word_variants_and_exact_fallback_image() {
    let (_, image) = donor();
    for file in PARENTS {
        let bytes = native(file);
        let p = Package::from_bytes(&bytes).unwrap();
        assert!(p.ink().unwrap().annotations().is_empty());
        for style in [
            ink::Style::Base(ink::BaseProfile::InkContent),
            ink::Style::Base(ink::BaseProfile::WordTextXml),
            ink::Style::drawing(
                ink::Placement::Inline(ink::Geometry::new(914400, 914400).unwrap()),
                ink::FallbackImage::from_bytes(image.clone()).unwrap(),
            ),
        ] {
            let drawing = style.placement().is_some();
            let payload = prepared();
            let mut e = p.edit_ink().unwrap();
            e.insert(
                ink::Destination::main(Position::new(0)),
                payload.clone(),
                style,
            )
            .unwrap();
            let c = e.commit().unwrap();
            let q = ink_publish(&p, &c);
            let annotations = q.ink().unwrap();
            assert_eq!(annotations.annotations().len(), 1);
            assert_eq!(annotations.annotations()[0].trace_count(), 2);
            assert_eq!(
                annotations.annotations()[0]
                    .contexts()
                    .next()
                    .unwrap()
                    .kind(),
                &ink::ContextKind::InkDrawing
            );
            let mut payloads = 0;
            let mut images = 0;
            for part in q.opc_package().iter_parts() {
                let data = q.opc_package().get_part(part.partname()).unwrap().blob();
                if data == payload.as_bytes() {
                    payloads += 1;
                }
                if data == image {
                    images += 1;
                }
            }
            assert_eq!(payloads, 1);
            assert_eq!(images, usize::from(drawing));
            let a = members(&p);
            let b = members(&q);
            let at = a[MAIN]
                .windows(b"</w:p>".len())
                .position(|w| w == b"</w:p>")
                .unwrap();
            assert!(b[MAIN].starts_with(&a[MAIN][..at]));
            assert!(b[MAIN].ends_with(&a[MAIN][at..]));
            for (name, data) in a {
                if ![MAIN, "word/_rels/document.xml.rels", "[Content_Types].xml"]
                    .contains(&name.as_str())
                {
                    assert!(b.get(&name) == Some(&data), "changed native member {name}");
                }
            }
            let mut e = q.edit_ink().unwrap();
            assert!(!e.replace(Position::new(0), payload, None).unwrap());
            assert!(e.commit().unwrap().patch().is_noop());
            let mut e = q.edit_ink().unwrap();
            assert!(e.remove(Position::new(0)).unwrap());
            let c = e.commit().unwrap();
            let removed = ink_publish(&q, &c);
            assert!(removed.ink().unwrap().annotations().is_empty());
            assert_eq!(
                members(&removed).keys().collect::<Vec<_>>(),
                members(&p).keys().collect::<Vec<_>>()
            );
        }
        assert_eq!(p.to_bytes().unwrap(), bytes);
        assert_eq!(native(file), bytes);
    }
}
