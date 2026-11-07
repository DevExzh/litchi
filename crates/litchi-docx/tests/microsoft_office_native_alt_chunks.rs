//! Opaque AltChunk derivatives of native parents; no foreign-content import or execution.
use litchi_docx::{
    Package,
    alt::{Data, Import, Kind, Target},
};
use litchi_opc::OpcPackage;
use std::{collections::BTreeMap, io::Cursor, path::Path};
const HTML:&[u8]=br#"<html><body><p>FOREIGN_PAYLOAD_060</p><img src="https://example.invalid/never-loaded.png"></body></html>"#;
const LINK: &str = "https://example.invalid/never-imported.html";
fn native(name: &str) -> Vec<u8> {
    std::fs::read(
        Path::new(env!("CARGO_MANIFEST_DIR"))
            .join("../../test-data/microsoft-office/word")
            .join(name),
    )
    .unwrap()
}
fn members(bytes: &[u8]) -> BTreeMap<String, Vec<u8>> {
    let r = litchi_opc::phys_pkg::OwnedPhysPkgReader::from_bytes(
        bytes.to_vec(),
        litchi_opc::SourceLimits::default(),
    )
    .unwrap();
    r.member_names()
        .unwrap()
        .into_iter()
        .map(|n| {
            let b = r.read_member(&n).unwrap();
            (n, b)
        })
        .collect()
}
type Relation = (String, String, String, String);
type Graph = BTreeMap<String, (String, Vec<Relation>)>;
fn graph(p: &OpcPackage) -> Graph {
    fn rels(r: &litchi_opc::Relationships) -> Vec<Relation> {
        let mut v = r
            .iter()
            .map(|r| {
                (
                    r.r_id().to_string(),
                    r.reltype().to_string(),
                    r.target_ref().to_string(),
                    format!("{:?}", r.target_mode()),
                )
            })
            .collect::<Vec<_>>();
        v.sort();
        v
    }
    let mut g = BTreeMap::from([("/".into(), (String::new(), rels(p.rels())))]);
    for part in p.iter_parts() {
        g.insert(
            part.partname().as_str().into(),
            (part.content_type().into(), rels(part.rels())),
        );
    }
    g
}
fn equivalent(a: &Package, b: &Package) {
    assert_eq!(graph(a.opc_package()), graph(b.opc_package()));
    for part in a.opc_package().iter_parts() {
        assert!(
            a.opc_package().get_part(part.partname()).unwrap().blob()
                == b.opc_package().get_part(part.partname()).unwrap().blob(),
            "changed {}",
            part.partname()
        );
    }
}
fn save(p: &Package) -> Package {
    let f = tempfile::NamedTempFile::with_suffix(".docx").unwrap();
    p.save(f.path()).unwrap();
    Package::open(f.path()).unwrap()
}
fn published(p: &Package, c: &litchi_docx::alt::Commit) -> Package {
    let r = save(c.package());
    assert!(
        members(
            &p.apply_alt_chunks(c.patch())
                .unwrap()
                .package()
                .to_bytes()
                .unwrap()
        ) == members(&r.to_bytes().unwrap())
    );
    let inverse = r
        .apply_alt_chunks(&c.patch().inverse())
        .unwrap()
        .into_package();
    equivalent(&inverse, p);
    let a = members(&p.to_bytes().unwrap());
    let b = members(&inverse.to_bytes().unwrap());
    assert_eq!(a.keys().collect::<Vec<_>>(), b.keys().collect::<Vec<_>>());
    let changed = a
        .keys()
        .filter(|name| a.get(*name) != b.get(*name))
        .cloned()
        .collect::<Vec<_>>();
    for name in changed {
        assert!(
            matches!(
                name.as_str(),
                "[Content_Types].xml" | "word/_rels/document.xml.rels"
            ),
            "unexpected inverse change {name}"
        );
    }
    r
}
fn strip_alts(bytes: &[u8]) -> (Vec<u8>, usize) {
    use quick_xml::events::Event;
    let mut reader = quick_xml::Reader::from_reader(bytes);
    let mut out = Vec::new();
    let mut depth = 0;
    let mut count = 0;
    loop {
        let start = reader.buffer_position() as usize;
        let event = reader.read_event().unwrap();
        let end = reader.buffer_position() as usize;
        match event {
            Event::Eof => break,
            Event::Start(ref e) if e.name().as_ref() == b"w:altChunk" => {
                assert_eq!(depth, 0);
                depth = 1;
                count += 1;
            },
            Event::Empty(ref e) if e.name().as_ref() == b"w:altChunk" => {
                assert_eq!(depth, 0);
                count += 1;
            },
            Event::Start(_) if depth > 0 => depth += 1,
            Event::End(_) if depth > 0 => depth -= 1,
            _ if depth == 0 => out.extend_from_slice(&bytes[start..end]),
            _ => {},
        }
    }
    assert_eq!(depth, 0);
    (out, count)
}
fn preserve_native(base: &Package, derived: &Package, alts: usize) {
    let a = members(&base.to_bytes().unwrap());
    let b = members(&derived.to_bytes().unwrap());
    for (n, data) in &a {
        if !matches!(
            n.as_str(),
            "[Content_Types].xml" | "word/document.xml" | "word/_rels/document.xml.rels"
        ) {
            assert!(b.get(n) == Some(data), "changed {n}");
        }
    }
    let (xml, count) = strip_alts(&b["word/document.xml"]);
    assert_eq!(count, alts);
    assert!(xml == a["word/document.xml"]);
    assert_eq!(
        derived.document().unwrap().text().unwrap(),
        base.document().unwrap().text().unwrap()
    );
    assert!(
        !derived
            .document()
            .unwrap()
            .text()
            .unwrap()
            .contains("FOREIGN_PAYLOAD_060")
    );
}
fn with_html(base: &Package) -> Package {
    let mut e = base.edit_alt_chunks();
    e.add(Import::data(Data::Html(HTML.to_vec())), None)
        .unwrap();
    published(base, &e.commit().unwrap())
}

#[test]
fn native_document_and_macro_alt_chunk_lifecycle_keeps_foreign_content_opaque() {
    for file in ["word-text-structure.docx", "word-recorded-macro.docm"] {
        let original = native(file);
        let base = Package::from_reader(Cursor::new(&original)).unwrap();
        assert!(base.document().unwrap().alts().unwrap().is_empty());
        let mut e = base.edit_alt_chunks();
        let html = e
            .add(Import::data(Data::Html(HTML.to_vec())), Some(false))
            .unwrap();
        let link = e.add(Import::link(LINK).unwrap(), None).unwrap();
        let text = e
            .add(
                Import::data(Data::Text(b"FOREIGN_PAYLOAD_060 text".to_vec())),
                None,
            )
            .unwrap();
        let p = published(&base, &e.commit().unwrap());
        preserve_native(&base, &p, 3);
        let d = p.document().unwrap();
        assert_eq!(
            d.alts().unwrap(),
            [html.clone(), link.clone(), text.clone()]
        );
        assert_eq!(d.resolve_alt(&html).unwrap().kind(), Kind::Html);
        assert_eq!(d.resolve_alt(&html).unwrap().bytes(), HTML);
        let html_part = d.resolve_alt(&html).unwrap().name().clone();
        assert!(matches!(d.alt_target(&link).unwrap(),Target::Link(uri) if uri==LINK));
        let mut e = p.edit_alt_chunks();
        e.move_to(2, 0).unwrap();
        let moved = published(&p, &e.commit().unwrap());
        assert_eq!(
            moved.document().unwrap().alts().unwrap(),
            [text, html.clone(), link]
        );
        preserve_native(&base, &moved, 3);
        let mut e = moved.edit_alt_chunks();
        let replacement = e
            .replace(
                1,
                Import::data(Data::Rtf(
                    br"{\rtf1 FOREIGN_PAYLOAD_060 replacement}".to_vec(),
                )),
                Some(true),
            )
            .unwrap();
        assert_eq!(replacement.relationship(), html.relationship());
        let updated = published(&moved, &e.commit().unwrap());
        // replace returns the removed anchor; its now-unreferenced target is
        // collected. Resolve the new anchor from the committed document.
        let fresh = updated.document().unwrap().alts().unwrap()[1].clone();
        assert_eq!(fresh.match_source(), Some(true));
        assert!(!updated.opc_package().contains_part(&html_part));
        assert!(
            updated
                .document()
                .unwrap()
                .resolve_alt(&replacement)
                .is_err()
        );
        assert_eq!(
            updated
                .document()
                .unwrap()
                .resolve_alt(&fresh)
                .unwrap()
                .kind(),
            Kind::Rtf
        );
        preserve_native(&base, &updated, 3);
        let mut e = updated.edit_alt_chunks();
        for index in (0..3).rev() {
            e.remove(index).unwrap();
        }
        let removed = published(&updated, &e.commit().unwrap());
        equivalent(&removed, &base);
        assert!(removed.document().unwrap().alts().unwrap().is_empty());
        assert!(native(file) == original);
    }
}

#[test]
fn native_alt_chunk_payload_source_overlay_changes_only_the_inert_target() {
    use litchi_docx::source_backed::{self, AltChunkSelector};
    for file in ["word-text-structure.docx", "word-recorded-macro.docm"] {
        let base = Package::from_bytes(&native(file)).unwrap();
        let p = with_html(&base);
        let input = p.to_bytes().unwrap();
        let name = p
            .document()
            .unwrap()
            .resolve_alt(&p.document().unwrap().alts().unwrap()[0])
            .unwrap()
            .name()
            .as_str()
            .trim_start_matches('/')
            .to_owned();
        let source = source_backed::Package::from_reader(Cursor::new(&input)).unwrap();
        let mut e = source.edit_alt_chunk_payload(AltChunkSelector::index(0));
        let replacement = b"<html><body>FOREIGN_PAYLOAD_060 replacement &amp; inert</body></html>";
        e.replace(Data::Html(replacement.to_vec())).unwrap();
        let c = e.commit().unwrap();
        let mut output = Vec::new();
        c.package().write_to(&mut output).unwrap();
        let a = members(&input);
        let b = members(&output);
        assert_eq!(a.keys().collect::<Vec<_>>(), b.keys().collect::<Vec<_>>());
        assert_eq!(
            a.keys()
                .filter(|n| a.get(*n) != b.get(*n))
                .cloned()
                .collect::<Vec<_>>(),
            std::slice::from_ref(&name)
        );
        assert_eq!(b[&name], replacement);
        let reopened = source_backed::Package::from_reader(Cursor::new(&output)).unwrap();
        let inverse = reopened
            .apply_alt_chunk_payload(&c.patch().inverse())
            .unwrap();
        let mut restored = Vec::new();
        inverse.package().write_to(&mut restored).unwrap();
        assert!(members(&restored) == a);
        let ordinary = Package::from_bytes(&output).unwrap();
        preserve_native(&base, &ordinary, 1);
        let mut e = source.edit_alt_chunk_payload(AltChunkSelector::index(0));
        e.replace(Data::Text(b"wrong media".to_vec())).unwrap();
        assert!(
            matches!(e.commit(), Err(litchi_docx::Error::Refused(refusal))
            if refusal.kind() == litchi_core::edit::RefusalKind::InvalidValue
                && refusal.reason() == "replacement media type family must match the existing altChunk target Part")
        );
        let mut untouched = Vec::new();
        source.write_to(&mut untouched).unwrap();
        assert!(untouched == input);
    }
}

#[test]
fn native_alt_chunk_bad_selectors_preserve_staging_and_stale_payloads_refuse() {
    let base = Package::from_bytes(&native("word-text-structure.docx")).unwrap();
    let mut e = base.edit_alt_chunks();
    let first = e
        .add(Import::data(Data::Html(HTML.to_vec())), None)
        .unwrap();
    assert!(matches!(
        e.remove(99),
        Err(litchi_docx::Error::OutOfBounds { .. })
    ));
    assert_eq!(e.chunks().unwrap(), [first]);
    let p = published(&base, &e.commit().unwrap());
    let input = p.to_bytes().unwrap();
    let source = litchi_docx::source_backed::Package::from_reader(Cursor::new(&input)).unwrap();
    let mut a =
        source.edit_alt_chunk_payload(litchi_docx::source_backed::AltChunkSelector::index(0));
    a.replace(Data::Html(b"first".to_vec())).unwrap();
    let a = a.commit().unwrap();
    let mut b =
        source.edit_alt_chunk_payload(litchi_docx::source_backed::AltChunkSelector::index(0));
    b.replace(Data::Html(b"different".to_vec())).unwrap();
    let b = b.commit().unwrap();
    assert!(matches!(
        b.package().apply_alt_chunk_payload(a.patch()),
        Err(litchi_docx::Error::PatchConflict {
            operation: "alternative-format payload",
            ..
        })
    ));
    let mut original = Vec::new();
    source.write_to(&mut original).unwrap();
    assert!(original == input);
}
