//! Explicit extension derivatives of native Word sources; no XPath, calendar or add-in execution.
use litchi_docx::{
    Error, Package,
    content_control::{self as sdt, FormattingAllowed as Allowed, PackageChecksumStatus as Status},
};
use litchi_opc::PackURI;
use std::{collections::BTreeMap, path::Path};
const MAIN: &str = "word/document.xml";
const HEADER: &str = "word/header1.xml";
const ID: &str = "{63000000-1111-4111-8111-111111111111}";
const W: &str = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
const MC: &str = "http://schemas.openxmlformats.org/markup-compatibility/2006";
const HASH: &str = "http://schemas.microsoft.com/office/word/2020/wordml/sdtdatahash";
const FORMAT: &str = "http://schemas.microsoft.com/office/word/2024/wordml/sdtformatlock";
fn native(name: &str) -> Vec<u8> {
    std::fs::read(
        Path::new(env!("CARGO_MANIFEST_DIR"))
            .join("../../test-data/microsoft-office/word")
            .join(name),
    )
    .unwrap()
}
fn members(p: &Package) -> BTreeMap<String, Vec<u8>> {
    let r = litchi_opc::phys_pkg::OwnedPhysPkgReader::from_bytes(
        p.to_bytes().unwrap(),
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
fn uri(part: &str) -> PackURI {
    PackURI::new(format!("/{part}")).unwrap()
}
fn rewrite(p: &Package, part: &str, xml: Vec<u8>) -> Package {
    let mut e = p.edit_opc();
    e.package_mut()
        .get_part_mut(&uri(part))
        .unwrap()
        .set_blob(xml);
    e.commit().unwrap().into_package()
}
fn append(p: &Package, part: &str, fragment: &str) -> Package {
    let source = String::from_utf8(members(p)[part].clone()).unwrap();
    let marker = if part == MAIN {
        "<w:sectPr"
    } else {
        "</w:hdr>"
    };
    let at = source.rfind(marker).unwrap();
    let mut xml = source[..at].to_owned();
    xml.push_str(fragment);
    xml.push_str(&source[at..]);
    rewrite(p, part, xml.into_bytes())
}
fn control(id: u32, checksum: Option<&str>, lock: &str) -> String {
    let hash = checksum.map_or(String::new(), |c| format!(" h:storeItemChecksum=\"{c}\""));
    format!(
        r#"<w:sdt xmlns:w="{W}" xmlns:mc="{MC}" xmlns:h="{HASH}" xmlns:f="{FORMAT}" mc:Ignorable="h f"><w:sdtPr><w:id w:val="{id}"/><w:lock w:val="{lock}"/><w:dataBinding w:xpath="/r" w:storeItemID="{ID}"{hash}/><w:text/></w:sdtPr><w:sdtContent/></w:sdt>"#
    )
}
fn prepared(file: &str, header: bool) -> (Package, Package) {
    let p = Package::from_bytes(&native(file)).unwrap();
    let mut e = p.edit_custom_xml();
    e.add(litchi_docx::custom_xml::NewStore {
        xml: b"<r/>".to_vec(),
        content_type: "application/xml".into(),
        id: ID.into(),
        schemas: vec![],
        conformance: litchi_ooxml_common::custom_xml::Conformance::Transitional,
    })
    .unwrap();
    let q = e.commit().unwrap().into_package();
    let q = append(
        &q,
        MAIN,
        &(control(63, Some("AsOd4Q=="), "contentLocked") + &control(64, None, "unlocked")),
    );
    let q = if header {
        append(
            &q,
            HEADER,
            &control(65, Some("AsOd4Q=="), "sdtContentLocked"),
        )
    } else {
        q
    };
    assert_eq!(
        q.document().unwrap().text().unwrap(),
        p.document().unwrap().text().unwrap()
    );
    (p, q)
}
fn strip_controls(bytes: &[u8]) -> Vec<u8> {
    use quick_xml::events::Event;
    let mut reader = quick_xml::Reader::from_reader(bytes);
    let mut out = Vec::new();
    let mut depth = 0;
    loop {
        let start = reader.buffer_position() as usize;
        let event = reader.read_event().unwrap();
        let end = reader.buffer_position() as usize;
        match event {
            Event::Eof => break,
            Event::Start(ref e) if e.name().as_ref() == b"w:sdt" => {
                assert_eq!(depth, 0);
                depth = 1;
            },
            Event::Start(_) if depth > 0 => depth += 1,
            Event::End(_) if depth > 0 => depth -= 1,
            _ if depth == 0 => out.extend_from_slice(&bytes[start..end]),
            _ => {},
        }
    }
    assert_eq!(depth, 0);
    out
}
fn preserve(p: &Package, q: &Package, parts: &[&str]) {
    let a = members(p);
    let b = members(q);
    assert_eq!(a.keys().collect::<Vec<_>>(), b.keys().collect::<Vec<_>>());
    for (n, v) in &a {
        if !parts.contains(&n.as_str()) {
            assert!(b.get(n) == Some(v), "changed {n}");
        } else if n == MAIN || n == HEADER {
            assert!(strip_controls(v) == strip_controls(&b[n]));
        }
    }
    assert_eq!(
        p.document().unwrap().text().unwrap(),
        q.document().unwrap().text().unwrap()
    );
}
fn publish(p: &Package, c: &sdt::Commit) -> Package {
    let file = tempfile::NamedTempFile::with_suffix(".docx").unwrap();
    c.package().save(file.path()).unwrap();
    let r = Package::open(file.path()).unwrap();
    assert!(members(&r) == members(c.package()));
    assert!(members(&p.apply_content_controls(c.patch()).unwrap().into_package()) == members(&r));
    assert!(
        members(
            &r.apply_content_controls(&c.patch().inverse())
                .unwrap()
                .into_package()
        ) == members(p)
    );
    r
}
fn formatting(p: &Package, at: usize) -> Option<Allowed> {
    let s = p.content_control_snapshot().unwrap();
    s.stories()
        .iter()
        .find(|s| s.part().as_str() == format!("/{MAIN}"))
        .unwrap()
        .snapshot()
        .inventory()
        .occurrences()[at]
        .control()
        .formatting_allowed()
}
#[test]
fn native_checksum_derivatives_refresh_declared_bindings_across_body_and_header() {
    let (_, p) = prepared("word-tables-layout.docx", true);
    let before = members(&p);
    let report = p.verify_content_control_checksums().unwrap();
    assert_eq!(report.len(), 3);
    assert_eq!(
        report
            .iter()
            .map(|r| (r.control_id().unwrap(), r.status().clone()))
            .collect::<Vec<_>>(),
        [
            (63, Status::Matches),
            (64, Status::Absent),
            (65, Status::Matches)
        ]
    );
    let store = report[0]
        .store_part()
        .unwrap()
        .as_str()
        .trim_start_matches('/')
        .to_owned();
    assert_eq!(before[&store], b"<r/>");
    let mut e = p.edit_custom_xml();
    e.set(ID, b"<r />".to_vec()).unwrap();
    let c = e.commit().unwrap();
    let file = tempfile::NamedTempFile::with_suffix(".docx").unwrap();
    c.package().save(file.path()).unwrap();
    let q = Package::open(file.path()).unwrap();
    preserve(&p, &q, &[MAIN, HEADER, &store]);
    let after = members(&q);
    let changed = before
        .keys()
        .filter(|n| before.get(*n) != after.get(*n))
        .cloned()
        .collect::<Vec<_>>();
    let mut expected = vec![store.clone(), MAIN.into(), HEADER.into()];
    expected.sort();
    assert_eq!(changed, expected);
    for part in [MAIN, HEADER] {
        assert_eq!(
            std::str::from_utf8(&after[part])
                .unwrap()
                .matches("hS38mw==")
                .count(),
            1
        );
    }
    let report = q.verify_content_control_checksums().unwrap();
    assert_eq!(
        report
            .iter()
            .map(|r| r.status().clone())
            .collect::<Vec<_>>(),
        [Status::Matches, Status::Absent, Status::Matches]
    );
    assert!(
        members(
            &q.apply_custom_xml(&c.patch().inverse())
                .unwrap()
                .into_package()
        ) == before
    );
    assert!(members(&p) == before);
}
#[test]
fn native_formatting_allowed_lifecycle_preserves_prior_staging_and_exact_output_limits() {
    for file in ["word-text-structure.docx", "word-recorded-macro.docm"] {
        let (_, p) = prepared(file, false);
        let before = members(&p);
        assert_eq!(formatting(&p, 0), None);
        let mut e = p.edit_content_controls();
        e.set_formatting_allowed(&uri(MAIN), 0, Some(Allowed::Allowed))
            .unwrap();
        assert!(
            matches!(e.set_formatting_allowed(&uri(MAIN),1,Some(Allowed::Allowed)),Err(Error::Invalid(m)) if m=="formattingAllowed applies only to contentLocked or sdtContentLocked")
        );
        let c = e.commit().unwrap();
        let added = publish(&p, &c);
        assert_eq!(formatting(&added, 0), Some(Allowed::Allowed));
        assert_eq!(formatting(&added, 1), None);
        preserve(&p, &added, &[MAIN]);
        assert!(matches!(
            added.apply_content_controls(c.patch()),
            Err(Error::PatchConflict { .. })
        ));
        let expected = members(&added)[MAIN].len();
        for (maximum, pass) in [(expected, true), (expected - 1, false)] {
            let mut e = p.edit_content_controls_with_options(sdt::EditOptions::default().limits(
                sdt::PackageLimits {
                    max_output_bytes: maximum,
                    ..Default::default()
                },
            ));
            e.set_formatting_allowed(&uri(MAIN), 0, Some(Allowed::Allowed))
                .unwrap();
            let result = e.commit();
            if pass {
                assert!(result.is_ok());
            } else {
                assert!(
                    matches!(result,Err(Error::ContentControlLimit{kind:sdt::LimitKind::Output,actual,maximum:m}) if actual==expected && m==maximum)
                );
            }
        }
        let mut e = added.edit_content_controls();
        e.set_formatting_allowed(&uri(MAIN), 0, Some(Allowed::Disallowed))
            .unwrap();
        let c = e.commit().unwrap();
        let updated = publish(&added, &c);
        assert_eq!(formatting(&updated, 0), Some(Allowed::Disallowed));
        preserve(&added, &updated, &[MAIN]);
        let mut e = updated.edit_content_controls();
        e.set_formatting_allowed(&uri(MAIN), 0, None).unwrap();
        let c = e.commit().unwrap();
        let removed = publish(&updated, &c);
        assert_eq!(formatting(&removed, 0), None);
        preserve(&updated, &removed, &[MAIN]);
        assert!(members(&p) == before);
    }
}
#[test]
fn native_checksum_reports_keep_absence_malformed_missing_and_mismatch_distinct() {
    let (_, p) = prepared("word-text-structure.docx", false);
    let original = members(&p);
    let xml = String::from_utf8(original[MAIN].clone()).unwrap();
    for (old, new, expected) in [
        ("AsOd4Q==", "not-base64", 0),
        (ID, "{64000000-1111-4111-8111-111111111111}", 1),
        ("AsOd4Q==", "AAAAAA==", 2),
    ] {
        let q = rewrite(&p, MAIN, xml.replacen(old, new, 1).into_bytes());
        let r = q.verify_content_control_checksums().unwrap();
        assert_eq!(r.len(), 2);
        assert_eq!(r[1].status(), &Status::Absent);
        match expected {
            0 => assert!(matches!(r[0].status(),Status::Malformed(v) if v.as_ref()=="not-base64")),
            1 => assert_eq!(r[0].status(), &Status::MissingStore),
            _ => assert!(
                matches!(r[0].status(),Status::Mismatch{expected,actual} if expected.word_value()==0 && actual.word_value()==0xe19d_c302)
            ),
        }
        let before = members(&q);
        let mut e = q.edit_content_controls();
        if expected == 1 {
            assert!(matches!(e.refresh_checksums(), Err(Error::PartNotFound(_))));
            assert!(e.commit().unwrap().patch().is_noop());
        } else {
            e.refresh_checksums().unwrap();
            let c = e.commit().unwrap();
            let fixed = publish(&q, &c);
            assert_eq!(
                fixed.verify_content_control_checksums().unwrap()[0].status(),
                &Status::Matches
            );
            preserve(&q, &fixed, &[MAIN]);
        }
        assert!(members(&q) == before);
    }
    assert!(members(&p) == original);
}
#[test]
fn native_calendar_and_web_link_derivatives_keep_mce_choices_and_inert_boolean_precedence() {
    use sdt::{Calendar, Inventory, WebExtensionBinding};
    let p = Package::from_bytes(&native("word-text-structure.docx")).unwrap();
    let before = members(&p);
    let mut controls = String::new();
    for (id, ns) in [
        (71, "http://schemas.microsoft.com/office/word/2010/wordml"),
        (72, "urn:unrecognized-choice"),
    ] {
        controls.push_str(&format!(r#"<w:sdt xmlns:w="{W}" xmlns:mc="{MC}" xmlns:c="{ns}" xmlns:x="http://schemas.microsoft.com/office/word/2012/wordml" mc:Ignorable="c x"><w:sdtPr><w:id w:val="{id}"/><x:webExtensionLinked w:val="true"/><x:webExtensionCreated w:val="false"/><w:date><mc:AlternateContent><mc:Choice Requires="c"><w:calendar w:val="umalqura"/></mc:Choice><mc:Fallback><w:calendar w:val="hijri"/></mc:Fallback></mc:AlternateContent></w:date></w:sdtPr><w:sdtContent/></w:sdt>"#));
    }
    let q = append(&p, MAIN, &controls);
    let m = members(&q);
    let parsed = Inventory::parse(&m[MAIN], &Default::default()).unwrap();
    assert_eq!(parsed.occurrences().len(), 2);
    for (item, calendar) in parsed
        .occurrences()
        .iter()
        .zip([Calendar::Umalqura, Calendar::Hijri])
    {
        let c = item.control();
        assert_eq!(c.date_calendar(), Some(&calendar));
        assert_eq!(c.web_extension_linked(), Some(true));
        assert_eq!(c.web_extension_created(), Some(false));
        assert_eq!(
            c.web_extension_binding(),
            Some(WebExtensionBinding::Created(false))
        );
    }
    assert!(strip_controls(&m[MAIN]) == before[MAIN]);
    let f = tempfile::NamedTempFile::with_suffix(".docx").unwrap();
    q.save(f.path()).unwrap();
    assert!(members(&Package::open(f.path()).unwrap()) == m);
    assert!(members(&p) == before);
}

#[test]
fn native_negative_footer_control_id_is_conformant_and_does_not_block_unrelated_store_edits() {
    let bytes = native("word-tables-layout.docx");
    let p = Package::from_bytes(&bytes).unwrap();
    let snapshot = p.content_control_snapshot().unwrap();
    let footer = snapshot
        .stories()
        .iter()
        .find(|s| s.part().as_str() == "/word/footer2.xml")
        .unwrap();
    let controls = footer.snapshot().inventory().occurrences();
    assert_eq!(controls.len(), 1);
    assert_eq!(
        controls[0].control().id_opt().map(i64::from),
        Some(-1_419_789_921)
    );
    assert!(footer.snapshot().inventory().tolerance_report().is_clean());
    let before = members(&p);
    let mut e = p.edit_custom_xml();
    e.add(litchi_docx::custom_xml::NewStore {
        xml: b"<r/>".to_vec(),
        content_type: "application/xml".into(),
        id: ID.into(),
        schemas: vec![],
        conformance: litchi_ooxml_common::custom_xml::Conformance::Transitional,
    })
    .unwrap();
    let c = e.commit().unwrap();
    let after = members(c.package());
    for (name, data) in before {
        if name.ends_with(".xml") && name.starts_with("word/") {
            assert!(
                after.get(&name) == Some(&data),
                "changed native story {name}"
            );
        }
    }
    assert!(p.to_bytes().unwrap() == bytes);
}
