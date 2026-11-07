//! Whole styles-with-effects resources copied from native styles; no visual cascade or rendering.
use litchi_docx::{
    Error, Package,
    styles::effects::{self, Limits, Owner, Resource},
};
use std::{collections::BTreeMap, path::Path};
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
fn resource(p: &Package, part: &str, name: &str) -> Resource {
    let xml = String::from_utf8(members(p)[part].clone()).unwrap();
    assert_eq!(xml.matches("<w:name w:val=\"Normal\"/>").count(), 1);
    Resource::new(
        xml.replacen(
            "<w:name w:val=\"Normal\"/>",
            &format!("<w:name w:val=\"{name}\"/>"),
            1,
        )
        .into_bytes(),
        &Limits::DEFAULT,
    )
    .unwrap()
}
fn published(p: &Package, c: &effects::Commit) -> Package {
    let f = tempfile::NamedTempFile::with_suffix(".docx").unwrap();
    c.package().save(f.path()).unwrap();
    let r = Package::open(f.path()).unwrap();
    assert!(members(&r) == members(c.package()));
    assert!(
        members(
            &p.apply_styles_with_effects(c.patch())
                .unwrap()
                .into_package()
        ) == members(&r)
    );
    assert!(
        members(
            &r.apply_styles_with_effects(&c.patch().inverse())
                .unwrap()
                .into_package()
        ) == members(p)
    );
    assert_eq!(
        r.document().unwrap().text().unwrap(),
        p.document().unwrap().text().unwrap()
    );
    r
}
fn unchanged(p: &Package, r: &Package, allowed: &[&str]) {
    let rels = |links: &litchi_opc::Relationships| {
        let mut result = links
            .iter()
            .filter(|link| link.reltype() != effects::RELATIONSHIP_TYPE)
            .map(|link| {
                (
                    link.r_id().to_owned(),
                    link.reltype().to_owned(),
                    link.target_ref().to_owned(),
                    link.is_external(),
                )
            })
            .collect::<Vec<_>>();
        result.sort();
        result
    };
    assert_eq!(rels(p.opc_package().rels()), rels(r.opc_package().rels()));
    for part in p.opc_package().iter_parts() {
        if part.content_type() == effects::CONTENT_TYPE {
            continue;
        }
        let other = r.opc_package().get_part(part.partname()).unwrap();
        assert_eq!(part.content_type(), other.content_type());
        assert_eq!(rels(part.rels()), rels(other.rels()));
    }

    let before = members(p);
    let after = members(r);
    for (name, bytes) in before {
        if !allowed.contains(&name.as_str()) {
            assert!(after.get(&name) == Some(&bytes), "changed {name}");
        }
    }
}
#[test]
fn native_document_and_macro_effects_resources_create_replace_remove_and_invert() {
    for (file, count) in [
        ("word-text-structure.docx", 43),
        ("word-recorded-macro.docm", 34),
    ] {
        let bytes = native(file);
        let p = Package::from_bytes(&bytes).unwrap();
        assert!(
            p.styles_with_effects(Owner::MainDocument)
                .unwrap()
                .is_empty()
        );
        let base = members(&p);
        let first = resource(&p, "word/styles.xml", "Native effects 日本語");
        assert_eq!(first.styles().len(), count);
        assert_eq!(
            first.styles().get_by_id("Normal").unwrap().name(),
            Some("Native effects 日本語")
        );
        let mut e = p.edit_styles_with_effects(Owner::MainDocument);
        e.put(first.clone());
        let c = e.commit().unwrap();
        let added = published(&p, &c);
        unchanged(
            &p,
            &added,
            &["[Content_Types].xml", "word/_rels/document.xml.rels"],
        );
        assert_eq!(
            members(&added)["word/stylesWithEffects.xml"],
            first.xml_bytes()
        );
        let mut same = added.edit_styles_with_effects(Owner::MainDocument);
        same.put(first);
        assert!(!same.commit().unwrap().diagnostics().changed());
        let second = resource(&p, "word/styles.xml", "Reviewed native effects");
        let mut e = added.edit_styles_with_effects(Owner::MainDocument);
        e.put(second.clone());
        let c = e.commit().unwrap();
        let changed = published(&added, &c);
        unchanged(&added, &changed, &["word/stylesWithEffects.xml"]);
        assert_eq!(
            members(&changed)["word/stylesWithEffects.xml"],
            second.xml_bytes()
        );
        assert!(
            matches!(changed.apply_styles_with_effects(c.patch()),Err(Error::InvalidFormat(m)) if m=="stylesWithEffects patch source does not match its owner, resource, or graph precondition")
        );
        let mut e = changed.edit_styles_with_effects(Owner::MainDocument);
        e.remove();
        let c = e.commit().unwrap();
        let removed = published(&changed, &c);
        assert!(
            removed
                .styles_with_effects(Owner::MainDocument)
                .unwrap()
                .is_empty()
        );
        unchanged(
            &p,
            &removed,
            &["[Content_Types].xml", "word/_rels/document.xml.rels"],
        );
        assert_eq!(
            members(&removed).keys().collect::<Vec<_>>(),
            base.keys().collect::<Vec<_>>()
        );
        assert!(p.to_bytes().unwrap() == bytes);
        assert!(native(file) == bytes);
    }
}
#[test]
fn native_glossary_effects_remain_independent_of_main_and_ordinary_styles() {
    let bytes = native("word-glossary.dotx");
    let p = Package::from_bytes(&bytes).unwrap();
    assert!(p.glossary_graph().unwrap().is_some());
    let mut e = p.edit_styles_with_effects(Owner::MainDocument);
    e.put(resource(&p, "word/styles.xml", "Main effects"));
    let c = e.commit().unwrap();
    let main = published(&p, &c);
    let mut e = main.edit_styles_with_effects(Owner::Glossary);
    e.put(resource(&p, "word/glossary/styles.xml", "Glossary effects"));
    let c = e.commit().unwrap();
    let both = published(&main, &c);
    unchanged(
        &main,
        &both,
        &[
            "[Content_Types].xml",
            "word/glossary/_rels/document.xml.rels",
        ],
    );
    assert_eq!(
        both.styles_with_effects(Owner::MainDocument)
            .unwrap()
            .resource()
            .unwrap()
            .styles()
            .get_by_id("Normal")
            .unwrap()
            .name(),
        Some("Main effects")
    );
    assert_eq!(
        both.styles_with_effects(Owner::Glossary)
            .unwrap()
            .resource()
            .unwrap()
            .styles()
            .get_by_id("Normal")
            .unwrap()
            .name(),
        Some("Glossary effects")
    );
    let mut e = both.edit_styles_with_effects(Owner::Glossary);
    e.remove();
    let c = e.commit().unwrap();
    let removed = published(&both, &c);
    assert!(
        removed
            .styles_with_effects(Owner::Glossary)
            .unwrap()
            .is_empty()
    );
    assert_eq!(
        members(&removed)["word/stylesWithEffects.xml"],
        members(&main)["word/stylesWithEffects.xml"]
    );
    unchanged(
        &p,
        &removed,
        &[
            "[Content_Types].xml",
            "word/_rels/document.xml.rels",
            "word/glossary/_rels/document.xml.rels",
        ],
    );
    assert!(p.to_bytes().unwrap() == bytes);
}
#[test]
fn native_effects_missing_glossary_and_exact_xml_budget_are_explicit() {
    let p = Package::from_bytes(&native("word-text-structure.docx")).unwrap();
    let before = p.to_bytes().unwrap();
    let r = resource(&p, "word/styles.xml", "Bounded effects");
    let mut e = p.edit_styles_with_effects(Owner::Glossary);
    e.put(r.clone());
    assert!(
        matches!(e.commit(),Err(Error::PartNotFound(m)) if m=="glossary document owner is required before publishing stylesWithEffects")
    );
    let len = r.xml_bytes().len();
    let d = Limits::DEFAULT;
    let limits = |bytes| {
        Limits::new(
            bytes,
            d.events(),
            d.depth(),
            d.nodes(),
            d.attributes(),
            d.attribute_bytes(),
        )
        .unwrap()
    };
    assert!(Resource::new(r.xml_bytes().to_vec(), &limits(len)).is_ok());
    assert!(
        matches!(Resource::new(r.xml_bytes().to_vec(),&limits(len-1)),Err(Error::StylesWithEffectsLimit{kind:effects::LimitKind::XmlBytes,actual,maximum}) if actual==len && maximum==len-1)
    );
    assert!(p.to_bytes().unwrap() == before);
}
