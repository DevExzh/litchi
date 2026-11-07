//! Native hyperlinks plus explicitly authored links in native header/footer stories.
use litchi_core::Position;
use litchi_docx::document::{Error as EditError, ParagraphHyperlinkAddress, Refusal};
use litchi_docx::package::StoryHyperlinkTextReplacement as Replacement;
use litchi_docx::{Error, Package};
use litchi_opc::{OpcPackage, PackURI};
const TEXT: &str = "word-text-structure.docx";
const LAYOUT: &str = "word-tables-layout.docx";
const MAIN: &str = "/word/document.xml";
const FOOT: &str = "/word/footer1.xml";
const HEAD: &str = "/word/header1.xml";
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

fn save(p: &Package) -> Package {
    let f = tempfile::NamedTempFile::with_suffix(".docx").unwrap();
    p.save(f.path()).unwrap();
    Package::open(f.path()).unwrap()
}
fn replacement(part: &str, paragraph: usize, text: &str) -> Replacement {
    Replacement::new(
        part,
        ParagraphHyperlinkAddress::new(Position::new(paragraph), Position::new(0)),
        text,
    )
}
fn changed(a: &[u8], b: &[u8]) -> Vec<String> {
    let a = members(a);
    let b = members(b);
    a.keys()
        .chain(b.keys())
        .collect::<std::collections::BTreeSet<_>>()
        .into_iter()
        .filter(|n| a.get(*n) != b.get(*n))
        .cloned()
        .collect()
}
fn derived_three_story_links() -> Package {
    let source = package(TEXT);
    let link = source
        .opc_package()
        .main_document_part()
        .unwrap()
        .rels()
        .get("rId11")
        .unwrap();
    assert!(link.is_external());
    assert_eq!(link.target_ref(), "https://example.com/litchi-fixture");
    let main = source.part_xml(MAIN).unwrap().unwrap();
    let xml = std::str::from_utf8(&main).unwrap();
    let start = xml.find("<w:hyperlink ").unwrap();
    let end = start + xml[start..].find("</w:hyperlink>").unwrap() + "</w:hyperlink>".len();
    let native_link = &xml[start..end];
    assert!(native_link.contains("Reference link"));
    let authored = native_link.replace("rId11", "rIdNativeLink");
    let original = native(LAYOUT);
    let mut opc = OpcPackage::from_bytes(&original, Default::default()).unwrap();
    for part in [MAIN, FOOT, HEAD] {
        let uri = PackURI::new(part).unwrap();
        let owner = opc.get_part_mut(&uri).unwrap();
        let mut xml = std::str::from_utf8(owner.blob()).unwrap().to_owned();
        let at = if part == FOOT {
            xml.rfind("</w:p>").unwrap()
        } else {
            xml.find("</w:p>").unwrap()
        };
        xml.insert_str(at, &authored);
        owner.set_blob(xml.into_bytes());
        assert!(owner.rels().get("rIdNativeLink").is_none());
        owner.rels_mut().add_relationship(
            litchi_opc::constants::relationship_type::HYPERLINK.into(),
            link.target_ref().into(),
            "rIdNativeLink".into(),
            true,
        );
    }
    let p = Package::from_opc_package(opc).unwrap();
    assert_eq!(native(LAYOUT), original);
    assert_members_equal(&source.to_bytes().unwrap(), &native(TEXT));
    p
}
#[test]
fn native_direct_hyperlink_batch_changes_only_its_text_and_replays_exactly() {
    let p = package(TEXT);
    let bytes = p.to_bytes().unwrap();
    let before = members(&bytes);
    let topology = p.story_inventory().unwrap().topology();
    let mut e = p.edit_stories();
    e.replace_hyperlink_texts(&[replacement(MAIN, 10, "Reviewed reference & source")])
        .unwrap();
    let c = e.commit().unwrap();
    let out = save(c.package());
    let after = out.to_bytes().unwrap();
    assert_eq!(changed(&bytes, &after), ["word/document.xml"]);
    let expected = std::str::from_utf8(&before["word/document.xml"])
        .unwrap()
        .replace("Reference link", "Reviewed reference &amp; source");
    assert_eq!(members(&after)["word/document.xml"], expected.as_bytes());
    assert_eq!(
        out.document().unwrap().paragraphs().unwrap()[10]
            .text()
            .unwrap(),
        "Reviewed reference & source"
    );
    assert_eq!(out.story_inventory().unwrap().topology(), topology);
    assert_members_equal(
        &p.apply_stories(c.patch())
            .unwrap()
            .package()
            .to_bytes()
            .unwrap(),
        &after,
    );
    assert!(matches!(
        out.apply_stories(c.patch()),
        Err(Error::PatchConflict { .. })
    ));
    assert_members_equal(
        &out.apply_stories(&c.patch().inverse())
            .unwrap()
            .package()
            .to_bytes()
            .unwrap(),
        &bytes,
    );
    let mut e = out.edit_stories();
    e.replace_hyperlink_texts(&[replacement(MAIN, 10, "Reviewed reference & source")])
        .unwrap();
    assert!(!e.commit().unwrap().diagnostics().changed());
    assert_members_equal(&p.to_bytes().unwrap(), &bytes);
}
#[test]
fn native_story_derivative_batch_preserves_every_unselected_byte_and_relationship() {
    let p = derived_three_story_links();
    let before = p.to_bytes().unwrap();
    let parts = members(&before);
    let topology = p.story_inventory().unwrap().topology();
    let mut e = p.edit_stories();
    let inventory = e
        .replace_hyperlink_texts(&[
            replacement(MAIN, 0, "Main & changed"),
            replacement(FOOT, 0, "Footer <changed>"),
            replacement(HEAD, 0, "Header changed"),
        ])
        .unwrap();
    assert_eq!(inventory.topology(), topology);
    let c = e.commit().unwrap();
    let out = save(c.package());
    let after = out.to_bytes().unwrap();
    assert_eq!(
        changed(&before, &after),
        ["word/document.xml", "word/footer1.xml", "word/header1.xml"]
    );
    let actual = members(&after);
    for (part, label) in [
        (MAIN, "Main &amp; changed"),
        (FOOT, "Footer &lt;changed&gt;"),
        (HEAD, "Header changed"),
    ] {
        let key = part.trim_start_matches('/');
        let input = std::str::from_utf8(&parts[key]).unwrap();
        assert_eq!(input.matches("Reference link").count(), 1);
        assert_eq!(
            actual[key],
            input.replace("Reference link", label).as_bytes()
        );
    }
    assert_eq!(out.story_inventory().unwrap().topology(), topology);
    assert_members_equal(
        &out.apply_stories(&c.patch().inverse())
            .unwrap()
            .package()
            .to_bytes()
            .unwrap(),
        &before,
    );
    assert!(matches!(
        out.apply_stories(c.patch()),
        Err(Error::PatchConflict { .. })
    ));
    assert_members_equal(&p.to_bytes().unwrap(), &before);
}
#[test]
fn native_story_batch_late_failure_and_ambiguous_selection_keep_prior_staging() {
    let p = derived_three_story_links();
    let before = p.to_bytes().unwrap();
    let mut e = p.edit_stories();
    e.replace_hyperlink_texts(&[replacement(MAIN, 0, "Already staged")])
        .unwrap();
    let staged = e.package().to_bytes().unwrap();
    let result = e.replace_hyperlink_texts(&[
        replacement(MAIN, 0, "must roll back"),
        replacement(HEAD, 99, "missing"),
    ]);
    assert!(matches!(
        result,
        Err(EditError::OutOfBounds { position: 99, .. })
    ));
    assert_members_equal(&e.package().to_bytes().unwrap(), &staged);
    for batch in [
        vec![],
        vec![replacement(MAIN, 0, "a"), replacement(MAIN, 0, "b")],
        vec![replacement(HEAD, 0, "a"), replacement(FOOT, 0, "b")],
    ] {
        assert!(matches!(
            e.replace_hyperlink_texts(&batch),
            Err(EditError::Refused {
                reason: Refusal::AmbiguousCompositeSelector,
                ..
            })
        ));
        assert_members_equal(&e.package().to_bytes().unwrap(), &staged);
    }
    let unreachable =
        e.replace_hyperlink_texts(&[replacement("/word/not-a-story.xml", 0, "absent")]);
    assert!(matches!(
        unreachable,
        Err(EditError::Document(Error::InvalidFormat(_)))
    ));
    assert_members_equal(&e.package().to_bytes().unwrap(), &staged);
    let out = save(e.commit().unwrap().package());
    assert_eq!(
        changed(&before, &out.to_bytes().unwrap()),
        ["word/document.xml"]
    );
    assert!(
        std::str::from_utf8(&members(&out.to_bytes().unwrap())["word/document.xml"])
            .unwrap()
            .contains("Already staged")
    );
    assert_members_equal(&p.to_bytes().unwrap(), &before);
}
