//! Native Word Custom XML derivatives and ordinary DOC edits preserve independent CFB owners.
#[path = "../../../tools/test_support/native_custom_xml.rs"]
mod stores;
use litchi_doc::Package;
#[test]
fn native_doc_custom_xml_lifecycle_and_text_edits_preserve_opaque_storage_metadata() {
    for name in [
        "word/word-text-structure.doc",
        "word/word-recorded-macro.dot",
    ] {
        let bytes = stores::native(name);
        let seeded = stores::exercise(&bytes);
        let p = Package::from_bytes(&seeded).unwrap();
        let original = Package::from_bytes(&bytes).unwrap();
        assert_eq!(
            p.document().text().unwrap(),
            original.document().text().unwrap()
        );
        assert_eq!(
            p.document().fib().is_template(),
            original.document().fib().is_template()
        );
        let mut raw = litchi_doc::raw::Reader::from_reader(std::io::Cursor::new(&seeded)).unwrap();
        let store = raw
            .custom_xml_data_store(Default::default())
            .unwrap()
            .unwrap();
        assert_eq!(store.items(), stores::snapshot(&seeded).unwrap().items());
        let mut e = p.edit().unwrap();
        e.replace_paragraph(litchi_core::Position::new(0), "CUSTOM_XML_PARENT_EDIT_064")
            .unwrap();
        let c = e.commit().unwrap();
        let out = c.package().to_bytes().unwrap();
        stores::preserved_owned(&seeded, &out);
        let before = stores::streams(&seeded);
        let after = stores::streams(&out);
        assert_eq!(
            before.keys().collect::<Vec<_>>(),
            after.keys().collect::<Vec<_>>()
        );
        for (path, data) in before {
            if !matches!(path.as_slice(),[name] if matches!(name.as_str(),"WordDocument"|"0Table"|"1Table"))
            {
                assert!(after.get(&path) == Some(&data), "changed {path:?}");
            }
        }

        let f = tempfile::NamedTempFile::with_suffix(".doc").unwrap();
        c.package().save(f.path()).unwrap();
        let reopened = Package::open(f.path()).unwrap();
        assert!(
            reopened
                .document()
                .text()
                .unwrap()
                .contains("CUSTOM_XML_PARENT_EDIT_064")
        );
        stores::preserved_owned(&seeded, &reopened.to_bytes().unwrap());
        assert!(
            reopened
                .apply(&c.patch().inverse())
                .unwrap()
                .package()
                .to_bytes()
                .unwrap()
                == seeded
        );
        assert!(p.to_bytes().unwrap() == seeded);
        assert!(stores::native(name) == bytes);
    }
}
