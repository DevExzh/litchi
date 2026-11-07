//! Native-parent symEx derivatives; raw run edits are detached, not package transactions.
use litchi_docx::{
    Error, Package,
    run_symbols::{Symbol, Symbols},
};
use std::{collections::BTreeMap, io::Cursor, path::Path};
const SYMBOLS: &str = r#"<w:r xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:sx="http://schemas.microsoft.com/office/word/2015/wordml/symex" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" mc:Ignorable="sx"><sx:symEx sx:font="Segoe UI Symbol" sx:char="0001F642"/><sx:symEx/></w:r><w:r xmlns:sx="urn:litchi:foreign-symbol" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" mc:Ignorable="sx"><sx:symEx sx:char="00000041"/></w:r>"#;
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
#[test]
fn native_symex_derivatives_keep_symbol_order_optional_values_and_detached_edits_inert() {
    for file in ["word-text-structure.docx", "word-recorded-macro.docm"] {
        let path = Path::new(env!("CARGO_MANIFEST_DIR"))
            .join("../../test-data/microsoft-office/word")
            .join(file);
        let bytes = std::fs::read(&path).unwrap();
        let p = Package::from_bytes(&bytes).unwrap();
        let original = members(&p);
        let document = String::from_utf8(original["word/document.xml"].clone()).unwrap();
        let at = document.find("</w:p>").unwrap();
        let xml = format!("{}{}{}", &document[..at], SYMBOLS, &document[at..]);
        let mut e = p.edit_opc();
        e.package_mut()
            .get_part_mut(&litchi_opc::PackURI::new("/word/document.xml").unwrap())
            .unwrap()
            .set_blob(xml.into_bytes());
        let q = e.commit().unwrap().into_package();
        let published = q.to_bytes().unwrap();
        let after = members(&q);
        for (name, data) in &original {
            if name != "word/document.xml" {
                assert!(after.get(name) == Some(data), "changed {name}");
            }
        }
        assert_eq!(
            String::from_utf8(after["word/document.xml"].clone())
                .unwrap()
                .replace(SYMBOLS, ""),
            document
        );
        let check = |mut runs: Vec<litchi_docx::paragraph::Run>| {
            let foreign = runs.pop().unwrap();
            assert!(foreign.symbols().unwrap().is_empty());
            let mut run = runs.pop().unwrap();
            let values = run.symbols().unwrap();
            assert_eq!(values.len(), 2);
            assert_eq!(values.get(0).unwrap().font(), Some("Segoe UI Symbol"));
            assert_eq!(values.get(0).unwrap().character(), Some(0x1f642));
            assert_eq!(values.get(1), Some(&Symbol::empty()));
            let mut replacement = Symbols::new();
            replacement
                .push(Symbol::from_parts(Some(String::new()), Some(0)).unwrap())
                .unwrap();
            replacement
                .push(Symbol::with_character(0x41).unwrap())
                .unwrap();
            run.set_symbols(replacement.clone()).unwrap();
            assert_eq!(run.symbols().unwrap(), replacement);
            run.clear_symbols().unwrap();
            assert!(run.symbols().unwrap().is_empty());
        };
        check(
            q.document().unwrap().paragraphs().unwrap()[0]
                .runs()
                .unwrap()
                .into_vec(),
        );
        let source =
            litchi_docx::source_backed::Package::from_reader(Cursor::new(&published)).unwrap();
        check(
            source.document().unwrap().paragraphs().unwrap()[0]
                .runs()
                .unwrap()
                .into_vec(),
        );
        assert!(q.to_bytes().unwrap() == published);
        let mut unchanged = Vec::new();
        source.write_to(&mut unchanged).unwrap();
        assert!(unchanged == published);
        for code in [0xd800, 0x110000] {
            assert!(
                matches!(Symbol::with_character(code),Err(Error::Invalid(message)) if message.contains("not a Unicode scalar value"))
            );
        }
        let f = tempfile::NamedTempFile::with_suffix(".docx").unwrap();
        q.save(f.path()).unwrap();
        assert!(members(&Package::open(f.path()).unwrap()) == after);
        assert!(p.to_bytes().unwrap() == bytes);
        assert!(std::fs::read(path).unwrap() == bytes);
    }
}

fn paragraph_symbols(paragraphs: impl IntoIterator<Item = litchi_docx::Paragraph>) -> usize {
    paragraphs
        .into_iter()
        .flat_map(|p| p.runs().unwrap())
        .map(|r| r.symbols().unwrap().len())
        .sum()
}
fn symbol_count(xml: &[u8]) -> usize {
    use quick_xml::{events::Event, name::ResolveResult};
    let mut reader = quick_xml::NsReader::from_reader(xml);
    let mut count = 0;
    loop {
        let (ns, event) = reader.read_resolved_event().unwrap();
        match event {
            Event::Start(e) | Event::Empty(e) => {
                if matches!(ns,ResolveResult::Bound(n) if n.as_ref()==litchi_docx::run_symbols::NAMESPACE.as_bytes())
                    && e.local_name().as_ref() == b"symEx"
                {
                    count += 1;
                }
            },
            Event::Eof => break,
            _ => {},
        }
    }
    count
}
#[test]
fn native_ignorable_symbols_survive_header_note_comment_and_glossary_projection() {
    for (file, part, kind) in [
        ("word-tables-layout.docx", "word/header1.xml", "header"),
        ("word-text-structure.docx", "word/footnotes.xml", "footnote"),
        ("word-tables-layout.docx", "word/endnotes.xml", "endnote"),
        ("word-review-controls.docx", "word/comments.xml", "comment"),
        (
            "word-glossary.dotx",
            "word/glossary/document.xml",
            "glossary",
        ),
    ] {
        let path = Path::new(env!("CARGO_MANIFEST_DIR"))
            .join("../../test-data/microsoft-office/word")
            .join(file);
        let bytes = std::fs::read(&path).unwrap();
        let p = Package::from_bytes(&bytes).unwrap();
        let before = members(&p);
        let xml = String::from_utf8(before[part].clone()).unwrap();
        let start = if kind == "footnote" || kind == "endnote" {
            xml.find("w:id=\"1\"").unwrap()
        } else {
            0
        };
        let at = start + xml[start..].find("</w:p>").unwrap();
        let changed = format!("{}{}{}", &xml[..at], SYMBOLS, &xml[at..]);
        let mut e = p.edit_opc();
        e.package_mut()
            .get_part_mut(&litchi_opc::PackURI::new(format!("/{part}")).unwrap())
            .unwrap()
            .set_blob(changed.into_bytes());
        let q = e.commit().unwrap().into_package();
        let document = q.document().unwrap();
        let count = match kind {
            "header" => paragraph_symbols(
                document
                    .headers()
                    .unwrap()
                    .iter()
                    .flat_map(|s| s.paragraphs().unwrap()),
            ),
            "footnote" => paragraph_symbols(
                document
                    .footnotes()
                    .unwrap()
                    .iter()
                    .flat_map(|n| n.paragraphs().unwrap()),
            ),
            "endnote" => paragraph_symbols(
                document
                    .endnotes()
                    .unwrap()
                    .iter()
                    .flat_map(|n| n.paragraphs().unwrap()),
            ),
            "comment" => document
                .comments()
                .unwrap()
                .iter()
                .map(|c| symbol_count(c.xml_bytes()))
                .sum(),
            _ => q
                .glossary()
                .unwrap()
                .unwrap()
                .0
                .entries()
                .iter()
                .filter_map(|e| e.body())
                .map(symbol_count)
                .sum(),
        };
        assert_eq!(count, 2, "{kind}");
        let after = members(&q);
        for (name, data) in &before {
            if name != part {
                assert!(after.get(name) == Some(data), "changed {name}");
            }
        }
        let output = q.to_bytes().unwrap();
        let source =
            litchi_docx::source_backed::Package::from_reader(Cursor::new(&output)).unwrap();
        if kind == "footnote" || kind == "endnote" {
            let notes = if kind == "footnote" {
                source.footnotes().unwrap()
            } else {
                source.endnotes().unwrap()
            };
            assert_eq!(
                paragraph_symbols(notes.iter().flat_map(|n| n.paragraphs().unwrap())),
                2
            );
        }
        if kind == "comment" {
            assert_eq!(
                source
                    .comments()
                    .unwrap()
                    .iter()
                    .map(|c| symbol_count(c.xml_bytes()))
                    .sum::<usize>(),
                2
            );
        }
        let f = tempfile::NamedTempFile::with_suffix(".docx").unwrap();
        q.save(f.path()).unwrap();
        assert!(members(&Package::open(f.path()).unwrap()) == after);
        assert!(p.to_bytes().unwrap() == bytes);
        assert!(std::fs::read(path).unwrap() == bytes);
    }
}
