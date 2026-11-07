//! Corpus evidence corrections, kept separate from positive semantic coverage.
#![cfg(feature = "encryption")]
use litchi_docx::Package;
use quick_xml::{events::Event, name::ResolveResult, reader::NsReader};
use std::path::Path;

const W: &[u8] = b"http://schemas.openxmlformats.org/wordprocessingml/2006/main";
const W14: &[u8] = b"http://schemas.microsoft.com/office/word/2010/wordml";
const W15: &[u8] = b"http://schemas.microsoft.com/office/word/2012/wordml";

#[test]
fn native_word_corpus_has_no_run_effects_or_extension_content_controls() {
    let root = Path::new(env!("CARGO_MANIFEST_DIR")).join("../../test-data/microsoft-office");
    let manifest: serde_json::Value =
        serde_json::from_slice(&std::fs::read(root.join("manifest.json")).unwrap()).unwrap();
    let mut count = 0;
    let mut encrypted = 0;
    let mut neighboring = [false; 3];
    for item in manifest["files"].as_array().unwrap() {
        let name = item["path"].as_str().unwrap();
        if !name.starts_with("word/")
            || !matches!(
                Path::new(name).extension().and_then(|e| e.to_str()),
                Some("docx" | "dotx" | "docm" | "dotm")
            )
        {
            continue;
        }
        let bytes = std::fs::read(root.join(name)).unwrap();
        let p = if bytes.starts_with(b"PK") {
            Package::from_bytes(&bytes).unwrap()
        } else {
            assert_eq!(name, "word/word-encrypted.docx");
            encrypted += 1;
            Package::from_bytes_with_password(&bytes, "Office42!").unwrap()
        };
        let mut effects = 0;
        let mut controls = 0;
        let mut ligatures = 0;
        let mut fields = 0;
        let mut shapes = 0;
        let mut xml_parts = 0;
        for part in p.opc_package().try_iter_parts() {
            let part = part.unwrap();
            if !part.partname().as_str().ends_with(".xml")
                && !part.content_type().ends_with("+xml")
                && !matches!(part.content_type(), "application/xml" | "text/xml")
            {
                continue;
            }
            xml_parts += 1;
            let mut reader = NsReader::from_reader(part.blob());
            loop {
                let (namespace, event) = reader.read_resolved_event().unwrap();
                match event {
                    Event::Start(e) | Event::Empty(e) => {
                        let ns = match namespace {
                            ResolveResult::Bound(ns) => ns.as_ref().to_vec(),
                            ResolveResult::Unbound => Vec::new(),
                            ResolveResult::Unknown(_) => {
                                panic!("unbound prefix in {name}: {}", part.partname())
                            },
                        };
                        let local = e.local_name();
                        let local = local.as_ref();
                        effects += usize::from(
                            ns == W14
                                && matches!(
                                    local,
                                    b"glow"
                                        | b"shadow"
                                        | b"reflection"
                                        | b"textOutline"
                                        | b"textFill"
                                        | b"scene3d"
                                        | b"props3d"
                                ),
                        );
                        controls += usize::from(
                            (ns == W14 && matches!(local, b"checkbox" | b"entityPicker"))
                                || (ns == W15
                                    && matches!(
                                        local,
                                        b"repeatingSection"
                                            | b"repeatingSectionItem"
                                            | b"appearance"
                                            | b"color"
                                            | b"dataBinding"
                                    ))
                                || (ns == W && local == b"dataBinding"),
                        );
                        if part.partname().as_str() == "/word/document.xml" {
                            ligatures += usize::from(ns == W14 && local == b"ligatures");
                            fields += usize::from(ns == W && local == b"ffData");
                            shapes += usize::from(ns==b"http://schemas.microsoft.com/office/word/2010/wordprocessingShape" && local==b"wsp");
                        }
                    },
                    Event::Eof => break,
                    _ => {},
                }
            }
        }
        assert!(xml_parts >= 5, "{name}: no meaningful XML inventory");
        assert_eq!(
            (effects, controls),
            (0, 0),
            "{name}: new feature evidence; revise the evidence review instead of treating absence as support"
        );
        if name == "word/word-recorded-macro.docm" {
            assert_eq!(ligatures, 9); // Five runs and four paragraph-mark properties.
            neighboring[0] = true;
        }
        if name == "word/word-legacy-form.docx" {
            assert_eq!(fields, 3);
            neighboring[1] = true;
        }
        if name == "word/word-vector-diagram.docx" {
            assert_eq!(shapes, 3);
            neighboring[2] = true;
        }
        count += 1;
    }
    assert_eq!((count, encrypted), (24, 1));
    assert_eq!(neighboring, [true; 3]);
}
