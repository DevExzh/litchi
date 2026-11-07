//! Native latent-style catalogs; password and SmartTag absence is evidence-only.
use litchi_core::Position;
use litchi_rtf::{Formatting, LatentStyles, Package};
use sha2::{Digest, Sha256};
use std::{collections::BTreeMap, path::Path};
fn root() -> std::path::PathBuf {
    Path::new(env!("CARGO_MANIFEST_DIR")).join("../../test-data/microsoft-office")
}
fn files() -> Vec<String> {
    let m: serde_json::Value =
        serde_json::from_slice(&std::fs::read(root().join("manifest.json")).unwrap()).unwrap();
    let files = m["files"]
        .as_array()
        .unwrap()
        .iter()
        .filter_map(|f| {
            let n = f["path"].as_str().unwrap();
            n.ends_with(".rtf").then(|| n.to_owned())
        })
        .collect::<Vec<_>>();
    assert_eq!(files.len(), 15);
    files
}
fn native(name: &str) -> Vec<u8> {
    std::fs::read(root().join("word").join(name)).unwrap()
}
fn hash(bytes: &[u8]) -> String {
    Sha256::digest(bytes)
        .iter()
        .map(|b| format!("{b:02x}"))
        .collect()
}
// Independent lexical inventory: escaped characters/hex bytes are not control
// words, and binary payloads are skipped rather than searched as RTF syntax.
fn controls(bytes: &[u8]) -> BTreeMap<String, usize> {
    let mut out = BTreeMap::new();
    let mut i = 0;
    while i < bytes.len() {
        if bytes[i] != b'\\' {
            i += 1;
            continue;
        }
        i += 1;
        if i == bytes.len() {
            break;
        }
        if bytes[i] == b'\'' {
            i += 3;
            continue;
        }
        if !bytes[i].is_ascii_alphabetic() {
            i += 1;
            continue;
        }
        let start = i;
        while i < bytes.len() && bytes[i].is_ascii_alphabetic() {
            i += 1;
        }
        let word = std::str::from_utf8(&bytes[start..i]).unwrap();
        *out.entry(word.to_owned()).or_insert(0) += 1;
        let parameter = i;
        if bytes.get(i) == Some(&b'-') {
            i += 1;
        }
        while i < bytes.len() && bytes[i].is_ascii_digit() {
            i += 1;
        }
        let number = &bytes[parameter..i];
        if bytes.get(i) == Some(&b' ') {
            i += 1;
        }
        if word == "bin" {
            let n = std::str::from_utf8(number)
                .unwrap()
                .parse::<usize>()
                .unwrap();
            i = i.checked_add(n).unwrap();
            assert!(i <= bytes.len());
        }
    }
    out
}
fn latent_group(bytes: &[u8]) -> &[u8] {
    let marker = br"{\*\latentstyles";
    let start = bytes
        .windows(marker.len())
        .position(|w| w == marker)
        .unwrap();
    let mut depth = 0usize;
    let mut i = start;
    while i < bytes.len() {
        match bytes[i] {
            b'\\' => {
                i += 2;
                continue;
            },
            b'{' => depth += 1,
            b'}' => {
                depth -= 1;
                if depth == 0 {
                    return &bytes[start..=i];
                }
            },
            _ => {},
        }
        i += 1;
    }
    panic!("unterminated latent-style source group")
}
fn catalog_hash(styles: &LatentStyles<'_>) -> String {
    let values = styles
        .exceptions
        .iter()
        .map(|e| {
            (
                e.name.as_ref(),
                e.locked.map(u8::from),
                e.semi_hidden.map(u8::from),
                e.unhide_when_used.map(u8::from),
                e.quick_format.map(u8::from),
                e.priority,
            )
        })
        .collect::<Vec<_>>();
    hash(&serde_json::to_vec(&values).unwrap())
}
#[test]
fn native_rtf_latent_catalogs_match_exact_defaults_order_and_values() {
    for file in files() {
        let bytes = std::fs::read(root().join(&file)).unwrap();
        let p = Package::from_bytes(&bytes).unwrap();
        let styles = p.raw().latent_styles().unwrap();
        assert_eq!(
            (
                styles.max_style_index,
                styles.locked_default,
                styles.semi_hidden_default,
                styles.unhide_when_used_default,
                styles.quick_format_default,
                styles.priority_default
            ),
            (
                376,
                Some(false),
                Some(false),
                Some(false),
                Some(false),
                Some(99)
            )
        );
        let (count, digest) = if file.ends_with("word-rtf-table-grid.rtf") {
            (
                332,
                "a507c51621e96ac8799ab450c8777b163d7f2c284708ff1477d38be3942dc240",
            )
        } else {
            (
                376,
                "451effee7ec167202747229c20ae4374c7b9bfc241a17a88b0a8649fa83d82a2",
            )
        };
        assert_eq!(styles.exceptions.len(), count, "{file}");
        assert_eq!(catalog_hash(styles), digest, "{file}");
        assert_eq!(styles.exceptions[0].name, "Normal");
        assert_eq!(styles.exceptions[0].priority, Some(0));
        assert_eq!(styles.exceptions[1].name, "heading 1");
        assert_eq!(styles.exceptions[1].priority, Some(9));
        assert_eq!(
            styles.exceptions[count - 5..]
                .iter()
                .map(|e| e.name.as_ref())
                .collect::<Vec<_>>(),
            [
                "Mention",
                "Smart Hyperlink",
                "Hashtag",
                "Unresolved Mention",
                "Smart Link"
            ]
        );
        for (name, priority) in [("Title", 10), ("Strong", 22), ("TOC Heading", 39)] {
            let e = styles.exceptions.iter().find(|e| e.name == name).unwrap();
            assert_eq!(e.priority, Some(priority));
            assert_eq!(e.quick_format, Some(true));
            assert_eq!(e.locked, Some(false));
        }
        assert!(p.raw().style_restrictions().is_empty());
        let controls = controls(&bytes);
        assert_eq!(controls.get("latentstyles"), Some(&1));
        for name in [
            "stylelock",
            "stylelockenforced",
            "stylelockbackcomp",
            "autofmtoverride",
        ] {
            assert!(!controls.contains_key(name));
        }
        assert!(p.to_bytes().unwrap() == bytes);
        assert!(std::fs::read(root().join(file)).unwrap() == bytes);
    }
}
#[test]
fn native_rtf_run_edits_preserve_the_complete_latent_style_header() {
    for file in ["word-rtf-minimal.rtf", "word-bidi-eastasia.rtf"] {
        let bytes = native(file);
        let p = Package::from_bytes(&bytes).unwrap();
        let original = p.raw().latent_styles().unwrap().clone();
        let mut e = p.edit();
        e.replace_body_run_texts(&[litchi_rtf::edit::RunTextReplacement::new(
            Position::new(0),
            Position::new(0),
            "RTF_STYLE_EDIT_070: café 日本語 🍎",
        )])
        .unwrap();
        let c = e
            .commit()
            .unwrap_or_else(|error| panic!("{file}: {error:?}"));
        let temp = tempfile::NamedTempFile::with_suffix(".rtf").unwrap();
        c.package().save(temp.path()).unwrap();
        let q = Package::open(temp.path()).unwrap();
        let output = q.to_bytes().unwrap();
        assert_eq!(q.raw().latent_styles(), Some(&original));
        assert_eq!(q.raw().style_restrictions(), p.raw().style_restrictions());
        assert_eq!(latent_group(&output), latent_group(&bytes));
        assert!(q.text().contains("RTF_STYLE_EDIT_070: café 日本語 🍎"));
        assert!(p.apply(c.patch()).unwrap().package().to_bytes().unwrap() == output);
        assert!(
            q.apply(&c.patch().inverse())
                .unwrap()
                .package()
                .to_bytes()
                .unwrap()
                == bytes
        );
        assert!(matches!(
            q.apply(c.patch()),
            Err(litchi_rtf::edit::Error::PatchConflict)
        ));
        assert!(p.to_bytes().unwrap() == bytes);
        assert!(native(file) == bytes);
    }
}
fn characters(tree: &litchi_rtf::raw::Tree<'_>) -> Vec<(char, Formatting)> {
    tree.blocks()
        .iter()
        .flat_map(|b| {
            b.text
                .chars()
                .filter(|c| *c != '\n')
                .map(move |c| (c, b.formatting))
        })
        .collect()
}
fn canonical(tree: &litchi_rtf::raw::Tree<'_>) -> Package {
    let mut file = tempfile::NamedTempFile::with_suffix(".rtf").unwrap();
    litchi_rtf::write::Writer::new(&mut file)
        .write_document(tree)
        .unwrap();
    Package::open(file.path()).unwrap()
}
fn unchanged_body(p: &Package, q: &Package) {
    assert_eq!(p.text(), q.text());
    assert_eq!(characters(p.raw()), characters(q.raw()));
    assert_eq!(
        p.raw()
            .paragraphs_with_content()
            .iter()
            .map(|p| p.properties)
            .collect::<Vec<_>>(),
        q.raw()
            .paragraphs_with_content()
            .iter()
            .map(|p| p.properties)
            .collect::<Vec<_>>()
    );
    assert_eq!(p.raw().font_table().fonts(), q.raw().font_table().fonts());
    assert_eq!(
        p.raw().color_table().colors(),
        q.raw().color_table().colors()
    );
    assert_eq!(p.raw().stylesheet(), q.raw().stylesheet());
    assert_eq!(p.raw().info(), q.raw().info());
    assert_eq!(p.raw().protection(), q.raw().protection());
    assert_eq!(p.raw().theme(), q.raw().theme());
    assert_eq!(p.raw().data_store(), q.raw().data_store());
    assert_eq!(p.raw().style_restrictions(), q.raw().style_restrictions());
}
#[test]
fn native_rtf_latent_model_edits_are_atomic_and_preserve_canonical_body_semantics() {
    for file in ["word-rtf-minimal.rtf", "word-bidi-eastasia.rtf"] {
        let bytes = native(file);
        let p = Package::from_bytes(&bytes).unwrap();
        let mut model =
            litchi_rtf::raw::Tree::parse_bytes(&bytes, litchi_rtf::read::OpenOptions::default())
                .unwrap();
        let mut changed = p.raw().latent_styles().unwrap().clone();
        changed.priority_default = Some(42);
        let title = changed
            .exceptions
            .iter_mut()
            .find(|e| e.name == "Title")
            .unwrap();
        title.priority = Some(11);
        title.quick_format = Some(false);
        model.set_latent_styles(changed.clone()).unwrap();
        for case in 0..3 {
            let mut invalid = changed.clone();
            let reason = match case {
                0 => {
                    invalid.exceptions[0].priority = Some(100);
                    "RTF latent-style priority must be in 0..=99"
                },
                1 => {
                    invalid.exceptions[1].name = invalid.exceptions[0].name.clone();
                    "RTF latent-style exception names must be unique"
                },
                _ => {
                    invalid.exceptions[0].name = "Invalid;Name".into();
                    "RTF latent-style exception name contains a forbidden delimiter/control"
                },
            };
            assert!(
                matches!(model.set_latent_styles(invalid),Err(litchi_rtf::Error::MalformedDocument(m)) if m==reason)
            );
            assert_eq!(model.latent_styles(), Some(&changed));
        }
        let q = canonical(&model);
        assert_eq!(q.raw().latent_styles(), Some(&changed));
        unchanged_body(&p, &q);
        model.clear_latent_styles();
        let q = canonical(&model);
        assert!(q.raw().latent_styles().is_none());
        assert!(!controls(&q.to_bytes().unwrap()).contains_key("latentstyles"));
        unchanged_body(&p, &q);
        assert!(p.to_bytes().unwrap() == bytes);
        assert!(native(file) == bytes);
    }
}
#[test]
fn native_rtf_corpus_has_protection_and_reservation_but_no_password_or_smarttag_payloads() {
    let mut protected = 0;
    let mut reservations = 0;
    for file in files() {
        let bytes = std::fs::read(root().join(&file)).unwrap();
        let p = Package::from_bytes(&bytes).unwrap();
        let controls = controls(&bytes);
        for word in [
            "password",
            "passwordhash",
            "xmlopen",
            "xmlclose",
            "xmlattr",
            "xmlattrns",
            "factoidname",
        ] {
            assert!(!controls.contains_key(word), "{file}: new {word} evidence");
        }
        assert!(p.raw().protection().password_hash.is_none());
        assert!(p.raw().protection().modern_password_hash().is_none());
        assert!(p.smart_tags().is_empty());
        if file.ends_with("word-document-protection.rtf") {
            assert_eq!(
                p.raw().protection().protection_type(),
                litchi_rtf::ProtectionType::Comments
            );
            assert_eq!(
                p.raw().protection().level,
                Some(litchi_rtf::ProtectionLevel::Level1)
            );
            protected += 1;
        } else if file.ends_with("word-legacy-form.rtf") {
            assert_eq!(
                p.raw().protection().protection_type(),
                litchi_rtf::ProtectionType::Forms
            );
            assert_eq!(
                p.raw().protection().level,
                Some(litchi_rtf::ProtectionLevel::Level2)
            );
            protected += 1;
        }
        if file.ends_with("word-write-reservation.rtf") {
            let r = p.raw().write_reservations();
            assert!(r.legacy.is_none());
            assert_eq!(r.hash.as_ref().unwrap().data.len(), 120);
            assert_eq!(controls.get("writereservhash"), Some(&1));
            reservations += 1;
        } else {
            assert!(p.raw().write_reservations().is_empty());
        }
        assert!(!p.text().is_empty());
        assert!(p.to_bytes().unwrap() == bytes);
        assert!(std::fs::read(root().join(file)).unwrap() == bytes);
    }
    assert_eq!((protected, reservations), (2, 1));
}

#[test]
fn native_rtf_tail_append_keeps_latent_catalogs_for_dependent_and_table_owners() {
    for file in ["word-layout-extensions.rtf", "word-rtf-table-grid.rtf"] {
        let bytes = native(file);
        let p = Package::from_bytes(&bytes).unwrap();
        if file == "word-layout-extensions.rtf" {
            let mut edit = p.edit();
            edit.replace_body_run_texts(&[litchi_rtf::edit::RunTextReplacement::new(
                Position::new(0),
                Position::new(0),
                "Refused dependent-body edit",
            )])
            .unwrap();
            assert!(
                matches!(edit.commit(),Err(litchi_rtf::edit::Error::Refused(r)) if r.kind()==litchi_core::edit::RefusalKind::UnsupportedContent && r.operation()=="edit" && r.reason()=="run-format editing refuses dependent body destinations or review content")
            );
        }
        let mut edit = p.edit_tail_append(litchi_rtf::tail_append::TailSelector::Body);
        edit.append_text_paragraphs(&["RTF_STYLE_TAIL_070: café 日本語 🍎"])
            .unwrap();
        let c = edit.commit().unwrap();
        let f = tempfile::NamedTempFile::with_suffix(".rtf").unwrap();
        c.package().save(f.path()).unwrap();
        let q = Package::open(f.path()).unwrap();
        let output = q.to_bytes().unwrap();
        assert_eq!(q.raw().latent_styles(), p.raw().latent_styles());
        assert_eq!(q.raw().style_restrictions(), p.raw().style_restrictions());
        assert_eq!(latent_group(&output), latent_group(&bytes));
        assert!(q.text().contains("RTF_STYLE_TAIL_070: café 日本語 🍎"));
        assert_eq!(q.paragraph_count(), p.paragraph_count() + 1);
        assert!(
            p.apply_tail_append(c.patch())
                .unwrap()
                .package()
                .to_bytes()
                .unwrap()
                == output
        );
        assert!(
            q.apply_tail_append(&c.patch().inverse())
                .unwrap()
                .package()
                .to_bytes()
                .unwrap()
                == bytes
        );
        assert!(p.to_bytes().unwrap() == bytes);
        assert!(native(file) == bytes);
    }
}
