//! Native Word font resources are read and edited as inert package data.
use litchi_docx::{
    Package,
    font::{self, Conformance, Embed, FontKey, Resource, Style},
};
use litchi_opc::PackURI;
use std::collections::BTreeMap;
fn native(name: &str) -> Package {
    Package::open(
        std::path::Path::new(env!("CARGO_MANIFEST_DIR"))
            .join("../../test-data/microsoft-office/word")
            .join(name),
    )
    .unwrap()
}
fn hash(bytes: &[u8]) -> String {
    use sha2::{Digest, Sha256};
    Sha256::digest(bytes)
        .iter()
        .map(|v| format!("{v:02x}"))
        .collect()
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
// Independently recorded native payload identities, including original keys.
struct Expected {
    name: &'static str,
    style: Style,
    key: &'static str,
    part: &'static str,
    bytes: usize,
    stored: &'static str,
    clear: &'static str,
}
const FULL: &[Expected] = &[
    Expected {
        name: "Aptos",
        style: Style::Regular,
        key: "{97061037-BF4F-EC4D-85B7-17BE7EB31062}",
        part: "/word/fonts/font1.odttf",
        bytes: 220792,
        stored: "5199e5ba0241e28e059805f4272ce8eb52c6a49444290de5069e10996c39d869",
        clear: "95980114fcfd42f2f9c446dae429b70582bf2f03097d68433ea9e7d85a49da0b",
    },
    Expected {
        name: "Aptos",
        style: Style::Italic,
        key: "{F80D6FAC-2739-9642-828A-7F5C21D9753E}",
        part: "/word/fonts/font2.odttf",
        bytes: 227828,
        stored: "96b70766d52e04e59dacd426490e5771e247361a3a75af62a572ec7f5769b127",
        clear: "888909eceb9acd84157078cc097c11f8a94a34ff3bd3fe08194ce7866f5998d1",
    },
    Expected {
        name: "Times New Roman",
        style: Style::Regular,
        key: "{B343B19C-7430-FF4E-9CF2-6B3300742D6D}",
        part: "/word/fonts/font3.odttf",
        bytes: 1188156,
        stored: "d638474b83a3e75342478731fa04c0f78d9bab1572ab747af826ace26414f81d",
        clear: "e77f429f5cde743c30216c2d5917706d09ddd938990f4f5ae98bfd4b5c7d0495",
    },
    Expected {
        name: "Times New Roman",
        style: Style::Bold,
        key: "{2A2F009B-A295-8945-ABC3-AB071A544A33}",
        part: "/word/fonts/font4.odttf",
        bytes: 1168472,
        stored: "be3e0ce880580fa0dbf3aa45f955f8cf8e02b805bd57a10cba0deaa1a5faf8c6",
        clear: "3750f5d563d062ed8c4e2a0516da3374b8fbcdedee0608e107efd97ded80bcae",
    },
    Expected {
        name: "Times New Roman",
        style: Style::Italic,
        key: "{038F4E0A-9EE6-FC49-A0A4-8E9051563447}",
        part: "/word/fonts/font5.odttf",
        bytes: 932820,
        stored: "774a395e3f547c92260ae5de4c91a6535508a67049a665945b456aaa4363b563",
        clear: "f957cbaf4af5b52260462b486dd10a01507184dc39cdbbb3364e428b800b37e6",
    },
    Expected {
        name: "Aptos Display",
        style: Style::Regular,
        key: "{52F08950-E6D2-894E-A22A-D29F742ADE50}",
        part: "/word/fonts/font6.odttf",
        bytes: 221112,
        stored: "bcce8c6dc0d3da59c96a686ce694e0a2583cef8b4ef37b9deed2da98d96633d8",
        clear: "af7a83e6f1d18b62535a8ee97939ab8653d5ac1fe1861cdb8fce829d3878d901",
    },
];
const SUBSET: &[Expected] = &[
    Expected {
        name: "Aptos",
        style: Style::Regular,
        key: "{B26F5D13-96A7-214E-AECA-DF7423233E66}",
        part: "/word/fonts/font1.odttf",
        bytes: 220792,
        stored: "363827cbc85176a20b67871666b84639e40da2edce3e8b3a11c799ea08643f46",
        clear: "95980114fcfd42f2f9c446dae429b70582bf2f03097d68433ea9e7d85a49da0b",
    },
    Expected {
        name: "Aptos",
        style: Style::Italic,
        key: "{FBB347AA-B0EA-9045-8E40-6550D1276269}",
        part: "/word/fonts/font2.odttf",
        bytes: 227828,
        stored: "69905247a209b9fd01d1dc6e1f2f6e9c770bfce2fbae1aafdaff850be3c059de",
        clear: "888909eceb9acd84157078cc097c11f8a94a34ff3bd3fe08194ce7866f5998d1",
    },
];

#[test]
fn native_full_and_subset_font_variants_keep_exact_faces_keys_and_clear_payloads() {
    for (name, expected, subset) in [
        ("word-embedded-fonts.docx", FULL, false),
        ("word-embedded-fonts-subset.docx", SUBSET, true),
    ] {
        let package = native(name);
        let table = package.fonts().unwrap().unwrap();
        assert_eq!(
            table.iter().map(|f| f.name()).collect::<Vec<_>>(),
            ["Aptos", "Times New Roman", "Aptos Display"]
        );
        let faces = table
            .iter()
            .flat_map(|font| font.embeds().iter().map(move |face| (font.name(), face)))
            .collect::<Vec<_>>();
        assert_eq!(faces.len(), expected.len());
        for ((font, face), expected) in faces.iter().zip(expected) {
            assert_eq!(*font, expected.name);
            assert_eq!(face.style(), expected.style);
            assert_eq!(face.key().unwrap().to_string(), expected.key);
            assert_eq!(face.subsetted(), None);
            let bytes = face.resource().unwrap().bytes();
            assert_eq!(bytes.len(), expected.bytes);
            assert_eq!(hash(bytes), expected.stored);
            let part = package
                .opc_package()
                .get_part(&PackURI::new(expected.part).unwrap())
                .unwrap();
            assert!(part.blob() == bytes);
            assert_eq!(
                part.content_type(),
                "application/vnd.openxmlformats-officedocument.obfuscatedFont"
            );
            let mut clear = bytes.to_vec();
            font::deobfuscate(&mut clear, face.key().unwrap()).unwrap();
            assert_eq!(hash(&clear), expected.clear);
            assert_eq!(&clear[..4], b"\x00\x01\x00\x00");
            font::obfuscate(&mut clear, face.key().unwrap()).unwrap();
            assert!(clear == bytes);
        }
        // These native faces omit w:subsetted. The global preference is a
        // separate saved setting, not an inferred face-level claim.
        let xml = package
            .opc_package()
            .get_part(&PackURI::new("/word/settings.xml").unwrap())
            .unwrap()
            .blob();
        let xml = std::str::from_utf8(xml).unwrap();
        assert!(xml.contains("<w:embedTrueTypeFonts/>"));
        assert_eq!(xml.contains("<w:saveSubsetFonts/>"), subset);
        let before = package.to_bytes().unwrap();
        let mut edit = package.edit_fonts();
        assert!(!edit.put(table, Conformance::Transitional).unwrap());
        let commit = edit.commit().unwrap();
        assert!(commit.package().to_bytes().unwrap() == before);
        let reopened = Package::from_bytes(&before).unwrap();
        assert_eq!(reopened.fonts().unwrap(), package.fonts().unwrap());
    }
    // Native subset-save retains the same clear Aptos programs under new keys.
    for (full, subset) in FULL[..2].iter().zip(SUBSET) {
        assert_eq!(full.clear, subset.clear);
        assert_ne!(full.stored, subset.stored);
    }
}

#[test]
fn native_font_replacement_reobfuscates_one_face_and_preserves_other_resources() {
    let package = native("word-embedded-fonts.docx");
    let before = package.to_bytes().unwrap();
    let mut table = package.fonts().unwrap().unwrap();
    let mut aptos = table.get("Aptos").unwrap().unwrap().clone();
    let original = aptos.embeds()[0].clone();
    let mut bytes = original.resource().unwrap().bytes().to_vec();
    font::deobfuscate(&mut bytes, original.key().unwrap()).unwrap();
    let clear_hash = hash(&bytes);
    let key = FontKey::new([17; 16]);
    font::obfuscate(&mut bytes, key).unwrap();
    aptos
        .put(Embed::new(
            Style::Regular,
            key,
            Resource::new(bytes.clone(), &font::Limits::DEFAULT).unwrap(),
        ))
        .unwrap();
    table.replace("Aptos", aptos).unwrap();
    let mut edit = package.edit_fonts();
    assert!(edit.put(table, Conformance::Transitional).unwrap());
    let commit = edit.commit().unwrap();
    let after = commit.package().to_bytes().unwrap();
    let output = tempfile::NamedTempFile::with_suffix(".docx").unwrap();
    commit.package().save(output.path()).unwrap();
    let reopened = Package::open(output.path()).unwrap();
    let fonts = reopened.fonts().unwrap().unwrap();
    let face = fonts
        .get("Aptos")
        .unwrap()
        .unwrap()
        .embeds()
        .iter()
        .find(|f| f.style() == Style::Regular)
        .unwrap();
    assert_eq!(face.key(), Some(key));
    assert!(face.resource().unwrap().bytes() == bytes);
    let mut clear = bytes;
    font::deobfuscate(&mut clear, key).unwrap();
    assert_eq!(hash(&clear), clear_hash);
    let a = members(&before);
    let b = members(&after);
    for (name, data) in &a {
        if !name.starts_with("word/fonts/")
            && ![
                "word/fontTable.xml",
                "word/_rels/fontTable.xml.rels",
                "[Content_Types].xml",
            ]
            .contains(&name.as_str())
        {
            assert!(b[name] == *data, "changed {name}");
        }
    }
    for expected in &FULL[1..] {
        assert!(
            b[expected.part.trim_start_matches('/')] == a[expected.part.trim_start_matches('/')]
        );
    }
    assert!(
        commit
            .package()
            .apply_fonts(&commit.patch().inverse())
            .unwrap()
            .package()
            .to_bytes()
            .unwrap()
            == before
    );
    assert!(package.to_bytes().unwrap() == before);
}

#[test]
fn native_unused_font_removal_collects_only_its_three_embedded_faces() {
    let package = native("word-embedded-fonts.docx");
    let before = package.to_bytes().unwrap();
    let mut table = package.fonts().unwrap().unwrap();
    let removed = table.remove("Times New Roman").unwrap().unwrap();
    assert_eq!(removed.embeds().len(), 3);
    let mut edit = package.edit_fonts();
    edit.put(table, Conformance::Transitional).unwrap();
    let commit = edit.commit().unwrap();
    let after = commit.package().to_bytes().unwrap();
    let reopened = Package::from_bytes(&after).unwrap();
    assert_eq!(
        reopened
            .fonts()
            .unwrap()
            .unwrap()
            .iter()
            .map(|f| f.name())
            .collect::<Vec<_>>(),
        ["Aptos", "Aptos Display"]
    );
    let a = members(&before);
    let b = members(&after);
    assert_eq!(
        a.keys()
            .filter(|n| !b.contains_key(*n))
            .map(String::as_str)
            .collect::<Vec<_>>(),
        [
            "word/fonts/font3.odttf",
            "word/fonts/font4.odttf",
            "word/fonts/font5.odttf"
        ]
    );
    assert!(b.keys().all(|n| a.contains_key(n)));
    for (name, data) in &b {
        if ![
            "word/fontTable.xml",
            "word/_rels/fontTable.xml.rels",
            "[Content_Types].xml",
        ]
        .contains(&name.as_str())
        {
            assert!(data == &a[name], "changed {name}");
        }
    }
    assert!(
        commit
            .package()
            .apply_fonts(&commit.patch().inverse())
            .unwrap()
            .package()
            .to_bytes()
            .unwrap()
            == before
    );
}
