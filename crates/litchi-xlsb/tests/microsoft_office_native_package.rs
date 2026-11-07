//! Native XLSB package, theme and property behavior; external content stays inert.
use litchi_ooxml_common::custom::{Props, Value};
use litchi_opc::{BlobPart, PackURI, TargetMode};
use litchi_xlsb::raw::{self, Encoder, Limits, Records, kind};
use litchi_xlsb::theme::{Color, Face, FontSet, Slot};
use litchi_xlsb::{Package, SourceBackedWorkbook};
use std::{collections::BTreeMap, io::Cursor};

const FILES: [&str; 3] = [
    "excel-values-formulas-chart.xlsb",
    "excel-metadata.xlsb",
    "excel-recorded-macro.xlsb",
];
const THEME: &str = "xl/theme/theme1.xml";
const FAMILY_ID: &str = "{2E142A2C-CD16-42D6-873A-C26D2A0506FA}";
const FAMILY_VARIANT: &str = "{1BDDFF52-6CD6-40A5-AB3C-68EB2F1E4D0A}";
fn native(name: &str) -> Vec<u8> {
    std::fs::read(
        std::path::Path::new(env!("CARGO_MANIFEST_DIR"))
            .join("../../test-data/microsoft-office/excel")
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
fn source_bytes(source: &SourceBackedWorkbook) -> Vec<u8> {
    let mut b = Vec::new();
    source.write_to(&mut b).unwrap();
    b
}
fn changed(a: &[u8], b: &[u8]) -> Vec<String> {
    let a = members(a);
    let b = members(b);
    let names: std::collections::BTreeSet<_> = a.keys().chain(b.keys()).collect();
    names
        .into_iter()
        .filter(|n| a.get(*n) != b.get(*n))
        .cloned()
        .collect()
}
fn fragment(xml: &[u8], name: &[u8]) -> Vec<u8> {
    use quick_xml::events::Event;
    let mut r = quick_xml::Reader::from_reader(xml);
    let mut start = None;
    let mut depth = 0usize;
    loop {
        let at = r.buffer_position() as usize;
        let e = r.read_event().unwrap();
        let end = r.buffer_position() as usize;
        match e {
            Event::Start(e) => {
                if start.is_none() && e.name().as_ref() == name {
                    start = Some((at, depth));
                }
                depth += 1;
            },
            Event::Empty(e) if start.is_none() && e.name().as_ref() == name => {
                return xml[at..end].to_vec();
            },
            Event::End(_) => {
                depth -= 1;
                if start.is_some_and(|(_, d)| d == depth) {
                    return xml[start.unwrap().0..end].to_vec();
                }
            },
            Event::Eof => panic!("missing XML fragment"),
            _ => {},
        }
    }
}
fn graph_elements(xml: &[u8]) -> Vec<(String, BTreeMap<String, String>)> {
    use quick_xml::events::Event;
    let mut r = quick_xml::Reader::from_reader(xml);
    let mut out = Vec::new();
    loop {
        match r.read_event().unwrap() {
            Event::Start(e) | Event::Empty(e) => {
                let attrs = e
                    .attributes()
                    .map(|a| {
                        let a = a.unwrap();
                        (
                            String::from_utf8(a.key.as_ref().to_vec()).unwrap(),
                            a.decoded_and_normalized_value(
                                quick_xml::XmlVersion::Explicit1_0,
                                r.decoder(),
                            )
                            .unwrap()
                            .into_owned(),
                        )
                    })
                    .collect();
                out.push((
                    String::from_utf8(e.name().as_ref().to_vec()).unwrap(),
                    attrs,
                ));
            },
            Event::Eof => break,
            Event::Decl(_) | Event::End(_) => {},
            Event::Text(t) => assert!(t.as_ref().iter().all(u8::is_ascii_whitespace)),
            _ => panic!("unexpected graph markup"),
        }
    }
    out.sort();
    out
}
fn content_types(parts: &BTreeMap<String, Vec<u8>>) -> BTreeMap<String, String> {
    let mut defaults = BTreeMap::new();
    let mut overrides = BTreeMap::new();
    for (name, attrs) in graph_elements(&parts["[Content_Types].xml"]) {
        match name.as_str() {
            "Default" => {
                assert_eq!(attrs.len(), 2);
                assert!(
                    defaults
                        .insert(attrs["Extension"].clone(), attrs["ContentType"].clone())
                        .is_none()
                );
            },
            "Override" => {
                assert_eq!(attrs.len(), 2);
                assert!(
                    overrides
                        .insert(attrs["PartName"].clone(), attrs["ContentType"].clone())
                        .is_none()
                );
            },
            "Types" => assert_eq!(
                attrs,
                BTreeMap::from([(
                    "xmlns".into(),
                    "http://schemas.openxmlformats.org/package/2006/content-types".into()
                )])
            ),
            _ => panic!("unexpected content-type element: {name}"),
        }
    }
    parts
        .keys()
        .filter(|n| *n != "[Content_Types].xml")
        .map(|n| {
            let kind = overrides
                .get(&format!("/{n}"))
                .or_else(|| defaults.get(n.rsplit('.').next().unwrap()))
                .unwrap();
            (n.clone(), kind.clone())
        })
        .collect()
}
fn assert_same_payloads_and_graph(before: &[u8], after: &[u8]) {
    let a = members(before);
    let b = members(after);
    assert_eq!(a.keys().collect::<Vec<_>>(), b.keys().collect::<Vec<_>>());
    // A default and per-part overrides can describe the same OPC typing.
    // Check every member's resolved type, not that serialization choice.
    assert_eq!(content_types(&a), content_types(&b));
    for (n, data) in a {
        if n == "[Content_Types].xml" {
            continue;
        }
        if n.ends_with(".rels") {
            assert_eq!(graph_elements(&data), graph_elements(&b[&n]), "{n}");
        } else {
            assert_eq!(data, b[&n], "{n}");
        }
    }
}
fn save(p: &Package) -> Package {
    let file = tempfile::NamedTempFile::with_suffix(".xlsb").unwrap();
    p.save(file.path()).unwrap();
    Package::open(file.path()).unwrap()
}
fn scripts(face: &Face, name: &str) -> String {
    face.scripts
        .iter()
        .find(|s| s.code == name)
        .unwrap()
        .typeface
        .clone()
}

#[test]
fn native_theme_values_match_eager_and_source_views_with_one_payload_load() {
    use litchi_drawingml::theme::System;
    for name in FILES {
        let bytes = native(name);
        let p = Package::from_bytes(&bytes).unwrap();
        let source = SourceBackedWorkbook::from_reader(Cursor::new(&bytes)).unwrap();
        let before = source.cache_diagnostics();
        let view = source.theme().unwrap().unwrap();
        let after = source.cache_diagnostics();
        assert_eq!(
            (
                after.cold_loads - before.cold_loads,
                after.retained_entries - before.retained_entries,
                after.retained_bytes - before.retained_bytes
            ),
            (1, 1, 8721)
        );
        let theme = p.theme().unwrap().unwrap();
        assert_eq!(view.theme(), theme.theme());
        assert_eq!(view.family(), theme.family());
        assert_eq!(view.source_xml(), members(&bytes)[THEME]);
        assert_eq!(theme.source_xml(), view.source_xml());
        assert_eq!(theme.part_name(), "/xl/theme/theme1.xml");
        assert_eq!(theme.theme().name, "Office Theme");
        assert_eq!(theme.theme().colors.name(), "Office");
        for (slot, value) in Slot::ALL.into_iter().zip([
            Color::system(System::WindowText, Some("000000")).unwrap(),
            Color::system(System::Window, Some("FFFFFF")).unwrap(),
            Color::rgb("0E2841").unwrap(),
            Color::rgb("E8E8E8").unwrap(),
            Color::rgb("156082").unwrap(),
            Color::rgb("E97132").unwrap(),
            Color::rgb("196B24").unwrap(),
            Color::rgb("0F9ED5").unwrap(),
            Color::rgb("A02B93").unwrap(),
            Color::rgb("4EA72E").unwrap(),
            Color::rgb("467886").unwrap(),
            Color::rgb("96607D").unwrap(),
        ]) {
            assert_eq!(theme.theme().colors.color(slot), Some(&value));
        }
        let fonts = &theme.theme().fonts;
        assert_eq!(fonts.name(), "Office");
        assert_eq!(fonts.major().latin, "Aptos Display");
        assert_eq!(fonts.minor().latin, "Aptos Narrow");
        for face in [fonts.major(), fonts.minor()] {
            assert_eq!(face.scripts.len(), 47);
            assert!(face.east_asian.is_empty() && face.complex_script.is_empty());
        }
        assert_eq!(scripts(fonts.major(), "Jpan"), "游ゴシック Light");
        assert_eq!(scripts(fonts.minor(), "Jpan"), "游ゴシック");
        let family = theme.family().unwrap();
        assert_eq!(family.name(), "Office Theme");
        assert_eq!(family.id().as_str(), FAMILY_ID);
        assert_eq!(family.variant_id().as_str(), FAMILY_VARIANT);
        source.theme().unwrap();
        assert_eq!(source.cache_diagnostics().cold_loads, after.cold_loads);
        assert_eq!(
            p.edit_theme()
                .unwrap()
                .commit()
                .unwrap()
                .package()
                .to_bytes()
                .unwrap(),
            bytes
        );
        assert_eq!(source_bytes(&source), bytes);
    }
}

#[test]
fn native_theme_name_palette_and_family_edits_preserve_fonts_effects_and_binary_parts() {
    for name in FILES {
        let bytes = native(name);
        let p = Package::from_bytes(&bytes).unwrap();
        let before = p.theme().unwrap().unwrap();
        let palette = before
            .theme()
            .colors
            .clone()
            .with(Slot::Accent1, Color::rgb("123456").unwrap());
        let mut family = before.family().unwrap().clone();
        family.set_name("Native family review").unwrap();
        let mut e = p.edit_theme().unwrap();
        e.set_name("Native theme review").unwrap();
        e.set_palette(palette.clone()).unwrap();
        e.set_family(family.clone()).unwrap();
        let staged = e.theme().unwrap().unwrap().clone();
        assert!(e.set_name("X".repeat(257)).is_err());
        assert_eq!(e.theme().unwrap(), Some(&staged));
        let commit = e.commit().unwrap();
        let after = save(commit.package());
        let output = after.to_bytes().unwrap();
        assert_eq!(changed(&bytes, &output), [THEME]);
        let read = after.theme().unwrap().unwrap();
        assert_eq!(read.theme(), &staged);
        assert_eq!(read.family(), Some(&family));
        assert_eq!(read.theme().colors, palette);
        for child in [
            b"a:fontScheme".as_slice(),
            b"a:fmtScheme",
            b"a:objectDefaults",
            b"a:extraClrSchemeLst",
        ] {
            assert_eq!(
                fragment(before.source_xml(), child),
                fragment(read.source_xml(), child)
            );
        }
        assert_eq!(
            members(
                &p.apply_theme(commit.patch())
                    .unwrap()
                    .package()
                    .to_bytes()
                    .unwrap()
            ),
            members(&output)
        );
        assert_eq!(
            members(
                &after
                    .apply_theme(&commit.patch().inverse())
                    .unwrap()
                    .package()
                    .to_bytes()
                    .unwrap()
            ),
            members(&bytes)
        );
        // panose is retained source data outside the writable font vocabulary.
        let fonts = FontSet::new(
            before.theme().fonts.name(),
            Face {
                latin: "Native Major".into(),
                ..before.theme().fonts.major().clone()
            },
            before.theme().fonts.minor().clone(),
        );
        let mut unsupported = p.edit_theme().unwrap();
        unsupported.set_fonts(fonts).unwrap();
        assert!(matches!(
            unsupported.commit(),
            Err(litchi_xlsb::package::PackageError::Drawing(
                litchi_drawingml::Error::Invalid(message)
            )) if message == "theme scheme contains an unsupported attribute"
        ));
        assert_eq!(p.to_bytes().unwrap(), bytes);
    }
}

#[test]
fn native_theme_family_and_owner_removal_restore_the_original_graph() {
    let bytes = native(FILES[0]);
    let p = Package::from_bytes(&bytes).unwrap();
    let original = p.theme().unwrap().unwrap();
    let mut e = p.edit_theme().unwrap();
    assert!(e.remove_family().unwrap());
    let removal = e.commit().unwrap();
    let removed = save(removal.package());
    assert_eq!(removed.theme().unwrap().unwrap().theme(), original.theme());
    assert!(removed.theme().unwrap().unwrap().family().is_none());
    assert_eq!(changed(&bytes, &removed.to_bytes().unwrap()), [THEME]);
    assert_eq!(
        members(
            &removed
                .apply_theme(&removal.patch().inverse())
                .unwrap()
                .package()
                .to_bytes()
                .unwrap()
        ),
        members(&bytes)
    );
    let mut e = p.edit_theme().unwrap();
    assert!(e.remove().unwrap());
    let removal = e.commit().unwrap();
    let empty = save(removal.package());
    assert!(empty.theme().unwrap().is_none());
    assert_eq!(
        changed(&bytes, &empty.to_bytes().unwrap()),
        ["[Content_Types].xml", "xl/_rels/workbook.bin.rels", THEME]
    );
    assert_same_payloads_and_graph(
        &bytes,
        &empty
            .apply_theme(&removal.patch().inverse())
            .unwrap()
            .package()
            .to_bytes()
            .unwrap(),
    );
    let mut create = empty.edit_theme().unwrap();
    create.set_theme(original.theme().clone()).unwrap();
    create
        .set_family(original.family().unwrap().clone())
        .unwrap();
    let created = create.commit().unwrap();
    let recreated = save(created.package());
    let theme = recreated.theme().unwrap().unwrap();
    assert_eq!(theme.theme(), original.theme());
    let family = theme.family().unwrap();
    let original_family = original.family().unwrap();
    assert_eq!(family.id(), original_family.id());
    assert_eq!(family.variant_id(), original_family.variant_id());
    assert_eq!(family.name(), original_family.name());
    assert_same_payloads_and_graph(
        &empty.to_bytes().unwrap(),
        &recreated
            .apply_theme(&created.patch().inverse())
            .unwrap()
            .package()
            .to_bytes()
            .unwrap(),
    );
    assert_eq!(p.to_bytes().unwrap(), bytes);
}

#[test]
fn native_properties_keep_saved_values_and_source_eager_noop_parity() {
    for (name, created, modified) in [
        (FILES[0], "2026-10-15T05:23:50Z", "2026-10-15T05:27:24Z"),
        (FILES[1], "2026-10-15T11:10:49Z", "2026-10-15T11:15:55Z"),
        (FILES[2], "2026-10-15T06:29:36Z", "2026-10-15T06:31:04Z"),
    ] {
        let bytes = native(name);
        let p = Package::from_bytes(&bytes).unwrap();
        let source = SourceBackedWorkbook::from_reader(Cursor::new(&bytes)).unwrap();
        let m = p.metadata().unwrap();
        assert_eq!(source.metadata().unwrap(), m);
        assert_eq!(m.author.as_deref(), Some("Ryker Zhu"));
        assert_eq!(m.last_modified_by, m.author);
        assert_eq!(m.title, None);
        assert_eq!(m.company.as_deref(), Some(""));
        assert_eq!(m.created.unwrap().to_string(), created);
        assert_eq!(m.modified.unwrap().to_string(), modified);
        let app = p.app_props().unwrap().unwrap();
        assert_eq!(source.app_props().unwrap(), Some(app.clone()));
        assert_eq!(
            app.application.as_deref(),
            Some("Microsoft Macintosh Excel")
        );
        assert_eq!(app.app_version.as_deref(), Some("16.0300"));
        assert_eq!(app.doc_security, Some(0));
        assert_eq!(
            (
                app.scale_crop,
                app.links_up_to_date,
                app.shared_doc,
                app.hyperlinks_changed
            ),
            (Some(false), Some(false), Some(false), Some(false))
        );
        let mut e = p.edit_metadata().unwrap();
        assert!(e.custom_props().is_empty());
        e.set(m.clone()).set_app_props(Some(app.clone()));
        let a = e.commit().unwrap();
        assert_eq!(a.package().to_bytes().unwrap(), bytes);
        let mut e = source.edit_metadata().unwrap();
        e.set(m).set_app_props(Some(app));
        let b = e.commit().unwrap();
        assert_eq!(source_bytes(b.package()), bytes);
        assert_eq!(a.patch(), b.patch());
    }
}

#[test]
fn native_metadata_crud_has_source_eager_patch_parity_and_preserves_app_vectors() {
    for name in FILES {
        let bytes = native(name);
        let p = Package::from_bytes(&bytes).unwrap();
        let source = SourceBackedWorkbook::from_reader(Cursor::new(&bytes)).unwrap();
        let original = p.metadata().unwrap();
        let mut custom = Props::new();
        custom.insert("NativeReviewed", Value::Bool(true)).unwrap();
        custom.insert("NativeCount", Value::I32(42)).unwrap();
        let mut eager = p.edit_metadata().unwrap();
        eager
            .update(|m| {
                m.title = Some("Native property review".into());
                m.company = Some("Fixture review".into());
            })
            .unwrap();
        eager.set_custom_props(custom.clone());
        let a = eager.commit().unwrap();
        let mut lazy = source.edit_metadata().unwrap();
        lazy.update(|m| {
            m.title = Some("Native property review".into());
            m.company = Some("Fixture review".into());
        })
        .unwrap();
        lazy.set_custom_props(custom);
        let b = lazy.commit().unwrap();
        assert_eq!(a.patch(), b.patch());
        let output = a.package().to_bytes().unwrap();
        // Graph writers may differ in whitespace and default/override choices. The
        // relationships/content types must agree and every payload is exact.
        assert_same_payloads_and_graph(&source_bytes(b.package()), &output);
        let after = save(a.package());
        assert_eq!(
            changed(&bytes, &output),
            [
                "[Content_Types].xml",
                "_rels/.rels",
                "docProps/app.xml",
                "docProps/core.xml",
                "docProps/custom.xml"
            ]
        );
        let before_parts = members(&bytes);
        let after_parts = members(&output);
        for name in [b"HeadingPairs".as_slice(), b"TitlesOfParts"] {
            assert_eq!(
                fragment(&before_parts["docProps/app.xml"], name),
                fragment(&after_parts["docProps/app.xml"], name)
            );
        }
        let mut expected = original;
        expected.title = Some("Native property review".into());
        expected.company = Some("Fixture review".into());
        assert_eq!(after.metadata().unwrap(), expected);
        assert_eq!(
            after
                .edit_metadata()
                .unwrap()
                .custom_props()
                .get("nativecount"),
            Some(&Value::I32(42))
        );
        assert_eq!(
            after
                .edit_metadata()
                .unwrap()
                .custom_props()
                .get("nativereviewed"),
            Some(&Value::Bool(true))
        );
        assert_same_payloads_and_graph(
            &bytes,
            &after
                .apply_metadata(&a.patch().inverse())
                .unwrap()
                .package()
                .to_bytes()
                .unwrap(),
        );
        let lazy_reopened = SourceBackedWorkbook::from_reader(Cursor::new(&output)).unwrap();
        assert_same_payloads_and_graph(
            &bytes,
            &source_bytes(
                lazy_reopened
                    .apply_metadata(&b.patch().inverse())
                    .unwrap()
                    .package(),
            ),
        );
        assert!(matches!(
            after.apply_metadata(a.patch()),
            Err(litchi_xlsb::metadata::Error::PatchConflict)
        ));
        let mut remove = after.edit_metadata().unwrap();
        remove.remove_custom_props().set_app_props(None);
        let removed = remove.commit().unwrap();
        let absent = save(removed.package());
        assert!(absent.app_props().unwrap().is_none());
        assert!(absent.edit_metadata().unwrap().custom_props().is_empty());
        assert_eq!(absent.metadata().unwrap().title, expected.title);
        assert_same_payloads_and_graph(
            &output,
            &absent
                .apply_metadata(&removed.patch().inverse())
                .unwrap()
                .package()
                .to_bytes()
                .unwrap(),
        );
        assert_eq!(p.to_bytes().unwrap(), bytes);
        assert_eq!(source_bytes(&source), bytes);
    }
}

#[test]
fn native_biff12_parts_reencode_exactly_and_bundle_records_match_the_package_graph() {
    for (name, member_count, counts, lengths, bundles) in [
        (
            FILES[0],
            25,
            [41, 58, 46, 210],
            [997, 978, 1274, 3179],
            vec![
                (
                    2,
                    "rId1",
                    "Sales chart",
                    "chartsheet",
                    "chartsheets/sheet1.bin",
                ),
                (1, "rId2", "Sheet1", "worksheet", "worksheets/sheet1.bin"),
                (3, "rId3", "Sheet2", "worksheet", "worksheets/sheet2.bin"),
            ],
        ),
        (
            FILES[1],
            30,
            [43, 47, 17, 70],
            [1127, 822, 563, 920],
            vec![
                (1, "rId1", "Sheet1", "worksheet", "worksheets/sheet1.bin"),
                (
                    2,
                    "rId2",
                    "Source_A_015",
                    "worksheet",
                    "worksheets/sheet2.bin",
                ),
                (
                    3,
                    "rId3",
                    "Source_B_015",
                    "worksheet",
                    "worksheets/sheet3.bin",
                ),
                (
                    4,
                    "rId4",
                    "Consolidated_015",
                    "worksheet",
                    "worksheets/sheet4.bin",
                ),
                (5, "rId5", "Print_015", "worksheet", "worksheets/sheet5.bin"),
            ],
        ),
        (
            FILES[2],
            13,
            [35, 45, 8, 48],
            [881, 711, 374, 630],
            vec![(1, "rId1", "Sheet1", "worksheet", "worksheets/sheet1.bin")],
        ),
    ] {
        let bytes = native(name);
        let parts = members(&bytes);
        assert_eq!(parts.len(), member_count);
        let p = Package::from_bytes(&bytes).unwrap();
        let wb = p
            .opc_package()
            .get_part(&PackURI::new("/xl/workbook.bin").unwrap())
            .unwrap();
        assert_eq!(
            wb.content_type(),
            "application/vnd.ms-excel.sheet.binary.macroEnabled.main"
        );
        let root = p.opc_package().rels().get("rId1").unwrap();
        assert_eq!(
            root.reltype(),
            "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument"
        );
        assert_eq!(root.target_ref(), "xl/workbook.bin");
        assert!(!root.is_external());
        for (i, (part, first, last)) in [
            ("xl/workbook.bin", 131, 132),
            ("xl/styles.bin", 278, 279),
            ("xl/sharedStrings.bin", 159, 160),
            ("xl/worksheets/sheet1.bin", 129, 130),
        ]
        .into_iter()
        .enumerate()
        {
            let data = &parts[part];
            assert_eq!(data.len(), lengths[i]);
            let records = Records::new(data, Limits::DEFAULT)
                .collect::<Result<Vec<_>, _>>()
                .unwrap();
            assert_eq!(records.len(), counts[i]);
            assert_eq!(records.first().unwrap().kind().get(), first);
            assert_eq!(records.last().unwrap().kind().get(), last);
            let mut output = Vec::new();
            let mut writer = Encoder::new(&mut output, Limits::DEFAULT);
            for record in &records {
                writer
                    .write_record(record.kind(), record.payload())
                    .unwrap();
            }
            assert_eq!(output, *data);
        }
        let data = &parts["xl/workbook.bin"];
        let mut actual = Vec::new();
        for record in Records::new(data, Limits::DEFAULT)
            .map(Result::unwrap)
            .filter(|r| r.kind() == kind::BUNDLE_SH)
        {
            let mut c = raw::Cursor::new(record.payload(), "native BrtBundleSh", Limits::DEFAULT);
            assert_eq!(c.read_u32().unwrap(), 0);
            let id = c.read_u32().unwrap();
            let rel = c.read_nullable_wide_string().unwrap().unwrap();
            let title = c.read_wide_string().unwrap();
            assert_eq!(c.remaining(), 0);
            actual.push((id, rel, title));
        }
        assert_eq!(
            actual,
            bundles
                .iter()
                .map(|(id, rel, title, _, _)| (*id, rel.to_string(), title.to_string()))
                .collect::<Vec<_>>()
        );
        for (_, id, _, kind, target) in bundles {
            let rel = wb.rels().get(id).unwrap();
            assert_eq!(
                rel.reltype(),
                format!(
                    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/{kind}"
                )
            );
            assert_eq!(rel.target_ref(), target);
            assert!(!rel.is_external());
            assert!(parts.contains_key(&format!("xl/{target}")));
        }
        let mut truncated = Records::new(&data[..data.len() - 1], Limits::DEFAULT);
        let error = truncated.find_map(Result::err).unwrap();
        assert!(matches!(
            error,
            raw::Error::Truncated {
                stage: raw::Stage::Length,
                ..
            }
        ));
        assert!(truncated.next().is_none());
        assert_eq!(p.to_bytes().unwrap(), bytes);
    }
}

#[test]
fn native_opc_sidecar_crud_and_failed_updates_preserve_the_workbook() {
    use litchi_xlsb::package::PackageError;
    let bytes = native(FILES[2]);
    let p = Package::from_bytes(&bytes).unwrap();
    let part = PackURI::new("/custom/native-check.bin").unwrap();
    let workbook = PackURI::new("/xl/workbook.bin").unwrap();
    let mut edit = p.edit_opc().unwrap();
    edit.update(|opc| {
        opc.try_add_part(Box::new(BlobPart::new(
            part.clone(),
            "application/octet-stream".into(),
            b"NATIVE_PACKAGE_CHECK".to_vec(),
        )))?;
        opc.rels_mut().try_add_relationship(
            "urn:litchi:native-check".into(),
            "custom/native-check.bin".into(),
            "rIdNativeCheck".into(),
            TargetMode::Internal,
        )?;
        Ok(())
    })
    .unwrap();
    let error = edit
        .update(|opc| -> Result<(), PackageError> {
            opc.get_part_mut(&workbook)?.set_blob(vec![0xff]);
            Err(PackageError::InvalidFormat("native rollback check"))
        })
        .unwrap_err();
    assert!(matches!(
        error,
        PackageError::InvalidFormat("native rollback check")
    ));
    assert_eq!(
        edit.opc().get_part(&workbook).unwrap().blob(),
        members(&bytes)["xl/workbook.bin"]
    );
    let created = edit.commit().unwrap();
    let after = save(created.package());
    let output = after.to_bytes().unwrap();
    assert_eq!(
        changed(&bytes, &output),
        [
            "[Content_Types].xml",
            "_rels/.rels",
            "custom/native-check.bin"
        ]
    );
    assert_same_payloads_and_graph(
        &bytes,
        &after
            .apply_opc(&created.patch().inverse())
            .unwrap()
            .package()
            .to_bytes()
            .unwrap(),
    );
    let mut update = after.edit_opc().unwrap();
    update
        .update(|opc| {
            opc.get_part_mut(&part)?
                .set_blob(b"UPDATED_NATIVE_CHECK".to_vec());
            Ok(())
        })
        .unwrap();
    let changed_commit = update.commit().unwrap();
    let updated = save(changed_commit.package());
    assert_eq!(
        changed(&output, &updated.to_bytes().unwrap()),
        ["custom/native-check.bin"]
    );
    let mut remove = updated.edit_opc().unwrap();
    remove
        .update(|opc| {
            assert!(opc.rels_mut().remove("rIdNativeCheck").is_some());
            assert!(opc.remove_part(&part));
            Ok(())
        })
        .unwrap();
    let removed = remove.commit().unwrap();
    let restored = save(removed.package());
    assert_same_payloads_and_graph(&bytes, &restored.to_bytes().unwrap());
    assert_same_payloads_and_graph(
        &updated.to_bytes().unwrap(),
        &restored
            .apply_opc(&removed.patch().inverse())
            .unwrap()
            .package()
            .to_bytes()
            .unwrap(),
    );
    // A successful low-level update still must pass workbook validation at
    // publication. Corrupt only a derivative, leaving the native source exact.
    let mut invalid = after.edit_opc().unwrap();
    invalid
        .update(|opc| {
            let mut truncated = opc.get_part(&workbook)?.blob().to_vec();
            truncated.pop();
            opc.get_part_mut(&workbook)?.set_blob(truncated);
            Ok(())
        })
        .unwrap();
    assert!(matches!(
        invalid.commit(),
        Err(PackageError::Wire(raw::Error::Truncated {
            stage: raw::Stage::Length,
            ..
        }))
    ));
    assert_eq!(after.to_bytes().unwrap(), output);
    assert_eq!(p.to_bytes().unwrap(), bytes);
}
