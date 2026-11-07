//! Native Word settings and OpenType extensions; fonts and macros stay inert.
use litchi_docx::font::open_type::{self, Ligatures, NumForm, NumSpacing, StyleSet, StyleSetId};
use litchi_docx::settings::{CompatibilitySetting, Guid, OnOff, Settings};
use litchi_docx::{Package, paragraph::Run};
use quick_xml::events::Event;
use std::collections::BTreeMap;

const VARIANTS: [&str; 4] = [
    "word-math-navigation.docx",
    "word-math-navigation.dotx",
    "word-recorded-macro.docm",
    "word-recorded-macro.dotm",
];
const COMPAT_URI: &str = "http://schemas.microsoft.com/office/word";
fn native(name: &str) -> Package {
    Package::open(
        std::path::Path::new(env!("CARGO_MANIFEST_DIR"))
            .join("../../test-data/microsoft-office/word")
            .join(name),
    )
    .unwrap()
}
fn members(p: &Package) -> BTreeMap<String, Vec<u8>> {
    let archive = litchi_opc::phys_pkg::OwnedPhysPkgReader::from_bytes(
        p.to_bytes().unwrap(),
        litchi_opc::SourceLimits::default(),
    )
    .unwrap();
    archive
        .member_names()
        .unwrap()
        .into_iter()
        .map(|name| {
            let bytes = archive.read_member(&name).unwrap();
            (name, bytes)
        })
        .collect()
}
fn assert_only_settings_changed(source: &Package, output: &Package) {
    let a = members(source);
    let b = members(output);
    assert_eq!(a.keys().collect::<Vec<_>>(), b.keys().collect::<Vec<_>>());
    assert_eq!(
        a.iter()
            .filter(|(name, data)| b[*name] != **data)
            .map(|(name, _)| name.as_str())
            .collect::<Vec<_>>(),
        ["word/settings.xml"]
    );
}
// Remove only the known extension elements, retaining all other source bytes.
fn without_settings_extensions(xml: &[u8]) -> Vec<u8> {
    without_elements(
        xml,
        &[
            b"w14:docId",
            b"w15:docId",
            b"w15:chartTrackingRefBased",
            b"w14:conflictMode",
            b"w14:discardImageEditingData",
            b"w14:defaultImageDpi",
        ],
    )
}
fn without_open_type(xml: &[u8]) -> Vec<u8> {
    without_elements(
        xml,
        &[
            b"w14:ligatures",
            b"w14:numForm",
            b"w14:numSpacing",
            b"w14:stylisticSets",
            b"w14:cntxtAlts",
        ],
    )
}
fn without_elements(xml: &[u8], names: &[&[u8]]) -> Vec<u8> {
    let mut reader = quick_xml::Reader::from_reader(xml);
    let mut output = Vec::new();
    let mut copied = 0;
    let mut depth = 0usize;
    let mut start = None;
    loop {
        let before = reader.buffer_position() as usize;
        let event = reader.read_event().unwrap();
        let after = reader.buffer_position() as usize;
        let owned = match &event {
            Event::Start(e) | Event::Empty(e) => names.contains(&e.name().as_ref()),
            _ => false,
        };
        match event {
            Event::Start(_) => {
                if owned {
                    assert!(start.is_none());
                    start = Some((before, depth));
                }
                depth += 1;
            },
            Event::Empty(_) if owned => {
                output.extend_from_slice(&xml[copied..before]);
                copied = after;
            },
            Event::End(_) => {
                depth -= 1;
                if start.is_some_and(|(_, d)| d == depth) {
                    let (begin, _) = start.take().unwrap();
                    output.extend_from_slice(&xml[copied..begin]);
                    copied = after;
                }
            },
            Event::Eof => break,
            _ => {},
        }
    }
    assert!(start.is_none());
    output.extend_from_slice(&xml[copied..]);
    output
}
fn assert_compatibility(package: &Package) {
    let s = package.document().unwrap().settings().unwrap().unwrap();
    assert_eq!(s.compatibility_mode(), Some(15));
    assert!(s.compatibility_options().is_empty());
    assert_eq!(
        s.compatibility_settings()
            .iter()
            .map(|v| (v.name(), v.uri(), v.value()))
            .collect::<Vec<_>>(),
        [
            ("compatibilityMode", COMPAT_URI, "15"),
            (
                "overrideTableStyleFontSizeAndJustification",
                COMPAT_URI,
                "1"
            ),
            ("enableOpenTypeFeatures", COMPAT_URI, "1"),
            ("doNotFlipMirrorIndents", COMPAT_URI, "1"),
            ("differentiateMultirowTableHeaders", COMPAT_URI, "1"),
            ("useWord2013TrackBottomHyphenation", COMPAT_URI, "0"),
        ]
    );
    assert!(
        s.compatibility_setting("enableOpenTypeFeatures", "urn:other")
            .is_none()
    );
}
fn reopened_settings_commit(
    source: &Package,
    commit: &litchi_docx::settings::extensions::Commit,
) -> Package {
    let file = tempfile::NamedTempFile::with_suffix(".docx").unwrap();
    commit.package().save(file.path()).unwrap();
    let reopened = Package::open(file.path()).unwrap();
    assert_only_settings_changed(source, &reopened);
    assert_compatibility(&reopened);
    assert!(matches!(
        reopened.apply_settings_extensions(commit.patch()),
        Err(litchi_docx::Error::PatchConflict { .. })
    ));
    let replayed = source
        .apply_settings_extensions(commit.patch())
        .unwrap()
        .into_package();
    assert_eq!(members(&replayed), members(&reopened));
    let restored = reopened
        .apply_settings_extensions(&commit.patch().inverse())
        .unwrap()
        .into_package();
    assert_eq!(members(&restored), members(source));
    reopened
}

#[test]
fn native_settings_ids_authored_defaults_and_compatibility_triples_are_distinct() {
    for (name, id, guid) in [
        (
            VARIANTS[0],
            0x338555c9,
            Some("{AA7E02D1-2356-234F-B79F-09B57E04FCB0}"),
        ),
        (VARIANTS[1], 0x338555c9, None),
        (
            VARIANTS[2],
            0x1eb37e66,
            Some("{AE65EE72-34DD-E848-9C96-D70156049F27}"),
        ),
        (VARIANTS[3], 0x1eb37e66, None),
    ] {
        let p = native(name);
        let source = p.to_bytes().unwrap();
        let snapshot = p.settings_extensions().unwrap();
        let e = snapshot.extensions();
        assert_eq!(e.document_id(), Some(id));
        assert_eq!(e.has_source_document_id(), guid.is_some());
        assert_eq!(
            e.source_document_id(),
            guid.map(|g| Guid::parse(g).unwrap()).as_ref()
        );
        assert_eq!(e.chart_tracking_ref_based(), Some(OnOff::default_on()));
        assert!(e.chart_tracking_ref_based().unwrap().effective());
        assert_eq!(e.conflict_mode(), None);
        assert_eq!(e.discard_image_editing_data(), None);
        assert_eq!(e.default_image_dpi(), None);
        assert_eq!(e.unknown().count(), 1);
        assert!(
            std::str::from_utf8(e.unknown().next().unwrap().xml())
                .unwrap()
                .contains("<m:mathPr")
        );
        assert_eq!(snapshot.xml_bytes(), members(&p)["word/settings.xml"]);
        assert_compatibility(&p);
        let mut edit = p.edit_settings_extensions();
        edit.set_document_id(Some(id)).unwrap();
        edit.set_chart_tracking_ref_based(Some(OnOff::default_on()))
            .unwrap();
        let noop = edit.commit().unwrap();
        assert!(!noop.diagnostics().changed());
        assert_eq!(noop.package().to_bytes().unwrap(), source);
    }
}

#[test]
fn native_settings_extension_updates_preserve_compatibility_math_and_macro_members() {
    for name in VARIANTS {
        let p = native(name);
        let source = p.to_bytes().unwrap();
        let before = p.settings_extensions().unwrap();
        let guid = Guid::parse("{12345678-1234-4321-ABCD-123456789ABC}").unwrap();
        let mut edit = p.edit_settings_extensions();
        edit.set_document_id(Some(0x12345678))
            .unwrap()
            .set_source_document_id(Some(guid))
            .unwrap()
            .set_chart_tracking_ref_based(Some(OnOff::off()))
            .unwrap()
            .set_conflict_mode(Some(OnOff::on()))
            .unwrap()
            .set_discard_image_editing_data(Some(OnOff::off()))
            .unwrap()
            .set_default_image_dpi(Some(144))
            .unwrap();
        let staged = edit.extensions().unwrap().clone();
        for invalid in [0, 0x80000000, u32::MAX] {
            assert!(
                matches!(edit.set_document_id(Some(invalid)), Err(litchi_docx::Error::Invalid(message)) if message == "Word settings paragraph context ID must be greater than 0 and less than 0x80000000")
            );
            assert_eq!(edit.extensions().unwrap(), &staged);
        }
        let commit = edit.commit().unwrap();
        let after = reopened_settings_commit(&p, &commit);
        let snapshot = after.settings_extensions().unwrap();
        assert_eq!(snapshot.extensions(), &staged);
        assert_eq!(
            without_settings_extensions(before.xml_bytes()),
            without_settings_extensions(snapshot.xml_bytes())
        );
        assert_eq!(p.to_bytes().unwrap(), source);
    }
}

#[test]
fn native_settings_extension_removal_and_reset_preserve_unowned_records() {
    for name in VARIANTS {
        let p = native(name);
        let original = p.to_bytes().unwrap();
        let before = p.settings_extensions().unwrap();
        let mut edit = p.edit_settings_extensions();
        edit.set_document_id(None)
            .unwrap()
            .set_source_document_id(None)
            .unwrap()
            .set_chart_tracking_ref_based(None)
            .unwrap();
        let commit = edit.commit().unwrap();
        let reopened = reopened_settings_commit(&p, &commit);
        let after = reopened.settings_extensions().unwrap();
        assert_eq!(after.extensions().document_id(), None);
        assert!(!after.extensions().has_source_document_id());
        assert_eq!(after.extensions().chart_tracking_ref_based(), None);
        assert_eq!(
            after.extensions().unknown().collect::<Vec<_>>(),
            before.extensions().unknown().collect::<Vec<_>>()
        );
        assert_eq!(
            after.xml_bytes(),
            without_settings_extensions(before.xml_bytes())
        );
        let mut discarded = p.edit_settings_extensions();
        discarded.set_default_image_dpi(Some(300)).unwrap();
        discarded.reset();
        assert_eq!(
            discarded.commit().unwrap().package().to_bytes().unwrap(),
            original
        );
    }
}

#[test]
fn native_compatibility_values_round_trip_through_the_typed_settings_codec() {
    // This is a detached typed codec result, not a source-package rewrite.
    let p = native(VARIANTS[0]);
    let original = p.to_bytes().unwrap();
    let xml = members(&p).remove("word/settings.xml").unwrap();
    let mut settings = Settings::parse(&xml).unwrap();
    assert_eq!(settings.compatibility_mode(), Some(15));
    let mut values = settings.compatibility_settings().to_vec();
    values[0] = CompatibilitySetting::new("compatibilityMode", COMPAT_URI, "14");
    values.retain(|v| v.name() != "enableOpenTypeFeatures");
    values.push(CompatibilitySetting::new(
        "nativeRegression",
        "urn:litchi:regression",
        "A & B",
    ));
    settings.set_compatibility_settings(values.clone());
    let fragment = settings.to_xml("w");
    let wrapped = format!(
        "<w:settings xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\">{fragment}</w:settings>"
    );
    let parsed = Settings::parse(wrapped.as_bytes()).unwrap();
    assert_eq!(parsed.compatibility_mode(), Some(14));
    assert_eq!(parsed.compatibility_settings(), values);
    assert_eq!(parsed.extensions(), settings.extensions());
    assert!(fragment.contains("A &amp; B"));
    assert_eq!(p.to_bytes().unwrap(), original);
}

#[test]
fn native_macro_runs_preserve_explicit_no_ligatures_and_detached_opentype_edits() {
    for name in [VARIANTS[2], VARIANTS[3]] {
        let p = native(name);
        let original = p.to_bytes().unwrap();
        let d = p.document().unwrap();
        let paragraphs = d.paragraphs().unwrap();
        assert_eq!(paragraphs.len(), 5);
        for (index, paragraph) in paragraphs.iter().enumerate() {
            let runs = paragraph.runs().unwrap();
            assert_eq!(runs.len(), 1);
            let run = &runs[0];
            let baseline = run.open_type().unwrap();
            assert_eq!(baseline.ligatures, Some(Ligatures::None));
            assert_eq!(baseline.num_form, None);
            assert_eq!(baseline.num_spacing, None);
            assert_eq!(baseline.cntxt_alts, None);
            assert!(!baseline.stylistic_sets_present());
            let snapshot = run.open_type_snapshot().unwrap();
            assert_eq!(snapshot.open_type(), &baseline);
            if index != 2 {
                continue;
            }
            assert_eq!(run.text().unwrap(), "MACRO_FORMAT_TARGET");
            assert_eq!(run.bold().unwrap(), Some(true));
            let mut detached = Run::new(snapshot.xml_bytes().to_vec());
            let mut changed = baseline.clone();
            changed.ligatures = Some(Ligatures::StandardContextual);
            changed.num_form = Some(NumForm::OldStyle);
            changed.num_spacing = Some(NumSpacing::Tabular);
            changed.cntxt_alts = Some(open_type::OnOff::off());
            changed
                .set_stylistic_sets(Some(vec![
                    StyleSet::new(StyleSetId::new(1).unwrap()),
                    StyleSet::new(StyleSetId::new(4).unwrap()).with_enabled(Some(false)),
                ]))
                .unwrap();
            detached.set_open_type(changed.clone()).unwrap();
            let reread = detached.open_type_snapshot().unwrap();
            assert_eq!(reread.open_type(), &changed);
            assert_eq!(
                without_open_type(snapshot.xml_bytes()),
                without_open_type(reread.xml_bytes())
            );
            assert_eq!(detached.text().unwrap(), run.text().unwrap());
            assert_eq!(detached.bold().unwrap(), Some(true));
            assert!(
                std::str::from_utf8(reread.xml_bytes())
                    .unwrap()
                    .contains("w:rFonts")
            );
            let mut reparsed = Run::new(reread.xml_bytes().to_vec());
            reparsed.set_open_type(baseline.clone()).unwrap();
            assert_eq!(reparsed.open_type().unwrap(), baseline);
            assert_eq!(reparsed.text().unwrap(), "MACRO_FORMAT_TARGET");
            assert_eq!(
                run.open_type_snapshot().unwrap().xml_bytes(),
                snapshot.xml_bytes()
            );
        }
        assert_eq!(p.to_bytes().unwrap(), original);
    }
}
