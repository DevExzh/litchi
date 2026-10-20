use litchi_docx::Package;
use litchi_docx::settings::extensions::Snapshot;
use litchi_docx::settings::{
    DocumentId, Extension, Extensions, Guid, OnOff, Settings, WORD_2010_NAMESPACE,
    WORD_2012_NAMESPACE,
};
use litchi_opc::PackURI;

const W: &str = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
const W14: &str = WORD_2010_NAMESPACE;
const W15: &str = WORD_2012_NAMESPACE;
const MC: &str = "http://schemas.openxmlformats.org/markup-compatibility/2006";

fn settings_xml(body: &str) -> String {
    format!(
        r#"<w:settings xmlns:w="{W}" xmlns:w14="{W14}" xmlns:w15="{W15}" xmlns:mc="{MC}" mc:Ignorable="w14 w15">{body}</w:settings>"#
    )
}

#[test]
fn parses_all_typed_extensions_in_order_and_round_trips_opaque_markup() {
    let guid = Guid::parse("{01234567-89AB-CDEF-0123-456789ABCDEF}").unwrap();
    let xml = settings_xml(
        r#"<w15:chartTrackingRefBased w:val="0"/><w14:docId w14:val="7F00AA10"/><x:future xmlns:x="urn:future"><x:child><![CDATA[a < b]]></x:child></x:future><w14:conflictMode/><w14:discardImageEditingData w14:val="0"/><w14:defaultImageDpi w14:val="-150"/><w15:docId w15:val="{01234567-89AB-CDEF-0123-456789ABCDEF}"/>"#,
    );

    let settings = Settings::parse(xml.as_bytes()).unwrap();
    let extensions = settings.extensions();
    assert_eq!(extensions.len(), 7);
    assert_eq!(extensions.chart_tracking_ref_based(), Some(OnOff::off()));
    assert_eq!(extensions.document_id(), Some(0x7F00AA10));
    assert_eq!(extensions.conflict_mode(), Some(OnOff::default_on()));
    assert_eq!(extensions.discard_image_editing_data(), Some(OnOff::off()));
    assert_eq!(extensions.default_image_dpi(), Some(-150));
    assert_eq!(extensions.source_document_id(), Some(&guid));
    assert!(extensions.has_source_document_id());
    assert_eq!(extensions.unknown().count(), 1);
    let unknown = extensions.unknown().next().unwrap().xml();
    assert!(unknown.starts_with(b"<x:future"));
    assert!(
        unknown
            .windows(b"xmlns:x=\"urn:future\"".len())
            .any(|window| { window == b"xmlns:x=\"urn:future\"" })
    );

    let fragment = settings.to_xml("w");
    let reparsed = Settings::parse(settings_xml(&fragment).as_bytes()).unwrap();
    assert_eq!(reparsed.extensions(), extensions);
    assert!(fragment.contains("w15:chartTrackingRefBased"));
    assert!(fragment.contains("w14:defaultImageDpi"));
    assert!(fragment.contains("x:future"));
}

#[test]
fn preserves_an_optional_source_document_id_without_a_guid() {
    let settings = Settings::parse(settings_xml(r#"<w15:docId/>"#).as_bytes()).unwrap();
    let extensions = settings.extensions();
    assert!(extensions.has_source_document_id());
    assert_eq!(extensions.source_document_id(), None);
    assert!(matches!(
        extensions.iter().next(),
        Some(Extension::DocumentId(DocumentId::Source(None)))
    ));
    assert!(extensions.to_xml("w").contains("<w15:docId"));
}

#[test]
fn authoring_uses_checked_bounded_values_and_prefix_free_types() {
    let guid = Guid::from_bytes([0xAB; 16]);
    let mut extensions = Extensions::new();
    extensions
        .set_chart_tracking_ref_based(Some(OnOff::off()))
        .unwrap()
        .set_document_id(Some(1))
        .unwrap()
        .set_source_document_id(Some(guid))
        .unwrap()
        .set_conflict_mode(Some(OnOff::on()))
        .unwrap()
        .set_discard_image_editing_data(Some(OnOff::off()))
        .unwrap()
        .set_default_image_dpi(Some(i32::MIN))
        .unwrap();

    assert!(extensions.document_id().is_some());
    assert_eq!(extensions.source_document_id(), Some(&guid));
    assert_eq!(extensions.default_image_dpi(), Some(i32::MIN));
    let fragment = extensions.to_xml("w");
    assert!(fragment.contains("w14:defaultImageDpi"));
    assert!(fragment.contains("w14:val=\"-2147483648\""));
}

#[test]
fn rejects_malformed_values_duplicate_known_extensions_and_nested_typed_markup() {
    let malformed = [
        r#"<w14:conflictMode w14:val="maybe"/>"#,
        r#"<w14:docId w14:val="00000000"/>"#,
        r#"<w14:docId w14:val="80000000"/>"#,
        r#"<w14:docId w14:val="1234"/>"#,
        r#"<w15:docId w15:val="{01234567-89ab-CDEF-0123-456789ABCDEF}"/>"#,
        r#"<w14:defaultImageDpi/>"#,
        r#"<w14:defaultImageDpi w14:val="2147483648"/>"#,
        r#"<w14:conflictMode><w14:child/></w14:conflictMode>"#,
        r#"<w14:conflictMode/><w14:conflictMode/>"#,
        r#"<x:future xmlns:x="urn:future"><x:child></x:future>"#,
    ];
    for body in malformed {
        assert!(
            Settings::parse(settings_xml(body).as_bytes()).is_err(),
            "malformed settings extension was accepted: {body}"
        );
    }
}

fn package_with_settings(xml: &str) -> Package {
    let package = Package::new().unwrap();
    let mut edit = package.edit_opc();
    let settings = PackURI::new("/word/settings.xml").unwrap();
    edit.package_mut()
        .get_part_mut(&settings)
        .unwrap()
        .set_blob(xml.as_bytes().to_vec());
    edit.commit().unwrap().into_package()
}

#[test]
fn settings_extension_snapshots_are_source_preserving_and_reversible() {
    let source = settings_xml(
        r#"  <w:zoom w:val="75"/>
        <w15:chartTrackingRefBased w:val="1"/>
        <x:future xmlns:x="urn:future"><x:payload>retain</x:payload></x:future>
        <w14:conflictMode w14:val="on"/>"#,
    );
    let package = package_with_settings(&source);
    let snapshot = package.settings_extensions().unwrap();
    assert_eq!(snapshot.xml_bytes(), source.as_bytes());
    assert_eq!(
        snapshot.extensions().chart_tracking_ref_based(),
        Some(OnOff::on())
    );
    assert_eq!(snapshot.extensions().conflict_mode(), Some(OnOff::on()));

    let no_op = package.edit_settings_extensions().commit().unwrap();
    assert!(!no_op.diagnostics().changed());
    assert_eq!(
        no_op.package().settings_extensions().unwrap().xml_bytes(),
        source.as_bytes()
    );

    let mut edit = package.edit_settings_extensions();
    edit.set_default_image_dpi(Some(144)).unwrap();
    let changed = edit.commit().unwrap();
    let changed_snapshot = changed.package().settings_extensions().unwrap();
    let output = changed_snapshot.xml_bytes();
    assert!(String::from_utf8_lossy(output).contains(r#"<w:zoom w:val="75"/>"#));
    assert!(String::from_utf8_lossy(output).contains(r#"w14:val="on""#));
    assert!(String::from_utf8_lossy(output).contains("x:payload>retain"));
    assert_eq!(
        Snapshot::from_xml(output.to_vec())
            .unwrap()
            .extensions()
            .default_image_dpi(),
        Some(144)
    );
    // The source package is never changed by an edit.
    assert_eq!(
        package.settings_extensions().unwrap().xml_bytes(),
        source.as_bytes()
    );

    let restored = changed
        .package()
        .apply_settings_extensions(&changed.patch().inverse())
        .unwrap();
    assert_eq!(
        restored
            .package()
            .settings_extensions()
            .unwrap()
            .xml_bytes(),
        source.as_bytes()
    );

    let mut chart_edit = package.edit_settings_extensions();
    chart_edit
        .set_chart_tracking_ref_based(Some(OnOff::off()))
        .unwrap();
    let chart_changed = chart_edit.commit().unwrap();
    let chart_snapshot = chart_changed.package().settings_extensions().unwrap();
    assert!(String::from_utf8_lossy(chart_snapshot.xml_bytes()).contains(
        r#"xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" w:val="false""#
    ));
    assert_eq!(
        chart_snapshot.extensions().chart_tracking_ref_based(),
        Some(OnOff::off())
    );
}

#[test]
fn settings_extension_transactions_distinguish_absence_default_on_and_explicit_off() {
    let source = package_with_settings(&settings_xml(""));
    assert_eq!(
        source
            .settings_extensions()
            .unwrap()
            .extensions()
            .conflict_mode(),
        None
    );

    let mut add = source.edit_settings_extensions();
    add.set_conflict_mode(Some(OnOff::default_on())).unwrap();
    let authored = add.commit().unwrap().into_package();
    let authored_snapshot = authored.settings_extensions().unwrap();
    assert_eq!(
        authored_snapshot.extensions().conflict_mode(),
        Some(OnOff::default_on())
    );
    assert!(String::from_utf8_lossy(authored_snapshot.xml_bytes()).contains(
        r#"<w14:conflictMode xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml"/>"#
    ));

    let mut turn_off = authored.edit_settings_extensions();
    turn_off.set_conflict_mode(Some(OnOff::off())).unwrap();
    let off = turn_off.commit().unwrap().into_package();
    let off_snapshot = off.settings_extensions().unwrap();
    assert_eq!(
        off_snapshot.extensions().conflict_mode(),
        Some(OnOff::off())
    );
    assert!(String::from_utf8_lossy(off_snapshot.xml_bytes()).contains(r#"w14:val="false""#));

    let mut clear = off.edit_settings_extensions();
    clear.set_conflict_mode(None).unwrap();
    let absent = clear.commit().unwrap().into_package();
    let absent_snapshot = absent.settings_extensions().unwrap();
    assert_eq!(absent_snapshot.extensions().conflict_mode(), None);
    assert!(!String::from_utf8_lossy(absent_snapshot.xml_bytes()).contains("conflictMode"));
}

#[test]
fn settings_extension_transactions_reject_stale_sources_and_invalid_changes_atomically() {
    let source = package_with_settings(&settings_xml(r#"<w14:conflictMode/>"#));
    let mut edit = source.edit_settings_extensions();
    edit.set_default_image_dpi(Some(96)).unwrap();
    let commit = edit.commit().unwrap();

    let alternate = package_with_settings(&settings_xml(r#"<w14:conflictMode w14:val="0"/>"#));
    assert!(alternate.apply_settings_extensions(commit.patch()).is_err());

    let source_extensions = source.settings_extensions().unwrap().extensions().clone();
    let mut invalid = source.edit_settings_extensions();
    assert!(invalid.set_document_id(Some(0)).is_err());
    assert_eq!(invalid.extensions().unwrap(), &source_extensions);
    assert_eq!(
        source.settings_extensions().unwrap().xml_bytes(),
        settings_xml(r#"<w14:conflictMode/>"#).as_bytes()
    );
}
