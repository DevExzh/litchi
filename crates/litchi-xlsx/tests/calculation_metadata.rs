use std::sync::Arc;

use litchi_opc::constants::{content_type as ct, relationship_type as rt};
use litchi_opc::{BlobPart, PackURI, Part, TargetMode};
use litchi_xlsx::calculation_properties::{
    Feature, Features, Limits, Mode, Patch, Properties, Snapshot,
};
use litchi_xlsx::formula::Formula;
use litchi_xlsx::{Error, Package};

const SIGNATURE_REL: &str =
    "http://schemas.openxmlformats.org/package/2006/relationships/digital-signature/signature";

fn feature_names(snapshot: &Snapshot) -> Vec<&str> {
    snapshot
        .features()
        .expect("calculation features")
        .iter()
        .map(Feature::as_str)
        .collect()
}

fn set_metadata(package: &mut Package, properties: Properties, features: Features) -> Patch {
    let mut edit = package
        .edit_calculation_metadata()
        .expect("calculation metadata edit");
    assert!(edit.set_properties(properties));
    assert!(edit.set_features(features));
    let commit = edit.commit().expect("calculation metadata commit");
    assert!(commit.changed());
    commit.patch().clone()
}

#[test]
fn absent_metadata_round_trips_typed_properties_and_inert_feature_occurrences() {
    let mut package = Package::create().expect("fresh package");
    let absent = package.calculation_metadata().expect("initial metadata");
    assert!(absent.properties().is_none());
    assert!(absent.features().is_none());

    let properties = Properties::new()
        .with_calculation_id(Some(2026))
        .with_calculation_mode(Some(Mode::Manual))
        .with_full_calculation_on_load(Some(true))
        .with_iteration_count(Some(37))
        .with_iteration_delta(Some(0.25))
        .expect("valid iteration delta")
        .with_calculation_completed(Some(false));
    let expected_names = ["Case", "case", "\u{8ba1}\u{7b97}<&\"'>", "Case"];
    let features = Features::try_from_vec(
        expected_names
            .iter()
            .map(|name| Feature::new(*name).expect("valid inert feature name"))
            .collect(),
    )
    .expect("nonempty features");
    set_metadata(&mut package, properties.clone(), features);

    let authored = package.calculation_metadata().expect("authored metadata");
    assert!(
        authored
            .properties()
            .expect("calcPr")
            .same_specification(&properties)
    );
    assert_eq!(feature_names(&authored), expected_names);
    assert_eq!(authored.features().unwrap().occurrence_count("Case"), 2);
    assert_eq!(authored.features().unwrap().occurrence_count("case"), 1);
    let xml = std::str::from_utf8(authored.source_xml()).expect("UTF-8 workbook XML");
    assert!(xml.contains("&lt;"));
    assert!(xml.contains("&amp;"));
    assert!(!xml.contains("\u{8ba1}\u{7b97}<&"));

    let directory = tempfile::tempdir().expect("temporary directory");
    let path = directory.path().join("calculation-metadata.xlsx");
    std::fs::write(&path, package.to_bytes().expect("save bytes")).expect("save package");
    let reopened = Package::open(&path).expect("reopen package");
    let reopened_metadata = reopened.calculation_metadata().expect("reopened metadata");
    assert!(
        reopened_metadata
            .properties()
            .expect("reopened calcPr")
            .same_specification(&properties)
    );
    assert_eq!(feature_names(&reopened_metadata), expected_names);

    let workbook = reopened.workbook().expect("workbook facade");
    let workbook_metadata = workbook
        .calculation_metadata()
        .expect("workbook calculation metadata");
    assert_eq!(feature_names(&workbook_metadata), expected_names);
    assert_eq!(
        workbook_metadata.properties().unwrap().calculation_id(),
        2026
    );
}

#[test]
fn patch_inverse_restores_exact_source_and_stale_application_is_atomic() {
    let baseline = Package::create().expect("baseline package");
    let baseline_xml = baseline
        .calculation_metadata()
        .expect("baseline metadata")
        .source_xml()
        .to_vec();
    let mut changed = baseline.clone();
    let patch = set_metadata(
        &mut changed,
        Properties::new().with_calculation_id(Some(17)),
        Features::new(Feature::new("X").unwrap()),
    );
    assert!(!patch.is_empty());
    assert!(patch.before().properties().is_none());
    assert_eq!(patch.after().properties().unwrap().calculation_id(), 17);

    changed
        .apply_calculation_metadata_patch(&patch.inverse())
        .expect("apply inverse");
    assert_eq!(
        changed
            .calculation_metadata()
            .expect("restored metadata")
            .source_xml(),
        baseline_xml
    );

    let mut stale = baseline;
    let mut stale_edit = stale.edit_calculation_metadata().expect("stale edit");
    stale_edit.set_properties(Properties::new().with_calculation_id(Some(99)));
    stale_edit.commit().expect("publish stale source");
    let stale_xml = stale
        .calculation_metadata()
        .expect("stale metadata")
        .source_xml()
        .to_vec();
    assert!(matches!(
        stale.apply_calculation_metadata_patch(&patch),
        Err(Error::PatchConflict { .. })
    ));
    assert_eq!(
        stale
            .calculation_metadata()
            .expect("unchanged stale metadata")
            .source_xml(),
        stale_xml
    );
}

#[test]
fn feature_name_and_output_limits_fail_without_partial_publication() {
    let mut package = Package::create().expect("package");
    set_metadata(
        &mut package,
        Properties::new(),
        Features::try_from_vec(vec![
            Feature::new("long").unwrap(),
            Feature::new("two").unwrap(),
        ])
        .unwrap(),
    );
    assert!(
        package
            .calculation_metadata_with_limits(&Limits::new().with_max_features(1).unwrap())
            .is_err()
    );
    assert!(
        package
            .calculation_metadata_with_limits(
                &Limits::new().with_max_feature_name_bytes(3).unwrap()
            )
            .is_err()
    );

    let before = package
        .calculation_metadata()
        .expect("before bounded edit")
        .source_xml()
        .to_vec();
    let limits = Limits::new().with_max_output_bytes(1).unwrap();
    let mut edit = package
        .edit_calculation_metadata_with_limits(&limits)
        .expect("bounded edit");
    edit.set_properties(Properties::new().with_calculation_id(Some(8)));
    assert!(edit.commit().is_err());
    assert_eq!(
        package
            .calculation_metadata()
            .expect("after rejected edit")
            .source_xml(),
        before
    );
}

#[test]
fn metadata_removal_restores_workbook_markup_and_preserves_unrelated_part_bytes() {
    let mut raw = Package::create().expect("package").into_plain_opc();
    let workbook_name = raw
        .main_document_part()
        .expect("workbook part")
        .partname()
        .clone();
    let workbook_xml = std::str::from_utf8(
        raw.get_part(&workbook_name)
            .expect("workbook")
            .blob(),
    )
    .expect("UTF-8 workbook")
    .replace(
        "</workbook>",
        r#"<extLst><ext uri="{KEEP}"><keep:data xmlns:keep="urn:keep">&amp;opaque</keep:data></ext></extLst></workbook>"#,
    )
    .into_bytes();
    raw.get_part_mut(&workbook_name)
        .expect("mutable workbook")
        .set_blob(workbook_xml.clone());
    let opaque_name = PackURI::new("/custom/opaque.bin").unwrap();
    let opaque = b"\0unrelated-package-bytes\xff".to_vec();
    raw.try_add_part(Box::new(BlobPart::new(
        opaque_name.clone(),
        "application/octet-stream".to_owned(),
        opaque.clone(),
    )))
    .unwrap();
    let mut package = Package::from_opc(raw).expect("adopt fixture");

    set_metadata(
        &mut package,
        Properties::new().with_calculation_id(Some(1)),
        Features::new(Feature::new("temporary").unwrap()),
    );
    let mut remove = package.edit_calculation_metadata().expect("remove edit");
    assert!(remove.remove_properties());
    assert!(remove.remove_features());
    remove.commit().expect("remove metadata");
    assert_eq!(
        package
            .calculation_metadata()
            .expect("removed metadata")
            .source_xml(),
        workbook_xml
    );
    assert_eq!(
        package
            .into_plain_opc()
            .get_part(&opaque_name)
            .expect("unrelated part")
            .blob(),
        opaque
    );
}

#[test]
fn signed_no_op_is_exact_but_changed_metadata_is_refused() {
    let mut raw = Package::create().expect("package").into_plain_opc();
    let origin_name = PackURI::new("/_xmlsignatures/origin.sigs").unwrap();
    let mut origin = BlobPart::new(
        origin_name,
        ct::OPC_DIGITAL_SIGNATURE_ORIGIN.to_owned(),
        b"origin".to_vec(),
    );
    origin
        .rels_mut()
        .try_add_relationship(
            SIGNATURE_REL.to_owned(),
            "sig1.xml".to_owned(),
            "rIdSignature".to_owned(),
            TargetMode::Internal,
        )
        .unwrap();
    raw.try_add_part(Box::new(origin)).unwrap();
    raw.try_add_part(Box::new(BlobPart::new(
        PackURI::new("/_xmlsignatures/sig1.xml").unwrap(),
        ct::OPC_DIGITAL_SIGNATURE_XMLSIGNATURE.to_owned(),
        b"signature".to_vec(),
    )))
    .unwrap();
    raw.rels_mut()
        .try_add_relationship(
            rt::DIGITAL_SIGNATURE_ORIGIN.to_owned(),
            "_xmlsignatures/origin.sigs".to_owned(),
            "rIdSignatureOrigin".to_owned(),
            TargetMode::Internal,
        )
        .unwrap();
    let mut package = Package::from_opc(raw).expect("signed package");
    let source = package
        .calculation_metadata()
        .unwrap()
        .source_arc()
        .unwrap();
    let no_op = package
        .edit_calculation_metadata()
        .unwrap()
        .commit()
        .expect("signed no-op");
    assert!(!no_op.changed());
    assert!(no_op.patch().is_empty());
    assert!(Arc::ptr_eq(
        &source,
        &no_op.snapshot().source_arc().unwrap()
    ));

    let mut changed = package.edit_calculation_metadata().unwrap();
    changed.set_properties(Properties::new().with_calculation_id(Some(1)));
    assert!(matches!(changed.commit(), Err(Error::Signed)));
    assert_eq!(
        package.calculation_metadata().unwrap().source_xml(),
        &*source
    );
}

#[test]
fn formula_edit_exposes_recalculation_invalidation_through_workbook_owner() {
    let mut package = Package::create().expect("package");
    let mut metadata = package.edit_calculation_metadata().unwrap();
    metadata.set_properties(
        Properties::new()
            .with_calculation_id(Some(123))
            .with_calculation_mode(Some(Mode::Manual))
            .with_full_calculation_on_load(Some(false))
            .with_calculation_completed(Some(true))
            .with_force_full_calculation(Some(false)),
    );
    metadata.commit().unwrap();
    let workbook = package.into_workbook().expect("workbook");
    let mut edit = workbook.edit().expect("workbook edit");
    edit.sheet("Sheet1")
        .expect("sheet lookup")
        .expect("sheet")
        .set("A1", Formula::new("1+1").expect("formula"))
        .expect("formula edit");
    let committed = edit.commit().expect("formula commit");
    let properties = committed
        .workbook()
        .calculation_metadata()
        .expect("metadata through workbook owner")
        .properties()
        .expect("invalidated calcPr")
        .clone();
    assert_eq!(properties.calculation_mode(), Mode::Manual);
    assert_eq!(properties.calculation_id(), 0);
    assert!(properties.full_calculation_on_load());
    assert!(!properties.calculation_completed());
    assert!(properties.force_full_calculation());
}

#[test]
fn mce_projected_metadata_is_readable_but_changed_edits_are_refused() {
    const MC: &str = "http://schemas.openxmlformats.org/markup-compatibility/2006";
    let mut raw = Package::create().unwrap().into_plain_opc();
    let workbook_name = raw.main_document_part().unwrap().partname().clone();
    let xml = std::str::from_utf8(raw.get_part(&workbook_name).unwrap().blob())
        .unwrap()
        .replacen("<workbook ", &format!(r#"<workbook xmlns:mc="{MC}" "#), 1)
        .replace(
            "</workbook>",
            r#"<mc:AlternateContent><mc:Choice Requires="future" xmlns:future="urn:future"><calcPr calcId="7"/></mc:Choice><mc:Fallback><calcPr calcId="42" calcMode="manual"/></mc:Fallback></mc:AlternateContent></workbook>"#,
        )
        .into_bytes();
    raw.get_part_mut(&workbook_name)
        .unwrap()
        .set_blob(xml.clone());
    let mut package = Package::from_opc(raw).expect("MCE fixture");
    assert_eq!(
        package
            .calculation_metadata()
            .expect("projected metadata")
            .properties()
            .unwrap()
            .calculation_id(),
        42
    );
    let mut edit = package.edit_calculation_metadata().unwrap();
    edit.edit_properties(|properties| {
        properties.set_calculation_id(Some(8));
        Ok(())
    })
    .unwrap();
    assert!(edit.commit().is_err());
    assert_eq!(package.calculation_metadata().unwrap().source_xml(), xml);
}

#[cfg(feature = "encryption")]
#[test]
fn encrypted_source_refuses_metadata_mutation_until_explicit_plain_reopen() {
    use litchi_xlsx::encryption::Mode as EncryptionMode;

    const PASSWORD: &str = "calculation-metadata-password";
    let plaintext = Package::create().unwrap().to_plain_bytes().unwrap();
    let encrypted =
        litchi_xlsx::encryption::encrypt(plaintext, PASSWORD, EncryptionMode::Standard).unwrap();
    let mut opened = Package::from_bytes_with_password(encrypted, PASSWORD).unwrap();
    assert!(matches!(
        opened.edit_calculation_metadata(),
        Err(Error::EncryptionPolicy { .. })
    ));

    let explicit_plaintext = opened.to_plain_bytes().expect("explicit plaintext output");
    let mut plain = Package::from_bytes(explicit_plaintext).expect("explicit plain reopen");
    let mut edit = plain.edit_calculation_metadata().expect("plain edit");
    edit.set_properties(Properties::new().with_calculation_id(Some(11)));
    edit.commit().expect("plain metadata commit");
    assert_eq!(
        plain
            .calculation_metadata()
            .unwrap()
            .properties()
            .unwrap()
            .calculation_id(),
        11
    );
}
