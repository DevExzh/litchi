use litchi_opc::constants::{content_type as ct, relationship_type as rt};
use litchi_opc::{BlobPart, OpcPackage, PackURI, TargetMode};
use litchi_xlsx::calculation_properties::{
    Feature, Features, Mode, Patch as CalculationPatch, Properties,
};
use litchi_xlsx::formula::Formula;
use litchi_xlsx::{Error, Package, Workbook};
use soapberry_zip::office::{ArchiveReader, StreamingArchiveWriter};

const CALC_CHAIN_CONTENT_TYPE: &str =
    "application/vnd.openxmlformats-officedocument.spreadsheetml.calcChain+xml";
const SIGNATURE_REL: &str =
    "http://schemas.openxmlformats.org/package/2006/relationships/digital-signature/signature";
const ORIGIN_BYTES: &[u8] = b"origin";
const SIGNATURE_BYTES: &[u8] = b"signature";

fn calculation_source() -> Workbook {
    let seed = Workbook::create().expect("seed workbook");
    let mut edit = seed.edit().expect("seed edit");
    edit.sheet("Sheet1")
        .expect("sheet lookup")
        .expect("sheet")
        .set("A1", 2_i32)
        .and_then(|sheet| sheet.set("B1", Formula::new("A1+1").expect("formula")))
        .expect("seed cells");
    let seed = edit.commit().expect("seed commit").into_workbook();

    let mut package =
        Package::from_bytes(seed.to_plain_bytes().expect("seed bytes")).expect("seed package");
    let mut metadata = package.edit_calculation_metadata().expect("metadata edit");
    assert!(
        metadata.set_properties(
            Properties::new()
                .with_calculation_id(Some(123))
                .with_calculation_mode(Some(Mode::Manual))
                .with_full_calculation_on_load(Some(false))
                .with_calculation_completed(Some(true))
                .with_force_full_calculation(Some(false)),
        )
    );
    metadata.commit().expect("metadata commit");

    let mut raw = package.into_plain_opc();
    let workbook_uri = raw
        .main_document_part()
        .expect("workbook part")
        .partname()
        .clone();
    raw.try_add_part(Box::new(BlobPart::new(
        PackURI::new("/xl/calcChain.xml").expect("chain URI"),
        CALC_CHAIN_CONTENT_TYPE.to_owned(),
        br#"<calcChain xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><c r="B1" i="1"/></calcChain>"#
            .to_vec(),
    )))
    .expect("chain part");
    raw.get_part_mut(&workbook_uri)
        .expect("mutable workbook part")
        .rels_mut()
        .try_add_relationship(
            rt::CALC_CHAIN.to_owned(),
            "calcChain.xml".to_owned(),
            "rIdCalculationChain".to_owned(),
            TargetMode::Internal,
        )
        .expect("chain relationship");
    Package::from_opc(raw)
        .expect("calculation package")
        .into_workbook()
        .expect("calculation workbook")
}

fn has_calculation_chain(workbook: &Workbook) -> bool {
    let bytes = workbook.to_plain_bytes().expect("workbook bytes");
    let package = OpcPackage::from_bytes(&bytes).expect("OPC package");
    let workbook_part = package.main_document_part().expect("workbook part");
    let relationship = workbook_part.rels().iter().find(|relationship| {
        matches!(
            relationship.reltype(),
            rt::CALC_CHAIN | rt::STRICT_CALC_CHAIN
        )
    });
    let Some(relationship) = relationship else {
        return false;
    };
    relationship
        .target_partname()
        .is_ok_and(|target| package.get_part(&target).is_ok())
}

fn assert_invalidated(workbook: &Workbook) {
    let properties = workbook
        .calculation_metadata()
        .expect("calculation metadata")
        .properties()
        .expect("calcPr")
        .clone();
    assert_eq!(properties.calculation_mode(), Mode::Manual);
    assert_eq!(properties.calculation_id(), 0);
    assert!(properties.full_calculation_on_load());
    assert!(!properties.calculation_completed());
    assert!(properties.force_full_calculation());
    assert!(!has_calculation_chain(workbook));
}

fn calculation_metadata_patches() -> (Package, CalculationPatch, CalculationPatch) {
    let baseline = Package::create().expect("baseline Package");
    let mut no_op_author = baseline.clone();
    let no_op = no_op_author
        .edit_calculation_metadata()
        .expect("no-op metadata edit")
        .commit()
        .expect("no-op metadata commit")
        .patch()
        .clone();
    assert!(no_op.is_empty());

    let mut changed_author = baseline.clone();
    let mut edit = changed_author
        .edit_calculation_metadata()
        .expect("changed metadata edit");
    assert!(edit.set_properties(Properties::new().with_calculation_id(Some(8675309)),));
    let changed = edit
        .commit()
        .expect("changed metadata commit")
        .patch()
        .clone();
    assert!(!changed.is_empty());
    (baseline, no_op, changed)
}

#[derive(Debug, Clone, Copy)]
enum ContentTransition {
    FormulaInsertion,
    FormulaReplacement,
    FormulaRemoval,
    FormulaToValue,
    InputValueChange,
}

#[test]
fn effective_cell_content_transitions_invalidate_flags_and_remove_calc_chain() {
    for transition in [
        ContentTransition::FormulaInsertion,
        ContentTransition::FormulaReplacement,
        ContentTransition::FormulaRemoval,
        ContentTransition::FormulaToValue,
        ContentTransition::InputValueChange,
    ] {
        let source = calculation_source();
        assert!(has_calculation_chain(&source), "{transition:?}");
        let mut edit = source.edit().expect("cell edit");
        let mut sheet = edit.sheet("Sheet1").expect("sheet lookup").expect("sheet");
        match transition {
            ContentTransition::FormulaInsertion => {
                sheet
                    .set("C1", Formula::new("A1*2").expect("formula"))
                    .expect("insert formula");
            },
            ContentTransition::FormulaReplacement => {
                sheet
                    .set("B1", Formula::new("A1+2").expect("formula"))
                    .expect("replace formula");
            },
            ContentTransition::FormulaRemoval => {
                sheet.clear("B1").expect("remove formula payload");
            },
            ContentTransition::FormulaToValue => {
                sheet.set("B1", 9_i32).expect("replace formula with value");
            },
            ContentTransition::InputValueChange => {
                sheet.set("A1", 3_i32).expect("change input value");
            },
        }
        let commit = edit
            .commit()
            .unwrap_or_else(|error| panic!("{transition:?}: {error}"));
        assert!(!commit.patch().is_empty(), "{transition:?}");
        assert_invalidated(commit.workbook());
        assert!(
            commit
                .workbook()
                .calculation_metadata()
                .expect("calculation metadata")
                .features()
                .is_none(),
            "{transition:?}",
        );
    }
}

#[test]
fn formula_invalidation_preserves_ordered_duplicate_calc_features_exactly() {
    let source = calculation_source();
    let mut package = Package::from_bytes(source.to_plain_bytes().expect("source bytes"))
        .expect("source Package");
    let mut metadata = package
        .edit_calculation_metadata()
        .expect("feature metadata edit");
    assert!(
        metadata.set_features(
            Features::try_from_vec(vec![
                Feature::new("Case").expect("feature"),
                Feature::new("case").expect("feature"),
                Feature::new("Case").expect("duplicate feature"),
            ])
            .expect("ordered features"),
        )
    );
    metadata.commit().expect("feature metadata commit");
    let source = package.into_workbook().expect("feature workbook");

    let mut edit = source.edit().expect("formula edit");
    edit.sheet("Sheet1")
        .expect("sheet lookup")
        .expect("sheet")
        .set("B1", Formula::new("A1+99").expect("replacement formula"))
        .expect("replace formula");
    let committed = edit.commit().expect("formula commit");
    assert_invalidated(committed.workbook());
    let metadata = committed
        .workbook()
        .calculation_metadata()
        .expect("invalidated metadata");
    let names = metadata
        .features()
        .expect("preserved calcFeatures")
        .iter()
        .map(Feature::as_str)
        .collect::<Vec<_>>();
    assert_eq!(names, ["Case", "case", "Case"]);
}

#[test]
fn style_only_edit_preserves_calc_flags_and_chain() {
    let source = calculation_source();
    let metadata_before = source
        .calculation_metadata()
        .expect("source metadata")
        .source_xml()
        .to_vec();
    let style = source.styles().expect("styles").base().expect("base style");
    let mut edit = source.edit().expect("style edit");
    edit.sheet("Sheet1")
        .expect("sheet lookup")
        .expect("sheet")
        .style("A1", &style)
        .expect("style cell");
    let commit = edit.commit().expect("style commit");

    assert!(!commit.patch().is_empty());
    assert!(has_calculation_chain(commit.workbook()));
    assert_eq!(
        commit
            .workbook()
            .calculation_metadata()
            .expect("preserved metadata")
            .source_xml(),
        metadata_before,
    );
}

#[test]
fn empty_and_effective_no_op_patches_preserve_exact_workbook_bytes() {
    let source = calculation_source();
    let source_bytes = source.to_plain_bytes().expect("source bytes");

    let empty = source
        .edit()
        .expect("empty edit")
        .commit()
        .expect("empty commit");
    assert!(empty.patch().is_empty());
    assert_eq!(empty.workbook().to_plain_bytes().unwrap(), source_bytes);
    let replayed_empty = source.apply(empty.patch()).expect("replay empty patch");
    assert_eq!(
        replayed_empty.workbook().to_plain_bytes().unwrap(),
        source_bytes
    );

    let mut edit = source.edit().expect("effective no-op edit");
    edit.sheet("Sheet1")
        .expect("sheet lookup")
        .expect("sheet")
        .set("A1", 2_i32)
        .expect("same input value");
    let no_op = edit.commit().expect("effective no-op commit");
    assert!(no_op.patch().is_empty());
    assert_eq!(no_op.workbook().to_plain_bytes().unwrap(), source_bytes);
    let replayed_no_op = source.apply(no_op.patch()).expect("replay no-op patch");
    assert_eq!(
        replayed_no_op.workbook().to_plain_bytes().unwrap(),
        source_bytes
    );
    assert!(has_calculation_chain(replayed_no_op.workbook()));
}

fn append_before_closing(xml: Vec<u8>, closing: &str, addition: &str) -> Vec<u8> {
    let mut xml = String::from_utf8(xml).expect("fixture XML is UTF-8");
    let position = xml.rfind(closing).expect("fixture XML closing element");
    xml.insert_str(position, addition);
    xml.into_bytes()
}

fn assert_inert_signature_graph(package: &Package) {
    let raw = package.clone().into_plain_opc();
    let origin_uri = PackURI::new("/_xmlsignatures/origin.sigs").expect("origin URI");
    let signature_uri = PackURI::new("/_xmlsignatures/sig1.xml").expect("signature URI");
    let root = raw
        .rels()
        .iter()
        .find(|relationship| relationship.reltype() == rt::DIGITAL_SIGNATURE_ORIGIN)
        .expect("signature-origin relationship");
    assert_eq!(root.target_ref(), "_xmlsignatures/origin.sigs");
    assert_eq!(root.target_mode(), TargetMode::Internal);

    let origin = raw.get_part(&origin_uri).expect("signature-origin part");
    assert_eq!(origin.content_type(), ct::OPC_DIGITAL_SIGNATURE_ORIGIN);
    assert_eq!(origin.blob(), ORIGIN_BYTES);
    let signature_relationship = origin
        .rels()
        .iter()
        .find(|relationship| relationship.reltype() == SIGNATURE_REL)
        .expect("signature relationship");
    assert_eq!(signature_relationship.target_ref(), "sig1.xml");
    assert_eq!(signature_relationship.target_mode(), TargetMode::Internal);

    let signature = raw.get_part(&signature_uri).expect("signature part");
    assert_eq!(
        signature.content_type(),
        ct::OPC_DIGITAL_SIGNATURE_XMLSIGNATURE
    );
    assert_eq!(signature.blob(), SIGNATURE_BYTES);
}

fn inert_signed_fixture(source: &[u8]) -> Package {
    let reader = ArchiveReader::new(source).expect("source ZIP");
    let names = reader.file_names().map(str::to_owned).collect::<Vec<_>>();
    assert!(!names.iter().any(|name| name.starts_with("_xmlsignatures/")));

    let content_types = format!(
        r#"<Override PartName="/_xmlsignatures/origin.sigs" ContentType="{}"/><Override PartName="/_xmlsignatures/sig1.xml" ContentType="{}"/>"#,
        ct::OPC_DIGITAL_SIGNATURE_ORIGIN,
        ct::OPC_DIGITAL_SIGNATURE_XMLSIGNATURE,
    );
    let root_relationship = format!(
        r#"<Relationship Id="rIdSignatureOrigin" Type="{}" Target="_xmlsignatures/origin.sigs"/>"#,
        rt::DIGITAL_SIGNATURE_ORIGIN,
    );
    let origin_relationships = format!(
        r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rIdSignature" Type="{SIGNATURE_REL}" Target="sig1.xml"/></Relationships>"#,
    );

    let mut writer = StreamingArchiveWriter::new();
    let mut saw_content_types = false;
    let mut saw_root_relationships = false;
    for name in names {
        let bytes = reader.read(&name).expect("source ZIP member");
        let bytes = match name.as_str() {
            "[Content_Types].xml" => {
                saw_content_types = true;
                append_before_closing(bytes, "</Types>", &content_types)
            },
            "_rels/.rels" => {
                saw_root_relationships = true;
                append_before_closing(bytes, "</Relationships>", &root_relationship)
            },
            _ => bytes,
        };
        writer
            .write_deflated(&name, &bytes)
            .expect("copy source ZIP member");
    }
    assert!(saw_content_types);
    assert!(saw_root_relationships);
    writer
        .write_deflated("_xmlsignatures/origin.sigs", ORIGIN_BYTES)
        .expect("raw signature-origin member");
    writer
        .write_deflated(
            "_xmlsignatures/_rels/origin.sigs.rels",
            origin_relationships.as_bytes(),
        )
        .expect("raw signature relationship member");
    writer
        .write_deflated("_xmlsignatures/sig1.xml", SIGNATURE_BYTES)
        .expect("raw malformed signature member");
    let bytes = writer.finish_to_bytes().expect("signed fixture ZIP");
    let package = Package::from_bytes(bytes).expect("signed fixture package");
    assert_inert_signature_graph(&package);
    package
}

#[test]
fn changed_public_patch_is_refused_on_workbook_from_signed_package() {
    let source = calculation_source();
    let source_bytes = source.to_plain_bytes().expect("source bytes");
    let mut edit = source.edit().expect("patch edit");
    edit.sheet("Sheet1")
        .expect("sheet lookup")
        .expect("sheet")
        .set("A1", 7_i32)
        .expect("change input");
    let patch = edit.commit().expect("patch commit").patch().clone();
    assert!(!patch.is_empty());

    let signed_package = inert_signed_fixture(&source_bytes);
    let signed = signed_package.workbook().expect("signed workbook facade");
    let signed_before = signed.to_plain_bytes().expect("signed bytes");
    assert!(matches!(signed.apply(&patch), Err(Error::Signed)));
    assert_eq!(signed.to_plain_bytes().unwrap(), signed_before);
    assert!(has_calculation_chain(&signed));
}

#[test]
fn calculation_metadata_patch_no_op_is_exact_but_change_is_refused_on_signed_package() {
    let (baseline, no_op, changed) = calculation_metadata_patches();
    let baseline_bytes = baseline.to_plain_bytes().expect("baseline bytes");
    let mut signed = inert_signed_fixture(&baseline_bytes);
    let before = signed.to_plain_bytes().expect("signed source bytes");

    signed
        .apply_calculation_metadata_patch(&no_op)
        .expect("signed exact no-op patch");
    assert_eq!(signed.to_plain_bytes().unwrap(), before);
    assert!(matches!(
        signed.apply_calculation_metadata_patch(&changed),
        Err(Error::Signed)
    ));
    assert_eq!(signed.to_plain_bytes().unwrap(), before);
}

#[cfg(feature = "encryption")]
#[test]
fn changed_public_patch_is_refused_on_encrypted_provenance_facade() {
    use litchi_xlsx::encryption::Mode as EncryptionMode;

    const PASSWORD: &str = "formula-security-password";
    let source = calculation_source();
    let source_bytes = source.to_plain_bytes().expect("source bytes");
    let mut edit = source.edit().expect("patch edit");
    edit.sheet("Sheet1")
        .expect("sheet lookup")
        .expect("sheet")
        .set("A1", 7_i32)
        .expect("change input");
    let patch = edit.commit().expect("patch commit").patch().clone();
    assert!(!patch.is_empty());

    let encrypted = Package::from_bytes(source_bytes)
        .expect("plain package")
        .to_encrypted(PASSWORD, EncryptionMode::Standard)
        .expect("encrypted bytes");
    let package = Package::from_bytes_with_password(encrypted, PASSWORD)
        .expect("encrypted provenance Package");
    let workbook = package.workbook().expect("encrypted workbook facade");
    assert_eq!(workbook.encryption(), Some(EncryptionMode::Standard));
    let before = workbook.to_plain_bytes().expect("clear inner bytes");
    assert!(matches!(
        workbook.apply(&patch),
        Err(Error::EncryptionPolicy {
            operation: "apply",
            ..
        })
    ));
    assert_eq!(workbook.to_plain_bytes().unwrap(), before);
    assert!(has_calculation_chain(&workbook));
}

#[cfg(feature = "encryption")]
#[test]
fn calculation_metadata_patches_are_refused_exactly_on_encrypted_provenance_package() {
    use litchi_xlsx::encryption::Mode as EncryptionMode;

    const PASSWORD: &str = "metadata-patch-security-password";
    let (baseline, no_op, changed) = calculation_metadata_patches();
    let encrypted = baseline
        .to_encrypted(PASSWORD, EncryptionMode::Standard)
        .expect("encrypted bytes");
    let mut package = Package::from_bytes_with_password(encrypted, PASSWORD)
        .expect("encrypted provenance Package");
    let before = package.to_plain_bytes().expect("clear inner source bytes");

    assert!(matches!(
        package.apply_calculation_metadata_patch(&changed),
        Err(Error::EncryptionPolicy {
            operation: "apply_calculation_metadata_patch",
            ..
        })
    ));
    assert_eq!(package.to_plain_bytes().unwrap(), before);
    assert!(matches!(
        package.apply_calculation_metadata_patch(&no_op),
        Err(Error::EncryptionPolicy {
            operation: "apply_calculation_metadata_patch",
            ..
        })
    ));
    assert_eq!(package.to_plain_bytes().unwrap(), before);
}

fn mce_projected_source() -> Workbook {
    const MC: &str = "http://schemas.openxmlformats.org/markup-compatibility/2006";
    let source = calculation_source();
    let mut package = Package::from_bytes(source.to_plain_bytes().expect("source bytes"))
        .expect("source package");
    let mut metadata = package
        .edit_calculation_metadata()
        .expect("metadata removal edit");
    assert!(metadata.remove_properties());
    metadata.commit().expect("metadata removal");

    let mut raw = package.into_plain_opc();
    let workbook_uri = raw
        .main_document_part()
        .expect("workbook part")
        .partname()
        .clone();
    let xml = std::str::from_utf8(raw.get_part(&workbook_uri).unwrap().blob())
        .expect("UTF-8 workbook XML")
        .replacen("<workbook ", &format!(r#"<workbook xmlns:mc="{MC}" "#), 1)
        .replace(
            "</workbook>",
            r#"<mc:AlternateContent><mc:Choice Requires="future" xmlns:future="urn:future"><calcPr calcId="7"/></mc:Choice><mc:Fallback><calcPr calcId="42" calcMode="manual"/></mc:Fallback></mc:AlternateContent></workbook>"#,
        )
        .into_bytes();
    raw.get_part_mut(&workbook_uri)
        .expect("mutable workbook part")
        .set_blob(xml);
    Package::from_opc(raw)
        .expect("MCE Package")
        .into_workbook()
        .expect("MCE workbook")
}

#[test]
fn recalculation_failure_on_mce_projected_calc_pr_is_atomic() {
    let source = mce_projected_source();
    assert_eq!(
        source
            .calculation_metadata()
            .expect("projected metadata")
            .properties()
            .expect("effective calcPr")
            .calculation_id(),
        42,
    );
    assert!(has_calculation_chain(&source));
    let before = source.to_plain_bytes().expect("source bytes");

    let mut edit = source.edit().expect("cell edit");
    edit.sheet("Sheet1")
        .expect("sheet lookup")
        .expect("sheet")
        .set("A1", 8_i32)
        .expect("change input");
    let error = edit
        .commit()
        .expect_err("MCE-projected calcPr must refuse rewrite");
    assert!(
        error
            .to_string()
            .contains("cannot rewrite calcPr projected through MCE markup")
    );
    assert_eq!(source.to_plain_bytes().unwrap(), before);
    assert!(has_calculation_chain(&source));
}
