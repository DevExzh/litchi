#![allow(
    clippy::expect_used,
    clippy::unwrap_used,
    reason = "integration tests use panic-on-failure extraction"
)]

//! Changed publications refuse a signed workbook (ADR 0006; unified-ops
//! record 0013). Removing a signature is explicit (`Package::unsign`, or
//! `OpcPackage::unsign` inside an `edit_opc` update); an exact no-op keeps
//! the source bytes, signature included.

use litchi_core::sheet::WorkbookTrait as _;
use std::io::Cursor;
use std::path::PathBuf;

use litchi_opc::{BlobPart, OpcError, OpcPackage, PackURI, PackageWriter, Part as _, TargetMode};
use litchi_xlsb::Package;
use litchi_xlsb::cell_values::{Reference, StyleIndex, Value};
use litchi_xlsb::cell_watches::Watch;
use litchi_xlsb::merged_cells::MergedCell;
use litchi_xlsb::package::PackageError;

fn fixture(name: &str) -> Vec<u8> {
    std::fs::read(
        PathBuf::from(env!("CARGO_MANIFEST_DIR"))
            .join("../../test-data/ooxml/xlsb")
            .join(name),
    )
    .expect("XLSB fixture")
}

/// `plain` with an OPC signature graph (origin part, one XML signature and
/// their relationships), written out and read back so the signature is an
/// ingressed source signature, as in a file a producer signed.
fn signed(plain: &[u8]) -> Vec<u8> {
    let source = OpcPackage::from_bytes(plain, litchi_opc::SourceLimits::default()).unwrap();
    let mut package = OpcPackage::new();
    for relationship in source.rels().iter() {
        package
            .rels_mut()
            .try_add_relationship(
                relationship.reltype().to_owned(),
                relationship.target_ref().to_owned(),
                relationship.r_id().to_owned(),
                relationship.target_mode(),
            )
            .unwrap();
    }
    let names: Vec<_> = source
        .iter_parts()
        .map(|part| part.partname().clone())
        .collect();
    for name in names {
        let part = source.get_part(&name).unwrap();
        let mut copy = BlobPart::new(
            part.partname().clone(),
            part.content_type().to_owned(),
            part.blob().to_vec(),
        );
        for relationship in part.rels().iter() {
            copy.rels_mut()
                .try_add_relationship(
                    relationship.reltype().to_owned(),
                    relationship.target_ref().to_owned(),
                    relationship.r_id().to_owned(),
                    relationship.target_mode(),
                )
                .unwrap();
        }
        package.try_add_part(Box::new(copy)).unwrap();
    }
    let origin = PackURI::new("/_xmlsignatures/origin.sigs").unwrap();
    let signature = PackURI::new("/_xmlsignatures/sig1.xml").unwrap();
    let mut origin_part = BlobPart::new(
        origin.clone(),
        litchi_opc::constants::content_type::OPC_DIGITAL_SIGNATURE_ORIGIN.to_owned(),
        Vec::new(),
    );
    origin_part
        .rels_mut()
        .try_add_relationship(
            "http://schemas.openxmlformats.org/package/2006/relationships/digital-signature/signature"
                .to_owned(),
            signature.relative_ref(origin.base_uri()),
            "rIdSignature".to_owned(),
            TargetMode::Internal,
        )
        .unwrap();
    package.try_add_part(Box::new(origin_part)).unwrap();
    package
        .try_add_part(Box::new(BlobPart::new(
            signature,
            litchi_opc::constants::content_type::OPC_DIGITAL_SIGNATURE_XMLSIGNATURE.to_owned(),
            b"<Signature xmlns=\"http://www.w3.org/2000/09/xmldsig#\"/>".to_vec(),
        )))
        .unwrap();
    package
        .rels_mut()
        .try_add_relationship(
            litchi_opc::constants::relationship_type::DIGITAL_SIGNATURE_ORIGIN.to_owned(),
            "_xmlsignatures/origin.sigs".to_owned(),
            "rIdSignatureOrigin".to_owned(),
            TargetMode::Internal,
        )
        .unwrap();
    let bytes = PackageWriter::to_bytes(&package).unwrap();
    assert!(
        OpcPackage::from_bytes(&bytes, litchi_opc::SourceLimits::default())
            .unwrap()
            .is_signed()
    );
    bytes
}

fn signed_package() -> (Vec<u8>, Package) {
    let bytes = signed(&fixture("Simple.xlsb"));
    let package = Package::from_bytes(&bytes).unwrap();
    assert!(package.is_signed());
    (bytes, package)
}

fn save(package: &Package) -> Vec<u8> {
    let mut output = Cursor::new(Vec::new());
    package.write_to(&mut output).unwrap();
    output.into_inner()
}

/// Whether `bytes` still carries any OPC signature infrastructure.
fn carries_signature(bytes: &[u8]) -> bool {
    let package = OpcPackage::from_bytes(bytes, litchi_opc::SourceLimits::default()).unwrap();
    package.is_signed()
        || package
            .iter_parts()
            .any(|part| part.partname().as_str().starts_with("/_xmlsignatures/"))
}

fn new_cell() -> (Reference, StyleIndex, Value) {
    (
        Reference::new(10_000, 100).unwrap(),
        StyleIndex::new(0).unwrap(),
        Value::InlineString("edited".to_owned()),
    )
}

#[track_caller]
fn assert_signed_refusal(error: &PackageError) {
    assert!(
        matches!(
            error,
            PackageError::Signed | PackageError::Opc(OpcError::SignedSourceRequiresExplicitPolicy)
        ),
        "expected a signature refusal, got {error:?}"
    );
}

#[test]
fn changed_cell_values_on_a_signed_workbook_are_refused() {
    let (bytes, package) = signed_package();
    let name = package.workbook().unwrap().worksheet_names()[0].clone();
    let mut edit = package
        .edit_cell_values(name)
        .unwrap()
        .resolve()
        .unwrap()
        .expect("worksheet selector");
    let (reference, style, value) = new_cell();
    edit.insert(reference, style, value).unwrap();
    let error = edit.commit().unwrap_err();
    assert!(matches!(error, PackageError::Signed), "{error:?}");
    // Nothing was published: the package still saves its signed source.
    assert!(package.is_signed());
    assert_eq!(save(&package), bytes);
}

#[test]
fn exact_noop_cell_values_keep_the_signature_bytes() {
    let (bytes, package) = signed_package();
    let name = package.workbook().unwrap().worksheet_names()[0].clone();
    let edit = package
        .edit_cell_values(name)
        .unwrap()
        .resolve()
        .unwrap()
        .expect("worksheet selector");
    let commit = edit.commit().unwrap();
    assert!(commit.patch().is_noop());
    assert_eq!(save(commit.package()), bytes);
}

#[test]
fn explicit_unsign_authorizes_a_changed_edit() {
    let (_, signed) = signed_package();
    let package = signed.unsign();
    let name = package.workbook().unwrap().worksheet_names()[0].clone();
    let mut edit = package
        .edit_cell_values(name)
        .unwrap()
        .resolve()
        .unwrap()
        .expect("worksheet selector");
    let (reference, style, value) = new_cell();
    edit.insert(reference, style, value).unwrap();
    let commit = edit.commit().unwrap();
    let output = save(commit.package());
    assert!(!carries_signature(&output));
    let reopened = Package::from_bytes(&output).unwrap();
    let cell = reopened.workbook().unwrap().cell_values(0).unwrap();
    assert!(cell.cell(reference).unwrap().is_some());
}

#[test]
fn workbook_structure_changes_are_refused_and_noops_kept() {
    let (bytes, package) = signed_package();
    let mut edit = package.edit().unwrap();
    edit.set_sheet_name(0, "Renamed".to_owned())
        .unwrap()
        .expect("resolved worksheet selector");
    let error = edit.commit().unwrap_err();
    assert!(matches!(error, PackageError::Signed), "{error:?}");

    let commit = package.edit().unwrap().commit().unwrap();
    assert!(!commit.diagnostics().changed());
    assert!(commit.patch().is_noop());
    let replayed = package.apply(commit.patch()).unwrap();
    assert_eq!(save(replayed.package()), bytes);

    let name = package.workbook().unwrap().worksheet_names()[0].clone();
    let mut same_name = package.edit().unwrap();
    same_name
        .set_sheet_name(0, name)
        .unwrap()
        .expect("resolved worksheet selector");
    let same_name_commit = same_name.commit().unwrap();
    assert!(!same_name_commit.diagnostics().changed());
    assert!(same_name_commit.patch().is_noop());
    assert_eq!(same_name_commit.patch().operation_count(), 1);
    let replayed = package.apply(same_name_commit.patch()).unwrap();
    assert!(replayed.package().is_signed());
    assert_eq!(save(replayed.package()), bytes);
}

#[test]
fn cell_watch_changes_are_refused() {
    let (bytes, package) = signed_package();
    let mut edit = package.edit_cell_watches().unwrap();
    let staged = edit.add(0, Watch::new(3, 5).unwrap());
    let error = staged.and_then(|()| edit.commit().map(|_| ())).unwrap_err();
    assert!(matches!(error, PackageError::Signed), "{error:?}");
    assert_eq!(save(&package), bytes);
}

#[test]
fn raw_opc_edits_keep_the_signature_policy() {
    let (bytes, package) = signed_package();
    // An operation that changes nothing keeps the signature byte for byte.
    let mut edit = package.edit_opc().unwrap();
    edit.update(|_| Ok(())).unwrap();
    assert_eq!(save(edit.commit().unwrap().package()), bytes);

    // Typed families refuse a change.
    let mut merged = package.edit_merged_cells().unwrap();
    let error = merged
        .set_merged_cell_ranges(0, &[MergedCell::new(20, 21, 0, 1)])
        .and_then(|_| merged.commit().map(|_| ()))
        .unwrap_err();
    assert_signed_refusal(&error);
    assert_eq!(save(&package), bytes);

    // Removing the signature inside the update is an explicit disposition.
    let uri = PackURI::new("/docProps/litchi-probe.xml").unwrap();
    let mut edit = package.edit_opc().unwrap();
    edit.update(|package| {
        package.unsign();
        package
            .try_add_part(Box::new(BlobPart::new(
                uri.clone(),
                "application/xml".to_owned(),
                b"<probe/>".to_vec(),
            )))
            .map_err(PackageError::from)
    })
    .unwrap();
    assert!(!carries_signature(&save(edit.commit().unwrap().package())));
}

#[test]
fn connections_refuse_a_signed_workbook() {
    use litchi_xlsb::package::connections::{
        CommandType, Connection, Connections, DbProperties, Properties, SourceType,
    };
    let (bytes, package) = signed_package();
    let mut edit = package.edit_connections().unwrap();
    edit.replace(Some(Connections {
        connections: vec![Connection {
            connection_id: 42,
            source_type: SourceType::Odbc,
            name: "Sales".to_owned(),
            properties: Properties::Database(DbProperties {
                command_type: CommandType::Sql,
                connection_string: "Driver={Generated};Server=example.invalid".to_owned(),
                command: Some("SELECT 1".to_owned()),
                server_command: None,
            }),
            ..Connection::default()
        }],
    }))
    .unwrap();
    let error = edit.commit().unwrap_err();
    assert!(matches!(error, PackageError::Signed), "{error:?}");
    let mut edit = package.edit_connections().unwrap();
    assert!(!edit.replace(None).unwrap());
    assert_eq!(save(edit.commit().unwrap().package()), bytes);
}

#[test]
fn threaded_comment_edits_refuse_a_signed_package() {
    let (bytes, package) = signed_package();
    let mut edit = package.edit_threaded_comments().unwrap();
    edit.set_worksheet(
        "/xl/worksheets/sheet1.bin",
        litchi_xlsb::comments::threaded::Comments::default(),
    )
    .unwrap();
    let error = edit.commit().unwrap_err();
    assert!(matches!(error, PackageError::Signed), "{error:?}");
    // Clearing an absent graph changes nothing and is not refused.
    let mut edit = package.edit_threaded_comments().unwrap();
    assert!(!edit.clear().unwrap());
    assert_eq!(save(edit.commit().unwrap().package()), bytes);
}

/// Follow-up F2 of record 0013: task panes and the Ribbon refuse a change to a
/// signed workbook with the typed signed error; the Ribbon used to remove the
/// signature.
mod task_panes_and_ribbon {
    use super::*;
    use litchi_ooxml_common::ribbon::{Family, Version};
    use litchi_ooxml_common::web;

    const UI_2007: &[u8] =
        br#"<customUI xmlns="http://schemas.microsoft.com/office/2006/01/customui"/>"#;
    const UI_2007_CHANGED: &[u8] =
        br#"<customUI xmlns="http://schemas.microsoft.com/office/2006/01/customui"><ribbon/></customUI>"#;

    fn panes(id: &str) -> web::Panes {
        let add_in = web::AddIn::new(
            id,
            web::Reference::new("ref-1", "1", web::Store::Registry).unwrap(),
        )
        .unwrap();
        let mut panes = web::Panes::new();
        panes.push(web::Pane::new(add_in)).unwrap();
        panes
    }

    /// `Simple.xlsb` edited by `edit` while unsigned, then signed.
    fn signed_package_with(edit: impl FnOnce(&Package) -> Package) -> (Vec<u8>, Package) {
        let plain = Package::from_bytes(&fixture("Simple.xlsb")).unwrap();
        let bytes = signed(&save(&edit(&plain)));
        let package = Package::from_bytes(&bytes).unwrap();
        assert!(package.is_signed());
        (bytes, package)
    }

    fn put_panes(package: &Package, id: &str) -> litchi_xlsb::package::PackageResult<Package> {
        let mut edit = package.edit_task_panes_with_options(
            litchi_xlsb::task_panes::EditOptions::default()
                .conformance(web::Conformance::Transitional),
        )?;
        edit.replace(panes(id))?;
        Ok(edit.commit()?.into_package())
    }

    fn clear_panes(package: &Package) -> litchi_xlsb::package::PackageResult<(bool, Package)> {
        let mut edit = package.edit_task_panes()?;
        let removed = edit.clear()?;
        Ok((removed, edit.commit()?.into_package()))
    }

    fn put_ribbon(package: &Package, xml: &[u8]) -> litchi_xlsb::package::PackageResult<Package> {
        let mut edit = package.edit_ribbon()?;
        edit.put(Version::V2007, xml.to_vec())?;
        Ok(edit.commit()?.into_package())
    }

    fn remove_ribbon(
        package: &Package,
        family: Family,
    ) -> litchi_xlsb::package::PackageResult<(bool, Package)> {
        let mut edit = package.edit_ribbon()?;
        let removed = edit.remove(family)?;
        Ok((removed, edit.commit()?.into_package()))
    }

    #[track_caller]
    fn assert_signed<T: std::fmt::Debug>(result: litchi_xlsb::package::PackageResult<T>) {
        let error = result.expect_err("a change to a signed workbook must be refused");
        assert!(matches!(error, PackageError::Signed), "{error:?}");
    }

    #[test]
    fn task_pane_changes_are_refused_up_front_and_noops_kept() {
        let (bytes, package) = signed_package();
        assert_signed(put_panes(&package, "add-in-1"));
        let (removed, same) = clear_panes(&package).unwrap();
        assert!(!removed);
        assert_eq!(save(&same), bytes);

        let (bytes, package) = signed_package_with(|plain| put_panes(plain, "add-in-1").unwrap());
        // Equal panes publish nothing.
        assert_eq!(save(&put_panes(&package, "add-in-1").unwrap()), bytes);
        assert_signed(put_panes(&package, "add-in-2"));
        assert_signed(clear_panes(&package));
        assert!(package.is_signed());
        assert_eq!(save(&package), bytes);

        let (removed, unsigned) = clear_panes(&package.unsign()).unwrap();
        assert!(removed);
        let output = save(&unsigned);
        assert!(!carries_signature(&output));
        let reopened = Package::from_bytes(&output).unwrap();
        assert!(reopened.task_panes().unwrap().is_none());
    }

    #[test]
    fn ribbon_changes_are_refused_up_front_and_noops_kept() {
        let (bytes, package) = signed_package();
        assert_signed(put_ribbon(&package, UI_2007));
        let (removed, same) = remove_ribbon(&package, Family::Legacy).unwrap();
        assert!(!removed);
        assert_eq!(save(&same), bytes);

        let (bytes, package) = signed_package_with(|plain| put_ribbon(plain, UI_2007).unwrap());
        // The stored version and bytes publish nothing; an absent family is
        // not removed.
        assert_eq!(save(&put_ribbon(&package, UI_2007).unwrap()), bytes);
        let (removed, same) = remove_ribbon(&package, Family::Modern).unwrap();
        assert!(!removed);
        assert_eq!(save(&same), bytes);
        assert_signed(put_ribbon(&package, UI_2007_CHANGED));
        assert_signed(remove_ribbon(&package, Family::Legacy));
        assert!(package.is_signed());
        assert_eq!(save(&package), bytes);

        let unsigned = put_ribbon(&package.unsign(), UI_2007_CHANGED).unwrap();
        let output = save(&unsigned);
        assert!(!carries_signature(&output));
        let reopened = Package::from_bytes(&output).unwrap();
        assert_eq!(
            reopened.ribbon().unwrap().legacy().unwrap().xml(),
            UI_2007_CHANGED
        );
    }
}

#[test]
fn workbook_projection_round_trip_preserves_signed_source() {
    let (bytes, package) = signed_package();
    let adopted = package.workbook().unwrap().to_package();
    assert!(adopted.opc_package().is_signed());
    assert_eq!(adopted.to_bytes().unwrap(), bytes);
    assert_eq!(save(&adopted), bytes);
}

#[test]
fn package_cell_removal_preserves_signed_noops_and_refuses_changed_publication() {
    let plain = Package::new().expect("new");
    let a1 = Reference::new(0, 0).expect("A1");
    let b1 = Reference::new(0, 1).expect("B1");
    let mut edit = plain
        .edit_cell_values(0)
        .expect("edit")
        .resolve()
        .expect("worksheet resolution")
        .expect("sheet");
    edit.set_scalar(a1, litchi_xlsb::cell_values::Scalar::Number(1.0))
        .expect("set");
    let bytes = signed(
        &edit
            .commit()
            .expect("commit")
            .package()
            .to_bytes()
            .expect("bytes"),
    );
    let source = Package::from_bytes(&bytes).expect("signed source");
    let mut noop = source
        .edit_cell_values(0)
        .expect("edit")
        .resolve()
        .expect("worksheet resolution")
        .expect("sheet");
    assert!(!noop.remove(b1).expect("missing"));
    let commit = noop.commit().expect("signed noop");
    assert!(commit.patch().is_noop());
    assert_eq!(commit.package().to_bytes().expect("bytes"), bytes);
    let mut edit = source
        .edit_cell_values(0)
        .expect("edit")
        .resolve()
        .expect("worksheet resolution")
        .expect("sheet");
    assert!(edit.remove(a1).expect("stage removal"));
    assert!(matches!(edit.commit(), Err(PackageError::Signed)));
    assert_eq!(source.to_bytes().expect("source unchanged"), bytes);
}
