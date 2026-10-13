#![allow(
    clippy::expect_used,
    clippy::unwrap_used,
    reason = "integration tests use panic-on-failure extraction"
)]

//! Changed publications refuse a signed workbook (ADR 0006; unified-ops
//! record 0013). Removing a signature is explicit (`Workbook::unsign`, or
//! `OpcPackage::unsign` inside `edit_opc`); an exact no-op keeps the source
//! bytes, signature included.

use litchi_core::sheet::WorkbookTrait as _;
use std::io::Cursor;
use std::path::PathBuf;

use litchi_opc::{BlobPart, OpcError, OpcPackage, PackURI, PackageWriter, Part as _, TargetMode};
use litchi_xlsb::cell_values::{Reference, StyleIndex, Value};
use litchi_xlsb::cell_watches::Watch;
use litchi_xlsb::merged_cells::MergedCell;
use litchi_xlsb::package::PackageError;
use litchi_xlsb::{Package, Workbook};

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
    let source = OpcPackage::from_bytes(plain).unwrap();
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
    assert!(OpcPackage::from_bytes(&bytes).unwrap().is_signed());
    bytes
}

fn signed_workbook() -> (Vec<u8>, Workbook) {
    let bytes = signed(&fixture("Simple.xlsb"));
    let workbook = Package::from_bytes(&bytes)
        .unwrap()
        .into_workbook()
        .unwrap();
    assert!(workbook.is_signed());
    (bytes, workbook)
}

fn save(workbook: &Workbook) -> Vec<u8> {
    let mut output = Cursor::new(Vec::new());
    workbook.to_package().write_to(&mut output).unwrap();
    output.into_inner()
}

/// Whether `bytes` still carries any OPC signature infrastructure.
fn carries_signature(bytes: &[u8]) -> bool {
    let package = OpcPackage::from_bytes(bytes).unwrap();
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
    let (bytes, workbook) = signed_workbook();
    let package = workbook.to_package();
    let name = workbook.worksheet_names()[0].clone();
    let mut edit = package
        .edit_cell_values(name)
        .unwrap()
        .expect("worksheet selector");
    let (reference, style, value) = new_cell();
    edit.insert(reference, style, value).unwrap();
    let error = edit.commit().unwrap_err();
    assert!(matches!(error, PackageError::Signed), "{error:?}");
    // Nothing was published: the workbook still saves its signed source.
    assert!(workbook.is_signed());
    assert_eq!(save(&workbook), bytes);
}

#[test]
fn exact_noop_cell_values_keep_the_signature_bytes() {
    let (bytes, mut workbook) = signed_workbook();
    let package = workbook.to_package();
    let name = workbook.worksheet_names()[0].clone();
    let edit = package
        .edit_cell_values(name)
        .unwrap()
        .expect("worksheet selector");
    let commit = edit.commit().unwrap();
    assert!(commit.patch().is_noop());
    workbook = commit.into_package().into_workbook().unwrap();
    assert_eq!(save(&workbook), bytes);
}

#[test]
fn explicit_unsign_authorizes_a_changed_edit() {
    let (_, mut workbook) = signed_workbook();
    workbook.unsign();
    let package = workbook.to_package();
    let name = workbook.worksheet_names()[0].clone();
    let mut edit = package
        .edit_cell_values(name)
        .unwrap()
        .expect("worksheet selector");
    let (reference, style, value) = new_cell();
    edit.insert(reference, style, value).unwrap();
    let commit = edit.commit().unwrap();
    workbook = commit.into_package().into_workbook().unwrap();
    let output = save(&workbook);
    assert!(!carries_signature(&output));
    let reopened = Package::from_bytes(&output)
        .unwrap()
        .into_workbook()
        .unwrap();
    let cell = reopened.cell_values(0).unwrap();
    assert!(cell.cell(reference).unwrap().is_some());
}

#[test]
fn workbook_structure_changes_are_refused_and_noops_kept() {
    use litchi_core::sheet::traits::WorkbookTrait as _;

    let (bytes, mut workbook) = signed_workbook();
    let mut edit = workbook.to_package().edit().unwrap();
    edit.set_sheet_name(0, "Renamed".to_owned())
        .unwrap()
        .expect("resolved worksheet selector");
    let error = edit.commit().unwrap_err();
    assert!(matches!(error, PackageError::Signed), "{error:?}");

    let commit = workbook.to_package().edit().unwrap().commit().unwrap();
    assert!(!commit.diagnostics().changed());
    assert!(commit.patch().is_noop());
    workbook = workbook
        .to_package()
        .apply(commit.patch())
        .unwrap()
        .into_workbook()
        .unwrap();
    assert_eq!(save(&workbook), bytes);

    let name = workbook.worksheet_names()[0].clone();
    let mut same_name = workbook.to_package().edit().unwrap();
    same_name
        .set_sheet_name(0, name)
        .unwrap()
        .expect("resolved worksheet selector");
    let same_name_commit = same_name.commit().unwrap();
    assert!(!same_name_commit.diagnostics().changed());
    assert!(same_name_commit.patch().is_noop());
    assert_eq!(same_name_commit.patch().operation_count(), 1);
    workbook = workbook
        .to_package()
        .apply(same_name_commit.patch())
        .unwrap()
        .into_workbook()
        .unwrap();
    assert!(workbook.is_signed());
    assert_eq!(save(&workbook), bytes);
}

#[test]
fn cell_watch_changes_are_refused() {
    let (bytes, mut workbook) = signed_workbook();
    let mut edit = workbook.edit_cell_watches(0).unwrap();
    edit.add(Watch::new(3, 5).unwrap()).unwrap();
    let commit = edit.commit().unwrap();
    let error = workbook.apply_cell_watches(0, &commit).unwrap_err();
    assert!(matches!(error, PackageError::Signed), "{error:?}");
    assert_eq!(save(&workbook), bytes);
}

#[test]
fn raw_opc_edits_keep_the_signature_policy() {
    let (bytes, mut workbook) = signed_workbook();
    // A closure that changes nothing keeps the signature byte for byte;
    // `edit_opc` used to drop it unconditionally.
    workbook.edit_opc(|_| Ok(())).unwrap();
    assert_eq!(save(&workbook), bytes);

    // Typed setters that publish through `edit_opc` refuse a change.
    let error = workbook
        .set_merged_cell_ranges(0, &[MergedCell::new(20, 21, 0, 1)])
        .unwrap_err();
    assert_signed_refusal(&error);
    let error = workbook.set_slicer_caches(Vec::new()).unwrap_err();
    assert_signed_refusal(&error);
    assert_eq!(save(&workbook), bytes);

    // Removing the signature inside the closure is an explicit disposition.
    let uri = PackURI::new("/docProps/litchi-probe.xml").unwrap();
    workbook
        .edit_opc(|package| {
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
    assert!(!carries_signature(&save(&workbook)));
}

#[test]
fn connections_refuse_a_signed_workbook() {
    use litchi_xlsb::package::connections::{
        CommandType, Connection, Connections, DbProperties, Properties, SourceType,
    };
    let (bytes, mut workbook) = signed_workbook();
    let error = workbook
        .set_connections(Connections {
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
        })
        .unwrap_err();
    assert!(matches!(error, PackageError::Signed), "{error:?}");
    assert!(!workbook.remove_connections().unwrap());
    assert_eq!(save(&workbook), bytes);
}

#[test]
fn raw_threaded_comment_stores_refuse_a_signed_package() {
    let bytes = signed(&fixture("Simple.xlsb"));
    // Owned ingress retains the exact source, so an unchanged signed package
    // republishes it.
    let mut package = OpcPackage::from_reader(Cursor::new(bytes.clone())).unwrap();
    let worksheet = PackURI::new("/xl/worksheets/sheet1.bin").unwrap();
    let error = litchi_xlsb::comments::threaded::store_on_worksheet(
        &mut package,
        &worksheet,
        &litchi_xlsb::comments::threaded::Comments::default(),
    )
    .unwrap_err();
    assert!(matches!(error, PackageError::Signed), "{error:?}");
    // Removing an absent graph changes nothing and is not refused.
    assert!(!litchi_xlsb::comments::threaded::remove_graph(&mut package).unwrap());
    assert_eq!(PackageWriter::to_bytes(&package).unwrap(), bytes);
}

/// Follow-up F2 of record 0013: task panes and the Ribbon refuse a change to a
/// signed workbook up front, with the typed signed error. Task panes used to
/// refuse with the shared OOXML owner's `SignedSourceRequiresExplicitPolicy`
/// wrapped as `PackageError::Common`; the Ribbon used to remove the signature.
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
    fn signed_workbook_with(edit: impl FnOnce(&mut Workbook)) -> (Vec<u8>, Workbook) {
        let mut plain = Package::from_bytes(&fixture("Simple.xlsb"))
            .unwrap()
            .into_workbook()
            .unwrap();
        edit(&mut plain);
        let bytes = signed(&save(&plain));
        let workbook = Package::from_bytes(&bytes)
            .unwrap()
            .into_workbook()
            .unwrap();
        assert!(workbook.is_signed());
        (bytes, workbook)
    }

    #[track_caller]
    fn assert_signed<T: std::fmt::Debug>(result: litchi_xlsb::package::PackageResult<T>) {
        let error = result.expect_err("a change to a signed workbook must be refused");
        assert!(matches!(error, PackageError::Signed), "{error:?}");
    }

    #[test]
    fn task_pane_changes_are_refused_up_front_and_noops_kept() {
        let (bytes, mut workbook) = signed_workbook();
        assert_signed(workbook.put_task_panes(panes("add-in-1"), web::Conformance::Transitional));
        assert!(!workbook.remove_task_panes().unwrap());
        assert_eq!(save(&workbook), bytes);

        let (bytes, mut workbook) = signed_workbook_with(|plain| {
            plain
                .put_task_panes(panes("add-in-1"), web::Conformance::Transitional)
                .unwrap();
        });
        // Equal panes publish nothing.
        workbook
            .put_task_panes(panes("add-in-1"), web::Conformance::Transitional)
            .unwrap();
        assert_eq!(save(&workbook), bytes);
        assert_signed(workbook.put_task_panes(panes("add-in-2"), web::Conformance::Transitional));
        assert_signed(workbook.remove_task_panes());
        assert!(workbook.is_signed());
        assert_eq!(save(&workbook), bytes);

        workbook.unsign();
        assert!(workbook.remove_task_panes().unwrap());
        let output = save(&workbook);
        assert!(!carries_signature(&output));
        let reopened = Package::from_bytes(&output)
            .unwrap()
            .into_workbook()
            .unwrap();
        assert!(reopened.task_panes().unwrap().is_none());
    }

    #[test]
    fn ribbon_changes_are_refused_up_front_and_noops_kept() {
        let (bytes, mut workbook) = signed_workbook();
        assert_signed(workbook.put_ribbon(Version::V2007, UI_2007.to_vec()));
        assert!(!workbook.remove_ribbon(Family::Legacy).unwrap());
        assert_eq!(save(&workbook), bytes);

        let (bytes, mut workbook) = signed_workbook_with(|plain| {
            plain.put_ribbon(Version::V2007, UI_2007.to_vec()).unwrap();
        });
        // The stored version and bytes publish nothing; an absent family is
        // not removed.
        workbook
            .put_ribbon(Version::V2007, UI_2007.to_vec())
            .unwrap();
        assert!(!workbook.remove_ribbon(Family::Modern).unwrap());
        assert_eq!(save(&workbook), bytes);
        assert_signed(workbook.put_ribbon(Version::V2007, UI_2007_CHANGED.to_vec()));
        assert_signed(workbook.remove_ribbon(Family::Legacy));
        assert!(workbook.is_signed());
        assert_eq!(save(&workbook), bytes);

        workbook.unsign();
        workbook
            .put_ribbon(Version::V2007, UI_2007_CHANGED.to_vec())
            .unwrap();
        let output = save(&workbook);
        assert!(!carries_signature(&output));
        let reopened = Package::from_bytes(&output)
            .unwrap()
            .into_workbook()
            .unwrap();
        assert_eq!(
            reopened.ribbon().unwrap().legacy().unwrap().xml(),
            UI_2007_CHANGED
        );
    }
}

#[test]
fn consuming_workbook_into_package_preserves_signed_source() {
    let (bytes, workbook) = signed_workbook();
    let package = workbook.into_package();
    assert!(package.opc_package().is_signed());
    assert_eq!(package.to_bytes().unwrap(), bytes);
    assert_eq!(save(&package.into_workbook().unwrap()), bytes);
}

#[test]
fn package_cell_removal_preserves_signed_noops_and_refuses_changed_publication() {
    let plain = Package::new().expect("new");
    let a1 = Reference::new(0, 0).expect("A1");
    let b1 = Reference::new(0, 1).expect("B1");
    let mut edit = plain.edit_cell_values(0).expect("edit").expect("sheet");
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
    let mut noop = source.edit_cell_values(0).expect("edit").expect("sheet");
    assert!(!noop.remove(b1).expect("missing"));
    let commit = noop.commit().expect("signed noop");
    assert!(commit.patch().is_noop());
    assert_eq!(commit.package().to_bytes().expect("bytes"), bytes);
    let mut edit = source.edit_cell_values(0).expect("edit").expect("sheet");
    assert!(edit.remove(a1).expect("stage removal"));
    assert!(matches!(edit.commit(), Err(PackageError::Signed)));
    assert_eq!(source.to_bytes().expect("source unchanged"), bytes);
}
