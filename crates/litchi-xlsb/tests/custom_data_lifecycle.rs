#![allow(
    clippy::expect_used,
    clippy::indexing_slicing,
    clippy::unwrap_used,
    reason = "Synthetic bounded OPC fixtures fail fast when their exact graph cannot be built."
)]

//! Public lifecycle coverage for the synthetic XLSB Custom Data profile.
//!
//! These fixtures are authored source/graph probes.  They make no claim that
//! a native producer emitted the package.  In particular, the connections
//! stream follows the pinned ExtConn14 grammar exactly:
//! `FRTBegin`, `BeginExtConn14`, `EndExtConn14`, `FRTEnd`.

use std::io::Cursor;
use std::num::{NonZeroU64, NonZeroUsize};

use litchi_core::{
    Budget, CancellationSource, ExecutionContext, ExecutionLimits, Limits as BudgetLimits,
};
use litchi_opc::phys_pkg::{PhysPkgReader, PhysPkgWriter};
use litchi_opc::{BlobPart, OpcPackage, PackURI, Part, XmlPart};
use litchi_xlsb::Package;
use litchi_xlsb::custom_data::{
    Commit, CustomData, Limits, Patch, Properties, RemovalDisposition, Snapshot, StorageId,
    StorageSelector, Transaction, write_properties,
};
use litchi_xlsb::package::PackageError;
use litchi_xlsb::raw::{Kind, Writer, kind as rt};
use litchi_xlsb::writer::{MutableWorksheet, WorkbookWriter};

const X14: &str = "http://schemas.microsoft.com/office/spreadsheetml/2009/9/main";
const WORKBOOK_URI: &str = "/xl/workbook.bin";
const CONNECTIONS_URI: &str = "/xl/connections.bin";
const CONNECTIONS_CONTENT_TYPE: &str = "application/vnd.ms-excel.connections";
const CONNECTIONS_RELATIONSHIP_TYPE: &str =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/connections";
const PROPERTIES_RELATIONSHIP_TYPE: &str =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/customDataProps";
const DATA_RELATIONSHIP_TYPE: &str =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/customData";
const PROPERTIES_CONTENT_TYPE: &str =
    "application/vnd.openxmlformats-officedocument.customDataProperties+xml";
const DATA_CONTENT_TYPE: &str = "application/binary";
const DATA_URI: &str = "/xl/customData/data1.bin";
const OPAQUE_URI: &str = "/xl/opaque.bin";
const OPAQUE_CONTENT_TYPE: &str = "application/octet-stream";
const SIGNATURE_ORIGIN_URI: &str = "/_xmlsignatures/origin.sigs";
const SIGNATURE_ORIGIN_RELATIONSHIP_TYPE: &str =
    "http://schemas.openxmlformats.org/package/2006/relationships/digital-signature/origin";
const SIGNATURE_ORIGIN_CONTENT_TYPE: &str =
    "application/vnd.openxmlformats-package.digital-signature-origin";

#[derive(Debug, Clone)]
struct Fixture {
    package: Package,
    payload: Vec<u8>,
    connections: Vec<u8>,
    opaque: Vec<u8>,
}

fn uri(value: &str) -> PackURI {
    PackURI::new(value).expect("synthetic fixture URI is valid")
}

fn wide(value: &str) -> Vec<u8> {
    let units = value.encode_utf16().collect::<Vec<_>>();
    let mut output = Vec::with_capacity(4 + units.len() * 2);
    output.extend_from_slice(
        &u32::try_from(units.len())
            .expect("synthetic fixture string length fits u32")
            .to_le_bytes(),
    );
    for unit in units {
        output.extend_from_slice(&unit.to_le_bytes());
    }
    output
}

fn record(kind: Kind, payload: &[u8], output: &mut Vec<u8>) {
    Writer::new(output)
        .write_record(kind, payload)
        .expect("synthetic BIFF12 record writes");
}

fn ext_conn_payload(connection_id: u32, name: &str) -> Vec<u8> {
    // BrtBeginExtConnection, §2.4.80.  The source type at offset 10 is
    // DBTOLEDB (5), which is required for an admitted ExtConn14 collection.
    let mut output = vec![7, 5, 2, 0];
    output.extend_from_slice(&30_u16.to_le_bytes());
    output.extend_from_slice(&0_u16.to_le_bytes());
    output.extend_from_slice(&(1_u16 << 3).to_le_bytes());
    output.extend_from_slice(&5_u32.to_le_bytes());
    output.extend_from_slice(&3_u32.to_le_bytes());
    output.extend_from_slice(&connection_id.to_le_bytes());
    output.push(0);
    output.extend_from_slice(&wide(name));
    output
}

fn ext_conn14_payload(culture: &str, uid: &str) -> Vec<u8> {
    let mut output = vec![0; 4];
    output.extend_from_slice(&wide(culture));
    output.extend_from_slice(&wide(uid));
    output
}

fn connections_stream(references: &[&str], opaque_frt: bool) -> Vec<u8> {
    let mut output = Vec::new();
    record(rt::BEGIN_EXT_CONNECTIONS, &[], &mut output);
    record(
        rt::BEGIN_EXT_CONNECTION,
        &ext_conn_payload(42, "Synthetic Custom Data"),
        &mut output,
    );
    for (ordinal, uid) in references.iter().enumerate() {
        // A malformed product/version header makes this balanced block opaque
        // while retaining the same complete ExtConn14 grammar for refusal
        // coverage.
        let frt = if opaque_frt {
            [1, 0, 0, 0x80]
        } else {
            [1, 0, 1, 0]
        };
        record(rt::FRT_BEGIN, &frt, &mut output);
        record(
            rt::BEGIN_EXT_CONN14,
            &ext_conn14_payload(if ordinal == 0 { "en-US" } else { "fr-FR" }, uid),
            &mut output,
        );
        record(rt::END_EXT_CONN14, &[], &mut output);
        record(rt::FRT_END, &[], &mut output);
    }
    record(rt::END_EXT_CONNECTION, &[], &mut output);
    record(rt::END_EXT_CONNECTIONS, &[], &mut output);
    output
}

fn properties_xml(uid: &str) -> Vec<u8> {
    format!(
        r#"<?xml version="1.0" encoding="UTF-8"?><!-- retained before root --><alias:datastoreItem xmlns:alias="{X14}" xmlns:s="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:v="urn:synthetic-vendor" id="{uid}"><?root-pi?><alias:extLst><?extension-pi?><s:ext uri="urn:synthetic"><v:opaque v:raw="&amp;raw"><v:leaf><![CDATA[<? retained-looking text ]]></v:leaf><!-- opaque comment --></v:opaque></s:ext></alias:extLst><!-- retained after extension --></alias:datastoreItem>"#
    )
    .into_bytes()
}

fn formatted_properties_xml(uid: &str) -> Vec<u8> {
    formatted_properties_xml_with_padding(uid, 256)
}

fn oversized_formatted_properties_xml(uid: &str) -> Vec<u8> {
    formatted_properties_xml_with_padding(uid, 1_024)
}

fn formatted_properties_xml_with_padding(uid: &str, padding_units: usize) -> Vec<u8> {
    let retained_root_padding = "  ".repeat(padding_units);
    format!(
        "<?xml version=\"1.0\" encoding=\"UTF-8\"?>\r\n<!-- retained before root -->\r\n<alias:datastoreItem xmlns:alias=\"{X14}\" xmlns:s=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" xmlns:v=\"urn:synthetic-vendor\" id=\"{uid}\">\r\n{retained_root_padding}<?root-pi?>\r\n  <alias:extLst>\r\n    <?extension-pi?>\r\n    <s:ext uri=\"urn:synthetic\">\r\n      <v:opaque v:raw=\"&amp;raw\">\r\n        <v:leaf><![CDATA[<? retained-looking text ]]></v:leaf>\r\n        <!-- opaque comment -->\r\n      </v:opaque>\r\n    </s:ext>\r\n  </alias:extLst>\r\n  <!-- retained after extension -->\r\n</alias:datastoreItem>\r\n"
    )
    .into_bytes()
}

fn add_relationship(
    package: &mut OpcPackage,
    source: &str,
    reltype: &str,
    target: &str,
    id: &str,
    external: bool,
) {
    package
        .get_part_mut(&uri(source))
        .expect("synthetic relationship source exists")
        .rels_mut()
        .add_relationship(
            reltype.to_owned(),
            target.to_owned(),
            id.to_owned(),
            external,
        );
}

fn synthetic_fixture(
    storage_ids: &[&str],
    references: &[&str],
    payload: &[u8],
    opaque_frt: bool,
) -> Fixture {
    let connections = connections_stream(references, opaque_frt);
    let opaque = b"opaque member source bytes".to_vec();

    let mut writer = WorkbookWriter::new();
    writer.add_worksheet(MutableWorksheet::new("Sheet1"));
    let mut output = Cursor::new(Vec::new());
    writer.save(&mut output).expect("base XLSB package writes");
    let base = Package::from_bytes(output.into_inner()).expect("base XLSB package reopens");
    let mut package = base.into_opc();

    for (index, uid) in storage_ids.iter().enumerate() {
        let properties_uri = format!("/xl/customData/props{}.xml", index + 1);
        let data_uri = format!("/xl/customData/data{}.bin", index + 1);
        let properties_part = XmlPart::new(
            uri(&properties_uri),
            PROPERTIES_CONTENT_TYPE.to_owned(),
            properties_xml(uid),
        );
        let mut properties_part = properties_part;
        properties_part.rels_mut().add_relationship(
            DATA_RELATIONSHIP_TYPE.to_owned(),
            format!("data{}.bin", index + 1),
            format!("rIdCustomData{}", index + 1),
            false,
        );
        package.add_part(Box::new(properties_part));
        package.add_part(Box::new(BlobPart::new(
            uri(&data_uri),
            DATA_CONTENT_TYPE.to_owned(),
            payload.to_vec(),
        )));
        add_relationship(
            &mut package,
            WORKBOOK_URI,
            PROPERTIES_RELATIONSHIP_TYPE,
            &format!("customData/props{}.xml", index + 1),
            &format!("rIdCustomDataProps{}", index + 1),
            false,
        );
    }
    package.add_part(Box::new(BlobPart::new(
        uri(CONNECTIONS_URI),
        CONNECTIONS_CONTENT_TYPE.to_owned(),
        connections.clone(),
    )));
    package.add_part(Box::new(BlobPart::new(
        uri(OPAQUE_URI),
        OPAQUE_CONTENT_TYPE.to_owned(),
        opaque.clone(),
    )));
    add_relationship(
        &mut package,
        WORKBOOK_URI,
        CONNECTIONS_RELATIONSHIP_TYPE,
        "connections.bin",
        "rIdConnections",
        false,
    );

    let mut package_bytes = Vec::new();
    package
        .to_stream(&mut package_bytes)
        .expect("synthetic source package serializes");
    let package = Package::from_bytes(package_bytes)
        .expect("synthetic XLSB custom-data graph reopens with source provenance");
    Fixture {
        package,
        payload: payload.to_vec(),
        connections,
        opaque,
    }
}

fn single_fixture() -> Fixture {
    synthetic_fixture(&["uid-A"], &["uid-A"], b"\0opaque\xffpayload\n", false)
}

fn many_reference_fixture() -> Fixture {
    synthetic_fixture(
        &["uid-A"],
        &["uid-A", "uid-A"],
        b"\0opaque\xffpayload\n",
        false,
    )
}

fn repeated_reference_fixture(count: usize) -> Fixture {
    let references = vec!["uid-A"; count];
    synthetic_fixture(&["uid-A"], &references, b"\0opaque\xffpayload\n", false)
}

fn retarget_fixture() -> Fixture {
    synthetic_fixture(
        &["uid-A", "uid-B"],
        &["uid-A", "uid-A"],
        b"\0opaque\xffpayload\n",
        false,
    )
}

fn opaque_frt_fixture() -> Fixture {
    synthetic_fixture(&["uid-A"], &["uid-A"], b"payload", true)
}

fn part_bytes(package: &Package, part_name: &str) -> Vec<u8> {
    package
        .opc_package()
        .get_part(&uri(part_name))
        .expect("synthetic fixture part exists")
        .blob()
        .to_vec()
}

fn package_bytes(package: &Package) -> Vec<u8> {
    package.to_bytes().expect("synthetic package serializes")
}

fn logical_package_size(package: &Package) -> usize {
    let opc = package.opc_package();
    let root = uri("/");
    let mut total = opc
        .source_content_types()
        .expect("source content types exist")
        .bytes()
        .len();
    total += opc
        .source_relationships(&root)
        .expect("source package relationships exist")
        .bytes()
        .len();
    for part in opc
        .try_iter_parts()
        .map(|part| part.expect("part payload decodes"))
    {
        total += part.blob().len();
        total += opc
            .source_relationships(part.partname())
            .expect("source part relationships exist")
            .bytes()
            .len();
    }
    total
}

fn max_relationship_xml_size(package: &Package) -> usize {
    let opc = package.opc_package();
    let root = uri("/");
    let mut maximum = opc
        .source_content_types()
        .expect("source content types exist")
        .bytes()
        .len()
        .max(
            opc.source_relationships(&root)
                .expect("source package relationships exist")
                .bytes()
                .len(),
        );
    for part in opc.iter_parts() {
        maximum = maximum.max(
            opc.source_relationships(part.partname())
                .expect("source part relationships exist")
                .bytes()
                .len(),
        );
    }
    maximum
}

fn relationship_xml_size(package: &Package, owner: &str) -> usize {
    package
        .opc_package()
        .source_relationships(&uri(owner))
        .expect("source relationship member exists")
        .bytes()
        .len()
}

fn package_with_part_bytes(package: Package, part_name: &str, bytes: Vec<u8>) -> Package {
    let mut opc = package.into_opc();
    opc.get_part_mut(&uri(part_name))
        .expect("synthetic stale part exists")
        .set_blob(bytes);
    Package::from_opc(opc).expect("modified synthetic package remains an XLSB package")
}

fn rewrite_physical_member(bytes: &[u8], wanted: &str, replacement: Vec<u8>) -> Vec<u8> {
    let reader = PhysPkgReader::new(bytes).expect("synthetic package is a physical ZIP");
    let mut writer = PhysPkgWriter::new();
    for member in reader
        .member_names()
        .expect("synthetic ZIP member names are readable")
    {
        let payload = if member == wanted {
            replacement.clone()
        } else {
            reader
                .read_member(&member)
                .expect("synthetic ZIP member bytes are readable")
        };
        writer
            .write(
                &PackURI::new(format!("/{member}"))
                    .expect("physical ZIP member maps to a package URI"),
                &payload,
            )
            .expect("synthetic physical package member writes");
    }
    writer.finish().expect("synthetic physical package closes")
}

fn source_preserving_fixture_with_properties(
    references: &[&str],
    source_properties: Vec<u8>,
) -> Fixture {
    let compact = synthetic_fixture(&["uid-A"], references, b"\0opaque\xffpayload\n", false);
    let physical = rewrite_physical_member(
        &package_bytes(&compact.package),
        "xl/customData/props1.xml",
        source_properties,
    );
    let package = Package::from_bytes(physical)
        .expect("formatted source XML reopens through the physical package reader");
    Fixture {
        package,
        payload: compact.payload,
        connections: compact.connections,
        opaque: compact.opaque,
    }
}

fn source_preserving_fixture_with_references(references: &[&str]) -> Fixture {
    source_preserving_fixture_with_properties(references, formatted_properties_xml("uid-A"))
}

fn source_preserving_fixture() -> Fixture {
    source_preserving_fixture_with_references(&["uid-A"])
}

fn unreferenced_source_preserving_fixture() -> Fixture {
    source_preserving_fixture_with_properties(&[], oversized_formatted_properties_xml("uid-A"))
}

fn managed_context() -> (CancellationSource, ExecutionContext) {
    let budget = Budget::root(
        "xlsb-custom-data-lifecycle-test",
        BudgetLimits::new(u64::MAX, u64::MAX, u64::MAX, u64::MAX, u64::MAX, u64::MAX),
    );
    let (cancellation_source, cancellation) = CancellationSource::pair();
    let limits = ExecutionLimits::new(
        NonZeroUsize::new(1).expect("one worker"),
        NonZeroUsize::new(1).expect("one in-flight task"),
        NonZeroU64::new(u64::MAX).expect("finite byte ceiling"),
        0,
    )
    .expect("managed execution limits");
    (
        cancellation_source,
        ExecutionContext::new(budget, cancellation, limits),
    )
}

fn assert_patch_conflict<T: std::fmt::Debug>(result: Result<T, PackageError>) {
    let diagnostic = format!("{result:?}");
    assert!(
        matches!(result, Err(PackageError::PatchConflict { .. })),
        "expected a source conflict, got {diagnostic}"
    );
}

#[test]
fn public_custom_data_crud_reads_typed_properties_and_inert_bytes() {
    let fixture = single_fixture();
    let package = &fixture.package;
    let snapshot: Snapshot = package.custom_data().expect("public Custom Data snapshot");
    let direct_snapshot = Snapshot::load(package.opc_package()).expect("direct snapshot load");
    assert_eq!(direct_snapshot.len(), snapshot.len());
    assert_eq!(snapshot.len(), 1);
    assert!(!snapshot.is_empty());
    let storage = snapshot.find("uid-A").expect("uid-A storage");
    assert_eq!(storage.id(), "uid-A");
    assert_eq!(storage.data(), fixture.payload.as_slice());
    assert_eq!(storage.properties().id, "uid-A");
    assert!(storage.properties().extension_list.is_some());
    assert_eq!(snapshot.connection_references("uid-A"), 1);
    assert!(!snapshot.has_opaque_reference_candidates());

    let properties: Properties = storage.properties().clone();
    let mut transaction: Transaction = package.edit_custom_data().expect("public edit transaction");
    let replacement = CustomData {
        properties,
        data: b"replacement".to_vec(),
    };
    assert!(
        transaction
            .set(StorageSelector::Id("uid-A"), replacement)
            .expect("replace complete storage")
    );
    assert!(
        transaction
            .set_data(
                StorageId::new("uid-A").expect("checked UID"),
                b"final".to_vec()
            )
            .expect("replace inert payload")
    );
    let inserted = transaction
        .insert(CustomData::new("uid-B", vec![7, 8]))
        .expect("insert storage");
    assert_eq!(inserted.as_str(), "uid-B");
    assert_eq!(transaction.storages().len(), 2);
    let upserted = transaction
        .upsert(CustomData::new("uid-B", vec![9]))
        .expect("upsert storage");
    assert_eq!(upserted.as_str(), "uid-B");
    let removed = transaction
        .remove(StorageSelector::Id("uid-B"))
        .expect("remove unreferenced storage")
        .expect("storage was present");
    assert_eq!(removed.id(), "uid-B");

    let commit: Commit = transaction.commit().expect("publish CRUD transaction");
    assert!(commit.changed());
    let changed = package
        .apply_custom_data(&commit)
        .expect("apply public Custom Data commit");
    let changed_snapshot = changed.custom_data().expect("read changed snapshot");
    assert_eq!(changed_snapshot.len(), 1);
    assert_eq!(changed_snapshot.find("uid-A").unwrap().data(), b"final");
    assert_eq!(changed_snapshot.connection_references("uid-A"), 1);
    assert_eq!(part_bytes(&changed, DATA_URI), b"final");
}

#[test]
fn rename_updates_all_effective_references_and_preserves_properties_payload() {
    let fixture = many_reference_fixture();
    let before_connections = part_bytes(&fixture.package, CONNECTIONS_URI);
    let before_payload = part_bytes(&fixture.package, "/xl/customData/data1.bin");
    let before_properties = part_bytes(&fixture.package, "/xl/customData/props1.xml");
    let mut transaction = fixture
        .package
        .edit_custom_data()
        .expect("rename transaction");
    assert!(
        transaction
            .rename(StorageSelector::Id("uid-A"), "renamed-A & <tag>")
            .expect("rename referenced storage")
    );
    let commit: Commit = transaction.commit().expect("publish rename");
    let patch: Patch = commit.patch().clone();
    let changed = fixture
        .package
        .apply_custom_data_patch(&patch)
        .expect("apply rename patch");
    let snapshot = changed.custom_data().expect("read renamed snapshot");
    assert!(snapshot.find("renamed-A & <tag>").is_some());
    assert_eq!(snapshot.connection_references("renamed-A & <tag>"), 2);
    assert_eq!(
        part_bytes(&changed, "/xl/customData/data1.bin"),
        before_payload
    );
    assert_ne!(
        part_bytes(&changed, "/xl/customData/props1.xml"),
        before_properties
    );
    assert_ne!(part_bytes(&changed, CONNECTIONS_URI), before_connections);
    assert!(
        String::from_utf8(part_bytes(&changed, "/xl/customData/props1.xml"))
            .expect("properties remain XML")
            .contains("opaque")
    );
}

#[test]
fn rename_to_a_later_uid_commits_across_sorted_catalog_and_inverts_exactly() {
    let fixture = synthetic_fixture(
        &["uid-A", "uid-B"],
        &["uid-A", "uid-B"],
        b"\0opaque\xffpayload\n",
        false,
    );
    let before = package_bytes(&fixture.package);
    let mut transaction = fixture
        .package
        .edit_custom_data()
        .expect("UID-order rename transaction");
    assert_eq!(transaction.storages()[0].id(), "uid-A");
    assert_eq!(transaction.storages()[1].id(), "uid-B");
    transaction
        .rename(StorageSelector::Id("uid-A"), "uid-Z")
        .expect("rename to later UID stages");
    let commit = transaction.commit().expect("UID-order rename commits");
    let changed = fixture
        .package
        .apply_custom_data(&commit)
        .expect("UID-order rename applies publicly");
    let snapshot = changed.custom_data().expect("read renamed catalog");
    let ids = snapshot
        .storages()
        .iter()
        .map(|storage| storage.id())
        .collect::<Vec<_>>();
    assert_eq!(ids, vec!["uid-B", "uid-Z"]);
    assert_eq!(snapshot.connection_references("uid-B"), 1);
    assert_eq!(snapshot.connection_references("uid-Z"), 1);
    assert_eq!(snapshot.connection_references("uid-A"), 0);

    let restored = changed
        .apply_custom_data_patch(&commit.patch().inverse())
        .expect("UID-order rename inverse applies");
    assert_eq!(package_bytes(&restored), before);
}

#[test]
fn inserting_an_earlier_uid_after_existing_storage_commits_and_inverts_exactly() {
    let fixture = synthetic_fixture(&["uid-Z"], &["uid-Z"], b"\0opaque\xffpayload\n", false);
    let before = package_bytes(&fixture.package);
    let mut transaction = fixture
        .package
        .edit_custom_data()
        .expect("UID-order insertion transaction");
    transaction
        .insert(CustomData::new("uid-A", b"inserted payload".to_vec()))
        .expect("insert earlier UID stages");
    let commit = transaction.commit().expect("UID-order insertion commits");
    let changed = fixture
        .package
        .apply_custom_data(&commit)
        .expect("UID-order insertion applies publicly");
    let snapshot = changed.custom_data().expect("read inserted catalog");
    let ids = snapshot
        .storages()
        .iter()
        .map(|storage| storage.id())
        .collect::<Vec<_>>();
    assert_eq!(ids, vec!["uid-A", "uid-Z"]);
    assert_eq!(snapshot.find("uid-A").unwrap().data(), b"inserted payload");
    assert_eq!(snapshot.connection_references("uid-Z"), 1);
    assert_eq!(snapshot.connection_references("uid-A"), 0);

    let restored = changed
        .apply_custom_data_patch(&commit.patch().inverse())
        .expect("UID-order insertion inverse applies");
    assert_eq!(package_bytes(&restored), before);
}

#[test]
fn source_formatted_properties_preserve_raw_opaque_xml_through_public_apply_and_inverse() {
    let fixture = source_preserving_fixture();
    let source_properties = formatted_properties_xml("uid-A");
    assert_eq!(
        part_bytes(&fixture.package, "/xl/customData/props1.xml"),
        source_properties
    );
    let source_snapshot = fixture
        .package
        .custom_data()
        .expect("formatted source Custom Data reads");
    let source_extension = source_snapshot
        .find("uid-A")
        .expect("formatted source storage")
        .properties()
        .extension_list
        .as_ref()
        .expect("formatted source extension")
        .xml
        .clone();
    for marker in [
        b"\r\n".as_slice(),
        b"<?extension-pi?>".as_slice(),
        b"<!-- opaque comment -->".as_slice(),
        b"<![CDATA[<? retained-looking text ]]>".as_slice(),
    ] {
        assert!(
            source_extension
                .windows(marker.len())
                .any(|window| window == marker),
            "formatted source extension retained marker {:?}",
            marker
        );
    }

    let mut transaction = fixture
        .package
        .edit_custom_data()
        .expect("formatted source rename transaction");
    transaction
        .rename(StorageSelector::Index(0), "renamed-source")
        .expect("formatted source rename stages");
    let commit = transaction
        .commit()
        .expect("formatted source rename commits");
    let changed = fixture
        .package
        .apply_custom_data(&commit)
        .expect("formatted source public apply");
    assert_eq!(
        part_bytes(&changed, "/xl/customData/props1.xml"),
        formatted_properties_xml("renamed-source")
    );
    let changed_properties = part_bytes(&changed, "/xl/customData/props1.xml");
    for marker in [
        b"\r\n".as_slice(),
        b"<?root-pi?>".as_slice(),
        b"<?extension-pi?>".as_slice(),
        b"<!-- opaque comment -->".as_slice(),
        b"<![CDATA[<? retained-looking text ]]>".as_slice(),
    ] {
        assert!(
            changed_properties
                .windows(marker.len())
                .any(|window| window == marker),
            "changed source XML retained marker {:?}",
            marker
        );
    }

    let restored = changed
        .apply_custom_data_patch(&commit.patch().inverse())
        .expect("formatted source inverse applies");
    assert_eq!(
        part_bytes(&restored, "/xl/customData/props1.xml"),
        source_properties
    );
    assert_eq!(package_bytes(&restored), package_bytes(&fixture.package));
}

#[test]
fn inverse_refuses_oversized_source_template_under_temporary_cap_atomically() {
    let fixture = unreferenced_source_preserving_fixture();
    let source_properties = part_bytes(&fixture.package, "/xl/customData/props1.xml");
    let properties = fixture
        .package
        .custom_data()
        .expect("oversized-template source reads")
        .find("uid-A")
        .expect("oversized-template storage")
        .properties()
        .clone();
    let canonical_properties =
        write_properties(&properties).expect("canonical restoration properties serialize");
    assert!(
        source_properties.len() > canonical_properties.len(),
        "formatted source must exceed canonical restoration: source={}, canonical={}",
        source_properties.len(),
        canonical_properties.len()
    );
    let temporary_cap = relationship_xml_size(&fixture.package, WORKBOOK_URI).max(
        fixture
            .package
            .opc_package()
            .source_content_types()
            .expect("source content types exist")
            .bytes()
            .len(),
    );
    assert!(source_properties.len() > temporary_cap);
    let limits = Limits::new().with_max_temporary_bytes(temporary_cap);
    let before = package_bytes(&fixture.package);
    let mut transaction = fixture
        .package
        .edit_custom_data_with_limits(limits)
        .expect("temporary cap admits source snapshot");
    transaction
        .remove(StorageSelector::Index(0))
        .expect("unreferenced storage removal stages")
        .expect("source storage exists");
    let commit = transaction
        .commit()
        .expect("removal does not rewrite oversized source template");
    let removed = fixture
        .package
        .apply_custom_data(&commit)
        .expect("removal applies under temporary cap");
    assert!(removed.custom_data().unwrap().is_empty());
    let removed_before_inverse = package_bytes(&removed);

    let inverse = removed.apply_custom_data_patch(&commit.patch().inverse());
    match inverse {
        Err(PackageError::LimitExceeded {
            resource,
            actual,
            maximum,
        }) => {
            assert!(
                resource.contains("temporary"),
                "oversized restoration should name temporary resource: {resource}"
            );
            assert!(actual > maximum);
        },
        other => panic!("oversized source restoration unexpectedly succeeded: {other:?}"),
    }
    assert_eq!(package_bytes(&removed), removed_before_inverse);
    assert_ne!(removed_before_inverse, before);
}

#[test]
fn referenced_remove_requires_disposition_and_detach_is_reversible() {
    let fixture = retarget_fixture();
    let before_bytes = package_bytes(&fixture.package);
    let before_connections = part_bytes(&fixture.package, CONNECTIONS_URI);

    let mut reject = fixture
        .package
        .edit_custom_data()
        .expect("reject transaction");
    assert!(matches!(
        reject.remove(StorageSelector::Id("uid-A")),
        Err(PackageError::CustomDataReferenced { connections: 2, .. })
    ));
    assert!(!reject.is_changed());

    let mut detach = fixture
        .package
        .edit_custom_data()
        .expect("detach transaction");
    detach
        .remove_with(
            StorageSelector::Id("uid-A"),
            RemovalDisposition::DetachConnections,
        )
        .expect("detach references");
    let detached = detach.commit().expect("publish detach");
    let detached_package = fixture
        .package
        .apply_custom_data(&detached)
        .expect("apply detach");
    assert!(
        detached_package
            .custom_data()
            .unwrap()
            .find("uid-A")
            .is_none()
    );
    assert_eq!(
        detached_package
            .custom_data()
            .unwrap()
            .connection_references("uid-A"),
        0
    );
    let restored = detached_package
        .apply_custom_data_patch(&detached.patch().inverse())
        .expect("apply detach inverse");
    assert_eq!(package_bytes(&restored), before_bytes);
    assert_eq!(part_bytes(&restored, CONNECTIONS_URI), before_connections);
}

#[test]
fn retarget_removal_updates_every_reference_to_the_remaining_storage() {
    let fixture = retarget_fixture();
    let mut retarget = fixture
        .package
        .edit_custom_data()
        .expect("retarget transaction");
    retarget
        .remove_with(
            StorageSelector::Id("uid-A"),
            RemovalDisposition::RetargetConnections("uid-B".to_owned()),
        )
        .expect("retarget references");
    let retargeted = retarget.commit().expect("publish retarget");
    let retargeted_package = fixture
        .package
        .apply_custom_data_patch(retargeted.patch())
        .expect("apply retarget");
    let retargeted_snapshot = retargeted_package.custom_data().unwrap();
    assert!(retargeted_snapshot.find("uid-A").is_none());
    assert_eq!(retargeted_snapshot.connection_references("uid-B"), 2);
}

#[test]
fn noop_commit_is_publicly_empty_and_patch_inverse_round_trips_source() {
    let fixture = single_fixture();
    let before = package_bytes(&fixture.package);
    let noop = fixture
        .package
        .edit_custom_data()
        .expect("no-op transaction")
        .commit()
        .expect("no-op commit");
    assert!(!noop.changed());
    assert!(noop.patch().is_empty());
    let unchanged = fixture
        .package
        .apply_custom_data(&noop)
        .expect("apply no-op commit");
    assert_eq!(package_bytes(&unchanged), before);

    let mut transaction = fixture
        .package
        .edit_custom_data()
        .expect("inverse transaction");
    transaction
        .set_data(StorageSelector::Index(0), b"changed".to_vec())
        .expect("stage data change");
    transaction
        .rename(StorageSelector::Index(0), "inverse-target")
        .expect("stage rename");
    let commit = transaction.commit().expect("publish inverse source");
    let changed = fixture
        .package
        .apply_custom_data_patch(commit.patch())
        .expect("apply forward patch");
    let restored = changed
        .apply_custom_data_patch(&commit.patch().inverse())
        .expect("apply inverse patch");
    assert_eq!(package_bytes(&restored), before);
}

#[test]
fn stale_member_and_connection_readsets_refuse_without_mutation() {
    let fixture = single_fixture();
    let mut transaction = fixture
        .package
        .edit_custom_data()
        .expect("stale source transaction");
    transaction
        .set_data(StorageSelector::Index(0), b"new source".to_vec())
        .expect("stage source mutation");
    let commit = transaction.commit().expect("source commit");

    let stale_data = package_with_part_bytes(
        fixture.package.clone(),
        "/xl/customData/data1.bin",
        b"independent data edit".to_vec(),
    );
    stale_data
        .custom_data()
        .expect("stale data package remains a readable Custom Data source");
    stale_data
        .custom_data()
        .expect("stale data source can be read repeatedly");
    let stale_data_bytes = package_bytes(&stale_data);
    assert_patch_conflict(stale_data.apply_custom_data_patch(commit.patch()));
    assert_eq!(package_bytes(&stale_data), stale_data_bytes);

    let mut changed_connections = fixture.connections.clone();
    let marker = changed_connections
        .windows(2)
        .position(|window| window == [b'S', 0])
        .expect("synthetic connection name marker");
    changed_connections[marker] = b'T';
    let stale_connections = package_with_part_bytes(
        fixture.package.clone(),
        CONNECTIONS_URI,
        changed_connections,
    );
    let stale_connections_bytes = package_bytes(&stale_connections);
    assert_patch_conflict(stale_connections.apply_custom_data_patch(commit.patch()));
    assert_eq!(package_bytes(&stale_connections), stale_connections_bytes);
}

#[test]
fn patch_preserves_unrelated_opaque_member_or_fails_closed() {
    let fixture = single_fixture();
    let mut transaction = fixture
        .package
        .edit_custom_data()
        .expect("opaque-member transaction");
    transaction
        .rename(StorageSelector::Index(0), "opaque-safe-rename")
        .expect("stage Custom Data rename");
    let commit = transaction.commit().expect("opaque-member source commit");

    let candidate_opaque = b"candidate opaque bytes remain".to_vec();
    assert_eq!(part_bytes(&fixture.package, OPAQUE_URI), fixture.opaque);
    let candidate = package_with_part_bytes(
        fixture.package.clone(),
        OPAQUE_URI,
        candidate_opaque.clone(),
    );
    match candidate.apply_custom_data_patch(commit.patch()) {
        Ok(applied) => {
            assert_eq!(part_bytes(&applied, OPAQUE_URI), candidate_opaque);
            match applied.apply_custom_data_patch(&commit.patch().inverse()) {
                Ok(restored) => assert_eq!(part_bytes(&restored, OPAQUE_URI), candidate_opaque),
                Err(_) => assert_eq!(part_bytes(&applied, OPAQUE_URI), candidate_opaque),
            }
        },
        Err(_) => assert_eq!(part_bytes(&candidate, OPAQUE_URI), candidate_opaque),
    }

    // Exercise the inverse against an independently edited after-state as
    // well.  The synthetic source splice is the same one the commit records,
    // while only the unrelated opaque member differs from the patch source.
    let renamed = package_with_part_bytes(
        package_with_part_bytes(
            fixture.package.clone(),
            "/xl/customData/props1.xml",
            properties_xml("opaque-safe-rename"),
        ),
        CONNECTIONS_URI,
        connections_stream(&["opaque-safe-rename"], false),
    );
    let inverse_candidate = package_with_part_bytes(renamed, OPAQUE_URI, candidate_opaque.clone());
    match inverse_candidate.apply_custom_data_patch(&commit.patch().inverse()) {
        Ok(restored) => assert_eq!(part_bytes(&restored, OPAQUE_URI), candidate_opaque),
        Err(_) => assert_eq!(part_bytes(&inverse_candidate, OPAQUE_URI), candidate_opaque),
    }
}

#[test]
fn rename_growth_accepts_exact_wire_and_xml_caps_and_rejects_one_under() {
    let fixture = repeated_reference_fixture(32);
    let long_id = "renamed-custom-data-uid-with-a-materially-longer-wire-value";
    let source_connections = part_bytes(&fixture.package, CONNECTIONS_URI);
    let source_properties = part_bytes(&fixture.package, "/xl/customData/props1.xml");
    let replacement_references = vec![long_id; 32];
    let expected_connections = connections_stream(&replacement_references, false);
    let expected_properties = properties_xml(long_id);
    assert_eq!(source_connections, fixture.connections);
    assert_eq!(source_properties.len(), properties_xml("uid-A").len());
    assert!(expected_connections.len() > source_connections.len());
    assert!(expected_properties.len() > source_properties.len());
    assert!(source_connections.len() >= source_properties.len());

    let exact_limits = Limits::new()
        .with_max_connections_bytes(expected_connections.len())
        .with_max_properties_xml_bytes(expected_properties.len())
        .with_max_temporary_bytes(expected_connections.len());
    let mut exact = fixture
        .package
        .edit_custom_data_with_limits(exact_limits)
        .expect("exact rename limits admit the source");
    exact
        .rename(StorageSelector::Index(0), long_id)
        .expect("stage exact-cap rename");
    assert!(
        exact
            .commit()
            .expect("exact wire/XML caps are inclusive")
            .changed()
    );

    let source_before_failure = package_bytes(&fixture.package);
    let under_connections = Limits::new()
        .with_max_connections_bytes(expected_connections.len() - 1)
        .with_max_properties_xml_bytes(expected_properties.len())
        .with_max_temporary_bytes(expected_connections.len());
    let mut under = fixture
        .package
        .edit_custom_data_with_limits(under_connections)
        .expect("one-under connections cap admits the source");
    under
        .rename(StorageSelector::Index(0), long_id)
        .expect("stage one-under connections rename");
    assert!(matches!(
        under.commit(),
        Err(PackageError::LimitExceeded { .. })
    ));
    assert_eq!(package_bytes(&fixture.package), source_before_failure);

    let under_properties = Limits::new()
        .with_max_connections_bytes(expected_connections.len())
        .with_max_properties_xml_bytes(expected_properties.len() - 1)
        .with_max_temporary_bytes(expected_connections.len());
    let mut under = fixture
        .package
        .edit_custom_data_with_limits(under_properties)
        .expect("one-under properties cap admits the source");
    under
        .rename(StorageSelector::Index(0), long_id)
        .expect("stage one-under properties rename");
    assert!(matches!(
        under.commit(),
        Err(PackageError::LimitExceeded { .. })
    ));
    assert_eq!(package_bytes(&fixture.package), source_before_failure);

    let temporary_source_cap = Limits::new()
        .with_max_connections_bytes(expected_connections.len())
        .with_max_properties_xml_bytes(expected_properties.len())
        .with_max_temporary_bytes(source_connections.len());
    let mut temporary = fixture
        .package
        .edit_custom_data_with_limits(temporary_source_cap)
        .expect("source-sized temporary cap admits the source");
    temporary
        .rename(StorageSelector::Index(0), long_id)
        .expect("stage source-sized temporary rename");
    assert!(matches!(
        temporary.commit(),
        Err(PackageError::LimitExceeded { .. })
    ));
    assert_eq!(package_bytes(&fixture.package), source_before_failure);
}

#[test]
fn final_output_cap_counts_mixed_shrink_and_growth_without_overcharging() {
    let original_payload = b"0123456789abcdefghijklmnopqrstuv";
    let fixture = synthetic_fixture(&["uid-A", "uid-B"], &["uid-A"], original_payload, false);
    let mut baseline = fixture
        .package
        .edit_custom_data()
        .expect("mixed output baseline transaction");
    baseline
        .set_data(StorageSelector::Id("uid-A"), b"a".to_vec())
        .expect("stage shrinking payload");
    baseline
        .set_data(
            StorageSelector::Id("uid-B"),
            b"0123456789abcdefghijklmnopqrstuv0123456789".to_vec(),
        )
        .expect("stage growing payload");
    let baseline_commit = baseline.commit().expect("mixed output baseline commit");
    let candidate = fixture
        .package
        .apply_custom_data(&baseline_commit)
        .expect("apply mixed output baseline");
    let source_size = logical_package_size(&fixture.package);
    let candidate_size = logical_package_size(&candidate);
    assert!(candidate_size < source_size);

    let exact_limits = Limits::new().with_max_output_bytes(candidate_size);
    let mut exact = fixture
        .package
        .edit_custom_data_with_limits(exact_limits)
        .expect("exact final-output cap admits the source");
    exact
        .set_data(StorageSelector::Id("uid-A"), b"a".to_vec())
        .expect("stage exact-cap shrinking payload");
    exact
        .set_data(
            StorageSelector::Id("uid-B"),
            b"0123456789abcdefghijklmnopqrstuv0123456789".to_vec(),
        )
        .expect("stage exact-cap growing payload");
    assert!(
        exact
            .commit()
            .expect("exact aggregate output cap is inclusive")
            .changed()
    );

    let source_before_failure = package_bytes(&fixture.package);
    let mut under = fixture
        .package
        .edit_custom_data_with_limits(Limits::new().with_max_output_bytes(candidate_size - 1))
        .expect("one-under final-output cap admits the source");
    under
        .set_data(StorageSelector::Id("uid-A"), b"a".to_vec())
        .expect("stage one-under shrinking payload");
    under
        .set_data(
            StorageSelector::Id("uid-B"),
            b"0123456789abcdefghijklmnopqrstuv0123456789".to_vec(),
        )
        .expect("stage one-under growing payload");
    assert!(matches!(
        under.commit(),
        Err(PackageError::LimitExceeded { .. })
    ));
    assert_eq!(package_bytes(&fixture.package), source_before_failure);
}

#[test]
fn relationship_xml_cap_includes_inserted_edges_at_exact_boundary() {
    let fixture = single_fixture();
    let mut baseline = fixture
        .package
        .edit_custom_data()
        .expect("relationship insertion baseline transaction");
    baseline
        .insert(CustomData::new("uid-B", b"inserted".to_vec()))
        .expect("stage relationship insertion");
    let baseline_commit = baseline.commit().expect("relationship insertion commit");
    let candidate = fixture
        .package
        .apply_custom_data(&baseline_commit)
        .expect("apply relationship insertion");
    let source_workbook_relationships = relationship_xml_size(&fixture.package, WORKBOOK_URI);
    let candidate_workbook_relationships = relationship_xml_size(&candidate, WORKBOOK_URI);
    assert!(candidate_workbook_relationships > source_workbook_relationships);
    let exact = max_relationship_xml_size(&candidate);
    assert!(exact > 0);

    let mut exact_transaction = fixture
        .package
        .edit_custom_data_with_limits(Limits::new().with_max_relationship_xml_bytes(exact))
        .expect("exact relationship cap admits the source");
    exact_transaction
        .insert(CustomData::new("uid-B", b"inserted".to_vec()))
        .expect("stage exact relationship insertion");
    assert!(
        exact_transaction
            .commit()
            .expect("exact relationship cap is inclusive")
            .changed()
    );

    let source_before_failure = package_bytes(&fixture.package);
    let mut under = fixture
        .package
        .edit_custom_data_with_limits(Limits::new().with_max_relationship_xml_bytes(exact - 1))
        .expect("one-under relationship cap admits the source");
    under
        .insert(CustomData::new("uid-B", b"inserted".to_vec()))
        .expect("stage one-under relationship insertion");
    let under_result = under.commit();
    assert!(
        matches!(under_result, Err(PackageError::LimitExceeded { .. })),
        "one-under relationship cap unexpectedly admitted: exact={exact}, source_workbook={source_workbook_relationships}, candidate_workbook={candidate_workbook_relationships}, result={under_result:?}"
    );
    assert_eq!(package_bytes(&fixture.package), source_before_failure);
}

#[test]
fn storage_cardinality_limit_admits_exact_capacity_and_rejects_one_under() {
    let fixture = synthetic_fixture(&["uid-Z"], &["uid-Z"], b"\0opaque\xffpayload\n", false);
    let exact_limits = Limits::new().with_max_storages(2);
    let mut exact = fixture
        .package
        .edit_custom_data_with_limits(exact_limits)
        .expect("exact storage cardinality admits source");
    exact
        .insert(CustomData::new("uid-A", b"exact-capacity".to_vec()))
        .expect("exact storage cardinality admits second entry");
    let commit = exact.commit().expect("exact storage cardinality commits");
    let candidate = fixture
        .package
        .apply_custom_data(&commit)
        .expect("exact storage cardinality applies");
    assert_eq!(
        candidate
            .custom_data_with_limits(exact_limits)
            .unwrap()
            .len(),
        2
    );
    assert!(matches!(
        candidate.custom_data_with_limits(Limits::new().with_max_storages(1)),
        Err(PackageError::LimitExceeded {
            actual: 2,
            maximum: 1,
            ..
        })
    ));

    let mut one_under = fixture
        .package
        .edit_custom_data_with_limits(Limits::new().with_max_storages(1))
        .expect("one-under storage cardinality admits source");
    assert!(matches!(
        one_under.insert(CustomData::new("uid-A", b"one-under".to_vec())),
        Err(PackageError::LimitExceeded {
            resource: "Custom Data storages",
            actual: 2,
            maximum: 1,
        })
    ));
    assert_eq!(one_under.storages().len(), 1);
    assert_eq!(one_under.storages()[0].id(), "uid-Z");
}

#[test]
fn signed_source_allows_noop_but_refuses_changed_public_commit() {
    let fixture = single_fixture();
    let mut opc = fixture.package.clone().into_opc();
    opc.add_part(Box::new(BlobPart::new(
        uri(SIGNATURE_ORIGIN_URI),
        SIGNATURE_ORIGIN_CONTENT_TYPE.to_owned(),
        Vec::new(),
    )));
    opc.rels_mut().add_relationship(
        SIGNATURE_ORIGIN_RELATIONSHIP_TYPE.to_owned(),
        "_xmlsignatures/origin.sigs".to_owned(),
        "rIdSignatureOrigin".to_owned(),
        false,
    );
    let signed = Package::from_opc(opc).expect("signed XLSB package validates");
    assert!(signed.opc_package().is_signed());

    let noop = signed
        .edit_custom_data()
        .expect("signed no-op transaction")
        .commit()
        .expect("signed no-op remains readable");
    assert!(!noop.changed());

    let mut transaction = signed
        .edit_custom_data()
        .expect("signed mutation transaction");
    transaction
        .set_data(StorageSelector::Index(0), b"signed mutation".to_vec())
        .expect("stage signed mutation");
    assert!(matches!(transaction.commit(), Err(PackageError::Signed)));
}

#[test]
fn opaque_frt_reference_is_readable_but_identity_mutation_refuses() {
    let fixture = opaque_frt_fixture();
    let snapshot = fixture.package.custom_data().expect("opaque FRT snapshot");
    assert!(snapshot.has_opaque_reference_candidates());
    assert_eq!(snapshot.connection_references("uid-A"), 0);
    let before = package_bytes(&fixture.package);

    let mut transaction = fixture
        .package
        .edit_custom_data()
        .expect("opaque FRT transaction");
    transaction
        .rename(StorageSelector::Index(0), "must-refuse")
        .expect("identity edit stages before closure proof");
    assert!(matches!(
        transaction.commit(),
        Err(PackageError::UnsupportedFeature(_))
    ));
    assert_eq!(package_bytes(&fixture.package), before);
}

#[test]
fn limits_and_cancellation_reject_before_publication_and_keep_staging_atomic() {
    let fixture = single_fixture();
    let connections_limit = Limits::new().with_max_connections_bytes(fixture.connections.len() - 1);
    assert!(matches!(
        fixture.package.custom_data_with_limits(connections_limit),
        Err(PackageError::LimitExceeded { .. })
    ));

    let uid_limit = Limits::new().with_max_uid_units(3);
    let uid_limited = fixture.package.custom_data_with_limits(uid_limit);
    let uid_diagnostic = format!("{uid_limited:?}");
    assert!(
        uid_limited.is_err(),
        "expected a UID limit refusal, got {uid_diagnostic}"
    );

    let mut limited = fixture
        .package
        .edit_custom_data_with_limits(Limits::new().with_max_output_bytes(1))
        .expect("bounded transaction starts from source");
    limited
        .set_data(
            StorageSelector::Index(0),
            b"staged within transaction".to_vec(),
        )
        .expect("staging does not publish output");
    let before_staged = limited.storages()[0].data().to_vec();
    assert!(matches!(
        limited.commit(),
        Err(PackageError::LimitExceeded { .. })
    ));
    assert_eq!(before_staged, b"staged within transaction");

    let (cancellation_source, context) = managed_context();
    let mut cancelled = fixture
        .package
        .edit_custom_data_with_limits_and_context(Limits::new(), context)
        .expect("cancellable transaction starts");
    let before_id = cancelled.storages()[0].id().to_owned();
    cancellation_source.cancel();
    assert!(matches!(
        cancelled.rename(StorageSelector::Index(0), "cancelled"),
        Err(PackageError::Execution(_))
    ));
    assert!(!cancelled.is_changed());
    assert_eq!(cancelled.storages()[0].id(), before_id);

    let (cancelled_before_read, context) = managed_context();
    cancelled_before_read.cancel();
    assert!(matches!(
        fixture
            .package
            .custom_data_with_limits_and_context(Limits::new(), context),
        Err(PackageError::Execution(_))
    ));
}
