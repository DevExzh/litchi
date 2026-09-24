#![allow(
    clippy::unwrap_used,
    reason = "focused integration tests use panic-on-failure assertions"
)]

//! Public fixtures and contract checks for the first pivot extension owner.
//!
//! The owner under test is the `pivotTableDefinition/extLst/ext` payload with
//! URI `{C510F80B-63DE-4267-81D5-13C33094786E}`.  Its admissibility is a
//! workbook relationship closure: the workbook's
//! `{983426D0-5260-488c-9760-48F4B6AC55F4}` extension must contain exactly one
//! relationship reference to the same pivot-table part, and that part must
//! not be reached from a worksheet.  The cache side of the closure is an
//! external cache source with the required `{ABF5C744-AB39-4b91-8756-CFA1BBC848D5}`
//! `pivotCacheIdVersion` extension and, when present, the external-source F057
//! `sourceConnection` closure through the workbook connections owner.
//!
//! These tests intentionally build schema-shaped synthetic OPC packages.  No
//! native Office interoperability claim is made here, and no `cacheField/items`
//! child is invented.  The first batch exercises scalar culture/format
//! metadata only; list mutation, pivot evaluation, refresh, and rendering are
//! outside the owner contract.

use std::collections::HashSet;
use std::sync::Arc;

use litchi_opc::constants::{content_type as ct, relationship_type as rt};
use litchi_opc::{
    BlobPart, OpcError, OpcPackage, PackURI, PackageWriter, ReadResource, TargetMode,
};
use litchi_xlsx::pivot::server_formats::table_data::load as load_pivot_table_data;
use litchi_xlsx::pivot::server_formats::{
    AttributeEdit, PivotTableSelector, ServerFormat, ServerFormatEdit, apply_patch, edit, load,
};
use litchi_xlsx::{Cell, Error, ReadLimits, Value, Workbook};
use quick_xml::events::Event;
use quick_xml::reader::NsReader;

const TRANSITIONAL_MAIN: &str = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
const STRICT_MAIN: &str = "http://purl.oclc.org/ooxml/spreadsheetml/main";
const TRANSITIONAL_REL: &str =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const STRICT_REL: &str = "http://purl.oclc.org/ooxml/officeDocument/relationships";
const EXT_NS: &str = "http://schemas.microsoft.com/office/spreadsheetml/2010/11/main";
const CACHE_SOURCE_EXT_NS: &str = "http://schemas.microsoft.com/office/spreadsheetml/2009/9/main";
const MCE_NS: &str = "http://schemas.openxmlformats.org/markup-compatibility/2006";

const SERVER_FORMATS_URI: &str = "{C510F80B-63DE-4267-81D5-13C33094786E}";
const PIVOT_TABLE_DATA_URI: &str = "{44433962-1CF7-4059-B4EE-95C3D5FFCF73}";
const PIVOT_TABLE_REFERENCES_URI: &str = "{983426D0-5260-488c-9760-48F4B6AC55F4}";
const PIVOT_CACHE_DEFINITION_URI: &str = "{725AE2AE-9491-48BE-B2B4-4EB974FC3084}";
const CACHE_ID_VERSION_URI: &str = "{ABF5C744-AB39-4b91-8756-CFA1BBC848D5}";
const CACHE_SOURCE_URI: &str = "{F057638F-6D5F-4E77-A914-E7F072B9BCA8}";

const WORKBOOK_URI: &str = "/xl/workbook.xml";
const SHEET_URI: &str = "/xl/worksheets/sheet1.xml";
const TABLE_URI: &str = "/xl/pivotTables/pivotTable1.xml";
const WORKSHEET_TABLE_URI: &str = "/xl/pivotTables/worksheetPivot.xml";
const CACHE_URI: &str = "/xl/pivotCache/pivotCacheDefinition1.xml";
const CONNECTIONS_URI: &str = "/xl/connections.xml";
const CONNECTIONS_CONTENT_TYPE: &str =
    "application/vnd.openxmlformats-officedocument.spreadsheetml.connections+xml";
const CONNECTIONS_REL: &str =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/connections";
const STRICT_CONNECTIONS_REL: &str =
    "http://purl.oclc.org/ooxml/officeDocument/relationships/connections";

#[derive(Clone, Copy, Debug, PartialEq, Eq)]
enum Dialect {
    Transitional,
    Strict,
}

impl Dialect {
    const fn main(self) -> &'static str {
        match self {
            Self::Transitional => TRANSITIONAL_MAIN,
            Self::Strict => STRICT_MAIN,
        }
    }

    const fn rel(self) -> &'static str {
        match self {
            Self::Transitional => TRANSITIONAL_REL,
            Self::Strict => STRICT_REL,
        }
    }

    const fn worksheet_rel(self) -> &'static str {
        match self {
            Self::Transitional => rt::WORKSHEET,
            Self::Strict => rt::STRICT_WORKSHEET,
        }
    }

    const fn pivot_table_rel(self) -> &'static str {
        match self {
            Self::Transitional => rt::PIVOT_TABLE,
            Self::Strict => rt::STRICT_PIVOT_TABLE,
        }
    }

    const fn pivot_cache_rel(self) -> &'static str {
        match self {
            Self::Transitional => rt::PIVOT_CACHE_DEFINITION,
            Self::Strict => rt::STRICT_PIVOT_CACHE_DEFINITION,
        }
    }

    const fn office_document_rel(self) -> &'static str {
        match self {
            Self::Transitional => rt::OFFICE_DOCUMENT,
            Self::Strict => rt::STRICT_OFFICE_DOCUMENT,
        }
    }

    const fn connections_rel(self) -> &'static str {
        match self {
            Self::Transitional => CONNECTIONS_REL,
            Self::Strict => STRICT_CONNECTIONS_REL,
        }
    }
}

#[derive(Clone, Debug)]
struct FixtureOptions {
    dialect: Dialect,
    ext_prefix: &'static str,
    root_prefix: &'static str,
    server_prefix: &'static str,
    server_uri: &'static str,
    server_count: &'static str,
    location_ref: &'static str,
    server_formats: String,
    pivot_table_exts: String,
    pivot_table_outer_exts: String,
    workbook_exts: String,
    table_attrs: String,
    table_children: String,
    cache_exts: String,
    worksheet_incoming: bool,
    worksheet_duplicate_name: bool,
    signed: bool,
}

impl Default for FixtureOptions {
    fn default() -> Self {
        Self {
            dialect: Dialect::Transitional,
            ext_prefix: "x15",
            root_prefix: "",
            server_prefix: "x15",
            server_uri: SERVER_FORMATS_URI,
            server_count: "2",
            location_ref: "A1:C5",
            server_formats: concat!(
                r##"<x15:serverFormat culture="en-US" format="#,##0.00"/>"##,
                r#"<x15:serverFormat culture="_x005F_x005F_005F_x005F_"/>"#,
            )
            .to_owned(),
            pivot_table_exts: String::new(),
            pivot_table_outer_exts: String::new(),
            workbook_exts: String::new(),
            table_attrs: String::new(),
            table_children: String::new(),
            cache_exts: String::new(),
            worksheet_incoming: false,
            worksheet_duplicate_name: false,
            signed: false,
        }
    }
}

#[derive(Clone, Debug)]
struct Fixture {
    bytes: Vec<u8>,
    workbook_xml: String,
    table_xml: String,
    cache_xml: String,
    connections_xml: String,
}

impl Fixture {
    fn build(options: FixtureOptions) -> Self {
        let main = options.dialect.main();
        let rel = options.dialect.rel();
        let ext_prefix = options.ext_prefix;
        let root_prefix = options.root_prefix;
        let server_prefix = options.server_prefix;
        let main_namespace = if root_prefix.is_empty() {
            format!(r#"xmlns="{main}""#)
        } else {
            format!(r#"xmlns:{root_prefix}="{main}""#)
        };

        let workbook_xml = format!(
            r#"<{root}workbook {main_namespace} xmlns:r="{rel}" xmlns:{ext_prefix}="{EXT_NS}" xmlns:mc="{MCE_NS}" mc:Ignorable="{ext_prefix}"><{root}sheets><{root}sheet name="Sheet1" sheetId="1" r:id="rIdSheet"/></{root}sheets><{root}pivotCaches><{root}pivotCache cacheId="7" r:id="rIdCache"/></{root}pivotCaches><{root}extLst><{root}ext uri="{PIVOT_TABLE_REFERENCES_URI}"><{ext_prefix}:pivotTableReferences><{ext_prefix}:pivotTableReference r:id="rIdPivot"/></{ext_prefix}:pivotTableReferences></{root}ext></{root}extLst>{workbook_exts}</{root}workbook>"#,
            root = root_prefix,
            main_namespace = main_namespace,
            rel = rel,
            ext_prefix = ext_prefix,
            workbook_exts = options.workbook_exts,
        );

        // Keep the cache definition valid for the Non-Worksheet closure: the
        // source is external, its named source connection resolves through
        // the workbook connections owner, and the required cache-ID-version
        // extension is present.  There is deliberately no cacheField/items
        // child.
        let cache_xml = format!(
            r#"<{root}pivotCacheDefinition {main_namespace} xmlns:{ext_prefix}="{EXT_NS}" xmlns:x14="{CACHE_SOURCE_EXT_NS}" xmlns:mc="{MCE_NS}" mc:Ignorable="{ext_prefix} x14"><{root}cacheSource type="external"><{root}extLst><{root}ext uri="{CACHE_SOURCE_URI}"><x14:sourceConnection name="canonical"/></{root}ext></{root}extLst></{root}cacheSource><{root}cacheFields count="0"/><{root}extLst><{root}ext uri="{CACHE_ID_VERSION_URI}"><{ext_prefix}:pivotCacheIdVersion cacheIdSupportedVersion="15" cacheIdCreatedVersion="15"/></{root}ext></{root}extLst>{cache_exts}</{root}pivotCacheDefinition>"#,
            root = root_prefix,
            main_namespace = main_namespace,
            ext_prefix = ext_prefix,
            cache_exts = options.cache_exts,
        );

        let connections_xml = format!(
            r#"<{root}connections xmlns="{main}"><{root}connection id="7" name="canonical" type="1" refreshedVersion="7"/></{root}connections>"#,
            root = root_prefix,
            main = main,
        );

        let table_xml = format!(
            r#"<{root}pivotTableDefinition {main_namespace} xmlns:{server_prefix}="{EXT_NS}" xmlns:mc="{MCE_NS}" mc:Ignorable="{server_prefix}" name="Pivot" cacheId="7"{table_attrs}><{root}location ref="{location_ref}" firstHeaderRow="1" firstDataRow="2" firstDataCol="1"/><{root}pivotFields count="0"/>{table_children}<{root}extLst><{root}ext uri="{server_uri}"><{server_prefix}:pivotTableServerFormats count="{server_count}">{server_formats}</{server_prefix}:pivotTableServerFormats>{pivot_table_exts}</{root}ext>{pivot_table_outer_exts}</{root}extLst></{root}pivotTableDefinition>"#,
            root = root_prefix,
            main_namespace = main_namespace,
            server_prefix = server_prefix,
            table_attrs = options.table_attrs,
            table_children = options.table_children,
            server_uri = options.server_uri,
            server_count = options.server_count,
            location_ref = options.location_ref,
            server_formats = options.server_formats,
            pivot_table_exts = options.pivot_table_exts,
            pivot_table_outer_exts = options.pivot_table_outer_exts,
        );

        // Change 0750: the eager writer refuses to publish XML that is not
        // well-formed. A table that carries a character reference XML
        // forbids, on purpose, is published as a well-formed stand-in and its
        // bytes are put back after publication, so the reader still sees them.
        let published_table_xml = table_xml.replace("&#x1;", "");

        let mut package = OpcPackage::new();
        package
            .try_add_part(Box::new(BlobPart::new(
                PackURI::new(WORKBOOK_URI).unwrap(),
                ct::SML_SHEET_MAIN.to_owned(),
                workbook_xml.as_bytes().to_vec(),
            )))
            .unwrap();
        package
            .try_add_part(Box::new(BlobPart::new(
                PackURI::new(SHEET_URI).unwrap(),
                ct::SML_WORKSHEET.to_owned(),
                format!(
                    r#"<worksheet xmlns="{main}"><sheetData/></worksheet>"#,
                    main = main
                )
                .into_bytes(),
            )))
            .unwrap();
        package
            .try_add_part(Box::new(BlobPart::new(
                PackURI::new(TABLE_URI).unwrap(),
                ct::SML_PIVOT_TABLE.to_owned(),
                published_table_xml.as_bytes().to_vec(),
            )))
            .unwrap();
        if options.worksheet_duplicate_name {
            package
                .try_add_part(Box::new(BlobPart::new(
                    PackURI::new(WORKSHEET_TABLE_URI).unwrap(),
                    ct::SML_PIVOT_TABLE.to_owned(),
                    published_table_xml.as_bytes().to_vec(),
                )))
                .unwrap();
        }
        package
            .try_add_part(Box::new(BlobPart::new(
                PackURI::new(CACHE_URI).unwrap(),
                ct::SML_PIVOT_CACHE_DEFINITION.to_owned(),
                cache_xml.as_bytes().to_vec(),
            )))
            .unwrap();
        package
            .try_add_part(Box::new(BlobPart::new(
                PackURI::new(CONNECTIONS_URI).unwrap(),
                CONNECTIONS_CONTENT_TYPE.to_owned(),
                connections_xml.as_bytes().to_vec(),
            )))
            .unwrap();

        let workbook_uri = PackURI::new(WORKBOOK_URI).unwrap();
        let sheet_uri = PackURI::new(SHEET_URI).unwrap();
        let table_uri = PackURI::new(TABLE_URI).unwrap();
        let workbook = package.get_part_mut(&workbook_uri).unwrap();
        workbook
            .rels_mut()
            .try_add_relationship(
                options.dialect.worksheet_rel().to_owned(),
                "worksheets/sheet1.xml".to_owned(),
                "rIdSheet".to_owned(),
                TargetMode::Internal,
            )
            .unwrap();
        workbook
            .rels_mut()
            .try_add_relationship(
                options.dialect.pivot_table_rel().to_owned(),
                "pivotTables/pivotTable1.xml".to_owned(),
                "rIdPivot".to_owned(),
                TargetMode::Internal,
            )
            .unwrap();
        workbook
            .rels_mut()
            .try_add_relationship(
                options.dialect.pivot_cache_rel().to_owned(),
                "pivotCache/pivotCacheDefinition1.xml".to_owned(),
                "rIdCache".to_owned(),
                TargetMode::Internal,
            )
            .unwrap();
        workbook
            .rels_mut()
            .try_add_relationship(
                options.dialect.connections_rel().to_owned(),
                "connections.xml".to_owned(),
                "rIdConnections".to_owned(),
                TargetMode::Internal,
            )
            .unwrap();

        // The normal worksheet relation is intentionally unrelated to the
        // pivot table.  A caller can opt into an incoming worksheet edge to
        // exercise the Non-Worksheet rejection path.
        if options.worksheet_incoming {
            package
                .get_part_mut(&sheet_uri)
                .unwrap()
                .rels_mut()
                .try_add_relationship(
                    options.dialect.pivot_table_rel().to_owned(),
                    "../pivotTables/pivotTable1.xml".to_owned(),
                    "rIdPivot".to_owned(),
                    TargetMode::Internal,
                )
                .unwrap();
        }
        if options.worksheet_duplicate_name {
            package
                .get_part_mut(&sheet_uri)
                .unwrap()
                .rels_mut()
                .try_add_relationship(
                    options.dialect.pivot_table_rel().to_owned(),
                    "../pivotTables/worksheetPivot.xml".to_owned(),
                    "rIdWorksheetPivot".to_owned(),
                    TargetMode::Internal,
                )
                .unwrap();
        }
        package
            .get_part_mut(&table_uri)
            .unwrap()
            .rels_mut()
            .try_add_relationship(
                options.dialect.pivot_cache_rel().to_owned(),
                "../pivotCache/pivotCacheDefinition1.xml".to_owned(),
                "rIdCache".to_owned(),
                TargetMode::Internal,
            )
            .unwrap();
        package.relate_to("xl/workbook.xml", options.dialect.office_document_rel());

        if options.signed {
            package
                .try_add_part(Box::new(BlobPart::new(
                    PackURI::new("/_xmlsignatures/origin.sigs").unwrap(),
                    ct::OPC_DIGITAL_SIGNATURE_ORIGIN.to_owned(),
                    b"<origin/>".to_vec(),
                )))
                .unwrap();
            package.relate_to("_xmlsignatures/origin.sigs", rt::DIGITAL_SIGNATURE_ORIGIN);
        }

        Self {
            bytes: with_unaudited_table(
                PackageWriter::to_bytes(&package).unwrap(),
                &published_table_xml,
                &table_xml,
            ),
            workbook_xml,
            table_xml,
            cache_xml,
            connections_xml,
        }
    }

    fn package(&self) -> OpcPackage {
        OpcPackage::from_bytes(&self.bytes).unwrap()
    }
}

/// Put the fixture's table bytes back when the writer published a stand-in.
fn with_unaudited_table(archive: Vec<u8>, published: &str, table_xml: &str) -> Vec<u8> {
    if published == table_xml {
        return archive;
    }
    let tables = [
        TABLE_URI.trim_start_matches('/'),
        WORKSHEET_TABLE_URI.trim_start_matches('/'),
    ];
    let reader = litchi_opc::phys_pkg::PhysPkgReader::new(&archive).unwrap();
    let mut writer = soapberry_zip::office::StreamingArchiveWriter::new();
    for name in reader.member_names().unwrap() {
        let data = if tables.contains(&name.as_str()) {
            table_xml.as_bytes().to_vec()
        } else {
            reader.read_member(&name).unwrap()
        };
        writer.write_deflated_sized(&name, &data).unwrap();
    }
    writer.finish_to_bytes().unwrap()
}

fn prefixed_server_format_fixture(dialect: Dialect, prefix: &'static str) -> Fixture {
    Fixture::build(FixtureOptions {
        dialect,
        ext_prefix: prefix,
        server_prefix: prefix,
        server_formats: format!(
            r#"<{prefix}:serverFormat culture="first"/><{prefix}:serverFormat culture="second"/>"#
        ),
        ..FixtureOptions::default()
    })
}

#[test]
fn workbook_facade_resolves_server_formats_by_semantic_selector() {
    let fixture = Fixture::build(FixtureOptions::default());
    let workbook = Workbook::from_bytes(fixture.bytes.clone()).unwrap();

    let by_name = workbook.pivot_table_server_formats("Pivot").unwrap();
    let by_position = workbook
        .pivot_server_formats(PivotTableSelector::Position(0))
        .unwrap();

    assert_eq!(by_name.table_name(), "Pivot");
    assert_eq!(by_name.cache_id(), 7);
    assert_eq!(by_name.formats(), by_position.formats());
    assert_eq!(
        workbook
            .pivot_table_server_formats_source("Pivot")
            .unwrap()
            .source_xml(),
        fixture.table_xml.as_bytes()
    );
}

#[test]
fn workbook_facade_scalar_edit_commit_save_reopen_and_inverse() {
    let fixture = Fixture::build(FixtureOptions::default());
    let workbook = Workbook::from_bytes(fixture.bytes).unwrap();
    let source_xml = workbook
        .pivot_table_server_formats_source("Pivot")
        .unwrap()
        .source_xml()
        .to_vec();

    let no_op = workbook
        .edit_pivot_table("Pivot")
        .unwrap()
        .commit()
        .unwrap();
    assert!(!no_op.changed());
    assert!(no_op.patch().is_empty());
    assert_eq!(no_op.patch().before(), no_op.patch().after());
    assert_eq!(
        no_op
            .workbook()
            .pivot_table_server_formats_source("Pivot")
            .unwrap()
            .source_xml(),
        source_xml.as_slice()
    );

    let mut edit = workbook.edit_pivot_table("Pivot").unwrap();
    assert!(
        edit.update_server_format(
            0,
            ServerFormatEdit {
                culture: AttributeEdit::Keep,
                format: AttributeEdit::Clear,
            },
        )
        .unwrap()
    );
    assert!(
        edit.update_server_format(
            1,
            ServerFormatEdit {
                culture: AttributeEdit::Set("fr-FR".to_owned()),
                format: AttributeEdit::Set("0.00".to_owned()),
            },
        )
        .unwrap()
    );
    let commit = edit.commit().unwrap();
    assert!(commit.changed());
    assert_eq!(
        commit.snapshot().formats(),
        &[
            ServerFormat::new(Some("en-US".to_owned()), None),
            ServerFormat::new(Some("fr-FR".to_owned()), Some("0.00".to_owned())),
        ]
    );

    let directory = tempfile::tempdir().unwrap();
    let path = directory.path().join("pivot-server-formats.xlsx");
    commit.workbook().save(&path).unwrap();
    let reopened = Workbook::open(&path).unwrap();
    assert_eq!(
        reopened.pivot_table("Pivot").unwrap().formats(),
        commit.snapshot().formats()
    );

    let restored = commit
        .workbook()
        .apply_pivot_table_server_formats_patch(&commit.patch().inverse())
        .unwrap();
    assert!(restored.changed());
    assert_eq!(
        restored
            .workbook()
            .pivot_table_server_formats_source("Pivot")
            .unwrap()
            .source_xml(),
        source_xml.as_slice()
    );
}

#[test]
fn workbook_facade_patch_is_stale_checked_and_composes_with_ordinary_edit() {
    let fixture = Fixture::build(FixtureOptions::default());
    let workbook = Workbook::from_bytes(fixture.bytes).unwrap();

    let mut pivot_edit = workbook.edit_pivot_table("Pivot").unwrap();
    pivot_edit.set_culture(0, Some("de-DE".to_owned())).unwrap();
    let pivot_commit = pivot_edit.commit().unwrap();

    let mut ordinary_edit = workbook.edit().unwrap();
    ordinary_edit
        .sheet("Sheet1")
        .unwrap()
        .unwrap()
        .set("A1", 42_i32)
        .unwrap();
    let ordinary_commit = ordinary_edit.commit().unwrap();
    let mut composed_edit = ordinary_commit
        .workbook()
        .edit_pivot_table("Pivot")
        .unwrap();
    composed_edit
        .set_culture(0, Some("de-DE".to_owned()))
        .unwrap();
    let composed = composed_edit.commit().unwrap();
    assert_eq!(
        composed.workbook().pivot_table("Pivot").unwrap().formats()[0]
            .culture
            .as_deref(),
        Some("de-DE")
    );
    assert!(matches!(
        composed
            .workbook()
            .sheet("Sheet1")
            .unwrap()
            .unwrap()
            .cell("A1")
            .unwrap()
            .stored(),
        Some(Cell::Value(Value::Number(number))) if number.as_str() == "42"
    ));

    let mut divergent_edit = workbook.edit_pivot_table("Pivot").unwrap();
    divergent_edit
        .set_culture(0, Some("ja-JP".to_owned()))
        .unwrap();
    let divergent = divergent_edit.commit().unwrap();
    assert!(matches!(
        divergent
            .workbook()
            .apply_pivot_table_server_formats_patch(pivot_commit.patch()),
        Err(Error::PatchConflict { .. })
    ));
}

#[cfg(feature = "encryption")]
#[test]
fn ordinary_workbook_patch_apply_refuses_encrypted_provenance() {
    use litchi_xlsx::encryption::Mode as EncryptionMode;

    const PASSWORD: &str = "pivot-server-formats-password";
    let fixture = Fixture::build(FixtureOptions::default());
    let workbook = Workbook::from_bytes(fixture.bytes).unwrap();
    let mut transaction = workbook.edit_pivot_table("Pivot").unwrap();
    transaction
        .set_culture(0, Some("de-DE".to_owned()))
        .unwrap();
    let patch = transaction.commit().unwrap().patch().clone();
    assert!(!patch.is_empty());

    let encrypted_bytes = workbook
        .to_encrypted(PASSWORD, EncryptionMode::Standard)
        .unwrap();
    let encrypted = Workbook::from_bytes_with_password(encrypted_bytes, PASSWORD).unwrap();
    let before = encrypted.to_plain_bytes().unwrap();
    assert!(matches!(
        patch.apply(&encrypted),
        Err(Error::EncryptionPolicy {
            operation: "apply_pivot_table_server_formats_patch",
            ..
        })
    ));
    assert_eq!(encrypted.to_plain_bytes().unwrap(), before);
}

fn relationship_signature(package: &OpcPackage, part: &str) -> HashSet<String> {
    package
        .get_part(&PackURI::new(part).unwrap())
        .unwrap()
        .rels()
        .iter()
        .map(|relationship| {
            format!(
                "{}\u{1f}{}\u{1f}{}\u{1f}{:?}",
                relationship.r_id(),
                relationship.reltype(),
                relationship.target_ref(),
                relationship.target_mode()
            )
        })
        .collect()
}

#[test]
fn fixture_closes_exact_workbook_reference_cache_and_nonworksheet_geometry() {
    let fixture = Fixture::build(FixtureOptions::default());
    let package = fixture.package();
    let workbook = package
        .get_part(&PackURI::new(WORKBOOK_URI).unwrap())
        .unwrap();
    let table = package.get_part(&PackURI::new(TABLE_URI).unwrap()).unwrap();
    let cache = package.get_part(&PackURI::new(CACHE_URI).unwrap()).unwrap();

    assert!(fixture.workbook_xml.contains(PIVOT_TABLE_REFERENCES_URI));
    assert_eq!(
        fixture
            .workbook_xml
            .matches(r#"<x15:pivotTableReference r:id="rIdPivot"/>"#)
            .count(),
        1
    );
    assert!(fixture.table_xml.contains(SERVER_FORMATS_URI));
    assert!(fixture.cache_xml.contains(CACHE_ID_VERSION_URI));
    assert!(fixture.cache_xml.contains(CACHE_SOURCE_URI));
    assert!(fixture.cache_xml.contains(r#"type="external""#));
    assert!(fixture.connections_xml.contains(r#"name="canonical""#));
    assert_eq!(
        workbook
            .rels()
            .iter()
            .filter(|relationship| relationship.r_id() == "rIdPivot")
            .count(),
        1
    );
    assert_eq!(
        workbook
            .rels()
            .get("rIdPivot")
            .unwrap()
            .target_partname()
            .unwrap(),
        PackURI::new(TABLE_URI).unwrap()
    );
    assert_eq!(
        table
            .rels()
            .get("rIdCache")
            .unwrap()
            .target_partname()
            .unwrap(),
        PackURI::new(CACHE_URI).unwrap()
    );
    assert!(
        package
            .iter_parts()
            .flat_map(|part| part.rels().iter())
            .filter(|relationship| !relationship.is_external())
            .all(|relationship| relationship.target_partname().unwrap()
                != PackURI::new(TABLE_URI).unwrap()
                || relationship.r_id() == "rIdPivot")
    );
    assert!(
        package
            .get_part(&PackURI::new(SHEET_URI).unwrap())
            .unwrap()
            .rels()
            .iter()
            .all(|relationship| relationship
                .target_partname()
                .map(|target| target != PackURI::new(TABLE_URI).unwrap())
                .unwrap_or(true))
    );
    assert!(cache.rels().iter().next().is_none());
    assert_eq!(relationship_signature(&package, WORKBOOK_URI).len(), 4);
}

#[test]
fn strict_and_transitional_roots_keep_the_same_extension_qnames_and_relationship_closure() {
    let transitional = Fixture::build(FixtureOptions::default());
    let strict = Fixture::build(FixtureOptions {
        dialect: Dialect::Strict,
        ..FixtureOptions::default()
    });

    for (fixture, main, rel) in [
        (&transitional, TRANSITIONAL_MAIN, TRANSITIONAL_REL),
        (&strict, STRICT_MAIN, STRICT_REL),
    ] {
        assert!(fixture.workbook_xml.contains(main));
        assert!(fixture.workbook_xml.contains(rel));
        assert!(fixture.table_xml.contains(EXT_NS));
        assert_eq!(
            relationship_signature(&fixture.package(), WORKBOOK_URI).len(),
            4
        );
        assert_eq!(
            load(&fixture.package(), "Pivot").unwrap().formats().len(),
            2
        );
    }
}

#[test]
fn u32_whitespace_and_strict_transitional_relationship_ids_are_accepted() {
    for dialect in [Dialect::Transitional, Dialect::Strict] {
        let fixture = Fixture::build(FixtureOptions {
            dialect,
            server_count: " \n2\t",
            ..FixtureOptions::default()
        });
        let snapshot = load(&fixture.package(), "Pivot").unwrap();
        assert_eq!(snapshot.formats().len(), 2);
        assert_eq!(snapshot.cache_id(), 7);
    }
}

#[test]
fn c510_uri_token_padding_and_numeric_lexicals_survive_structural_rewrite() {
    let padded_count = " \n2\t";
    for padded_uri in [
        " \n{C510F80B-63DE-4267-81D5-13C33094786E}\t ",
        "&#x20;{C510F80B-63DE-4267-81D5-13C33094786E}&#x9;",
    ] {
        let fixture = Fixture::build(FixtureOptions {
            server_uri: padded_uri,
            server_count: padded_count,
            ..FixtureOptions::default()
        });
        let mut package = fixture.package();
        let before = table_blob(&package);
        let source = String::from_utf8(before.clone()).unwrap();
        assert!(source.contains(&format!(r#"uri="{padded_uri}""#)));
        assert!(source.contains(&format!(r#"count="{padded_count}""#)));
        assert_eq!(load(&package, "Pivot").unwrap().formats().len(), 2);

        let no_op = edit(&mut package, "Pivot").unwrap().commit().unwrap();
        assert!(!no_op.changed());
        assert_eq!(table_blob(&package), before);

        let mut transaction = edit(&mut package, "Pivot").unwrap();
        transaction.reorder_server_formats(&[1, 0]).unwrap();
        let commit = transaction.commit().unwrap();
        assert!(commit.changed());
        let changed = String::from_utf8(table_blob(&package)).unwrap();
        assert!(changed.contains(&format!(r#"uri="{padded_uri}""#)));
        assert!(changed.contains(&format!(r#"count="{padded_count}""#)));
        assert_eq!(load(&package, "Pivot").unwrap().formats().len(), 2);

        commit.patch().inverse().apply(&mut package).unwrap();
        assert_eq!(table_blob(&package), before);
    }
}

#[test]
fn c510_uri_internal_whitespace_is_not_a_token_match() {
    for server_uri in [
        "{C510F80B-63DE-4267 81D5-13C33094786E}",
        "{C510F80B-63DE-4267&#x20;81D5-13C33094786E}",
    ] {
        let fixture = Fixture::build(FixtureOptions {
            server_uri,
            ..FixtureOptions::default()
        });
        assert!(!fixture.table_xml.contains(SERVER_FORMATS_URI));
        assert!(load(&fixture.package(), "Pivot").is_err());
    }
}

#[test]
fn unknown_ext_uri_with_the_same_qname_is_not_a_server_format_owner() {
    let fixture = Fixture::build(FixtureOptions {
        server_uri: "{UNKNOWN-SERVER-FORMAT-URI}",
        ..FixtureOptions::default()
    });
    assert!(fixture.table_xml.contains("{UNKNOWN-SERVER-FORMAT-URI}"));
    assert!(
        fixture
            .table_xml
            .contains(r#"<x15:pivotTableServerFormats count="2">"#)
    );
    assert!(!fixture.table_xml.contains(SERVER_FORMATS_URI));
}

#[test]
fn duplicate_known_payloads_and_mce_branches_are_retained_as_source_diagnostics() {
    let fixture = Fixture::build(FixtureOptions {
        pivot_table_exts: concat!(
            r#"<x15:pivotTableServerFormats count="1"><x15:serverFormat/></x15:pivotTableServerFormats>"#,
            r#"<!-- retained sibling -->"#,
        )
        .to_owned(),
        workbook_exts: r#"<mc:AlternateContent><mc:Choice Requires="x15"><x15:future/></mc:Choice><mc:Fallback><x15:futureFallback/></mc:Fallback></mc:AlternateContent>"#.to_owned(),
        ..FixtureOptions::default()
    });
    assert!(fixture.table_xml.contains(SERVER_FORMATS_URI));
    assert!(fixture.table_xml.contains("retained sibling"));
    assert!(fixture.workbook_xml.contains("AlternateContent"));
    assert!(fixture.workbook_xml.contains("futureFallback"));
}

#[test]
fn mce_selected_owner_is_readable_but_scalar_publication_is_refused() {
    let fixture = Fixture::build(FixtureOptions::default());
    let mut package = fixture.package();
    let table_uri = PackURI::new(TABLE_URI).unwrap();
    let table = package.get_part_mut(&table_uri).unwrap();
    let source = std::str::from_utf8(table.blob()).unwrap();
    let wrapped = source
        .replace(
            r#"<x15:pivotTableServerFormats count="2">"#,
            r#"<mc:AlternateContent><mc:Choice Requires="x15"><x15:pivotTableServerFormats count="2">"#,
        )
        .replace(
            "</x15:pivotTableServerFormats>",
            "</x15:pivotTableServerFormats></mc:Choice><mc:Fallback><x15:futureFallback/></mc:Fallback></mc:AlternateContent>",
        );
    assert!(wrapped.contains("<mc:AlternateContent>"));
    table.set_blob(wrapped.into_bytes());

    let snapshot = load(&package, "Pivot").unwrap();
    assert!(snapshot.has_ambiguous_mce_owner());
    assert_eq!(snapshot.formats().len(), 2);
    let before = table_blob(&package);
    assert!(edit(&mut package, "Pivot").is_err());
    assert_eq!(table_blob(&package), before);
}

#[test]
fn no_invalid_cache_field_items_fixture_is_synthesized() {
    let fixture = Fixture::build(FixtureOptions::default());
    assert!(!fixture.cache_xml.contains("<items"));
    assert!(!fixture.cache_xml.contains(":items"));
    assert!(fixture.cache_xml.contains("<cacheFields count=\"0\""));
}

#[test]
fn optional_x14_pivot_cache_id_may_be_absent_when_its_owner_is_present() {
    let fixture = Fixture::build(FixtureOptions {
        cache_exts: format!(
            r#"<ext uri="{PIVOT_CACHE_DEFINITION_URI}"><x14:pivotCacheDefinition/></ext>"#
        ),
        ..FixtureOptions::default()
    });
    assert!(fixture.cache_xml.contains(PIVOT_CACHE_DEFINITION_URI));
    assert!(!fixture.cache_xml.contains("pivotCacheId=\""));
    assert_eq!(load(&fixture.package(), "Pivot").unwrap().cache_id(), 7);
}

fn table_blob(package: &OpcPackage) -> Vec<u8> {
    part_blob(package, TABLE_URI)
}

fn part_blob(package: &OpcPackage, part: &str) -> Vec<u8> {
    package
        .get_part(&PackURI::new(part).unwrap())
        .unwrap()
        .blob()
        .to_vec()
}

fn replace_default_server_format_children(package: &mut OpcPackage, replacement: &str) {
    // PackageWriter deliberately rejects non-compact publication input.  Open
    // the compact fixture first, then replace only the retained table bytes so
    // the owner scanner sees physical XML whitespace/comments rather than a
    // normalized reconstruction.
    let original = r##"<x15:serverFormat culture="en-US" format="#,##0.00"/><x15:serverFormat culture="_x005F_x005F_005F_x005F_"/>"##;
    replace_table_fragment(package, original, replacement);
}

fn replace_table_fragment(package: &mut OpcPackage, original: &str, replacement: &str) {
    let table_uri = PackURI::new(TABLE_URI).unwrap();
    let table = package.get_part_mut(&table_uri).unwrap();
    let source = String::from_utf8(table.blob().to_vec()).unwrap();
    let changed = source.replace(original, replacement);
    assert_ne!(changed, source);
    table.set_blob(changed.into_bytes());
}

fn oversized_extension(owner_uri: &str, payload: &str) -> String {
    format!(
        r#"<ext uri="{owner_uri}">{payload}<!--{}--></ext>"#,
        "x".repeat(1024 * 1024)
    )
}

fn oversized_payload(payload: &str) -> String {
    let close = payload
        .rfind("</")
        .expect("payload fixture must have an explicit closing tag");
    format!(
        "{}<!--{}-->{}",
        &payload[..close],
        "x".repeat(1024 * 1024),
        &payload[close..]
    )
}

fn append_owner_extension(package: &mut OpcPackage, part: &str, owner_uri: &str, duplicate: &str) {
    let part_uri = PackURI::new(part).unwrap();
    let target = package.get_part_mut(&part_uri).unwrap();
    let mut source = String::from_utf8(target.blob().to_vec()).unwrap();
    let owner_open = format!(r#"<ext uri="{owner_uri}">"#);
    let owner_start = source
        .find(&owner_open)
        .unwrap_or_else(|| panic!("fixture source missing owner: {owner_uri}"));
    let owner_end = owner_start
        + source[owner_start..]
            .find("</ext>")
            .expect("fixture owner missing closing ext")
        + "</ext>".len();
    source.insert_str(owner_end, duplicate);
    target.set_blob(source.into_bytes());
}

fn append_owner_payload(package: &mut OpcPackage, part: &str, owner_uri: &str, payload: &str) {
    let part_uri = PackURI::new(part).unwrap();
    let target = package.get_part_mut(&part_uri).unwrap();
    let mut source = String::from_utf8(target.blob().to_vec()).unwrap();
    let owner_open = format!(r#"<ext uri="{owner_uri}">"#);
    let owner_start = source
        .find(&owner_open)
        .unwrap_or_else(|| panic!("fixture source missing owner: {owner_uri}"));
    let owner_end = owner_start
        + source[owner_start..]
            .find("</ext>")
            .expect("fixture owner missing closing ext");
    source.insert_str(owner_end, payload);
    target.set_blob(source.into_bytes());
}

fn assert_fragment_resource_limit<T>(result: Result<T, Error>, owner: &str) {
    match result {
        Err(Error::ResourceLimit(limit)) => {
            assert!(
                limit.observed > limit.limit,
                "{owner} refusal must report an observed size above its limit: {limit:?}"
            );
            assert_eq!(
                limit.limit,
                1024 * 1024,
                "{owner} refusal must carry the one MiB fragment limit: {limit:?}"
            );
            assert!(
                limit
                    .scope
                    .to_ascii_lowercase()
                    .contains(&owner.to_ascii_lowercase()),
                "{owner} refusal lost its owner scope: {limit:?}"
            );
        },
        Err(error) => {
            panic!("expected typed resource refusal for oversized {owner} extension, got {error:?}")
        },
        Ok(_) => panic!("accepted oversized {owner} extension"),
    }
}

#[derive(Clone, Copy, Debug, Default)]
struct RelationshipInventory {
    parts: usize,
    total_xml_bytes: usize,
    max_xml_bytes: usize,
    total_xml_events: usize,
    max_xml_events: usize,
    total_edges: usize,
    max_edges: usize,
    graph_nodes: usize,
    max_target_bytes: usize,
    max_part_name_bytes: usize,
    max_relationship_member_name_bytes: usize,
}

fn relationship_xml_event_count(bytes: &[u8]) -> usize {
    let mut reader = NsReader::from_reader(bytes);
    reader.config_mut().trim_text(false);
    reader.config_mut().check_end_names = true;
    let mut events = 0;
    loop {
        events += 1;
        if matches!(reader.read_event().unwrap(), Event::Eof) {
            return events;
        }
    }
}

fn relationship_inventory(package: &OpcPackage) -> RelationshipInventory {
    let mut owners = vec![PackURI::new("/").unwrap()];
    owners.extend(package.iter_parts().map(|part| part.partname().clone()));

    let mut inventory = RelationshipInventory::default();
    let mut graph_nodes = HashSet::new();
    for owner in owners {
        if owner.as_str() != "/" {
            inventory.max_part_name_bytes = inventory.max_part_name_bytes.max(owner.as_str().len());
        }
        inventory.max_relationship_member_name_bytes = inventory
            .max_relationship_member_name_bytes
            .max(owner.rels_uri().unwrap().membername().len());

        let relationships = if owner.as_str() == "/" {
            package.rels()
        } else {
            package.get_part(&owner).unwrap().rels()
        };
        inventory.total_edges += relationships.len();
        inventory.max_edges = inventory.max_edges.max(relationships.len());
        for relationship in relationships.iter() {
            inventory.max_target_bytes = inventory
                .max_target_bytes
                .max(relationship.target_ref().len());
            if !relationship.is_external() {
                let target = relationship.target_partname().unwrap();
                inventory.max_target_bytes = inventory.max_target_bytes.max(target.as_str().len());
                graph_nodes.insert(target.as_str().to_ascii_lowercase());
            }
        }

        let source = package.source_relationships(&owner).unwrap();
        if source.member_present() {
            inventory.parts += 1;
            inventory.total_xml_bytes += source.bytes().len();
            inventory.max_xml_bytes = inventory.max_xml_bytes.max(source.bytes().len());
            let events = relationship_xml_event_count(source.bytes());
            inventory.total_xml_events += events;
            inventory.max_xml_events = inventory.max_xml_events.max(events);
        }
    }
    inventory.graph_nodes = graph_nodes.len();
    inventory
}

fn add_unknown_relationships(package: &mut OpcPackage) {
    let root_target = format!("unknown-root-target-{}", "r".repeat(512));
    package
        .rels_mut()
        .try_add_relationship(
            "urn:test:pivot-root-unknown".to_owned(),
            root_target,
            "rIdUnknownRoot".to_owned(),
            TargetMode::Internal,
        )
        .unwrap();

    let sheet_target = format!("../unknown-sheet-target-{}", "s".repeat(32));
    package
        .get_part_mut(&PackURI::new(SHEET_URI).unwrap())
        .unwrap()
        .rels_mut()
        .try_add_relationship(
            "urn:test:pivot-sheet-unknown".to_owned(),
            sheet_target,
            "rIdUnknownSheet".to_owned(),
            TargetMode::Internal,
        )
        .unwrap();
}

fn assert_read_limit<T>(result: Result<T, Error>, resource: ReadResource) {
    assert!(matches!(
        result,
        Err(Error::Package(OpcError::ReadLimit {
            resource: actual, ..
        })) if actual == resource
    ));
}

fn pivot_data_extension(child: &str) -> String {
    format!(
        r#"<ext uri="{PIVOT_TABLE_DATA_URI}"><x15:pivotTableData>{child}</x15:pivotTableData></ext>"#
    )
}

fn pivot_value_cell(in_value: &str) -> String {
    format!(r#"<x15:pivotRow><x15:c><x15:x in="{in_value}"/></x15:c></x15:pivotRow>"#)
}

#[test]
fn typed_reads_are_selector_first_and_close_the_logical_cache_identity() {
    let fixture = Fixture::build(FixtureOptions::default());
    let package = fixture.package();

    let by_name = load(&package, PivotTableSelector::Name("Pivot")).unwrap();
    assert_eq!(by_name.table_name(), "Pivot");
    assert_eq!(by_name.cache_id(), 7);
    assert_eq!(by_name.formats().len(), 2);
    assert_eq!(
        by_name.formats()[0],
        ServerFormat::new(Some("en-US".to_owned()), Some("#,##0.00".to_owned()))
    );
    assert_eq!(
        by_name.formats()[1],
        ServerFormat::new(Some("_x005F_005F_".to_owned()), None)
    );

    let by_position = load(&package, PivotTableSelector::Position(0)).unwrap();
    assert_eq!(by_position.formats(), by_name.formats());
    assert!(load(&package, TABLE_URI).is_err());
    assert!(load(&package, PivotTableSelector::Position(1)).is_err());
}

#[test]
fn server_format_children_use_the_extension_namespace_and_allow_legal_spacing() {
    let fixture = Fixture::build(FixtureOptions::default());
    let mut package = fixture.package();
    let spaced_children = concat!(
        "<!-- before -->\n  ",
        r#"<x15:serverFormat culture="en-US"/>"#,
        "\n  <!-- between -->\n\t",
        r#"<x15:serverFormat format="0.00"/>"#,
        "\n",
    );
    replace_default_server_format_children(&mut package, spaced_children);

    let snapshot = load(&package, "Pivot").unwrap();
    assert_eq!(snapshot.formats().len(), 2);
    assert_eq!(snapshot.formats()[0].culture.as_deref(), Some("en-US"));
    assert_eq!(snapshot.formats()[1].format.as_deref(), Some("0.00"));
}

#[test]
fn core_namespace_child_and_nested_server_format_content_are_rejected() {
    let core_child = Fixture::build(FixtureOptions {
        server_formats: r#"<serverFormat culture="en-US"/><serverFormat/>"#.to_owned(),
        ..FixtureOptions::default()
    });
    assert!(load(&core_child.package(), "Pivot").is_err());

    let nested_leaf = Fixture::build(FixtureOptions {
        server_formats: r#"<x15:serverFormat><x15:opaque/></x15:serverFormat><x15:serverFormat/>"#
            .to_owned(),
        ..FixtureOptions::default()
    });
    assert!(load(&nested_leaf.package(), "Pivot").is_err());

    let non_whitespace_text = Fixture::build(FixtureOptions {
        server_formats: r#"<x15:serverFormat>text</x15:serverFormat><x15:serverFormat/>"#
            .to_owned(),
        ..FixtureOptions::default()
    });
    assert!(load(&non_whitespace_text.package(), "Pivot").is_err());

    let whitespace_leaf = Fixture::build(FixtureOptions::default());
    let mut whitespace_leaf = whitespace_leaf.package();
    replace_default_server_format_children(
        &mut whitespace_leaf,
        r#"<x15:serverFormat> </x15:serverFormat><x15:serverFormat/>"#,
    );
    assert!(load(&whitespace_leaf, "Pivot").is_err());

    let cdata_leaf = Fixture::build(FixtureOptions::default());
    let mut cdata_leaf = cdata_leaf.package();
    replace_default_server_format_children(
        &mut cdata_leaf,
        r#"<x15:serverFormat><![CDATA[text]]></x15:serverFormat><x15:serverFormat/>"#,
    );
    assert!(load(&cdata_leaf, "Pivot").is_err());
}

#[test]
fn pivot_server_formats_parent_rejects_literal_cdata_references_as_content() {
    for marker in ["<![CDATA[&#x20;]]>", "<![CDATA[&#x9;]]>"] {
        let fixture = Fixture::build(FixtureOptions {
            server_formats: format!(
                r#"{marker}<x15:serverFormat culture="en-US"/><x15:serverFormat/>"#
            ),
            ..FixtureOptions::default()
        });
        assert!(
            load(&fixture.package(), "Pivot").is_err(),
            "accepted literal CDATA character-reference content: {marker}"
        );
    }
}

#[test]
fn oversized_duplicate_pivot_table_references_owner_or_payload_is_refused_before_diagnostics() {
    let fixture = Fixture::build(FixtureOptions::default());
    let duplicate = oversized_extension(
        PIVOT_TABLE_REFERENCES_URI,
        r#"<x15:pivotTableReferences><x15:pivotTableReference r:id="rIdPivot"/></x15:pivotTableReferences>"#,
    );
    let mut package = fixture.package();
    append_owner_extension(
        &mut package,
        WORKBOOK_URI,
        PIVOT_TABLE_REFERENCES_URI,
        &duplicate,
    );
    assert_fragment_resource_limit(load(&package, "Pivot"), "pivotTableReferences");

    let mut payload_duplicate = fixture.package();
    append_owner_payload(
        &mut payload_duplicate,
        WORKBOOK_URI,
        PIVOT_TABLE_REFERENCES_URI,
        &oversized_payload(
            r#"<x15:pivotTableReferences><x15:pivotTableReference r:id="rIdPivot"/></x15:pivotTableReferences>"#,
        ),
    );
    assert_fragment_resource_limit(load(&payload_duplicate, "Pivot"), "pivotTableReferences");
}

#[test]
fn third_oversized_pivot_table_reference_payload_is_capped_in_diagnostic_mode() {
    let payload = r#"<x15:pivotTableReferences><x15:pivotTableReference r:id="rIdPivot"/></x15:pivotTableReferences>"#;
    let c444 = format!(
        r#"<ext uri="{PIVOT_TABLE_DATA_URI}"><x15:pivotTableData rowCount="1" columnCount="1" cacheId="7"><x15:pivotRow r="0"><x15:c/></x15:pivotRow></x15:pivotTableData></ext>"#
    );
    let fixture = Fixture::build(FixtureOptions {
        table_children: r#"<rowItems count="1"/><colItems count="1"/>"#.to_owned(),
        pivot_table_outer_exts: c444,
        ..FixtureOptions::default()
    });
    let mut package = fixture.package();
    append_owner_payload(
        &mut package,
        WORKBOOK_URI,
        PIVOT_TABLE_REFERENCES_URI,
        payload,
    );
    append_owner_payload(
        &mut package,
        WORKBOOK_URI,
        PIVOT_TABLE_REFERENCES_URI,
        &oversized_payload(payload),
    );

    // C444 uses the shared graph's diagnostic/reference path, where duplicate
    // C983 payloads are retained read-only.  The third recognized payload must
    // still be charged before that diagnostic result is published.
    assert_fragment_resource_limit(
        load_pivot_table_data(&package, "Pivot"),
        "pivotTableReferences",
    );
}

#[test]
fn strict_root_and_alternate_prefixes_keep_typed_values() {
    let fixture = Fixture::build(FixtureOptions {
        dialect: Dialect::Strict,
        ext_prefix: "p",
        server_prefix: "p",
        server_formats: concat!(
            r##"<p:serverFormat culture="&quot;fr-FR&quot;" format="_x005F_x0041_"/>"##,
            r#"<p:serverFormat culture="_x0001_"/>"#,
        )
        .to_owned(),
        ..FixtureOptions::default()
    });
    let snapshot = load(&fixture.package(), "Pivot").unwrap();
    assert_eq!(
        snapshot.formats(),
        [
            ServerFormat::new(Some("\"fr-FR\"".to_owned()), Some("_x0041_".to_owned())),
            ServerFormat::new(Some("\u{1}".to_owned()), None),
        ]
    );
}

#[test]
fn scalar_set_replace_clear_preserves_children_count_and_opaque_siblings() {
    let fixture = Fixture::build(FixtureOptions {
        pivot_table_exts:
            r#"<!--opaque sibling--><x15:unknown xmlns:x15="urn:unknown" data="keep"/>"#.to_owned(),
        workbook_exts: r#"<mc:AlternateContent><mc:Choice Requires="x15"><x15:future/></mc:Choice><mc:Fallback><x15:futureFallback/></mc:Fallback></mc:AlternateContent>"#.to_owned(),
        cache_exts: r#"<extLst><ext uri="{opaque-cache}"><x15:futureCache/></ext></extLst>"#.to_owned(),
        ..FixtureOptions::default()
    });
    let original = fixture.package();
    let original_xml = table_blob(&original);
    let original_owner_start = load(&original, "Pivot").unwrap().source_owner_range().start;
    let original_workbook = part_blob(&original, WORKBOOK_URI);
    let original_sheet = part_blob(&original, SHEET_URI);
    let original_cache = part_blob(&original, CACHE_URI);
    let original_connections = part_blob(&original, CONNECTIONS_URI);
    let original_workbook_rels = relationship_signature(&original, WORKBOOK_URI);
    let original_table_rels = relationship_signature(&original, TABLE_URI);
    let mut package = original.clone();

    let mut transaction = edit(&mut package, "Pivot").unwrap();
    assert!(
        transaction
            .set_culture(0, Some("de-DE".to_owned()))
            .unwrap()
    );
    assert!(transaction.clear_format(0).unwrap());
    assert!(transaction.set_format(1, Some("0.00%".to_owned())).unwrap());
    let commit = transaction.commit().unwrap();
    assert!(commit.changed());
    let changed_xml = table_blob(&package);
    assert_ne!(changed_xml, original_xml);
    assert!(
        changed_xml
            .windows(b"count=\"2\"".len())
            .any(|window| { window == b"count=\"2\"" })
    );
    assert!(
        changed_xml
            .windows(b"opaque sibling".len())
            .any(|window| { window == b"opaque sibling" })
    );
    assert!(
        changed_xml
            .windows(b"data=\"keep\"".len())
            .any(|window| { window == b"data=\"keep\"" })
    );
    assert!(
        changed_xml
            .windows(b"ref=\"A1:C5\"".len())
            .any(|window| { window == b"ref=\"A1:C5\"" })
    );
    assert_eq!(
        load(&package, "Pivot").unwrap().source_owner_range().start,
        original_owner_start
    );
    assert_eq!(part_blob(&package, WORKBOOK_URI), original_workbook);
    assert_eq!(part_blob(&package, SHEET_URI), original_sheet);
    assert_eq!(part_blob(&package, CACHE_URI), original_cache);
    assert_eq!(part_blob(&package, CONNECTIONS_URI), original_connections);
    assert_eq!(
        relationship_signature(&package, WORKBOOK_URI),
        original_workbook_rels
    );
    assert_eq!(
        relationship_signature(&package, TABLE_URI),
        original_table_rels
    );

    let after = load(&package, "Pivot").unwrap();
    assert_eq!(after.formats()[0].culture.as_deref(), Some("de-DE"));
    assert_eq!(after.formats()[0].format, None);
    assert_eq!(after.formats()[1].format.as_deref(), Some("0.00%"));
    assert_eq!(after.formats().len(), 2);

    let mut replay = original.clone();
    apply_patch(&mut replay, commit.patch()).unwrap();
    assert_eq!(table_blob(&replay), changed_xml);
    commit.patch().inverse().apply(&mut replay).unwrap();
    assert_eq!(table_blob(&replay), original_xml);
}

#[test]
fn changed_scalar_preserves_unchanged_entity_lexical_attributes() {
    let fixture = Fixture::build(FixtureOptions {
        server_formats:
            r#"<x15:serverFormat culture="en-US" format="&#34;&#10;"/><x15:serverFormat/>"#
                .to_owned(),
        ..FixtureOptions::default()
    });
    let mut package = fixture.package();
    let before = table_blob(&package);
    assert!(
        String::from_utf8(before.clone())
            .unwrap()
            .contains(r#"format="&#34;&#10;""#)
    );
    assert_eq!(
        load(&package, "Pivot").unwrap().formats()[0]
            .format
            .as_deref(),
        Some("\"\n")
    );

    let mut transaction = edit(&mut package, "Pivot").unwrap();
    assert!(
        transaction
            .set_culture(0, Some("de-DE".to_owned()))
            .unwrap()
    );
    transaction.commit().unwrap();
    let changed = String::from_utf8(table_blob(&package)).unwrap();
    assert!(changed.contains(r#"culture="de-DE""#));
    assert!(changed.contains(r#"format="&#34;&#10;""#));
    assert_ne!(changed.as_bytes(), before.as_slice());
}

#[test]
fn explicit_keep_set_clear_distinguishes_absent_attributes() {
    let fixture = Fixture::build(FixtureOptions::default());
    let mut package = fixture.package();
    let mut transaction = edit(&mut package, "Pivot").unwrap();

    assert!(
        transaction
            .update_server_format(
                0,
                ServerFormatEdit {
                    culture: AttributeEdit::Keep,
                    format: AttributeEdit::Set("0".to_owned()),
                },
            )
            .unwrap()
    );
    assert!(
        transaction
            .update_server_format(
                1,
                ServerFormatEdit {
                    culture: AttributeEdit::Clear,
                    format: AttributeEdit::Set("0.00".to_owned()),
                },
            )
            .unwrap()
    );
    let commit = transaction.commit().unwrap();
    assert!(commit.changed());
    let snapshot = load(&package, "Pivot").unwrap();
    assert_eq!(snapshot.formats()[0].culture.as_deref(), Some("en-US"));
    assert_eq!(snapshot.formats()[0].format.as_deref(), Some("0"));
    assert_eq!(snapshot.formats()[1].culture, None);
    assert_eq!(snapshot.formats()[1].format.as_deref(), Some("0.00"));
}

#[test]
fn exact_no_op_and_inverse_do_not_regenerate_the_owner() {
    let fixture = Fixture::build(FixtureOptions::default());
    let mut package = fixture.package();
    let before_snapshot = load(&package, "Pivot").unwrap();
    let before_source = before_snapshot.source_arc();
    let before = table_blob(&package);
    let mut transaction = edit(&mut package, "Pivot").unwrap();
    assert!(
        !transaction
            .set_server_format(
                0,
                ServerFormat::new(Some("en-US".to_owned()), Some("#,##0.00".to_owned())),
            )
            .unwrap()
    );
    let commit = transaction.commit().unwrap();
    assert!(!commit.changed());
    assert!(commit.patch().is_empty());
    assert!(Arc::ptr_eq(&before_source, &commit.snapshot().source_arc()));
    assert_eq!(table_blob(&package), before);
    commit.patch().inverse().apply(&mut package).unwrap();
    assert_eq!(table_blob(&package), before);
}

#[test]
fn structural_list_crud_preserves_order_count_comments_and_prefix_in_both_dialects() {
    for (dialect, prefix) in [(Dialect::Transitional, "x15"), (Dialect::Strict, "p")] {
        let fixture = prefixed_server_format_fixture(dialect, prefix);
        let mut package = fixture.package();
        let original = format!(
            r#"<{prefix}:serverFormat culture="first"/><{prefix}:serverFormat culture="second"/>"#
        );
        let with_comments = format!(
            r#"<!-- before -->
  <{prefix}:serverFormat culture="first"/>
  <!-- between -->
  <{prefix}:serverFormat culture="second"/>
  <!-- after -->"#
        );
        replace_table_fragment(&mut package, &original, &with_comments);

        let before = table_blob(&package);
        let mut transaction = edit(&mut package, "Pivot").unwrap();
        transaction
            .insert_server_format(
                1,
                ServerFormat::new(Some("inserted".to_owned()), Some("0.00%".to_owned())),
            )
            .unwrap();
        transaction.move_server_format(2, 0).unwrap();
        assert_eq!(
            transaction.remove_server_format(1).unwrap(),
            ServerFormat::new(Some("first".to_owned()), None)
        );
        transaction
            .push_server_format(ServerFormat::new(Some("tail".to_owned()), None))
            .unwrap();
        transaction.reorder_server_formats(&[2, 0, 1]).unwrap();
        let commit = transaction.commit().unwrap();
        assert!(commit.changed());

        let expected = [
            ServerFormat::new(Some("tail".to_owned()), None),
            ServerFormat::new(Some("second".to_owned()), None),
            ServerFormat::new(Some("inserted".to_owned()), Some("0.00%".to_owned())),
        ];
        assert_eq!(load(&package, "Pivot").unwrap().formats(), expected);
        let changed = String::from_utf8(table_blob(&package)).unwrap();
        assert!(changed.contains(r#"count="3""#));
        assert!(changed.contains("<!-- before -->"));
        assert!(changed.contains("<!-- between -->"));
        assert!(changed.contains("<!-- after -->"));
        assert!(changed.contains(&format!("<{prefix}:serverFormat culture=\"tail\"/>")));
        assert!(changed.contains(&format!("<{prefix}:serverFormat culture=\"second\"/>")));

        // This source intentionally retains physical comments/spacing.  The
        // ordinary PackageWriter publication contract is compact-only, while
        // the owner under test must still read and rewrite the retained XML.
        assert_eq!(load(&package, "Pivot").unwrap().formats(), expected);

        commit.patch().inverse().apply(&mut package).unwrap();
        assert_eq!(table_blob(&package), before);
    }
}

#[test]
fn ordinary_workbook_facade_supports_structural_list_edits_and_exact_inverse() {
    for (dialect, prefix) in [(Dialect::Transitional, "x15"), (Dialect::Strict, "p")] {
        let fixture = prefixed_server_format_fixture(dialect, prefix);
        let workbook = Workbook::from_bytes(fixture.bytes.clone()).unwrap();
        let original = workbook
            .pivot_table_server_formats_source("Pivot")
            .unwrap()
            .source_xml()
            .to_vec();

        let mut transaction = workbook.edit_pivot_table("Pivot").unwrap();
        transaction
            .insert_server_format(1, ServerFormat::new(Some("middle".to_owned()), None))
            .unwrap();
        transaction.move_server_format(2, 0).unwrap();
        assert_eq!(
            transaction.remove_server_format(1).unwrap(),
            ServerFormat::new(Some("first".to_owned()), None)
        );
        let commit = transaction.commit().unwrap();
        assert!(commit.changed());
        assert_eq!(
            commit.snapshot().formats(),
            &[
                ServerFormat::new(Some("second".to_owned()), None),
                ServerFormat::new(Some("middle".to_owned()), None),
            ]
        );

        let directory = tempfile::tempdir().unwrap();
        let path = directory.path().join("pivot-server-formats-list.xlsx");
        commit.workbook().save(&path).unwrap();
        let reopened = Workbook::open(&path).unwrap();
        assert_eq!(
            reopened.pivot_table("Pivot").unwrap().formats(),
            commit.snapshot().formats()
        );

        let restored = commit
            .workbook()
            .apply_pivot_table_server_formats_patch(&commit.patch().inverse())
            .unwrap();
        assert!(restored.changed());
        assert_eq!(
            restored
                .workbook()
                .pivot_table_server_formats_source("Pivot")
                .unwrap()
                .source_xml(),
            original.as_slice()
        );
    }
}

#[test]
fn equal_typed_reorder_retains_each_lexical_leaf_source() {
    let fixture = Fixture::build(FixtureOptions {
        server_formats: format!(
            r#"<x15:serverFormat culture="same" format="0"/><p:serverFormat xmlns:p="{EXT_NS}" format="0" culture="&#115;ame"/>"#
        ),
        ..FixtureOptions::default()
    });
    let mut package = fixture.package();
    let before = table_blob(&package);
    let snapshot = load(&package, "Pivot").unwrap();
    assert_eq!(snapshot.formats()[0], snapshot.formats()[1]);

    let mut transaction = edit(&mut package, "Pivot").unwrap();
    transaction.reorder_server_formats(&[1, 0]).unwrap();
    let commit = transaction.commit().unwrap();
    assert!(commit.changed());
    let changed = String::from_utf8(table_blob(&package)).unwrap();
    let prefixed = changed.find(r#"<p:serverFormat xmlns:p="#).unwrap();
    let defaulted = changed.find(r#"<x15:serverFormat culture="same""#).unwrap();
    assert!(prefixed < defaulted);
    assert!(changed.contains(r#"format="0" culture="&#115;ame""#));
    assert_eq!(
        load(&package, "Pivot").unwrap().formats(),
        snapshot.formats()
    );

    commit.patch().inverse().apply(&mut package).unwrap();
    assert_eq!(table_blob(&package), before);
}

#[test]
fn insert_remove_is_an_exact_noop_but_remove_insert_is_a_source_change() {
    let fixture = Fixture::build(FixtureOptions::default());
    let mut package = fixture.package();
    let before = table_blob(&package);
    let inserted = ServerFormat::new(Some("new".to_owned()), Some("0.00".to_owned()));
    let mut transaction = edit(&mut package, "Pivot").unwrap();
    transaction
        .insert_server_format(1, inserted.clone())
        .unwrap();
    assert_eq!(transaction.remove_server_format(1).unwrap(), inserted);
    let commit = transaction.commit().unwrap();
    assert!(!commit.changed());
    assert!(commit.patch().is_empty());
    assert_eq!(table_blob(&package), before);

    let fixture = Fixture::build(FixtureOptions {
        server_formats:
            r#"<x15:serverFormat culture="&#115;ame"/><x15:serverFormat culture="other"/>"#
                .to_owned(),
        ..FixtureOptions::default()
    });
    let mut package = fixture.package();
    let before = table_blob(&package);
    let expected = load(&package, "Pivot").unwrap().formats().to_vec();
    let mut transaction = edit(&mut package, "Pivot").unwrap();
    let removed = transaction.remove_server_format(0).unwrap();
    assert_eq!(removed, ServerFormat::new(Some("same".to_owned()), None));
    transaction.insert_server_format(0, removed).unwrap();
    let commit = transaction.commit().unwrap();
    assert!(commit.changed());
    assert!(!commit.patch().is_empty());
    assert_eq!(load(&package, "Pivot").unwrap().formats(), expected);
    assert_ne!(table_blob(&package), before);
}

#[test]
fn structural_list_cardinality_and_invalid_operations_are_failure_atomic() {
    let fixture = Fixture::build(FixtureOptions::default());
    let mut package = fixture.package();
    let before = table_blob(&package);
    let mut transaction = edit(&mut package, "Pivot").unwrap();
    assert!(
        transaction
            .insert_server_format(3, ServerFormat::new(None, None))
            .is_err()
    );
    assert!(transaction.remove_server_format(2).is_err());
    assert!(transaction.move_server_format(0, 2).is_err());
    assert!(transaction.reorder_server_formats(&[0, 0]).is_err());
    drop(transaction);
    assert_eq!(table_blob(&package), before);

    let mut package = fixture.package();
    let before = table_blob(&package);
    let mut transaction = edit(&mut package, "Pivot").unwrap();
    transaction.remove_server_format(0).unwrap();
    assert!(transaction.remove_server_format(0).is_err());
    drop(transaction);
    assert_eq!(table_blob(&package), before);

    let mut package = fixture.package();
    replace_table_fragment(&mut package, r#" count="2">"#, ">");
    let malformed = table_blob(&package);
    assert!(edit(&mut package, "Pivot").is_err());
    assert!(load(&package, "Pivot").is_err());
    assert_eq!(table_blob(&package), malformed);
}

#[test]
fn proven_index_references_are_remapped_and_invalid_removals_are_refused() {
    let fixture = Fixture::build(FixtureOptions {
        pivot_table_outer_exts: pivot_data_extension(&pivot_value_cell("1")),
        ..FixtureOptions::default()
    });
    let mut package = fixture.package();
    let mut transaction = edit(&mut package, "Pivot").unwrap();
    transaction
        .insert_server_format(0, ServerFormat::new(Some("new".to_owned()), None))
        .unwrap();
    let commit = transaction.commit().unwrap();
    assert!(commit.changed());
    let changed = String::from_utf8(table_blob(&package)).unwrap();
    assert!(changed.contains(r#"<x15:x in="2"/>"#));
    assert!(!changed.contains(r#"<x15:x in="1"/>"#));
    assert_eq!(load(&package, "Pivot").unwrap().formats().len(), 3);

    let fixture = Fixture::build(FixtureOptions {
        pivot_table_outer_exts: pivot_data_extension(&pivot_value_cell("1")),
        ..FixtureOptions::default()
    });
    let mut package = fixture.package();
    let mut transaction = edit(&mut package, "Pivot").unwrap();
    transaction.reorder_server_formats(&[1, 0]).unwrap();
    let commit = transaction.commit().unwrap();
    assert!(commit.changed());
    let changed = String::from_utf8(table_blob(&package)).unwrap();
    assert!(changed.contains(r#"<x15:x in="0"/>"#));
    assert!(!changed.contains(r#"<x15:x in="1"/>"#));

    let fixture = Fixture::build(FixtureOptions {
        pivot_table_outer_exts: pivot_data_extension(&pivot_value_cell("1")),
        ..FixtureOptions::default()
    });
    let mut package = fixture.package();
    let before = table_blob(&package);
    let mut transaction = edit(&mut package, "Pivot").unwrap();
    transaction.remove_server_format(1).unwrap();
    assert!(transaction.commit().is_err());
    assert_eq!(table_blob(&package), before);

    let fixture = Fixture::build(FixtureOptions {
        pivot_table_outer_exts: pivot_data_extension(&pivot_value_cell("2")),
        ..FixtureOptions::default()
    });
    let mut package = fixture.package();
    let before = table_blob(&package);
    assert!(edit(&mut package, "Pivot").is_err());
    assert_eq!(table_blob(&package), before);
}

#[test]
fn opaque_namespaced_wrapped_and_mce_index_references_refuse_structural_edits() {
    let namespaced = format!(
        r#"<ext uri="{PIVOT_TABLE_DATA_URI}"><x15:pivotTableData><x15:pivotRow><x15:c><x15:x xmlns:p="{EXT_NS}" p:in="1"/></x15:c></x15:pivotRow></x15:pivotTableData></ext>"#
    );
    let wrapped = format!(
        r#"<ext uri="{PIVOT_TABLE_DATA_URI}"><x15:pivotTableData><x15:wrapper>{}</x15:wrapper></x15:pivotTableData></ext>"#,
        pivot_value_cell("1")
    );
    let mce = format!(
        r#"<ext uri="{PIVOT_TABLE_DATA_URI}"><mc:AlternateContent><mc:Choice Requires="x15"><x15:pivotTableData>{}</x15:pivotTableData></mc:Choice><mc:Fallback><x15:futureFallback/></mc:Fallback></mc:AlternateContent></ext>"#,
        pivot_value_cell("1")
    );

    for (label, opaque_reference) in [
        ("namespaced", namespaced),
        ("wrapped", wrapped),
        ("mce", mce),
    ] {
        let fixture = Fixture::build(FixtureOptions {
            pivot_table_outer_exts: opaque_reference,
            ..FixtureOptions::default()
        });
        let mut package = fixture.package();
        let snapshot = load(&package, "Pivot").unwrap();
        assert!(
            snapshot.has_opaque_index_references(),
            "{label} reference must remain readable as a source diagnostic"
        );
        let before = table_blob(&package);
        let mut transaction = edit(&mut package, "Pivot").unwrap();
        transaction
            .insert_server_format(0, ServerFormat::new(Some("new".to_owned()), None))
            .unwrap();
        assert!(
            transaction.commit().is_err(),
            "{label} reference must refuse structural edits"
        );
        assert_eq!(table_blob(&package), before);
    }
}

#[test]
fn unchanged_reference_indices_retain_whitespace_and_entity_lexical_spelling() {
    for raw_index in [" 01 ", "&#48;1"] {
        let fixture = Fixture::build(FixtureOptions {
            server_count: "3",
            server_formats: concat!(
                r#"<x15:serverFormat culture="first"/>"#,
                r#"<x15:serverFormat culture="referenced"/>"#,
                r#"<x15:serverFormat culture="third"/>"#,
            )
            .to_owned(),
            pivot_table_outer_exts: pivot_data_extension(&pivot_value_cell(raw_index)),
            ..FixtureOptions::default()
        });
        let mut package = fixture.package();
        let source = String::from_utf8(table_blob(&package)).unwrap();
        assert!(source.contains(&format!(r#"in="{raw_index}""#)));

        let mut transaction = edit(&mut package, "Pivot").unwrap();
        transaction
            .insert_server_format(3, ServerFormat::new(Some("appended".to_owned()), None))
            .unwrap();
        transaction.reorder_server_formats(&[0, 1, 3, 2]).unwrap();
        let commit = transaction.commit().unwrap();
        assert!(commit.changed());

        let changed = String::from_utf8(table_blob(&package)).unwrap();
        assert!(
            changed.contains(&format!(r#"in="{raw_index}""#)),
            "structural edits that keep the referenced index must preserve raw lexical in"
        );
        assert_eq!(load(&package, "Pivot").unwrap().formats().len(), 4);
    }
}

#[test]
fn structural_list_aggregate_cap_is_checked_before_publication() {
    let fixture = Fixture::build(FixtureOptions::default());
    let inserted = ServerFormat::new(
        Some("inserted-list-value".to_owned()),
        Some("0.00%".to_owned()),
    );
    let source_total = fixture
        .package()
        .try_iter_parts()
        .map(|part| part.expect("part payload decodes"))
        .map(|part| part.blob().len() as u64)
        .sum::<u64>();

    let mut probe = fixture.package();
    let mut probe_transaction = edit(&mut probe, "Pivot").unwrap();
    probe_transaction
        .insert_server_format(1, inserted.clone())
        .unwrap();
    probe_transaction.commit().unwrap();
    let final_total = probe
        .try_iter_parts()
        .map(|part| part.expect("part payload decodes"))
        .map(|part| part.blob().len() as u64)
        .sum::<u64>();
    assert!(final_total > source_total);

    let exact_limits = ReadLimits::builder()
        .max_total_part_bytes(final_total)
        .unwrap()
        .build()
        .unwrap();
    let mut exact = OpcPackage::from_bytes_with_limits(&fixture.bytes, exact_limits).unwrap();
    let mut exact_transaction = edit(&mut exact, "Pivot").unwrap();
    exact_transaction
        .insert_server_format(1, inserted.clone())
        .unwrap();
    assert!(exact_transaction.commit().is_ok());
    assert_eq!(
        exact
            .try_iter_parts()
            .map(|part| part.expect("part payload decodes"))
            .map(|part| part.blob().len() as u64)
            .sum::<u64>(),
        final_total
    );

    let one_under_limits = ReadLimits::builder()
        .max_total_part_bytes(final_total - 1)
        .unwrap()
        .build()
        .unwrap();
    let mut one_under =
        OpcPackage::from_bytes_with_limits(&fixture.bytes, one_under_limits).unwrap();
    let before = table_blob(&one_under);
    let mut one_under_transaction = edit(&mut one_under, "Pivot").unwrap();
    one_under_transaction
        .insert_server_format(1, inserted)
        .unwrap();
    assert!(one_under_transaction.commit().is_err());
    assert_eq!(table_blob(&one_under), before);
}

#[test]
fn stale_signed_and_limited_publications_refuse_before_mutation() {
    let fixture = Fixture::build(FixtureOptions::default());
    let mut source = fixture.package();
    let mut transaction = edit(&mut source, "Pivot").unwrap();
    transaction
        .set_culture(0, Some("de-DE".to_owned()))
        .unwrap();
    let commit = transaction.commit().unwrap();

    let mut stale = fixture.package();
    let stale_xml = table_blob(&stale);
    let stale_rewritten = String::from_utf8(stale_xml.clone())
        .unwrap()
        .replace("#,##0.00", "#,##0.0")
        .into_bytes();
    stale
        .get_part_mut(&PackURI::new(TABLE_URI).unwrap())
        .unwrap()
        .set_blob(stale_rewritten);
    let stale_before = table_blob(&stale);
    assert!(commit.patch().apply(&mut stale).is_err());
    assert_eq!(table_blob(&stale), stale_before);

    let mut stale_connections = fixture.package();
    let connections_uri = PackURI::new(CONNECTIONS_URI).unwrap();
    let connections = stale_connections.get_part_mut(&connections_uri).unwrap();
    let stale_connection_xml = std::str::from_utf8(connections.blob()).unwrap().replace(
        r#"name="canonical""#,
        r#"name="canonical" description="stale-source""#,
    );
    connections.set_blob(stale_connection_xml.into_bytes());
    let stale_connections_before = part_blob(&stale_connections, CONNECTIONS_URI);
    assert!(commit.patch().apply(&mut stale_connections).is_err());
    assert_eq!(
        part_blob(&stale_connections, CONNECTIONS_URI),
        stale_connections_before
    );

    let signed_fixture = Fixture::build(FixtureOptions {
        signed: true,
        ..FixtureOptions::default()
    });
    let mut signed = signed_fixture.package();
    let signed_before = table_blob(&signed);
    let mut signed_transaction = edit(&mut signed, "Pivot").unwrap();
    signed_transaction
        .set_culture(0, Some("de-DE".to_owned()))
        .unwrap();
    assert!(signed_transaction.commit().is_err());
    assert_eq!(table_blob(&signed), signed_before);

    let limited_fixture = Fixture::build(FixtureOptions::default());
    let source_package = limited_fixture.package();
    let max_source = source_package
        .try_iter_parts()
        .map(|part| part.expect("part payload decodes"))
        .map(|part| part.blob().len())
        .max()
        .unwrap();
    let limits = ReadLimits::builder()
        .max_part_bytes(max_source.saturating_mul(16) as u64)
        .unwrap()
        .build()
        .unwrap();
    let mut limited = OpcPackage::from_bytes_with_limits(&limited_fixture.bytes, limits).unwrap();
    let limited_before = table_blob(&limited);
    let mut limited_transaction = edit(&mut limited, "Pivot").unwrap();
    limited_transaction
        .set_format(1, Some("z".repeat(max_source.saturating_mul(32))))
        .unwrap();
    assert!(limited_transaction.commit().is_err());
    assert_eq!(table_blob(&limited), limited_before);
}

#[test]
fn ordinary_patch_rejects_stale_workbook_root_namespace_and_mce_context() {
    let fixture = Fixture::build(FixtureOptions::default());
    let workbook = Workbook::from_bytes(fixture.bytes.clone()).unwrap();
    let mut transaction = workbook.edit_pivot_table("Pivot").unwrap();
    transaction
        .set_culture(0, Some("de-DE".to_owned()))
        .unwrap();
    let patch = transaction.commit().unwrap().patch().clone();

    let mut stale_root_package = fixture.package();
    let workbook_uri = PackURI::new(WORKBOOK_URI).unwrap();
    let workbook_part = stale_root_package.get_part_mut(&workbook_uri).unwrap();
    let stale_root = std::str::from_utf8(workbook_part.blob()).unwrap().replace(
        &format!(r#"xmlns="{TRANSITIONAL_MAIN}""#),
        &format!(r#"xmlns="{STRICT_MAIN}""#),
    );
    assert_ne!(stale_root, fixture.workbook_xml);
    workbook_part.set_blob(stale_root.into_bytes());
    let stale_root =
        Workbook::from_bytes(PackageWriter::to_bytes(&stale_root_package).unwrap()).unwrap();
    let stale_root_before = stale_root
        .pivot_table_server_formats_source("Pivot")
        .unwrap()
        .source_xml()
        .to_vec();
    assert!(patch.apply(&stale_root).is_err());
    assert_eq!(
        stale_root
            .pivot_table_server_formats_source("Pivot")
            .unwrap()
            .source_xml(),
        stale_root_before.as_slice()
    );

    let mut stale_mce_package = fixture.package();
    let workbook_part = stale_mce_package.get_part_mut(&workbook_uri).unwrap();
    let stale_mce = std::str::from_utf8(workbook_part.blob()).unwrap().replace(
        &format!(r#"xmlns:mc="{MCE_NS}" mc:Ignorable="x15""#),
        &format!(r#"xmlns:x14="{CACHE_SOURCE_EXT_NS}" xmlns:mc="{MCE_NS}" mc:Ignorable="x15 x14""#),
    );
    assert_ne!(stale_mce, fixture.workbook_xml);
    workbook_part.set_blob(stale_mce.into_bytes());
    let stale_mce =
        Workbook::from_bytes(PackageWriter::to_bytes(&stale_mce_package).unwrap()).unwrap();
    let stale_mce_before = stale_mce
        .pivot_table_server_formats_source("Pivot")
        .unwrap()
        .source_xml()
        .to_vec();
    assert!(patch.apply(&stale_mce).is_err());
    assert_eq!(
        stale_mce
            .pivot_table_server_formats_source("Pivot")
            .unwrap()
            .source_xml(),
        stale_mce_before.as_slice()
    );
}

#[test]
fn duplicate_known_uri_unknown_uri_and_wrong_qname_are_rejected_or_inert() {
    let duplicate = Fixture::build(FixtureOptions {
        pivot_table_exts: r#"<x15:pivotTableServerFormats count="1"><x15:serverFormat/></x15:pivotTableServerFormats>"#
            .to_owned(),
        ..FixtureOptions::default()
    });
    assert!(load(&duplicate.package(), "Pivot").is_err());

    let unknown = Fixture::build(FixtureOptions {
        server_uri: "{UNKNOWN-SERVER-FORMAT-URI}",
        ..FixtureOptions::default()
    });
    assert!(load(&unknown.package(), "Pivot").is_err());

    let wrong_qname = Fixture::build(FixtureOptions {
        server_formats:
            r#"<serverFormat xmlns="urn:other" culture="en-US"/> <serverFormat xmlns="urn:other"/>"#
                .replace("> <", "><"),
        ..FixtureOptions::default()
    });
    assert!(load(&wrong_qname.package(), "Pivot").is_err());

    let nested_impostor = Fixture::build(FixtureOptions {
        server_uri: "{UNKNOWN-SERVER-FORMAT-URI}",
        pivot_table_exts: r#"<x15:wrapper><x15:pivotTableServerFormats count="1"><x15:serverFormat/></x15:pivotTableServerFormats></x15:wrapper>"#
            .to_owned(),
        ..FixtureOptions::default()
    });
    assert!(load(&nested_impostor.package(), "Pivot").is_err());
}

#[test]
fn count_child_and_pivot_value_index_bounds_are_checked_before_typed_use() {
    for (count, server_formats) in [
        ("0", r#"<x15:serverFormat/> <x15:serverFormat/>"#),
        ("1", r#"<x15:serverFormat/> <x15:serverFormat/>"#),
        ("3", r#"<x15:serverFormat/> <x15:serverFormat/>"#),
    ] {
        let fixture = Fixture::build(FixtureOptions {
            server_count: count,
            server_formats: server_formats.replace("> <", "><"),
            ..FixtureOptions::default()
        });
        assert!(
            load(&fixture.package(), "Pivot").is_err(),
            "count {count} must match the two retained children"
        );
    }

    let in_range = Fixture::build(FixtureOptions {
        pivot_table_outer_exts: pivot_data_extension(&pivot_value_cell("1")),
        ..FixtureOptions::default()
    });
    let snapshot = load(&in_range.package(), "Pivot").unwrap();
    assert!(!snapshot.has_diagnostic_index_boundary());

    let boundary = Fixture::build(FixtureOptions {
        pivot_table_outer_exts: pivot_data_extension(&pivot_value_cell("2")),
        ..FixtureOptions::default()
    });
    let snapshot = load(&boundary.package(), "Pivot").unwrap();
    assert!(snapshot.has_diagnostic_index_boundary());
    let mut boundary_package = boundary.package();
    assert!(edit(&mut boundary_package, "Pivot").is_err());

    let out_of_range = Fixture::build(FixtureOptions {
        pivot_table_outer_exts: pivot_data_extension(&pivot_value_cell("3")),
        ..FixtureOptions::default()
    });
    assert!(load(&out_of_range.package(), "Pivot").is_err());
}

#[test]
fn diagnostic_index_requires_exact_pivot_row_c_x_owner_chain() {
    let wrong_ancestry = Fixture::build(FixtureOptions {
        pivot_table_outer_exts: format!(
            r#"<ext uri="{PIVOT_TABLE_DATA_URI}"><x15:wrapper>{}</x15:wrapper></ext>"#,
            pivot_value_cell("2")
        ),
        ..FixtureOptions::default()
    });
    let wrong_ancestry_snapshot = load(&wrong_ancestry.package(), "Pivot").unwrap();
    assert!(!wrong_ancestry_snapshot.has_diagnostic_index_boundary());

    let wrong_uri = Fixture::build(FixtureOptions {
        pivot_table_outer_exts: format!(
            r#"<ext uri="{{not-pivot-data}}"><x15:pivotTableData>{}</x15:pivotTableData></ext>"#,
            pivot_value_cell("2")
        ),
        ..FixtureOptions::default()
    });
    let wrong_uri_snapshot = load(&wrong_uri.package(), "Pivot").unwrap();
    assert!(!wrong_uri_snapshot.has_diagnostic_index_boundary());

    let lookalike_qnames = Fixture::build(FixtureOptions {
        pivot_table_outer_exts: pivot_data_extension(
            r#"<x15:pivotRows><x15:cell><x15:xx in="2"/></x15:cell></x15:pivotRows>"#,
        ),
        ..FixtureOptions::default()
    });
    let lookalike_snapshot = load(&lookalike_qnames.package(), "Pivot").unwrap();
    assert!(!lookalike_snapshot.has_diagnostic_index_boundary());
}

#[test]
fn nonworksheet_eligibility_and_cache_closure_are_enforced() {
    let incoming = Fixture::build(FixtureOptions {
        worksheet_incoming: true,
        ..FixtureOptions::default()
    });
    assert!(load(&incoming.package(), "Pivot").is_err());

    assert!(
        load(
            &Fixture::build(FixtureOptions {
                table_attrs: r#" enableEdit="true""#.to_owned(),
                ..FixtureOptions::default()
            })
            .package(),
            "Pivot"
        )
        .is_err()
    );
    for table_children in [
        r#"<pivotEdits/>"#.to_owned(),
        r#"<pivotChanges/>"#.to_owned(),
        r#"<conditionalFormats/>"#.to_owned(),
    ] {
        assert!(
            load(
                &Fixture::build(FixtureOptions {
                    table_children,
                    ..FixtureOptions::default()
                })
                .package(),
                "Pivot"
            )
            .is_err()
        );
    }
    assert!(
        load(
            &Fixture::build(FixtureOptions {
                location_ref: "B2:C5",
                ..FixtureOptions::default()
            })
            .package(),
            "Pivot"
        )
        .is_err()
    );
    let duplicate_name = Fixture::build(FixtureOptions {
        worksheet_duplicate_name: true,
        ..FixtureOptions::default()
    });
    assert!(load(&duplicate_name.package(), "Pivot").is_err());

    let missing_cache_version = Fixture::build(FixtureOptions {
        cache_exts: String::new(),
        ..FixtureOptions::default()
    });
    // The required extension is part of the base cache XML; remove it from
    // the generated source so this is a genuine closure failure.
    let mut package = missing_cache_version.package();
    let cache_uri = PackURI::new(CACHE_URI).unwrap();
    let cache = package.get_part_mut(&cache_uri).unwrap();
    let without_version = std::str::from_utf8(cache.blob())
        .unwrap()
        .replace(
            &format!(
                r#"<extLst><ext uri="{CACHE_ID_VERSION_URI}"><x15:pivotCacheIdVersion cacheIdSupportedVersion="15" cacheIdCreatedVersion="15"/></ext></extLst>"#
            ),
            "",
    );
    cache.set_blob(without_version.into_bytes());
    assert!(load(&package, "Pivot").is_err());

    let wrong_source = Fixture::build(FixtureOptions::default());
    let mut package = wrong_source.package();
    let cache = package.get_part_mut(&cache_uri).unwrap();
    let wrong_type = std::str::from_utf8(cache.blob())
        .unwrap()
        .replace(r#"type="external""#, r#"type="worksheet""#);
    cache.set_blob(wrong_type.into_bytes());
    assert!(load(&package, "Pivot").is_err());

    let duplicate_cache_version = Fixture::build(FixtureOptions {
        cache_exts: format!(
            r#"<extLst><ext uri="{CACHE_ID_VERSION_URI}"><x15:pivotCacheIdVersion cacheIdSupportedVersion="15" cacheIdCreatedVersion="15"/></ext></extLst>"#
        ),
        ..FixtureOptions::default()
    });
    assert!(load(&duplicate_cache_version.package(), "Pivot").is_err());

    let oversized_cache_version = Fixture::build(FixtureOptions::default());
    let mut package = oversized_cache_version.package();
    let cache = package.get_part_mut(&cache_uri).unwrap();
    let oversized = std::str::from_utf8(cache.blob()).unwrap().replace(
        r#"cacheIdSupportedVersion="15""#,
        r#"cacheIdSupportedVersion="256""#,
    );
    cache.set_blob(oversized.into_bytes());
    assert!(load(&package, "Pivot").is_err());

    let missing_cache_version_attribute = Fixture::build(FixtureOptions::default());
    let mut package = missing_cache_version_attribute.package();
    let cache = package.get_part_mut(&cache_uri).unwrap();
    let missing_attribute = std::str::from_utf8(cache.blob())
        .unwrap()
        .replace(r#" cacheIdCreatedVersion="15""#, "");
    cache.set_blob(missing_attribute.into_bytes());
    assert!(load(&package, "Pivot").is_err());

    let explicit_connection_id = Fixture::build(FixtureOptions::default());
    let mut package = explicit_connection_id.package();
    let cache = package.get_part_mut(&cache_uri).unwrap();
    let explicit_id = std::str::from_utf8(cache.blob()).unwrap().replace(
        r#"<cacheSource type="external""#,
        r#"<cacheSource type="external" connectionId="7""#,
    );
    cache.set_blob(explicit_id.into_bytes());
    assert!(load(&package, "Pivot").is_ok());

    let mismatched_connection_id = Fixture::build(FixtureOptions::default());
    let mut package = mismatched_connection_id.package();
    let cache = package.get_part_mut(&cache_uri).unwrap();
    let mismatched_id = std::str::from_utf8(cache.blob()).unwrap().replace(
        r#"<cacheSource type="external""#,
        r#"<cacheSource type="external" connectionId="8""#,
    );
    cache.set_blob(mismatched_id.into_bytes());
    assert!(load(&package, "Pivot").is_err());

    let mismatched_cache_id = Fixture::build(FixtureOptions::default());
    let mut package = mismatched_cache_id.package();
    let table_uri = PackURI::new(TABLE_URI).unwrap();
    let table = package.get_part_mut(&table_uri).unwrap();
    let mismatched =
        std::str::from_utf8(table.blob())
            .unwrap()
            .replacen(r#"cacheId="7""#, r#"cacheId="8""#, 1);
    table.set_blob(mismatched.into_bytes());
    assert!(load(&package, "Pivot").is_err());

    let mismatched_optional_cache_id = Fixture::build(FixtureOptions::default());
    let mut package = mismatched_optional_cache_id.package();
    let cache = package.get_part_mut(&cache_uri).unwrap();
    let mismatched_optional = std::str::from_utf8(cache.blob()).unwrap().replace(
        "<pivotCacheDefinition ",
        r#"<pivotCacheDefinition pivotCacheId="8" "#,
    );
    cache.set_blob(mismatched_optional.into_bytes());
    assert!(load(&package, "Pivot").is_err());

    let core_optional_cache_id = Fixture::build(FixtureOptions::default());
    let mut package = core_optional_cache_id.package();
    let cache = package.get_part_mut(&cache_uri).unwrap();
    let core_optional = std::str::from_utf8(cache.blob()).unwrap().replace(
        "<pivotCacheDefinition ",
        r#"<pivotCacheDefinition pivotCacheId="7" "#,
    );
    cache.set_blob(core_optional.into_bytes());
    assert!(load(&package, "Pivot").is_err());

    let invalid_core_cache_id = Fixture::build(FixtureOptions::default());
    let mut package = invalid_core_cache_id.package();
    let cache = package.get_part_mut(&cache_uri).unwrap();
    let invalid_core = std::str::from_utf8(cache.blob()).unwrap().replace(
        "<pivotCacheDefinition ",
        r#"<pivotCacheDefinition cacheId="7" "#,
    );
    cache.set_blob(invalid_core.into_bytes());
    assert!(load(&package, "Pivot").is_err());

    let matching_extension_cache_id = Fixture::build(FixtureOptions::default());
    let mut package = matching_extension_cache_id.package();
    let cache = package.get_part_mut(&cache_uri).unwrap();
    let matching_extension = std::str::from_utf8(cache.blob())
        .unwrap()
        .replace(
            &format!(r#"<extLst><ext uri="{CACHE_ID_VERSION_URI}">"#),
            &format!(
                r#"<extLst><ext uri="{PIVOT_CACHE_DEFINITION_URI}"><x14:pivotCacheDefinition pivotCacheId="7"/></ext><ext uri="{CACHE_ID_VERSION_URI}">"#
            ),
        );
    cache.set_blob(matching_extension.into_bytes());
    assert!(load(&package, "Pivot").is_ok());

    let mismatched_extension_cache_id = Fixture::build(FixtureOptions::default());
    let mut package = mismatched_extension_cache_id.package();
    let cache = package.get_part_mut(&cache_uri).unwrap();
    let mismatched_extension = std::str::from_utf8(cache.blob())
        .unwrap()
        .replace(
            &format!(r#"<extLst><ext uri="{CACHE_ID_VERSION_URI}">"#),
            &format!(
                r#"<extLst><ext uri="{PIVOT_CACHE_DEFINITION_URI}"><x14:pivotCacheDefinition pivotCacheId="8"/></ext><ext uri="{CACHE_ID_VERSION_URI}">"#
            ),
        );
    cache.set_blob(mismatched_extension.into_bytes());
    assert!(load(&package, "Pivot").is_err());

    let absent_source_connection_extension = Fixture::build(FixtureOptions::default());
    let mut package = absent_source_connection_extension.package();
    let cache = package.get_part_mut(&cache_uri).unwrap();
    let without_source = std::str::from_utf8(cache.blob())
        .unwrap()
        .replace(
            &format!(
                r#"<extLst><ext uri="{CACHE_SOURCE_URI}"><x14:sourceConnection name="canonical"/></ext></extLst>"#
            ),
            "",
        );
    cache.set_blob(without_source.into_bytes());
    let workbook_uri = PackURI::new(WORKBOOK_URI).unwrap();
    package
        .get_part_mut(&workbook_uri)
        .unwrap()
        .rels_mut()
        .remove("rIdConnections")
        .unwrap();
    assert!(package.remove_part(&PackURI::new(CONNECTIONS_URI).unwrap()));
    assert!(load(&package, "Pivot").is_ok());

    let numeric_only_connection_id = Fixture::build(FixtureOptions::default());
    let mut package = numeric_only_connection_id.package();
    let cache = package.get_part_mut(&cache_uri).unwrap();
    let numeric_only = std::str::from_utf8(cache.blob())
        .unwrap()
        .replace(
            &format!(
                r#"<extLst><ext uri="{CACHE_SOURCE_URI}"><x14:sourceConnection name="canonical"/></ext></extLst>"#
            ),
            "",
        )
        .replace(
            r#"<cacheSource type="external""#,
            r#"<cacheSource type="external" connectionId="7""#,
        );
    cache.set_blob(numeric_only.into_bytes());
    assert!(load(&package, "Pivot").is_ok());

    let mismatched_numeric_only_connection_id = Fixture::build(FixtureOptions::default());
    let mut package = mismatched_numeric_only_connection_id.package();
    let cache = package.get_part_mut(&cache_uri).unwrap();
    let mismatched_numeric_only = std::str::from_utf8(cache.blob())
        .unwrap()
        .replace(
            &format!(
                r#"<extLst><ext uri="{CACHE_SOURCE_URI}"><x14:sourceConnection name="canonical"/></ext></extLst>"#
            ),
            "",
        )
        .replace(
            r#"<cacheSource type="external""#,
            r#"<cacheSource type="external" connectionId="8""#,
        );
    cache.set_blob(mismatched_numeric_only.into_bytes());
    assert!(load(&package, "Pivot").is_err());

    let missing_source_connection = Fixture::build(FixtureOptions::default());
    let mut package = missing_source_connection.package();
    let cache = package.get_part_mut(&cache_uri).unwrap();
    let without_source_name = std::str::from_utf8(cache.blob())
        .unwrap()
        .replace(r#"<x14:sourceConnection name="canonical"/>"#, "");
    cache.set_blob(without_source_name.into_bytes());
    assert!(load(&package, "Pivot").is_err());

    let wrong_source_connection = Fixture::build(FixtureOptions::default());
    let mut package = wrong_source_connection.package();
    let cache = package.get_part_mut(&cache_uri).unwrap();
    let wrong_name = std::str::from_utf8(cache.blob())
        .unwrap()
        .replace(r#"name="canonical""#, r#"name="missing""#);
    cache.set_blob(wrong_name.into_bytes());
    assert!(load(&package, "Pivot").is_err());

    let duplicate_source_connection = Fixture::build(FixtureOptions::default());
    let mut package = duplicate_source_connection.package();
    let cache = package.get_part_mut(&cache_uri).unwrap();
    let duplicate_name = std::str::from_utf8(cache.blob()).unwrap().replace(
        r#"<x14:sourceConnection name="canonical"/>"#,
        r#"<x14:sourceConnection name="canonical"/><x14:sourceConnection name="canonical"/>"#,
    );
    cache.set_blob(duplicate_name.into_bytes());
    assert!(load(&package, "Pivot").is_err());
}

#[test]
fn caller_limits_reject_before_server_format_collection_allocation() {
    let fixture = Fixture::build(FixtureOptions {
        server_count: "4294967295",
        ..FixtureOptions::default()
    });
    let package = fixture.package();
    assert!(load(&package, "Pivot").is_err());

    let tiny = ReadLimits::builder()
        .max_part_bytes(1)
        .unwrap()
        .build()
        .unwrap();
    assert!(OpcPackage::from_bytes_with_limits(&fixture.bytes, tiny).is_err());
}

#[test]
fn exact_source_aggregate_cap_admits_equal_size_edit_but_rejects_growth() {
    let fixture = Fixture::build(FixtureOptions {
        server_formats:
            r##"<x15:serverFormat culture="en-US" format="#,##0.00"/><x15:serverFormat culture="fr-FR" format="&#34;&#10;"/>"##
                .to_owned(),
        ..FixtureOptions::default()
    });
    let source_total = fixture
        .package()
        .try_iter_parts()
        .map(|part| part.expect("part payload decodes"))
        .map(|part| part.blob().len() as u64)
        .sum::<u64>();

    // Probe the exact final geometry without a restrictive aggregate cap. A
    // one-under cap still admits the source bytes, so rejection demonstrates
    // candidate preflight rather than package ingress rejection.
    let planned_format = format!("A&B<C\"D'_x0001_\u{1}{}", "x".repeat(2048));
    let mut probe = fixture.package();
    let mut probe_edit = edit(&mut probe, "Pivot").unwrap();
    probe_edit
        .set_format(0, Some(planned_format.clone()))
        .unwrap();
    probe_edit.commit().unwrap();
    let probe_xml = String::from_utf8(table_blob(&probe)).unwrap();
    assert!(probe_xml.contains("&amp;"));
    assert!(probe_xml.contains("&lt;"));
    assert!(probe_xml.contains("&quot;"));
    assert!(probe_xml.contains("&apos;"));
    assert!(probe_xml.contains("_x005F_x0001_"));
    assert!(probe_xml.contains("_x0001_"));
    assert!(probe_xml.contains(r#"format="&#34;&#10;""#));
    let exact_final_total = probe
        .try_iter_parts()
        .map(|part| part.expect("part payload decodes"))
        .map(|part| part.blob().len() as u64)
        .sum::<u64>();
    assert!(exact_final_total > source_total + 1);

    let exact_limits = ReadLimits::builder()
        .max_total_part_bytes(exact_final_total)
        .unwrap()
        .build()
        .unwrap();
    let mut exact = OpcPackage::from_bytes_with_limits(&fixture.bytes, exact_limits).unwrap();
    let mut exact_edit = edit(&mut exact, "Pivot").unwrap();
    exact_edit
        .set_format(0, Some(planned_format.clone()))
        .unwrap();
    assert!(exact_edit.commit().is_ok());
    assert_eq!(
        exact
            .try_iter_parts()
            .map(|part| part.expect("part payload decodes"))
            .map(|part| part.blob().len() as u64)
            .sum::<u64>(),
        exact_final_total
    );
    assert_eq!(table_blob(&exact), table_blob(&probe));
    assert!(
        String::from_utf8(table_blob(&exact))
            .unwrap()
            .contains(r#"format="&#34;&#10;""#)
    );

    let one_under_limits = ReadLimits::builder()
        .max_total_part_bytes(exact_final_total - 1)
        .unwrap()
        .build()
        .unwrap();
    let mut one_under =
        OpcPackage::from_bytes_with_limits(&fixture.bytes, one_under_limits).unwrap();
    let before = table_blob(&one_under);
    let mut one_under_edit = edit(&mut one_under, "Pivot").unwrap();
    one_under_edit.set_format(0, Some(planned_format)).unwrap();
    assert!(one_under_edit.commit().is_err());
    assert_eq!(table_blob(&one_under), before);
}

#[test]
fn caller_encoded_attribute_limit_admits_exact_output_and_rejects_one_under() {
    let fixture = Fixture::build(FixtureOptions {
        server_formats:
            r##"<x15:serverFormat culture="en-US" format="#,##0.00"/><x15:serverFormat culture="fr-FR" format="&#34;&#10;"/>"##
                .to_owned(),
        ..FixtureOptions::default()
    });
    let planned_format = format!("A&B<C\"D'_x0001_\u{1}{}", "z".repeat(256));

    let mut generous = fixture.package();
    let mut generous_edit = edit(&mut generous, "Pivot").unwrap();
    generous_edit
        .set_format(0, Some(planned_format.clone()))
        .unwrap();
    generous_edit.commit().unwrap();
    let generous_xml = String::from_utf8(table_blob(&generous)).unwrap();
    let marker = r#"culture="en-US" format=""#;
    let value_start = generous_xml.find(marker).unwrap() + marker.len();
    let value_end = value_start + generous_xml[value_start..].find('"').unwrap();
    let encoded_value = &generous_xml[value_start..value_end];
    let encoded_attribute_bytes = b"format".len() + encoded_value.len();
    assert!(encoded_value.contains("&amp;"));
    assert!(encoded_value.contains("&quot;"));
    assert!(encoded_value.contains("_x005F_x0001_"));
    assert!(generous_xml.contains(r#"format="&#34;&#10;""#));

    let exact_limits = ReadLimits::builder()
        .max_xml_attribute_bytes(encoded_attribute_bytes)
        .unwrap()
        .max_relationship_target_bytes(encoded_attribute_bytes)
        .unwrap()
        .build()
        .unwrap();
    let mut exact = OpcPackage::from_bytes_with_limits(&fixture.bytes, exact_limits).unwrap();
    let mut exact_edit = edit(&mut exact, "Pivot").unwrap();
    exact_edit
        .set_format(0, Some(planned_format.clone()))
        .unwrap();
    exact_edit.commit().unwrap();
    assert_eq!(table_blob(&exact), table_blob(&generous));

    let one_under_limits = ReadLimits::builder()
        .max_xml_attribute_bytes(encoded_attribute_bytes - 1)
        .unwrap()
        .max_relationship_target_bytes(encoded_attribute_bytes - 1)
        .unwrap()
        .build()
        .unwrap();
    let mut one_under =
        OpcPackage::from_bytes_with_limits(&fixture.bytes, one_under_limits).unwrap();
    let before = table_blob(&one_under);
    let mut one_under_edit = edit(&mut one_under, "Pivot").unwrap();
    one_under_edit.set_format(0, Some(planned_format)).unwrap();
    assert!(one_under_edit.commit().is_err());
    assert_eq!(table_blob(&one_under), before);
}

#[test]
fn retained_xml_event_depth_and_attribute_limits_are_honored_by_owner_scan() {
    let event_fixture = Fixture::build(FixtureOptions {
        table_children: "<x15:opaque/>".repeat(64),
        ..FixtureOptions::default()
    });
    let event_limits = ReadLimits::builder()
        .max_xml_events(64)
        .unwrap()
        .build()
        .unwrap();
    let event_package =
        OpcPackage::from_bytes_with_limits(&event_fixture.bytes, event_limits).unwrap();
    assert!(load(&event_package, "Pivot").is_err());

    let mut nested = String::new();
    for _ in 0..8 {
        nested.push_str("<x15:opaque>");
    }
    for _ in 0..8 {
        nested.push_str("</x15:opaque>");
    }
    let depth_fixture = Fixture::build(FixtureOptions {
        table_children: nested,
        ..FixtureOptions::default()
    });
    let depth_limits = ReadLimits::builder()
        .max_xml_depth(4)
        .unwrap()
        .build()
        .unwrap();
    let depth_package =
        OpcPackage::from_bytes_with_limits(&depth_fixture.bytes, depth_limits).unwrap();
    assert!(load(&depth_package, "Pivot").is_err());

    let attribute_fixture = Fixture::build(FixtureOptions {
        table_attrs: format!(r#" opaque="{}""#, "x".repeat(1024)),
        ..FixtureOptions::default()
    });
    let attribute_limits = ReadLimits::builder()
        .max_xml_attribute_bytes(512)
        .unwrap()
        .max_relationship_target_bytes(512)
        .unwrap()
        .build()
        .unwrap();
    let attribute_package =
        OpcPackage::from_bytes_with_limits(&attribute_fixture.bytes, attribute_limits).unwrap();
    assert!(load(&attribute_package, "Pivot").is_err());

    let escaped_attribute_fixture = Fixture::build(FixtureOptions {
        table_attrs: format!(r#" opaque="{}""#, "&#34;".repeat(110)),
        ..FixtureOptions::default()
    });
    let escaped_attribute_limits = ReadLimits::builder()
        .max_xml_attribute_bytes(512)
        .unwrap()
        .max_relationship_target_bytes(512)
        .unwrap()
        .build()
        .unwrap();
    let escaped_attribute_package = OpcPackage::from_bytes_with_limits(
        &escaped_attribute_fixture.bytes,
        escaped_attribute_limits,
    )
    .unwrap();
    // The decoded value is only 110 bytes, but its raw lexical attribute is
    // over the caller limit.  The owner scan must reject before unescaping or
    // collecting the server-format values.
    assert!(load(&escaped_attribute_package, "Pivot").is_err());
}

#[test]
fn removed_signature_graph_still_requires_explicit_edit_policy() {
    let fixture = Fixture::build(FixtureOptions {
        signed: true,
        ..FixtureOptions::default()
    });
    let mut package = fixture.package();
    let signature_id = package
        .rels()
        .iter()
        .find(|relationship| relationship.reltype() == rt::DIGITAL_SIGNATURE_ORIGIN)
        .map(|relationship| relationship.r_id().to_owned())
        .unwrap();
    package.rels_mut().remove(&signature_id).unwrap();
    assert!(package.remove_part(&PackURI::new("/_xmlsignatures/origin.sigs").unwrap()));
    assert!(!package.is_signed());
    assert!(package.requires_signature_edit_policy());

    let before = table_blob(&package);
    let mut transaction = edit(&mut package, "Pivot").unwrap();
    transaction
        .set_culture(0, Some("de-DE".to_owned()))
        .unwrap();
    assert!(transaction.commit().is_err());
    assert_eq!(table_blob(&package), before);
}

#[test]
fn greater_than_65535_server_format_text_is_allowed_within_caller_limits() {
    let long_value = "x".repeat(65_536);
    let fixture = Fixture::build(FixtureOptions {
        server_formats: format!(r#"<x15:serverFormat culture="{long_value}"/><x15:serverFormat/>"#),
        ..FixtureOptions::default()
    });
    let limits = ReadLimits::builder()
        .max_xml_attribute_bytes(128 * 1024)
        .unwrap()
        .build()
        .unwrap();
    let package = OpcPackage::from_bytes_with_limits(&fixture.bytes, limits).unwrap();
    let snapshot = load(&package, "Pivot").unwrap();
    assert_eq!(
        snapshot.formats()[0].culture.as_ref().unwrap().len(),
        65_536
    );
}

#[test]
fn unpaired_surrogate_is_rejected_but_escaped_control_and_literal_are_retained() {
    let surrogate = Fixture::build(FixtureOptions {
        server_formats: r#"<x15:serverFormat culture="_xD800_"/><x15:serverFormat/>"#.to_owned(),
        ..FixtureOptions::default()
    });
    assert!(load(&surrogate.package(), "Pivot").is_err());

    let control = Fixture::build(FixtureOptions {
        server_formats:
            r#"<x15:serverFormat culture="_x0001_"/><x15:serverFormat culture="_x005F_x0041_"/>"#
                .to_owned(),
        ..FixtureOptions::default()
    });
    let snapshot = load(&control.package(), "Pivot").unwrap();
    assert_eq!(snapshot.formats()[0].culture.as_deref(), Some("\u{1}"));
    assert_eq!(snapshot.formats()[1].culture.as_deref(), Some("_x0041_"));
}

#[test]
fn numeric_xml_controls_are_rejected_in_typed_and_opaque_attributes() {
    let typed_numeric_control = Fixture::build(FixtureOptions {
        server_formats: r#"<x15:serverFormat culture="&#x1;"/><x15:serverFormat/>"#.to_owned(),
        ..FixtureOptions::default()
    });
    assert!(load(&typed_numeric_control.package(), "Pivot").is_err());

    let opaque_numeric_control = Fixture::build(FixtureOptions {
        table_attrs: r#" foreign="&#x1;""#.to_owned(),
        ..FixtureOptions::default()
    });
    assert!(load(&opaque_numeric_control.package(), "Pivot").is_err());

    let typed_spreadsheet_escape = Fixture::build(FixtureOptions {
        server_formats: r#"<x15:serverFormat culture="_x0001_"/><x15:serverFormat/>"#.to_owned(),
        ..FixtureOptions::default()
    });
    let snapshot = load(&typed_spreadsheet_escape.package(), "Pivot").unwrap();
    assert_eq!(snapshot.formats()[0].culture.as_deref(), Some("\u{1}"));

    let opaque_spreadsheet_escape = Fixture::build(FixtureOptions {
        table_attrs: r#" foreign="_x0001_""#.to_owned(),
        ..FixtureOptions::default()
    });
    assert!(load(&opaque_spreadsheet_escape.package(), "Pivot").is_ok());
}

#[test]
fn changed_writes_escape_decoded_controls_and_literal_xstring_markers() {
    let fixture = Fixture::build(FixtureOptions::default());
    let mut package = fixture.package();
    let mut transaction = edit(&mut package, "Pivot").unwrap();
    assert!(
        transaction
            .set_culture(0, Some("\u{1}".to_owned()))
            .unwrap()
    );
    assert!(
        transaction
            .set_format(1, Some("_x0041_".to_owned()))
            .unwrap()
    );
    let commit = transaction.commit().unwrap();
    assert!(commit.changed());
    let changed = String::from_utf8(table_blob(&package)).unwrap();
    assert!(changed.contains(r#"culture="_x0001_""#));
    assert!(changed.contains(r#"format="_x005F_x0041_""#));
    let snapshot = load(&package, "Pivot").unwrap();
    assert_eq!(snapshot.formats()[0].culture.as_deref(), Some("\u{1}"));
    assert_eq!(snapshot.formats()[1].format.as_deref(), Some("_x0041_"));
}

#[test]
fn retained_relationship_limits_account_for_mutated_root_and_unknown_edges() {
    let fixture = Fixture::build(FixtureOptions::default());
    let base = fixture.package();
    let base_inventory = relationship_inventory(&base);
    let mut final_package = fixture.package();
    add_unknown_relationships(&mut final_package);
    let final_inventory = relationship_inventory(&final_package);

    // The sheet mutation creates a new relationship member.  The two new
    // internal edges resolve to distinct missing targets, so both are charged
    // as graph nodes even though neither target is an OPC part.
    assert_eq!(base_inventory.parts + 1, final_inventory.parts);
    assert_eq!(base_inventory.total_edges + 2, final_inventory.total_edges);
    assert_eq!(base_inventory.graph_nodes + 2, final_inventory.graph_nodes);
    assert!(final_inventory.max_edges >= base_inventory.max_edges);
    assert!(final_inventory.max_xml_events >= base_inventory.max_xml_events);
    assert!(base_inventory.total_xml_bytes < final_inventory.total_xml_bytes - 1);
    assert!(base_inventory.total_xml_events < final_inventory.total_xml_events - 1);
    assert!(base_inventory.max_target_bytes < final_inventory.max_target_bytes - 1);

    let cases = [
        (
            ReadResource::RelationshipParts,
            ReadLimits::builder()
                .max_relationship_parts(final_inventory.parts)
                .unwrap(),
            ReadLimits::builder()
                .max_relationship_parts(final_inventory.parts - 1)
                .unwrap(),
        ),
        (
            ReadResource::TotalRelationshipXmlBytes,
            ReadLimits::builder()
                .max_total_relationship_xml_bytes(final_inventory.total_xml_bytes)
                .unwrap(),
            ReadLimits::builder()
                .max_total_relationship_xml_bytes(final_inventory.total_xml_bytes - 1)
                .unwrap(),
        ),
        (
            ReadResource::TotalRelationshipXmlEvents,
            ReadLimits::builder()
                .max_total_relationship_xml_events(final_inventory.total_xml_events)
                .unwrap(),
            ReadLimits::builder()
                .max_total_relationship_xml_events(final_inventory.total_xml_events - 1)
                .unwrap(),
        ),
        (
            ReadResource::TotalRelationships,
            ReadLimits::builder()
                .max_total_relationships(final_inventory.total_edges)
                .unwrap(),
            ReadLimits::builder()
                .max_total_relationships(final_inventory.total_edges - 1)
                .unwrap(),
        ),
        (
            ReadResource::RelationshipGraphNodes,
            ReadLimits::builder()
                .max_relationship_graph_nodes(final_inventory.graph_nodes)
                .unwrap(),
            ReadLimits::builder()
                .max_relationship_graph_nodes(final_inventory.graph_nodes - 1)
                .unwrap(),
        ),
    ];

    for (resource, exact_builder, under_builder) in cases {
        let exact =
            OpcPackage::from_bytes_with_limits(&fixture.bytes, exact_builder.build().unwrap())
                .unwrap();
        let mut exact = exact;
        add_unknown_relationships(&mut exact);
        load(&exact, "Pivot").unwrap();

        let under =
            OpcPackage::from_bytes_with_limits(&fixture.bytes, under_builder.build().unwrap())
                .unwrap();
        let mut under = under;
        add_unknown_relationships(&mut under);
        assert_read_limit(load(&under, "Pivot"), resource);
    }

    // The authored root target is deliberately longer than every source
    // target.  Both the XML attribute and resolved PackURI forms are charged,
    // so exact admission and one-byte-under rejection exercise the retained
    // target policy after the mutation.
    let target_exact = final_inventory.max_target_bytes;
    let target_exact_builder = ReadLimits::builder()
        .max_relationship_target_bytes(target_exact)
        .unwrap();
    let target_under_builder = ReadLimits::builder()
        .max_relationship_target_bytes(target_exact - 1)
        .unwrap();
    let mut exact =
        OpcPackage::from_bytes_with_limits(&fixture.bytes, target_exact_builder.build().unwrap())
            .unwrap();
    add_unknown_relationships(&mut exact);
    load(&exact, "Pivot").unwrap();

    let mut under =
        OpcPackage::from_bytes_with_limits(&fixture.bytes, target_under_builder.build().unwrap())
            .unwrap();
    add_unknown_relationships(&mut under);
    assert_read_limit(load(&under, "Pivot"), ReadResource::RelationshipTargetBytes);
}

#[test]
fn retained_relationship_xml_part_limit_has_exact_and_one_under_mutation_bounds() {
    let fixture = Fixture::build(FixtureOptions::default());
    let mut final_package = fixture.package();
    add_unknown_relationships(&mut final_package);
    let final_inventory = relationship_inventory(&final_package);
    assert!(final_inventory.max_xml_bytes > 1);

    let exact = ReadLimits::builder()
        .max_relationship_xml_bytes(final_inventory.max_xml_bytes)
        .unwrap()
        .build()
        .unwrap();
    let mut exact_package = OpcPackage::from_bytes_with_limits(&fixture.bytes, exact).unwrap();
    add_unknown_relationships(&mut exact_package);
    load(&exact_package, "Pivot").unwrap();

    let under = ReadLimits::builder()
        .max_relationship_xml_bytes(final_inventory.max_xml_bytes - 1)
        .unwrap()
        .build()
        .unwrap();
    let mut under_package = OpcPackage::from_bytes_with_limits(&fixture.bytes, under).unwrap();
    add_unknown_relationships(&mut under_package);
    assert_read_limit(
        load(&under_package, "Pivot"),
        ReadResource::RelationshipXmlBytes,
    );
}

#[test]
fn retained_relationship_archive_member_bound_uses_public_source_capture() {
    let fixture = Fixture::build(FixtureOptions::default());
    let mut package = fixture.package();
    add_unknown_relationships(&mut package);
    let inventory = relationship_inventory(&package);
    let owner = PackURI::new(SHEET_URI).unwrap();
    let member_len = owner.rels_uri().unwrap().membername().len();
    assert!(member_len > 1);
    assert_eq!(owner.as_str().len(), owner.membername().len() + 1);
    assert!(inventory.max_part_name_bytes >= owner.as_str().len());
    assert!(inventory.max_relationship_member_name_bytes >= member_len);

    let exact = ReadLimits::builder()
        .max_archive_member_name_bytes(member_len as u64)
        .unwrap()
        .build()
        .unwrap();
    assert!(
        package
            .source_relationships_with_limits(&owner, exact)
            .is_ok()
    );

    let under = ReadLimits::builder()
        .max_archive_member_name_bytes((member_len - 1) as u64)
        .unwrap()
        .build()
        .unwrap();
    assert!(matches!(
        package.source_relationships_with_limits(&owner, under),
        Err(OpcError::ReadLimit {
            resource: ReadResource::ArchiveMemberNameBytes,
            actual,
            maximum,
        }) if actual == member_len as u64 && maximum == (member_len - 1) as u64
    ));
}
