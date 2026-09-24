#![allow(
    dead_code,
    clippy::unwrap_used,
    clippy::expect_used,
    reason = "synthetic OPC contract tests intentionally use panic-on-fixture-failure"
)]

//! Synthetic source-bound coverage for the public `pivotTableData` facade.
//!
//! The package builder below follows the established pivot server-format and
//! cached-unique-name fixtures.  It deliberately contains the complete local
//! C444/C983/725/ABF5 graph, with optional F057/connection-ID routes.  These
//! are schema and source-preservation tests; they do not claim native Office
//! producer interoperability.

use litchi_opc::constants::{content_type as ct, relationship_type as rt};
use litchi_opc::{BlobPart, OpcPackage, PackURI, PackageWriter, TargetMode};
use litchi_xlsx::pivot::{
    PivotCacheId, PivotCellType, PivotCellValueEdit, PivotTableDataDiagnostic,
    PivotTableDataLimits, PivotTableSelector, PivotValueAttributeEdit, PivotValueCellExtraEdit,
    edit_pivot_table_data, edit_pivot_table_data_with_limits, load_pivot_table_data,
    load_pivot_table_data_with_limits,
};
use litchi_xlsx::{Error, ReadLimits, Workbook};

const TRANSITIONAL_MAIN: &str = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
const STRICT_MAIN: &str = "http://purl.oclc.org/ooxml/spreadsheetml/main";
const TRANSITIONAL_REL: &str =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const STRICT_REL: &str = "http://purl.oclc.org/ooxml/officeDocument/relationships";
const EXT_NS: &str = "http://schemas.microsoft.com/office/spreadsheetml/2010/11/main";
const CACHE_EXT_NS: &str = "http://schemas.microsoft.com/office/spreadsheetml/2009/9/main";
const MCE_NS: &str = "http://schemas.openxmlformats.org/markup-compatibility/2006";

const C444_URI: &str = "{44433962-1CF7-4059-B4EE-95C3D5FFCF73}";
const C510_URI: &str = "{C510F80B-63DE-4267-81D5-13C33094786E}";
const C983_URI: &str = "{983426D0-5260-488c-9760-48F4B6AC55F4}";
const PIVOT_CACHE_DEFINITION_URI: &str = "{725AE2AE-9491-48BE-B2B4-4EB974FC3084}";
const ABF5_URI: &str = "{ABF5C744-AB39-4b91-8756-CFA1BBC848D5}";
const F057_URI: &str = "{F057638F-6D5F-4E77-A914-E7F072B9BCA8}";

const WORKBOOK_URI: &str = "/xl/workbook.xml";
const SHEET_URI: &str = "/xl/worksheets/sheet1.xml";
const TABLE_URI: &str = "/xl/pivotTables/pivotTable1.xml";
const CACHE_URI: &str = "/xl/pivotCache/pivotCacheDefinition1.xml";
const CONNECTIONS_URI: &str = "/xl/connections.xml";
const WORKSHEET_PIVOT_URI_PREFIX: &str = "/xl/pivotTables/worksheetPivot";
const WORKSHEET_TABLE_URI: &str = "/xl/tables/table1.xml";

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

    const fn connections_rel(self) -> &'static str {
        match self {
            Self::Transitional => {
                "http://schemas.openxmlformats.org/officeDocument/2006/relationships/connections"
            },
            Self::Strict => "http://purl.oclc.org/ooxml/officeDocument/relationships/connections",
        }
    }

    const fn office_document_rel(self) -> &'static str {
        match self {
            Self::Transitional => rt::OFFICE_DOCUMENT,
            Self::Strict => rt::STRICT_OFFICE_DOCUMENT,
        }
    }
}

#[derive(Clone, Copy, Debug, PartialEq, Eq)]
enum CacheRoute {
    F057Name,
    ExplicitId,
    BothConsistent,
    BothMismatch,
    EmptyF057Name,
    ExplicitIdZero,
    MissingF057,
    NoConnectionRoute,
}

#[derive(Clone, Debug)]
struct FixtureOptions {
    dialect: Dialect,
    root_prefix: &'static str,
    extension_prefix: &'static str,
    cache_prefix: &'static str,
    c444_uri: &'static str,
    c510_uri: &'static str,
    c983_uri: &'static str,
    cache_ext_uri: &'static str,
    abf5_uri: &'static str,
    f057_uri: &'static str,
    route: CacheRoute,
    table_name: &'static str,
    table_exts: String,
    workbook_exts: String,
    cache_exts: String,
    server_formats: String,
    server_count: &'static str,
    data_payload: String,
    cache_id_lexical: &'static str,
    cache_definition_id: Option<&'static str>,
    duplicate_cache_definition: bool,
    duplicate_workbook_reference: bool,
    workbook_mce: bool,
    cache_mce: bool,
    worksheet_pivot_names: Vec<String>,
    worksheet_table_name: Option<String>,
    signed: bool,
}

impl Default for FixtureOptions {
    fn default() -> Self {
        Self {
            dialect: Dialect::Transitional,
            root_prefix: "",
            extension_prefix: "x15",
            cache_prefix: "x14",
            c444_uri: C444_URI,
            c510_uri: C510_URI,
            c983_uri: C983_URI,
            cache_ext_uri: PIVOT_CACHE_DEFINITION_URI,
            abf5_uri: ABF5_URI,
            f057_uri: F057_URI,
            route: CacheRoute::F057Name,
            table_name: "Pivot",
            table_exts: String::new(),
            workbook_exts: String::new(),
            cache_exts: String::new(),
            server_formats:
                r#"<x15:serverFormat culture="first"/><x15:serverFormat culture="second"/>"#
                    .to_owned(),
            server_count: "2",
            data_payload: r#"<x15:pivotTableData rowCount="2" columnCount="2" cacheId="7"><x15:pivotRow count="2" r="0"><x15:c i="0" t="str"><x15:v>alpha</x15:v><x15:x in="+0" bc="0000000A" fc="000000FF" i="1" un="false" st="0" b="true"/></x15:c><x15:c i="1" t="b"><x15:v>true</x15:v></x15:c></x15:pivotRow><x15:pivotRow count="2" r="1"><x15:c i="0" t="e"><x15:v>#N/A</x15:v></x15:c><x15:c i="1" t="bl"><x15:v/></x15:c></x15:pivotRow></x15:pivotTableData>"#.to_owned(),
            cache_id_lexical: "7",
            cache_definition_id: Some("7"),
            duplicate_cache_definition: false,
            duplicate_workbook_reference: false,
            workbook_mce: false,
            cache_mce: false,
            worksheet_pivot_names: Vec::new(),
            worksheet_table_name: None,
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
        let root = if options.root_prefix.is_empty() {
            String::new()
        } else {
            format!("{}:", options.root_prefix)
        };
        let main_decl = if options.root_prefix.is_empty() {
            format!(r#"xmlns="{main}""#)
        } else {
            format!(r#"xmlns:{}="{main}""#, options.root_prefix)
        };
        let ext = options.extension_prefix;
        let cache = options.cache_prefix;
        let cache_id = options.cache_id_lexical;
        let server_formats = options.server_formats.replace("x15:", &format!("{ext}:"));
        let data_payload = options
            .data_payload
            .replace("x15:", &format!("{ext}:"))
            .replace("cacheId=\"7\"", &format!("cacheId=\"{cache_id}\""));

        let workbook_reference = if options.workbook_mce {
            format!(
                r#"<mc:AlternateContent><mc:Choice Requires="{ext}"><{root}ext uri="{c983_uri}"><{ext}:pivotTableReferences><{ext}:pivotTableReference r:id="rIdPivot"/></{ext}:pivotTableReferences></{root}ext></mc:Choice><mc:Fallback><{root}ext uri="{{opaque-c983-fallback}}"><{ext}:future/></{root}ext></mc:Fallback></mc:AlternateContent>"#,
                ext = ext,
                root = root,
                c983_uri = options.c983_uri,
            )
        } else {
            format!(
                r#"<{root}ext uri="{c983_uri}"><{ext}:pivotTableReferences><{ext}:pivotTableReference r:id="rIdPivot"/></{ext}:pivotTableReferences></{root}ext>"#,
                root = root,
                ext = ext,
                c983_uri = options.c983_uri,
            )
        };
        let duplicate_workbook_reference = if options.duplicate_workbook_reference {
            format!(
                r#"<{root}ext uri="{c983_uri}"><{ext}:pivotTableReferences><{ext}:pivotTableReference r:id="rIdPivot"/></{ext}:pivotTableReferences></{root}ext>"#,
                root = root,
                ext = ext,
                c983_uri = options.c983_uri,
            )
        } else {
            String::new()
        };

        let workbook_xml = format!(
            r#"<{root}workbook {main_decl} xmlns:r="{rel}" xmlns:{ext}="{EXT_NS}" xmlns:mc="{MCE_NS}" mc:Ignorable="{ext}"><{root}sheets><{root}sheet name="Sheet1" sheetId="1" r:id="rIdSheet"/></{root}sheets><{root}pivotCaches><{root}pivotCache cacheId="{cache_id}" r:id="rIdCache"/></{root}pivotCaches><{root}extLst>{workbook_reference}{duplicate_workbook_reference}</{root}extLst>{workbook_exts}</{root}workbook>"#,
            root = root,
            main_decl = main_decl,
            rel = rel,
            ext = ext,
            cache_id = cache_id,
            workbook_reference = workbook_reference,
            duplicate_workbook_reference = duplicate_workbook_reference,
            workbook_exts = options.workbook_exts,
        );

        let cache_source = match options.route {
            CacheRoute::F057Name | CacheRoute::EmptyF057Name => format!(
                r#"<{root}cacheSource type="external"><{root}extLst><{root}ext uri="{f057_uri}"><{cache}:sourceConnection name="{name}"/></{root}ext></{root}extLst></{root}cacheSource>"#,
                root = root,
                f057_uri = options.f057_uri,
                cache = cache,
                name = if options.route == CacheRoute::EmptyF057Name {
                    ""
                } else {
                    "canonical"
                },
            ),
            CacheRoute::ExplicitId | CacheRoute::ExplicitIdZero => format!(
                r#"<{root}cacheSource type="external" connectionId="{id}"/>"#,
                root = root,
                id = if options.route == CacheRoute::ExplicitIdZero {
                    "0"
                } else {
                    "7"
                },
            ),
            CacheRoute::BothConsistent => format!(
                r#"<{root}cacheSource type="external" connectionId="7"><{root}extLst><{root}ext uri="{f057_uri}"><{cache}:sourceConnection name="canonical"/></{root}ext></{root}extLst></{root}cacheSource>"#,
                root = root,
                f057_uri = options.f057_uri,
                cache = cache,
            ),
            CacheRoute::BothMismatch => format!(
                r#"<{root}cacheSource type="external" connectionId="8"><{root}extLst><{root}ext uri="{f057_uri}"><{cache}:sourceConnection name="canonical"/></{root}ext></{root}extLst></{root}cacheSource>"#,
                root = root,
                f057_uri = options.f057_uri,
                cache = cache,
            ),
            CacheRoute::MissingF057 => format!(
                r#"<{root}cacheSource type="external" connectionId="7"/>"#,
                root = root
            ),
            CacheRoute::NoConnectionRoute => {
                format!(r#"<{root}cacheSource type="external"/>"#, root = root)
            },
        };

        let cache_definition_id = options
            .cache_definition_id
            .map(|value| format!(r#" pivotCacheId="{value}""#))
            .unwrap_or_default();
        let cache_definition_ext = format!(
            r#"<{root}ext uri="{cache_ext_uri}"><{cache}:pivotCacheDefinition{cache_definition_id}/></{root}ext>"#,
            root = root,
            cache_ext_uri = options.cache_ext_uri,
            cache = cache,
            cache_definition_id = cache_definition_id,
        );
        let cache_definition_ext = if options.cache_mce {
            format!(
                r#"<mc:AlternateContent><mc:Choice Requires="{cache}">{cache_definition_ext}</mc:Choice><mc:Fallback><{root}ext uri="{{opaque-725-fallback}}"><{cache}:future/></{root}ext></mc:Fallback></mc:AlternateContent>"#,
                cache = cache,
                cache_definition_ext = cache_definition_ext,
                root = root,
            )
        } else {
            cache_definition_ext
        };
        let duplicate_cache_definition = if options.duplicate_cache_definition {
            format!(
                r#"<{root}ext uri="{cache_ext_uri}"><{cache}:pivotCacheDefinition{cache_definition_id}/></{root}ext>"#,
                root = root,
                cache_ext_uri = options.cache_ext_uri,
                cache = cache,
                cache_definition_id = cache_definition_id,
            )
        } else {
            String::new()
        };
        let cache_xml = format!(
            r#"<{root}pivotCacheDefinition {main_decl} xmlns:{ext}="{EXT_NS}" xmlns:{cache}="{CACHE_EXT_NS}" xmlns:mc="{MCE_NS}" mc:Ignorable="{ext} {cache}">{cache_source}<{root}cacheFields count="0"/><{root}extLst>{cache_definition_ext}{duplicate_cache_definition}<{root}ext uri="{abf5_uri}"><{ext}:pivotCacheIdVersion cacheIdSupportedVersion="15" cacheIdCreatedVersion="15"/></{root}ext>{cache_exts}</{root}extLst></{root}pivotCacheDefinition>"#,
            root = root,
            main_decl = main_decl,
            ext = ext,
            cache = cache,
            cache_source = cache_source,
            cache_definition_ext = cache_definition_ext,
            duplicate_cache_definition = duplicate_cache_definition,
            abf5_uri = options.abf5_uri,
            cache_exts = options.cache_exts,
        );

        let connections_xml = format!(
            r#"<{root}connections {main_decl}><{root}connection id="7" name="canonical" type="1" refreshedVersion="7"/><{root}connection id="0" name="" type="1" refreshedVersion="7"/></{root}connections>"#,
            root = root,
            main_decl = main_decl,
        );

        let table_sibling_ext = if options.table_exts.is_empty() {
            String::new()
        } else {
            format!(
                r#"<{root}ext uri="{{opaque-table-sibling}}">{table_exts}</{root}ext>"#,
                root = root,
                table_exts = options.table_exts,
            )
        };

        let table_xml = format!(
            r#"<{root}pivotTableDefinition {main_decl} xmlns:{ext}="{EXT_NS}" xmlns:mc="{MCE_NS}" mc:Ignorable="{ext}" name="{table_name}" cacheId="{cache_id}"><{root}location ref="A1:B3" firstHeaderRow="1" firstDataRow="2" firstDataCol="1"/><{root}pivotFields count="2"><{root}pivotField/><{root}pivotField/></{root}pivotFields><{root}rowItems count="2"><{root}i/><{root}i/></{root}rowItems><{root}colItems count="2"><{root}i/><{root}i/></{root}colItems><{root}extLst><{root}ext uri="{c510_uri}"><{ext}:pivotTableServerFormats count="{server_count}">{server_formats}</{ext}:pivotTableServerFormats></{root}ext><{root}ext uri="{c444_uri}">{data_payload}</{root}ext>{table_sibling_ext}</{root}extLst></{root}pivotTableDefinition>"#,
            root = root,
            main_decl = main_decl,
            ext = ext,
            table_name = options.table_name,
            cache_id = cache_id,
            server_count = options.server_count,
            server_formats = server_formats,
            data_payload = data_payload,
            c510_uri = options.c510_uri,
            c444_uri = options.c444_uri,
            table_sibling_ext = table_sibling_ext,
        );

        let worksheet_table_parts = options
            .worksheet_table_name
            .as_ref()
            .map(|_| {
                format!(
                    r#"<{root}tableParts count="1"><{root}tablePart r:id="rIdWorksheetTable"/></{root}tableParts>"#,
                    root = root,
                )
            })
            .unwrap_or_default();
        let worksheet_pivot_parts = if options.worksheet_pivot_names.is_empty() {
            String::new()
        } else {
            let mut parts = format!(
                r#"<{root}pivotTableParts count="{count}">"#,
                count = options.worksheet_pivot_names.len(),
                root = root,
            );
            for (index, _) in options.worksheet_pivot_names.iter().enumerate() {
                parts.push_str(&format!(
                    r#"<{root}pivotTablePart r:id="rIdWorksheetPivot{}"/>"#,
                    index + 1,
                    root = root,
                ));
            }
            parts.push_str(&format!(r#"</{root}pivotTableParts>"#, root = root));
            parts
        };
        let worksheet_xml = format!(
            r#"<{root}worksheet {main_decl} xmlns:r="{rel}"><{root}sheetData/>{worksheet_table_parts}{worksheet_pivot_parts}</{root}worksheet>"#,
            root = root,
            main_decl = main_decl,
            rel = rel,
            worksheet_table_parts = worksheet_table_parts,
            worksheet_pivot_parts = worksheet_pivot_parts,
        );

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
                worksheet_xml.into_bytes(),
            )))
            .unwrap();
        package
            .try_add_part(Box::new(BlobPart::new(
                PackURI::new(TABLE_URI).unwrap(),
                ct::SML_PIVOT_TABLE.to_owned(),
                table_xml.as_bytes().to_vec(),
            )))
            .unwrap();
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
                "application/vnd.openxmlformats-officedocument.spreadsheetml.connections+xml"
                    .to_owned(),
                connections_xml.as_bytes().to_vec(),
            )))
            .unwrap();

        for (index, name) in options.worksheet_pivot_names.iter().enumerate() {
            let uri = format!("{WORKSHEET_PIVOT_URI_PREFIX}{}.xml", index + 1);
            let pivot_xml = format!(
                r#"<{root}pivotTableDefinition {main_decl} name="{name}" cacheId="{cache_id}"><{root}location ref="A1:B3"/><{root}pivotFields count="0"/><{root}rowItems count="0"/><{root}colItems count="0"/></{root}pivotTableDefinition>"#,
                root = root,
                main_decl = main_decl,
                name = name,
                cache_id = cache_id,
            );
            package
                .try_add_part(Box::new(BlobPart::new(
                    PackURI::new(&uri).unwrap(),
                    ct::SML_PIVOT_TABLE.to_owned(),
                    pivot_xml.into_bytes(),
                )))
                .unwrap();
        }
        if let Some(name) = options.worksheet_table_name.as_ref() {
            let table_xml = format!(
                r#"<table xmlns="{main}" id="1" name="{name}" displayName="{name}" ref="A1:A2"><tableColumns count="1"><tableColumn id="1" name="Column1"/></tableColumns></table>"#,
                main = main,
                name = name,
            );
            package
                .try_add_part(Box::new(BlobPart::new(
                    PackURI::new(WORKSHEET_TABLE_URI).unwrap(),
                    ct::SML_TABLE.to_owned(),
                    table_xml.into_bytes(),
                )))
                .unwrap();
        }

        let workbook_uri = PackURI::new(WORKBOOK_URI).unwrap();
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
        package
            .get_part_mut(&PackURI::new(TABLE_URI).unwrap())
            .unwrap()
            .rels_mut()
            .try_add_relationship(
                options.dialect.pivot_cache_rel().to_owned(),
                "../pivotCache/pivotCacheDefinition1.xml".to_owned(),
                "rIdCache".to_owned(),
                TargetMode::Internal,
            )
            .unwrap();
        let worksheet = package
            .get_part_mut(&PackURI::new(SHEET_URI).unwrap())
            .unwrap();
        for (index, _) in options.worksheet_pivot_names.iter().enumerate() {
            worksheet
                .rels_mut()
                .try_add_relationship(
                    options.dialect.pivot_table_rel().to_owned(),
                    format!("../pivotTables/worksheetPivot{}.xml", index + 1),
                    format!("rIdWorksheetPivot{}", index + 1),
                    TargetMode::Internal,
                )
                .unwrap();
        }
        if options.worksheet_table_name.is_some() {
            worksheet
                .rels_mut()
                .try_add_relationship(
                    match options.dialect {
                        Dialect::Transitional => rt::TABLE,
                        Dialect::Strict => rt::STRICT_TABLE,
                    }
                    .to_owned(),
                    "../tables/table1.xml".to_owned(),
                    "rIdWorksheetTable".to_owned(),
                    TargetMode::Internal,
                )
                .unwrap();
        }
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
            bytes: PackageWriter::to_bytes(&package).unwrap(),
            workbook_xml,
            table_xml,
            cache_xml,
            connections_xml,
        }
    }

    fn workbook(&self) -> Workbook {
        Workbook::from_bytes(self.bytes.clone()).unwrap()
    }

    fn package(&self) -> OpcPackage {
        OpcPackage::from_bytes(&self.bytes).unwrap()
    }
}

fn table_blob(package: &OpcPackage) -> Vec<u8> {
    package
        .get_part(&PackURI::new(TABLE_URI).unwrap())
        .unwrap()
        .blob()
        .to_vec()
}

fn workbook_table_blob(workbook: &Workbook) -> Vec<u8> {
    let bytes = workbook.to_bytes().unwrap();
    table_blob(&OpcPackage::from_bytes(&bytes).unwrap())
}

fn workbook_blob(workbook: &Workbook) -> Vec<u8> {
    let bytes = workbook.to_bytes().unwrap();
    OpcPackage::from_bytes(&bytes)
        .unwrap()
        .get_part(&PackURI::new(WORKBOOK_URI).unwrap())
        .unwrap()
        .blob()
        .to_vec()
}

fn workbook_member_blob(workbook: &Workbook, uri: &str) -> Vec<u8> {
    let bytes = workbook.to_bytes().unwrap();
    OpcPackage::from_bytes(&bytes)
        .unwrap()
        .get_part(&PackURI::new(uri).unwrap())
        .unwrap()
        .blob()
        .to_vec()
}

fn fixture_part_xml(fixture: &Fixture, uri: &str) -> String {
    match uri {
        WORKBOOK_URI => fixture.workbook_xml.clone(),
        TABLE_URI => fixture.table_xml.clone(),
        CACHE_URI => fixture.cache_xml.clone(),
        CONNECTIONS_URI => fixture.connections_xml.clone(),
        _ => panic!("fixture has no tracked XML Part {uri}"),
    }
}

fn set_fixture_part_xml(fixture: &mut Fixture, uri: &str, xml: String) {
    let mut package = fixture.package();
    package
        .get_part_mut(&PackURI::new(uri).unwrap())
        .unwrap()
        .set_blob(xml.as_bytes().to_vec());
    fixture.bytes = PackageWriter::to_bytes(&package).unwrap();
    match uri {
        WORKBOOK_URI => fixture.workbook_xml = xml,
        TABLE_URI => fixture.table_xml = xml,
        CACHE_URI => fixture.cache_xml = xml,
        CONNECTIONS_URI => fixture.connections_xml = xml,
        _ => panic!("fixture has no tracked XML Part {uri}"),
    }
}

fn replace_fixture_part_xml(fixture: &mut Fixture, uri: &str, from: &str, to: &str) {
    let source = fixture_part_xml(fixture, uri);
    assert_eq!(
        source.matches(from).count(),
        1,
        "expected one {from:?} occurrence in {uri}"
    );
    set_fixture_part_xml(fixture, uri, source.replacen(from, to, 1));
}

fn fixture_with_local_v_namespace(default_namespace: bool) -> Fixture {
    let mut fixture = Fixture::build(FixtureOptions::default());
    let replacement = if default_namespace {
        format!(r#"<v xmlns="{EXT_NS}">alpha</v>"#)
    } else {
        format!(r#"<x15:v xmlns:x15="{EXT_NS}">alpha</x15:v>"#)
    };
    replace_fixture_part_xml(
        &mut fixture,
        TABLE_URI,
        "<x15:v>alpha</x15:v>",
        &replacement,
    );
    fixture
}

fn fixture_with_duplicate_c444_owner() -> Fixture {
    let mut fixture = Fixture::build(FixtureOptions::default());
    let source = fixture.table_xml.clone();
    let owner_start = format!(r#"<ext uri="{C444_URI}">"#);
    let start = source.find(&owner_start).unwrap();
    let end = start + source[start..].find("</ext>").unwrap() + "</ext>".len();
    let owner = source[start..end].to_owned();
    let mut updated = source.clone();
    updated.insert_str(end, &owner);
    set_fixture_part_xml(&mut fixture, TABLE_URI, updated);
    fixture
}

fn fixture_with_duplicate_c444_payload() -> Fixture {
    let mut fixture = Fixture::build(FixtureOptions::default());
    let source = fixture.table_xml.clone();
    let start = source.find("<x15:pivotTableData").unwrap();
    let end = start
        + source[start..].find("</x15:pivotTableData>").unwrap()
        + "</x15:pivotTableData>".len();
    let payload = source[start..end].to_owned();
    let mut updated = source.clone();
    updated.insert_str(end, &payload);
    set_fixture_part_xml(&mut fixture, TABLE_URI, updated);
    fixture
}

fn oversized_fragment_comment() -> String {
    format!("<!--{}-->", "x".repeat(1024 * 1024))
}

fn fixture_with_oversized_duplicate_c444_owner() -> Fixture {
    let mut fixture = Fixture::build(FixtureOptions::default());
    let source = fixture.table_xml.clone();
    let owner_start = format!(r#"<ext uri="{C444_URI}">"#);
    let start = source.find(&owner_start).unwrap();
    let end = start + source[start..].find("</ext>").unwrap() + "</ext>".len();
    let owner = source[start..end].to_owned();
    let oversized_owner = owner.replacen(
        "</ext>",
        &format!("{}</ext>", oversized_fragment_comment()),
        1,
    );
    let mut updated = source;
    updated.insert_str(end, &oversized_owner);
    set_fixture_part_xml(&mut fixture, TABLE_URI, updated);
    fixture
}

fn fixture_with_oversized_duplicate_c444_payload() -> Fixture {
    let mut fixture = Fixture::build(FixtureOptions::default());
    let source = fixture.table_xml.clone();
    let start = source.find("<x15:pivotTableData").unwrap();
    let end = start
        + source[start..].find("</x15:pivotTableData>").unwrap()
        + "</x15:pivotTableData>".len();
    let payload = source[start..end].to_owned();
    let oversized_payload = payload.replacen(
        "</x15:pivotTableData>",
        &format!("{}</x15:pivotTableData>", oversized_fragment_comment()),
        1,
    );
    let mut updated = source;
    updated.insert_str(end, &oversized_payload);
    set_fixture_part_xml(&mut fixture, TABLE_URI, updated);
    fixture
}

fn assert_refused_or_diagnostic_read_only(workbook: &Workbook) {
    match workbook.pivot_table_data("Pivot") {
        Err(_) => {},
        Ok(Some(view)) => {
            assert_ne!(
                view.diagnostic_status(),
                PivotTableDataDiagnostic::None,
                "malformed recognized closure became silently editable"
            );
            assert!(!view.is_editable());
            assert!(workbook.edit_pivot_table_data("Pivot").is_err());
        },
        Ok(None) => panic!("recognized C444 owner disappeared"),
    }
}

fn assert_diagnostic_read_only(workbook: &Workbook) {
    let view = workbook
        .pivot_table_data("Pivot")
        .unwrap()
        .expect("recognized C444 owner should remain readable");
    assert_ne!(view.diagnostic_status(), PivotTableDataDiagnostic::None);
    assert!(!view.is_editable());
    assert!(workbook.edit_pivot_table_data("Pivot").is_err());
}

fn read_u16(bytes: &[u8], offset: usize) -> u16 {
    u16::from_le_bytes([bytes[offset], bytes[offset + 1]])
}

fn read_u32(bytes: &[u8], offset: usize) -> u32 {
    u32::from_le_bytes([
        bytes[offset],
        bytes[offset + 1],
        bytes[offset + 2],
        bytes[offset + 3],
    ])
}

fn mark_zip_entry_encrypted(mut bytes: Vec<u8>, wanted: &str) -> Vec<u8> {
    let eocd = bytes
        .windows(4)
        .rposition(|window| window == 0x0605_4b50_u32.to_le_bytes())
        .unwrap();
    let count = read_u16(&bytes, eocd + 10) as usize;
    let central_offset = read_u32(&bytes, eocd + 16) as usize;
    let mut cursor = central_offset;
    for _ in 0..count {
        assert_eq!(&bytes[cursor..cursor + 4], &0x0201_4b50_u32.to_le_bytes());
        let name_len = read_u16(&bytes, cursor + 28) as usize;
        let extra_len = read_u16(&bytes, cursor + 30) as usize;
        let comment_len = read_u16(&bytes, cursor + 32) as usize;
        let name_start = cursor + 46;
        let name = &bytes[name_start..name_start + name_len];
        let local_offset = read_u32(&bytes, cursor + 42) as usize;
        if name == wanted.as_bytes() {
            let central_flags = read_u16(&bytes, cursor + 8) | 1;
            bytes[cursor + 8..cursor + 10].copy_from_slice(&central_flags.to_le_bytes());
            let local_flags = read_u16(&bytes, local_offset + 6) | 1;
            bytes[local_offset + 6..local_offset + 8].copy_from_slice(&local_flags.to_le_bytes());
            return bytes;
        }
        cursor += 46 + name_len + extra_len + comment_len;
    }
    panic!("missing ZIP member {wanted}");
}

fn assert_data_error(options: FixtureOptions) {
    let fixture = Fixture::build(options);
    let result = fixture.workbook().pivot_table_data("Pivot");
    assert!(
        result.is_err(),
        "malformed pivotTableData was accepted: {result:?}"
    );
}

fn assert_resource_limit(error: &Error, expected_scope: &str) {
    match error {
        Error::ResourceLimit(limit) => assert!(
            limit.scope.contains(expected_scope),
            "resource limit scope {:?} did not contain {expected_scope:?}",
            limit.scope
        ),
        other => panic!("expected a typed resource refusal, got {other:?}"),
    }
}

fn retained_limit_admits(package: &OpcPackage, maximum: usize) -> bool {
    let limits = PivotTableDataLimits::new().with_max_retained_bytes(maximum);
    match load_pivot_table_data_with_limits(package, "Pivot", &limits) {
        Ok(_) => true,
        Err(Error::ResourceLimit(_)) => false,
        Err(error) => panic!("retained-limit search hit an unrelated error: {error:?}"),
    }
}

fn minimum_semantic_limit(mut admits: impl FnMut(usize) -> bool) -> usize {
    let mut lower = 0usize;
    let mut upper = PivotTableDataLimits::new().max_retained_bytes();
    assert!(admits(upper));
    while upper.saturating_sub(lower) > 1 {
        let midpoint = lower + (upper - lower) / 2;
        if admits(midpoint) {
            upper = midpoint;
        } else {
            lower = midpoint;
        }
    }
    upper
}

fn minimum_retained_limit(package: &OpcPackage) -> usize {
    minimum_semantic_limit(|maximum| retained_limit_admits(package, maximum))
}

fn transaction_limit_admits(fixture: &Fixture, maximum: usize) -> bool {
    let workbook = fixture.workbook();
    let limits = PivotTableDataLimits::new().with_max_retained_bytes(maximum);
    match workbook.edit_pivot_table_data_with_limits("Pivot", &limits) {
        Ok(_) => true,
        Err(Error::ResourceLimit(_)) => false,
        Err(error) => panic!("transaction-limit search hit an unrelated error: {error:?}"),
    }
}

fn minimum_transaction_limit(fixture: &Fixture) -> usize {
    minimum_semantic_limit(|maximum| transaction_limit_admits(fixture, maximum))
}

fn setter_limit_admits(fixture: &Fixture, replacement: &str, maximum: usize) -> bool {
    let workbook = fixture.workbook();
    let limits = PivotTableDataLimits::new().with_max_retained_bytes(maximum);
    let mut edit = match workbook.edit_pivot_table_data_with_limits("Pivot", &limits) {
        Ok(edit) => edit,
        Err(Error::ResourceLimit(_)) => return false,
        Err(error) => panic!("setter-limit search hit an unrelated error: {error:?}"),
    };
    match edit.set_text((0, 0), replacement) {
        Ok(_) => true,
        Err(Error::ResourceLimit(_)) => false,
        Err(error) => panic!("setter-limit search hit an unrelated error: {error:?}"),
    }
}

fn minimum_setter_limit(fixture: &Fixture, replacement: &str) -> usize {
    minimum_semantic_limit(|maximum| setter_limit_admits(fixture, replacement, maximum))
}

fn owned_text_with_capacity(length: usize, capacity: usize) -> PivotCellValueEdit {
    let mut value = String::with_capacity(capacity);
    value.push_str(&"q".repeat(length));
    assert!(value.capacity() >= capacity);
    PivotCellValueEdit::text(value)
}

fn owned_setter_limit_admits(
    fixture: &Fixture,
    length: usize,
    capacity: usize,
    maximum: usize,
) -> bool {
    let workbook = fixture.workbook();
    let limits = PivotTableDataLimits::new().with_max_retained_bytes(maximum);
    let mut edit = match workbook.edit_pivot_table_data_with_limits("Pivot", &limits) {
        Ok(edit) => edit,
        Err(Error::ResourceLimit(_)) => return false,
        Err(error) => panic!("owned setter-limit search hit an unrelated error: {error:?}"),
    };
    match edit.set_value((0, 0), owned_text_with_capacity(length, capacity)) {
        Ok(_) => true,
        Err(Error::ResourceLimit(_)) => false,
        Err(error) => panic!("owned setter-limit search hit an unrelated error: {error:?}"),
    }
}

fn minimum_owned_setter_limit(fixture: &Fixture, length: usize, capacity: usize) -> usize {
    minimum_semantic_limit(|maximum| owned_setter_limit_admits(fixture, length, capacity, maximum))
}

fn repeated_replacement_limit_admits(
    fixture: &Fixture,
    first: &str,
    second: &str,
    maximum: usize,
) -> bool {
    let workbook = fixture.workbook();
    let limits = PivotTableDataLimits::new().with_max_retained_bytes(maximum);
    let mut edit = match workbook.edit_pivot_table_data_with_limits("Pivot", &limits) {
        Ok(edit) => edit,
        Err(Error::ResourceLimit(_)) => return false,
        Err(error) => panic!("repeated-limit search hit an unrelated error: {error:?}"),
    };
    match edit.set_text((0, 0), first) {
        Ok(_) => {},
        Err(Error::ResourceLimit(_)) => return false,
        Err(error) => panic!("repeated-limit first replacement failed: {error:?}"),
    }
    match edit.set_text((0, 0), "short") {
        Ok(_) => {},
        Err(Error::ResourceLimit(_)) => return false,
        Err(error) => panic!("repeated-limit shrink replacement failed: {error:?}"),
    }
    match edit.set_text((0, 0), second) {
        Ok(_) => {},
        Err(Error::ResourceLimit(_)) => return false,
        Err(error) => panic!("repeated-limit second replacement failed: {error:?}"),
    }
    match edit.commit() {
        Ok(_) => true,
        Err(Error::ResourceLimit(_)) => false,
        Err(error) => panic!("repeated-limit publication failed: {error:?}"),
    }
}

fn minimum_repeated_replacement_limit(fixture: &Fixture, first: &str, second: &str) -> usize {
    minimum_semantic_limit(|maximum| {
        repeated_replacement_limit_admits(fixture, first, second, maximum)
    })
}

fn low_level_transaction_limit_admits(fixture: &Fixture, maximum: usize) -> bool {
    let mut package = fixture.package();
    let limits = PivotTableDataLimits::new().with_max_retained_bytes(maximum);
    match edit_pivot_table_data_with_limits(&mut package, "Pivot", &limits) {
        Ok(_) => true,
        Err(Error::ResourceLimit(_)) => false,
        Err(error) => {
            panic!("low-level transaction-limit search hit an unrelated error: {error:?}")
        },
    }
}

fn minimum_low_level_transaction_limit(fixture: &Fixture) -> usize {
    minimum_semantic_limit(|maximum| low_level_transaction_limit_admits(fixture, maximum))
}

fn low_level_setter_limit_admits(fixture: &Fixture, replacement: &str, maximum: usize) -> bool {
    let mut package = fixture.package();
    let limits = PivotTableDataLimits::new().with_max_retained_bytes(maximum);
    let mut edit = match edit_pivot_table_data_with_limits(&mut package, "Pivot", &limits) {
        Ok(edit) => edit,
        Err(Error::ResourceLimit(_)) => return false,
        Err(error) => panic!("low-level setter-limit search hit an unrelated error: {error:?}"),
    };
    match edit.set_text((0, 0), replacement) {
        Ok(_) => true,
        Err(Error::ResourceLimit(_)) => false,
        Err(error) => panic!("low-level setter-limit search hit an unrelated error: {error:?}"),
    }
}

fn minimum_low_level_setter_limit(fixture: &Fixture, replacement: &str) -> usize {
    minimum_semantic_limit(|maximum| low_level_setter_limit_admits(fixture, replacement, maximum))
}

fn fixture_with_large_selected_part_relationships() -> Fixture {
    let base = Fixture::build(FixtureOptions::default());
    let mut package = base.package();
    for (uri, label) in [
        (TABLE_URI, "table"),
        (CACHE_URI, "cache"),
        (WORKBOOK_URI, "workbook"),
    ] {
        let part = package.get_part_mut(&PackURI::new(uri).unwrap()).unwrap();
        for index in 0..16 {
            part.rels_mut()
                .try_add_relationship(
                    "http://example.test/opaque".to_owned(),
                    format!("https://example.test/{label}/{index}/{}", "x".repeat(3_900)),
                    format!("rIdOpaque{label}{index}"),
                    TargetMode::External,
                )
                .unwrap();
        }
    }

    for (uri, target, label) in [
        (
            "/xl/opaque/incoming-cache.bin",
            "../pivotCache/pivotCacheDefinition1.xml",
            "cache",
        ),
        (
            "/xl/opaque/incoming-workbook.bin",
            "../workbook.xml",
            "workbook",
        ),
    ] {
        package
            .try_add_part(Box::new(BlobPart::new(
                PackURI::new(uri).unwrap(),
                "application/octet-stream".to_owned(),
                b"opaque incoming source".to_vec(),
            )))
            .unwrap();
        let part = package.get_part_mut(&PackURI::new(uri).unwrap()).unwrap();
        part.rels_mut()
            .try_add_relationship(
                format!("http://example.test/incoming/{}", "r".repeat(4_096)),
                target.to_owned(),
                format!("rIdIncoming{label}"),
                TargetMode::Internal,
            )
            .unwrap();
    }
    Fixture {
        bytes: PackageWriter::to_bytes(&package).unwrap(),
        ..base
    }
}

fn default_data_payload_with(value: &str) -> String {
    format!(
        r#"<x15:pivotTableData rowCount="2" columnCount="2" cacheId="7"><x15:pivotRow count="2" r="0"><x15:c i="0" t="str"><x15:v>{value}</x15:v><x15:x in="+0" bc="0000000A" fc="000000FF" i="1" un="false" st="0" b="true"/></x15:c><x15:c i="1" t="b"><x15:v>true</x15:v></x15:c></x15:pivotRow><x15:pivotRow count="2" r="1"><x15:c i="0" t="e"><x15:v>#N/A</x15:v></x15:c><x15:c i="1" t="bl"><x15:v/></x15:c></x15:pivotRow></x15:pivotTableData>"#
    )
}

#[test]
fn fixture_has_complete_source_graph_before_public_facade_tests() {
    let fixture = Fixture::build(FixtureOptions::default());
    assert!(fixture.workbook_xml.contains(C983_URI));
    assert!(fixture.table_xml.contains(C444_URI));
    assert!(fixture.table_xml.contains(C510_URI));
    assert!(fixture.cache_xml.contains(PIVOT_CACHE_DEFINITION_URI));
    assert!(fixture.cache_xml.contains(ABF5_URI));
    assert!(fixture.cache_xml.contains(F057_URI));
    assert!(fixture.connections_xml.contains("canonical"));
    assert!(fixture.package().iter_parts().next().is_some());
    let _ = fixture.workbook();
}

#[test]
fn public_facade_reads_rows_and_authored_coordinates() {
    let fixture = Fixture::build(FixtureOptions::default());
    let workbook = fixture.workbook();
    let view = workbook.pivot_table_data("Pivot").unwrap().unwrap();

    assert_eq!(view.table_name(), "Pivot");
    assert_eq!(view.cache_id(), PivotCacheId::from(7));
    assert_eq!(view.row_count(), 2);
    assert_eq!(view.column_count(), 2);

    let rows = view.rows();
    assert_eq!(rows.len(), 2);
    let first = &rows[0];
    assert_eq!(first.row_index(), Some(0));
    let cells = first.cells();
    let first_cell = &cells[0];
    assert_eq!(first_cell.column_index(), Some(0));
    assert_eq!(first_cell.value_text(), "alpha");
    assert_eq!(cells.len(), 2);
    assert_eq!(rows[1].cells().len(), 2);
}

#[test]
fn public_facade_reads_all_cell_kinds_and_extra_flags() {
    let fixture = Fixture::build(FixtureOptions::default());
    let workbook = fixture.workbook();
    let view = workbook.pivot_table_data("Pivot").unwrap().unwrap();

    let text = view.cell((0, 0)).unwrap().unwrap();
    assert_eq!(text.kind(), PivotCellType::Text);
    assert_eq!(text.value_text(), "alpha");
    assert_eq!(text.address(), Some((0, 0).into()));
    let extra = text.extra().unwrap();
    assert_eq!(extra.format_index, Some(0));
    assert_eq!(extra.background_color, Some(0x0A));
    assert_eq!(extra.foreground_color, Some(0x00FF));
    assert_eq!(extra.italic, Some(true));
    assert_eq!(extra.underline, Some(false));
    assert_eq!(extra.strike, Some(false));
    assert_eq!(extra.bold, Some(true));

    let boolean = view.cell((0, 1)).unwrap().unwrap();
    assert_eq!(boolean.kind(), PivotCellType::Boolean);
    assert_eq!(boolean.value_text(), "true");
    assert!(boolean.extra().is_none());

    let error = view.cell((1, 0)).unwrap().unwrap();
    assert_eq!(error.kind(), PivotCellType::Error);
    assert_eq!(error.value_text(), "#N/A");

    let blank = view.cell((1, 1)).unwrap().unwrap();
    assert_eq!(blank.kind(), PivotCellType::Blank);
    assert_eq!(blank.value_text(), "");
    assert!(blank.extra().is_none());
    assert_eq!(view.diagnostic_status(), PivotTableDataDiagnostic::None);
    assert!(view.is_editable());
}

#[test]
fn strict_transitional_and_prefix_aliases_use_the_same_public_routes() {
    for (dialect, root_prefix, extension_prefix, cache_prefix) in [
        (Dialect::Transitional, "", "x15", "x14"),
        (Dialect::Strict, "s", "x16", "x17"),
    ] {
        let fixture = Fixture::build(FixtureOptions {
            dialect,
            root_prefix,
            extension_prefix,
            cache_prefix,
            ..FixtureOptions::default()
        });
        let view = fixture
            .workbook()
            .pivot_table_data("Pivot")
            .unwrap()
            .unwrap();
        assert_eq!(view.table_name(), "Pivot");
        assert_eq!(view.cache_id(), PivotCacheId::from(7));
        assert_eq!(view.cell((0, 0)).unwrap().unwrap().value_text(), "alpha");
    }
}

#[test]
fn external_cache_source_routes_are_generic_and_do_not_require_model_extensions() {
    for route in [
        CacheRoute::F057Name,
        CacheRoute::ExplicitId,
        CacheRoute::BothConsistent,
        CacheRoute::EmptyF057Name,
        CacheRoute::ExplicitIdZero,
        CacheRoute::MissingF057,
        CacheRoute::NoConnectionRoute,
    ] {
        let fixture = Fixture::build(FixtureOptions {
            route,
            ..FixtureOptions::default()
        });
        let view = fixture.workbook().pivot_table_data("Pivot").unwrap();
        assert!(view.is_some(), "cache route {route:?} was not accepted");
    }

    let mismatch = Fixture::build(FixtureOptions {
        route: CacheRoute::BothMismatch,
        ..FixtureOptions::default()
    });
    assert!(mismatch.workbook().pivot_table_data("Pivot").is_err());
}

#[test]
fn external_cache_without_f057_or_connection_id_ignores_unrelated_connections() {
    let mut malformed = Fixture::build(FixtureOptions {
        route: CacheRoute::NoConnectionRoute,
        ..FixtureOptions::default()
    });
    set_fixture_part_xml(
        &mut malformed,
        CONNECTIONS_URI,
        "<notConnections/>".to_owned(),
    );
    let workbook = malformed.workbook();
    let view = workbook.pivot_table_data("Pivot").unwrap().unwrap();
    assert_eq!(view.diagnostic_status(), PivotTableDataDiagnostic::None);
    assert!(view.is_editable());
    assert_eq!(view.cell((0, 0)).unwrap().unwrap().value_text(), "alpha");

    let source = Fixture::build(FixtureOptions {
        route: CacheRoute::NoConnectionRoute,
        ..FixtureOptions::default()
    })
    .workbook();
    let mut edit = source.edit_pivot_table_data("Pivot").unwrap();
    assert!(edit.set_text((0, 0), "unrouted-patch").unwrap());
    let committed = edit.commit().unwrap();
    let patch = committed.patch().clone();

    let mut candidate_package = OpcPackage::from_bytes(&source.to_bytes().unwrap()).unwrap();
    let candidate_connections = String::from_utf8(
        candidate_package
            .get_part(&PackURI::new(CONNECTIONS_URI).unwrap())
            .unwrap()
            .blob()
            .to_vec(),
    )
    .unwrap()
    .replace("refreshedVersion=\"7\"", "refreshedVersion=\"8\"");
    candidate_package
        .get_part_mut(&PackURI::new(CONNECTIONS_URI).unwrap())
        .unwrap()
        .set_blob(candidate_connections.into_bytes());
    let candidate =
        Workbook::from_bytes(PackageWriter::to_bytes(&candidate_package).unwrap()).unwrap();
    let applied = patch.apply(&candidate).unwrap();
    assert!(applied.changed());
    assert_eq!(
        applied
            .workbook()
            .pivot_table_data("Pivot")
            .unwrap()
            .unwrap()
            .cell((0, 0))
            .unwrap()
            .unwrap()
            .value_text(),
        "unrouted-patch"
    );
    let preserved_connections =
        String::from_utf8(workbook_member_blob(applied.workbook(), CONNECTIONS_URI)).unwrap();
    assert!(preserved_connections.contains(r#"refreshedVersion="8""#));
}

#[test]
fn authored_counts_coordinates_and_omissions_are_checked_without_dense_assumptions() {
    let omitted = Fixture::build(FixtureOptions {
        data_payload: r#"<x15:pivotTableData rowCount="2" columnCount="2" cacheId="7"><x15:pivotRow count="2"><x15:c t="str"><x15:v>alpha</x15:v><x15:x in="+0"/></x15:c><x15:c t="b"><x15:v>true</x15:v></x15:c></x15:pivotRow><x15:pivotRow count="2"><x15:c t="e"><x15:v>#N/A</x15:v></x15:c><x15:c t="bl"><x15:v/></x15:c></x15:pivotRow></x15:pivotTableData>"#.to_owned(),
        ..FixtureOptions::default()
    });
    let workbook = omitted.workbook();
    let view = workbook.pivot_table_data("Pivot").unwrap().unwrap();
    assert_eq!(view.rows()[0].row_index(), None);
    assert_eq!(view.rows()[0].cells()[0].column_index(), None);
    assert_eq!(view.cell((0, 0)).unwrap(), None);
    assert!(
        workbook
            .edit_pivot_table_data("Pivot")
            .unwrap()
            .set_value((0, 0), PivotCellValueEdit::text("unresolvable"),)
            .is_err()
    );

    assert_data_error(FixtureOptions {
        data_payload: default_data_payload_with("alpha").replace("r=\"1\"", "r=\"2\""),
        ..FixtureOptions::default()
    });
    assert_data_error(FixtureOptions {
        data_payload: default_data_payload_with("alpha").replace("i=\"1\"", "i=\"2\""),
        ..FixtureOptions::default()
    });
    assert_data_error(FixtureOptions {
        data_payload: default_data_payload_with("alpha").replace(
            "<x15:pivotRow count=\"2\" r=\"1\">",
            "<x15:pivotRow count=\"2\" r=\"0\">",
        ),
        ..FixtureOptions::default()
    });
    assert_data_error(FixtureOptions {
        data_payload: default_data_payload_with("alpha")
            .replace("<x15:c i=\"1\" t=\"b\">", "<x15:c i=\"0\" t=\"b\">"),
        ..FixtureOptions::default()
    });
    assert_data_error(FixtureOptions {
        data_payload: default_data_payload_with("alpha")
            .replace("rowCount=\"2\"", "rowCount=\"1\""),
        ..FixtureOptions::default()
    });
    assert_data_error(FixtureOptions {
        data_payload: default_data_payload_with("alpha").replace(
            "<x15:pivotRow count=\"2\" r=\"0\">",
            "<x15:pivotRow count=\"1\" r=\"0\">",
        ),
        ..FixtureOptions::default()
    });
    assert_data_error(FixtureOptions {
        data_payload: default_data_payload_with("alpha")
            .replace("columnCount=\"2\"", "columnCount=\"3\""),
        ..FixtureOptions::default()
    });
}

#[test]
fn direct_ancestry_is_required_and_mce_owner_is_readable_but_not_editable() {
    let wrapped = Fixture::build(FixtureOptions {
        data_payload: r#"<x15:wrapper><x15:pivotTableData rowCount="2" columnCount="2" cacheId="7"><x15:pivotRow count="2" r="0"><x15:c i="0" t="str"><x15:v>alpha</x15:v></x15:c><x15:c i="1" t="b"><x15:v>true</x15:v></x15:c></x15:pivotRow><x15:pivotRow count="2" r="1"><x15:c i="0" t="e"><x15:v>#N/A</x15:v></x15:c><x15:c i="1" t="bl"><x15:v/></x15:c></x15:pivotRow></x15:pivotTableData></x15:wrapper>"#.to_owned(),
        ..FixtureOptions::default()
    });
    assert!(wrapped.workbook().pivot_table_data("Pivot").is_err());

    let mce = Fixture::build(FixtureOptions {
        data_payload: r#"<mc:AlternateContent><mc:Choice Requires="x15"><x15:pivotTableData rowCount="2" columnCount="2" cacheId="7"><x15:pivotRow count="2" r="0"><x15:c i="0" t="str"><x15:v>alpha</x15:v><x15:x in="+0"/></x15:c><x15:c i="1" t="b"><x15:v>true</x15:v></x15:c></x15:pivotRow><x15:pivotRow count="2" r="1"><x15:c i="0" t="e"><x15:v>#N/A</x15:v></x15:c><x15:c i="1" t="bl"><x15:v/></x15:c></x15:pivotRow></x15:pivotTableData></mc:Choice></mc:AlternateContent>"#.to_owned(),
        ..FixtureOptions::default()
    });
    let workbook = mce.workbook();
    let view = workbook.pivot_table_data("Pivot").unwrap().unwrap();
    assert_eq!(
        view.diagnostic_status(),
        PivotTableDataDiagnostic::MceAmbiguous
    );
    assert!(!view.is_editable());
    assert!(workbook.edit_pivot_table_data("Pivot").is_err());
}

#[test]
fn c444_value_cdata_is_readable_source_preserving_and_scalar_editable() {
    let fixture = Fixture::build(FixtureOptions {
        data_payload: FixtureOptions::default().data_payload.replace(
            "<x15:v>alpha</x15:v>",
            "<x15:v><![CDATA[alpha & beta]]></x15:v>",
        ),
        ..FixtureOptions::default()
    });
    let workbook = fixture.workbook();
    let view = workbook.pivot_table_data("Pivot").unwrap().unwrap();
    assert_eq!(
        view.cell((0, 0)).unwrap().unwrap().value_text(),
        "alpha & beta"
    );

    let before = table_blob(&fixture.package());
    let mut noop = workbook.edit_pivot_table_data("Pivot").unwrap();
    assert!(
        !noop
            .set_value((0, 0), PivotCellValueEdit::text("alpha & beta"))
            .unwrap()
    );
    let noop_commit = noop.commit().unwrap();
    assert!(!noop_commit.changed());
    assert_eq!(
        table_blob(&OpcPackage::from_bytes(&noop_commit.workbook().to_bytes().unwrap()).unwrap()),
        before
    );

    let mut edit = workbook.edit_pivot_table_data("Pivot").unwrap();
    edit.set_value((0, 0), PivotCellValueEdit::text("changed & value"))
        .unwrap();
    let commit = edit.commit().unwrap();
    let changed = commit
        .workbook()
        .pivot_table_data("Pivot")
        .unwrap()
        .unwrap();
    assert_eq!(
        changed.cell((0, 0)).unwrap().unwrap().value_text(),
        "changed & value"
    );
}

#[test]
fn cache_ids_accept_xml_unsigned_lexical_forms_and_preserve_the_source() {
    for (lexical, semantic) in [(" \n+7\t", 7), ("-0", 0)] {
        let fixture = Fixture::build(FixtureOptions {
            cache_id_lexical: lexical,
            cache_definition_id: Some(lexical),
            ..FixtureOptions::default()
        });
        let workbook = fixture.workbook();
        let view = workbook.pivot_table_data("Pivot").unwrap().unwrap();
        assert_eq!(view.cache_id(), PivotCacheId::from(semantic));
        assert_eq!(view.cell((0, 0)).unwrap().unwrap().value_text(), "alpha");

        let output = String::from_utf8(workbook_blob(&workbook)).unwrap();
        assert!(output.contains(&format!(r#"cacheId="{lexical}""#)));
    }
}

#[test]
fn cache_725_identity_gaps_are_readable_diagnostic_and_read_only() {
    for options in [
        FixtureOptions {
            cache_definition_id: None,
            ..FixtureOptions::default()
        },
        FixtureOptions {
            duplicate_cache_definition: true,
            ..FixtureOptions::default()
        },
        FixtureOptions {
            cache_definition_id: Some("8"),
            ..FixtureOptions::default()
        },
    ] {
        let workbook = Fixture::build(options).workbook();
        let view = workbook.pivot_table_data("Pivot").unwrap().unwrap();
        assert_eq!(
            view.diagnostic_status(),
            PivotTableDataDiagnostic::CacheClosureUnresolved
        );
        assert!(!view.is_editable());
        assert!(workbook.edit_pivot_table_data("Pivot").is_err());
    }
}

#[test]
fn duplicate_known_workbook_owner_is_readable_diagnostic_and_read_only() {
    let workbook = Fixture::build(FixtureOptions {
        duplicate_workbook_reference: true,
        ..FixtureOptions::default()
    })
    .workbook();
    let view = workbook.pivot_table_data("Pivot").unwrap().unwrap();
    assert_eq!(
        view.diagnostic_status(),
        PivotTableDataDiagnostic::CacheClosureUnresolved
    );
    assert!(!view.is_editable());
    assert!(workbook.edit_pivot_table_data("Pivot").is_err());
}

#[test]
fn malformed_f057_and_c983_closure_shapes_refuse_or_become_diagnostic_read_only() {
    let mut f057_unknown_attribute = Fixture::build(FixtureOptions::default());
    let f057_owner = format!(r#"<ext uri="{F057_URI}">"#);
    let f057_owner_with_unknown_attribute = format!(r#"<ext uri="{F057_URI}" opaque="yes">"#);
    replace_fixture_part_xml(
        &mut f057_unknown_attribute,
        CACHE_URI,
        &f057_owner,
        &f057_owner_with_unknown_attribute,
    );

    let mut f057_unknown_child = Fixture::build(FixtureOptions::default());
    let f057_shape =
        format!(r#"<ext uri="{F057_URI}"><x14:sourceConnection name="canonical"/></ext>"#);
    let f057_shape_with_unknown_child = format!(
        r#"<ext uri="{F057_URI}"><x14:sourceConnection name="canonical"/><x14:opaque/></ext>"#
    );
    replace_fixture_part_xml(
        &mut f057_unknown_child,
        CACHE_URI,
        &f057_shape,
        &f057_shape_with_unknown_child,
    );

    let mut c983_unknown_child = Fixture::build(FixtureOptions::default());
    replace_fixture_part_xml(
        &mut c983_unknown_child,
        WORKBOOK_URI,
        r#"<x15:pivotTableReference r:id="rIdPivot"/></x15:pivotTableReferences>"#,
        r#"<x15:pivotTableReference r:id="rIdPivot"/><x15:opaque/></x15:pivotTableReferences>"#,
    );

    let mut c983_unknown_reference_attribute = Fixture::build(FixtureOptions::default());
    replace_fixture_part_xml(
        &mut c983_unknown_reference_attribute,
        WORKBOOK_URI,
        r#"<x15:pivotTableReference r:id="rIdPivot"/>"#,
        r#"<x15:pivotTableReference r:id="rIdPivot" opaque="yes"/>"#,
    );

    for fixture in [
        f057_unknown_attribute,
        f057_unknown_child,
        c983_unknown_child,
        c983_unknown_reference_attribute,
    ] {
        assert_refused_or_diagnostic_read_only(&fixture.workbook());
    }
}

#[test]
fn nonwhitespace_text_or_cdata_in_recognized_c444_particles_is_refused() {
    let markers = [
        format!(r#"<ext uri="{C444_URI}">"#),
        r#"<x15:pivotTableData rowCount="2" columnCount="2" cacheId="7">"#.to_owned(),
        r#"<x15:pivotRow count="2" r="0">"#.to_owned(),
        r#"<x15:c i="0" t="str">"#.to_owned(),
    ];
    for marker in markers {
        for content in [
            "nonwhitespace",
            "<![CDATA[nonwhitespace]]>",
            "<![CDATA[&#x20;]]>",
        ] {
            let mut fixture = Fixture::build(FixtureOptions::default());
            let replacement = format!("{marker}{content}");
            replace_fixture_part_xml(&mut fixture, TABLE_URI, &marker, &replacement);
            assert_refused_or_diagnostic_read_only(&fixture.workbook());
        }
    }
}

#[test]
fn nonwhitespace_text_or_cdata_in_empty_c983_reference_is_refused() {
    for content in ["nonwhitespace", "<![CDATA[nonwhitespace]]>"] {
        let mut fixture = Fixture::build(FixtureOptions::default());
        let replacement = format!(
            r#"<x15:pivotTableReference r:id="rIdPivot">{content}</x15:pivotTableReference>"#
        );
        replace_fixture_part_xml(
            &mut fixture,
            WORKBOOK_URI,
            r#"<x15:pivotTableReference r:id="rIdPivot"/>"#,
            &replacement,
        );
        assert_refused_or_diagnostic_read_only(&fixture.workbook());
    }
}

#[test]
fn duplicate_c444_owner_and_payload_remain_diagnostic_read_only() {
    for fixture in [
        fixture_with_duplicate_c444_owner(),
        fixture_with_duplicate_c444_payload(),
    ] {
        assert_diagnostic_read_only(&fixture.workbook());
    }
}

#[test]
fn oversized_ignored_duplicate_c444_owner_or_payload_is_refused() {
    for fixture in [
        fixture_with_oversized_duplicate_c444_owner(),
        fixture_with_oversized_duplicate_c444_payload(),
    ] {
        assert!(
            fixture.workbook().pivot_table_data("Pivot").is_err(),
            "an oversized recognized duplicate must be refused before it is ignored"
        );
    }
}

#[test]
fn entire_registered_closure_extensions_under_mce_are_readable_diagnostic_and_read_only() {
    for options in [
        FixtureOptions {
            workbook_mce: true,
            ..FixtureOptions::default()
        },
        FixtureOptions {
            cache_mce: true,
            ..FixtureOptions::default()
        },
    ] {
        let workbook = Fixture::build(options).workbook();
        let view = workbook.pivot_table_data("Pivot").unwrap().unwrap();
        assert_eq!(
            view.diagnostic_status(),
            PivotTableDataDiagnostic::MceAmbiguous
        );
        assert!(!view.is_editable());
        assert!(workbook.edit_pivot_table_data("Pivot").is_err());
    }
}

#[test]
fn worksheet_pivot_name_duplicates_are_rejected_before_nonworksheet_selection() {
    let fixture = Fixture::build(FixtureOptions {
        worksheet_pivot_names: vec!["WorksheetDup".to_owned(), "WorksheetDup".to_owned()],
        ..FixtureOptions::default()
    });
    assert!(fixture.workbook().pivot_table_data("Pivot").is_err());
}

#[test]
fn ordinary_worksheet_table_name_collision_refuses_name_selector_but_position_remains_clear() {
    let workbook = Fixture::build(FixtureOptions {
        worksheet_table_name: Some("Pivot".to_owned()),
        ..FixtureOptions::default()
    })
    .workbook();
    assert!(workbook.pivot_table_data("Pivot").is_err());
    let view = workbook
        .pivot_table_data(PivotTableSelector::Position(0))
        .unwrap()
        .unwrap();
    assert_eq!(view.table_name(), "Pivot");
}

#[test]
fn unknown_c444_siblings_survive_read_and_scalar_publication() {
    let fixture = Fixture::build(FixtureOptions {
        table_exts: r#"<x15:opaqueData keep="yes"><x15:unknown/></x15:opaqueData>"#.to_owned(),
        ..FixtureOptions::default()
    });
    let workbook = fixture.workbook();
    assert!(workbook.pivot_table_data("Pivot").unwrap().is_some());
    let mut edit = workbook.edit_pivot_table_data("Pivot").unwrap();
    edit.set_value((0, 0), PivotCellValueEdit::text("changed"))
        .unwrap();
    let commit = edit.commit().unwrap();
    let table = workbook_table_blob(commit.workbook());
    let text = String::from_utf8(table).unwrap();
    assert!(text.contains(r#"opaqueData keep="yes""#));
    assert!(text.contains("<x15:unknown/>"));
}

#[test]
fn local_v_namespace_declarations_survive_scalar_edit_with_semantic_qname() {
    for default_namespace in [false, true] {
        let fixture = fixture_with_local_v_namespace(default_namespace);
        let workbook = fixture.workbook();
        let mut edit = workbook.edit_pivot_table_data("Pivot").unwrap();
        assert!(edit.set_text((0, 0), "namespace-edited").unwrap());
        let commit = edit.commit().unwrap();
        let view = commit
            .workbook()
            .pivot_table_data("Pivot")
            .unwrap()
            .unwrap();
        assert_eq!(
            view.cell((0, 0)).unwrap().unwrap().value_text(),
            "namespace-edited"
        );

        let table = String::from_utf8(workbook_table_blob(commit.workbook())).unwrap();
        let expected = if default_namespace {
            format!(r#"<v xmlns="{EXT_NS}">namespace-edited</v>"#)
        } else {
            format!(r#"<x15:v xmlns:x15="{EXT_NS}">namespace-edited</x15:v>"#)
        };
        assert!(
            table.contains(&expected),
            "local namespace declaration or semantic QName was lost: {table}"
        );
    }
}

#[test]
fn recognized_uri_tokens_trim_xml_schema_whitespace_without_rewriting_lexical_source() {
    let fixture = Fixture::build(FixtureOptions {
        c444_uri: " \n{44433962-1CF7-4059-B4EE-95C3D5FFCF73}&#x9; ",
        c510_uri: " \n{C510F80B-63DE-4267-81D5-13C33094786E}&#x9; ",
        c983_uri: " \n{983426D0-5260-488c-9760-48F4B6AC55F4}&#x9; ",
        cache_ext_uri: " \n{725AE2AE-9491-48BE-B2B4-4EB974FC3084}&#x9; ",
        abf5_uri: " \n{ABF5C744-AB39-4b91-8756-CFA1BBC848D5}&#x9; ",
        f057_uri: " \n{F057638F-6D5F-4E77-A914-E7F072B9BCA8}&#x9; ",
        ..FixtureOptions::default()
    });
    let source_package = fixture.package();
    let snapshot = load_pivot_table_data(&source_package, "Pivot").unwrap();
    let source = std::str::from_utf8(snapshot.source_xml()).unwrap();
    assert!(source.contains("uri=\" \n{44433962-1CF7-4059-B4EE-95C3D5FFCF73}&#x9; \""));
    assert!(source.contains("in=\"+0\""));

    let borrowed = Workbook::from_slice(&fixture.bytes).unwrap();
    let before = borrowed.to_bytes().unwrap();
    let mut edit = borrowed.edit_pivot_table_data("Pivot").unwrap();
    assert!(
        !edit
            .set_value((0, 0), PivotCellValueEdit::text("alpha"))
            .unwrap()
    );
    let committed = edit.commit().unwrap();
    assert!(!committed.changed());
    assert!(committed.patch().is_empty());
    assert_eq!(committed.workbook().to_bytes().unwrap(), before);
}

#[test]
fn missing_c444_is_none_and_c510_index_routes_are_diagnostic_read_only() {
    let missing = Fixture::build(FixtureOptions {
        c444_uri: "{11111111-1111-1111-1111-111111111111}",
        ..FixtureOptions::default()
    });
    assert_eq!(missing.workbook().pivot_table_data("Pivot").unwrap(), None);

    let missing_c510 = Fixture::build(FixtureOptions {
        c510_uri: "{22222222-2222-2222-2222-222222222222}",
        ..FixtureOptions::default()
    });
    let workbook = missing_c510.workbook();
    let view = workbook.pivot_table_data("Pivot").unwrap().unwrap();
    assert_eq!(
        view.diagnostic_status(),
        PivotTableDataDiagnostic::ServerFormatIndexUnresolved
    );
    assert!(!view.is_editable());
    assert!(workbook.edit_pivot_table_data("Pivot").is_err());

    let count_mismatch = Fixture::build(FixtureOptions {
        server_count: "3",
        ..FixtureOptions::default()
    });
    let view = count_mismatch
        .workbook()
        .pivot_table_data("Pivot")
        .unwrap()
        .unwrap();
    assert_eq!(
        view.diagnostic_status(),
        PivotTableDataDiagnostic::ServerFormatIndexUnresolved
    );

    let out_of_range_index = Fixture::build(FixtureOptions {
        data_payload: default_data_payload_with("alpha").replace("in=\"+0\"", "in=\"2\""),
        ..FixtureOptions::default()
    });
    let view = out_of_range_index
        .workbook()
        .pivot_table_data("Pivot")
        .unwrap()
        .unwrap();
    assert_eq!(
        view.diagnostic_status(),
        PivotTableDataDiagnostic::ServerFormatIndexUnresolved
    );

    let negative_zero = Fixture::build(FixtureOptions {
        data_payload: default_data_payload_with("alpha").replace("in=\"+0\"", "in=\"-0\""),
        ..FixtureOptions::default()
    });
    let view = negative_zero
        .workbook()
        .pivot_table_data("Pivot")
        .unwrap()
        .unwrap();
    assert_eq!(view.diagnostic_status(), PivotTableDataDiagnostic::None);
    assert_eq!(
        view.cell((0, 0))
            .unwrap()
            .unwrap()
            .extra()
            .unwrap()
            .format_index,
        Some(0)
    );
    assert!(view.is_editable());
}

#[test]
fn existing_x_allows_safe_insertion_of_absent_extra_attributes() {
    let fixture = Fixture::build(FixtureOptions {
        data_payload: default_data_payload_with("alpha").replace(
            "<x15:x in=\"+0\" bc=\"0000000A\" fc=\"000000FF\" i=\"1\" un=\"false\" st=\"0\" b=\"true\"/>",
            "<x15:x in=\"+0\"/>",
        ),
        ..FixtureOptions::default()
    });
    let workbook = fixture.workbook();
    let mut edit = workbook.edit_pivot_table_data("Pivot").unwrap();
    assert!(
        edit.set_extra(
            (0, 0),
            PivotValueCellExtraEdit {
                background_color: PivotValueAttributeEdit::Set("0000000A".to_owned()),
                foreground_color: PivotValueAttributeEdit::Set("000000FF".to_owned()),
                italic: PivotValueAttributeEdit::Set(true),
                underline: PivotValueAttributeEdit::Set(false),
                strike: PivotValueAttributeEdit::Set(false),
                bold: PivotValueAttributeEdit::Set(true),
            },
        )
        .unwrap()
    );
    let commit = edit.commit().unwrap();
    let view = commit
        .workbook()
        .pivot_table_data("Pivot")
        .unwrap()
        .unwrap();
    let extra = view.cell((0, 0)).unwrap().unwrap().extra().unwrap();
    assert_eq!(extra.background_color, Some(0x0A));
    assert_eq!(extra.foreground_color, Some(0x00FF));
    assert_eq!(extra.italic, Some(true));
    assert_eq!(extra.underline, Some(false));
    assert_eq!(extra.strike, Some(false));
    assert_eq!(extra.bold, Some(true));
    let table = String::from_utf8(workbook_table_blob(commit.workbook())).unwrap();
    assert!(table.contains(r#"bc="0000000A""#));
    assert!(table.contains(r#"fc="000000FF""#));
}

#[test]
fn st_unsigned_int_hex_requires_exactly_eight_digits() {
    for (from, to) in [
        (r#"bc="0000000A""#, r#"bc="000000A""#),
        (r#"bc="0000000A""#, r#"bc="00000000A""#),
        (r#"fc="000000FF""#, r#"fc="000000F""#),
        (r#"fc="000000FF""#, r#"fc="0000000FF""#),
    ] {
        let mut fixture = Fixture::build(FixtureOptions::default());
        replace_fixture_part_xml(&mut fixture, TABLE_URI, from, to);
        assert_refused_or_diagnostic_read_only(&fixture.workbook());
    }
}

#[test]
fn public_scalar_edits_cover_text_boolean_error_blank_and_existing_extra_attributes() {
    let fixture = Fixture::build(FixtureOptions {
        table_exts: r#"<x15:opaqueData keep="yes"/>"#.to_owned(),
        ..FixtureOptions::default()
    });
    let original = fixture.workbook();
    let original_bytes = original.to_bytes().unwrap();
    let original_table = workbook_table_blob(&original);
    let mut edit = original.edit_pivot_table_data("Pivot").unwrap();
    assert!(
        edit.set_value((0, 0), PivotCellValueEdit::text("beta & 😀"))
            .unwrap()
    );
    assert!(
        edit.set_value((0, 1), PivotCellValueEdit::Boolean(false))
            .unwrap()
    );
    assert!(
        edit.set_value((1, 0), PivotCellValueEdit::error("#DIV/0!"))
            .unwrap()
    );
    assert!(!edit.set_blank((1, 1)).unwrap());
    let extra = PivotValueCellExtraEdit {
        background_color: PivotValueAttributeEdit::Set("0000BEEF".to_owned()),
        foreground_color: PivotValueAttributeEdit::Clear,
        italic: PivotValueAttributeEdit::Clear,
        underline: PivotValueAttributeEdit::Set(true),
        strike: PivotValueAttributeEdit::Keep,
        bold: PivotValueAttributeEdit::Set(false),
    };
    assert!(edit.set_extra((0, 0), extra).unwrap());
    let commit = edit.commit().unwrap();
    assert!(commit.changed());
    assert!(!commit.patch().is_empty());
    assert_eq!(original.to_bytes().unwrap(), original_bytes);

    let edited = commit.workbook();
    let view = edited.pivot_table_data("Pivot").unwrap().unwrap();
    assert_eq!(
        view.cell((0, 0)).unwrap().unwrap().value_text(),
        "beta & 😀"
    );
    assert_eq!(view.cell((0, 1)).unwrap().unwrap().value_text(), "false");
    assert_eq!(view.cell((1, 0)).unwrap().unwrap().value_text(), "#DIV/0!");
    let edited_extra = view.cell((0, 0)).unwrap().unwrap().extra().unwrap();
    assert_eq!(edited_extra.background_color, Some(0xBEEF));
    assert_eq!(edited_extra.foreground_color, None);
    assert_eq!(edited_extra.italic, None);
    assert_eq!(edited_extra.underline, Some(true));
    assert_eq!(edited_extra.strike, Some(false));
    assert_eq!(edited_extra.bold, Some(false));

    let table_text = String::from_utf8(workbook_table_blob(edited)).unwrap();
    assert!(table_text.contains(r#"opaqueData keep="yes"/>"#));
    assert!(table_text.contains(r#"in="+0""#));
    assert!(table_text.contains(r#"bc="0000BEEF""#));
    assert_ne!(workbook_table_blob(edited), original_table);

    let mut invalid = edited.edit_pivot_table_data("Pivot").unwrap();
    assert!(
        invalid
            .set_value((0, 0), PivotCellValueEdit::Boolean(true))
            .is_err()
    );
    assert!(
        invalid
            .set_value((0, 1), PivotCellValueEdit::text("wrong kind"))
            .is_err()
    );
    assert!(
        invalid
            .set_value((1, 0), PivotCellValueEdit::error("#REF!"))
            .is_err()
    );
    assert!(
        invalid
            .set_extra(
                (0, 0),
                PivotValueCellExtraEdit {
                    background_color: PivotValueAttributeEdit::Set("not-hex".to_owned()),
                    ..PivotValueCellExtraEdit::keep()
                },
            )
            .is_err()
    );
    assert!(
        invalid
            .set_value((9, 9), PivotCellValueEdit::Blank)
            .is_err()
    );
}

#[test]
fn numeric_and_datetime_cells_are_readable_but_scalar_editing_is_explicitly_unsupported() {
    let fixture = Fixture::build(FixtureOptions {
        data_payload: r#"<x15:pivotTableData rowCount="2" columnCount="2" cacheId="7"><x15:pivotRow count="2" r="0"><x15:c i="0" t="n"><x15:v>12.5</x15:v></x15:c><x15:c i="1" t="d"><x15:v>2026-09-12T00:00:00</x15:v></x15:c></x15:pivotRow><x15:pivotRow count="2" r="1"><x15:c i="0" t="e"><x15:v>#N/A</x15:v></x15:c><x15:c i="1" t="bl"><x15:v/></x15:c></x15:pivotRow></x15:pivotTableData>"#.to_owned(),
        ..FixtureOptions::default()
    });
    let workbook = fixture.workbook();
    let view = workbook.pivot_table_data("Pivot").unwrap().unwrap();
    assert_eq!(
        view.cell((0, 0)).unwrap().unwrap().kind(),
        PivotCellType::Number
    );
    assert_eq!(
        view.cell((0, 1)).unwrap().unwrap().kind(),
        PivotCellType::DateTime
    );
    let mut edit = workbook.edit_pivot_table_data("Pivot").unwrap();
    assert!(matches!(
        edit.set_value((0, 0), PivotCellValueEdit::text("13")),
        Err(Error::Unsupported { .. })
    ));
    assert!(matches!(
        edit.set_value((0, 1), PivotCellValueEdit::Blank),
        Err(Error::Unsupported { .. })
    ));
}

#[test]
fn workbook_patch_is_fresh_source_bound_stale_atomic_and_exactly_reversible() {
    let fixture = Fixture::build(FixtureOptions::default());
    let base = fixture.workbook();
    let base_bytes = base.to_bytes().unwrap();
    let mut edit = base.edit_pivot_table_data("Pivot").unwrap();
    edit.set_value((0, 0), PivotCellValueEdit::text("changed"))
        .unwrap();
    let committed = edit.commit().unwrap();
    assert!(committed.changed());
    let patch = committed.patch().clone();

    let fresh = Fixture::build(FixtureOptions::default()).workbook();
    let forwarded = fresh.apply_pivot_table_data_patch(&patch).unwrap();
    assert!(forwarded.changed());
    assert_eq!(
        forwarded.workbook().to_bytes().unwrap(),
        committed.workbook().to_bytes().unwrap()
    );

    let restored = committed
        .workbook()
        .apply_pivot_table_data_patch(&patch.inverse())
        .unwrap();
    assert_eq!(restored.workbook().to_bytes().unwrap(), base_bytes);

    let mut stale_package = fixture.package();
    let stale_connections = String::from_utf8(
        stale_package
            .get_part(&PackURI::new(CONNECTIONS_URI).unwrap())
            .unwrap()
            .blob()
            .to_vec(),
    )
    .unwrap()
    .replace("refreshedVersion=\"7\"", "refreshedVersion=\"8\"");
    stale_package
        .get_part_mut(&PackURI::new(CONNECTIONS_URI).unwrap())
        .unwrap()
        .set_blob(stale_connections.into_bytes());
    let stale = Workbook::from_bytes(PackageWriter::to_bytes(&stale_package).unwrap()).unwrap();
    let stale_before = stale.to_bytes().unwrap();
    assert!(stale.apply_pivot_table_data_patch(&patch).is_err());
    assert_eq!(stale.to_bytes().unwrap(), stale_before);
}

#[test]
fn workbook_patch_apply_binds_to_reopened_source_and_inverse() {
    let fixture = Fixture::build(FixtureOptions::default());
    let source = fixture.workbook();
    let source_bytes = source.to_bytes().unwrap();
    let mut edit = source.edit_pivot_table_data("Pivot").unwrap();
    assert!(edit.set_text((0, 0), "reopened-patch").unwrap());
    let committed = edit.commit().unwrap();
    let patch = committed.patch().clone();

    let mut candidate_package = OpcPackage::from_bytes(&source_bytes).unwrap();
    candidate_package
        .try_add_part(Box::new(BlobPart::new(
            PackURI::new("/xl/opaque.bin").unwrap(),
            "application/octet-stream".to_owned(),
            b"candidate-opaque".to_vec(),
        )))
        .unwrap();
    let reopened =
        Workbook::from_bytes(PackageWriter::to_bytes(&candidate_package).unwrap()).unwrap();
    let applied = patch.apply(&reopened).unwrap();
    assert!(applied.changed());
    assert_eq!(
        applied
            .workbook()
            .pivot_table_data("Pivot")
            .unwrap()
            .unwrap()
            .cell((0, 0))
            .unwrap()
            .unwrap()
            .value_text(),
        "reopened-patch"
    );
    assert_eq!(
        workbook_member_blob(applied.workbook(), "/xl/opaque.bin"),
        b"candidate-opaque"
    );

    let restored = applied.patch().inverse().apply(applied.workbook()).unwrap();
    assert!(restored.changed());
    assert_eq!(
        restored
            .workbook()
            .pivot_table_data("Pivot")
            .unwrap()
            .unwrap()
            .cell((0, 0))
            .unwrap()
            .unwrap()
            .value_text(),
        "alpha"
    );
    assert_eq!(
        workbook_member_blob(restored.workbook(), "/xl/opaque.bin"),
        b"candidate-opaque"
    );
}

#[test]
fn signed_scalar_publication_refuses_before_mutating_source() {
    let fixture = Fixture::build(FixtureOptions {
        signed: true,
        ..FixtureOptions::default()
    });
    let workbook = fixture.workbook();
    let before = workbook.to_bytes().unwrap();
    let mut edit = workbook.edit_pivot_table_data("Pivot").unwrap();
    edit.set_value((0, 0), PivotCellValueEdit::text("signed change"))
        .unwrap();
    assert!(matches!(edit.commit(), Err(Error::Signed)));
    assert_eq!(workbook.to_bytes().unwrap(), before);
}

#[test]
fn st_xstring_utf16_decoding_and_reencoding_preserve_literal_escape_intent() {
    let fixture = Fixture::build(FixtureOptions {
        data_payload: default_data_payload_with("A_x005F_x0041_ &amp; 😀"),
        ..FixtureOptions::default()
    });
    let workbook = fixture.workbook();
    let view = workbook.pivot_table_data("Pivot").unwrap().unwrap();
    assert_eq!(
        view.cell((0, 0)).unwrap().unwrap().value_text(),
        "A_x0041_ & 😀"
    );

    let mut literal = workbook.edit_pivot_table_data("Pivot").unwrap();
    literal
        .set_value((0, 0), PivotCellValueEdit::text("_x0041_"))
        .unwrap();
    let literal_commit = literal.commit().unwrap();
    let literal_table = String::from_utf8(workbook_table_blob(literal_commit.workbook())).unwrap();
    assert!(literal_table.contains(">_x005F_x0041_</x15:v>"));

    let mut control = literal_commit
        .workbook()
        .edit_pivot_table_data("Pivot")
        .unwrap();
    control
        .set_value((0, 0), PivotCellValueEdit::text("\u{1}"))
        .unwrap();
    let control_commit = control.commit().unwrap();
    let control_table = String::from_utf8(workbook_table_blob(control_commit.workbook())).unwrap();
    assert!(control_table.contains(">_x0001_</x15:v>"));
    assert_eq!(
        control_commit
            .workbook()
            .pivot_table_data("Pivot")
            .unwrap()
            .unwrap()
            .cell((0, 0))
            .unwrap()
            .unwrap()
            .value_text(),
        "\u{1}"
    );
}

#[test]
fn utf16_text_limit_counts_code_units_and_rejects_over_limit_source() {
    let at_limit = Fixture::build(FixtureOptions {
        data_payload: default_data_payload_with(&"a".repeat(65_535)),
        ..FixtureOptions::default()
    });
    assert_eq!(
        at_limit
            .workbook()
            .pivot_table_data("Pivot")
            .unwrap()
            .unwrap()
            .cell((0, 0))
            .unwrap()
            .unwrap()
            .value_text()
            .encode_utf16()
            .count(),
        65_535
    );

    let supplementary = Fixture::build(FixtureOptions {
        data_payload: default_data_payload_with(&("😀".repeat(32_767) + "a")),
        ..FixtureOptions::default()
    });
    let generous_limits = ReadLimits::builder()
        .max_xml_attribute_bytes(256 * 1024)
        .unwrap()
        .build()
        .unwrap();
    assert_eq!(
        Workbook::from_bytes_with_limits(supplementary.bytes.clone(), generous_limits)
            .unwrap()
            .pivot_table_data("Pivot")
            .unwrap()
            .unwrap()
            .cell((0, 0))
            .unwrap()
            .unwrap()
            .value_text()
            .encode_utf16()
            .count(),
        65_535
    );

    let over_limit = Fixture::build(FixtureOptions {
        data_payload: default_data_payload_with(&"a".repeat(65_536)),
        ..FixtureOptions::default()
    });
    assert!(over_limit.workbook().pivot_table_data("Pivot").is_err());

    let over_supplementary = Fixture::build(FixtureOptions {
        data_payload: default_data_payload_with(&"😀".repeat(32_768)),
        ..FixtureOptions::default()
    });
    assert!(
        Workbook::from_bytes_with_limits(over_supplementary.bytes, generous_limits)
            .unwrap()
            .pivot_table_data("Pivot")
            .is_err()
    );
}

#[test]
fn borrowed_workbook_limit_constructor_accepts_valid_source_and_enforces_event_budget() {
    let fixture = Fixture::build(FixtureOptions::default());
    let generous = ReadLimits::builder()
        .max_xml_attribute_bytes(256 * 1024)
        .unwrap()
        .build()
        .unwrap();
    let workbook = Workbook::from_slice_with_limits(&fixture.bytes, generous).unwrap();
    assert_eq!(
        workbook
            .pivot_table_data("Pivot")
            .unwrap()
            .unwrap()
            .cell((0, 0))
            .unwrap()
            .unwrap()
            .value_text(),
        "alpha"
    );

    let event_limits = ReadLimits::builder()
        .max_xml_events(8)
        .unwrap()
        .build()
        .unwrap();
    assert!(Workbook::from_slice_with_limits(&fixture.bytes, event_limits).is_err());
}

#[test]
fn borrowed_low_level_scalar_setters_preflight_utf16_and_preserve_source() {
    let fixture = Fixture::build(FixtureOptions {
        table_exts: r#"<x15:opaqueData keep="yes"><x15:unknown/></x15:opaqueData>"#.to_owned(),
        ..FixtureOptions::default()
    });
    let generous = ReadLimits::builder()
        .max_xml_attribute_bytes(256 * 1024)
        .unwrap()
        .build()
        .unwrap();
    let mut package = OpcPackage::from_bytes_with_limits(&fixture.bytes, generous).unwrap();
    let original = table_blob(&package);

    let mut noop = edit_pivot_table_data(&mut package, "Pivot").unwrap();
    assert!(!noop.set_text((0, 0), "alpha").unwrap());
    let noop_commit = noop.commit().unwrap();
    assert!(!noop_commit.changed());
    assert_eq!(table_blob(&package), original);

    let at_limit = format!("{}a", "😀".repeat(32_767));
    assert_eq!(at_limit.encode_utf16().count(), 65_535);
    let mut edit = edit_pivot_table_data(&mut package, "Pivot").unwrap();
    assert!(edit.set_text((0, 0), at_limit.as_str()).unwrap());
    assert!(edit.set_error((1, 0), "#DIV/0!").unwrap());
    let commit = edit.commit().unwrap();
    assert!(commit.changed());
    let table = String::from_utf8(table_blob(&package)).unwrap();
    assert!(table.contains(r#"opaqueData keep="yes""#));
    assert!(table.contains("<x15:unknown/>") || table.contains("<x15:unknown />"));
    assert!(table.contains(r#"in="+0""#));
    let snapshot = load_pivot_table_data(&package, "Pivot").unwrap();
    assert_eq!(
        snapshot
            .cell((0, 0))
            .unwrap()
            .unwrap()
            .value_text()
            .encode_utf16()
            .count(),
        65_535
    );
    assert_eq!(
        snapshot.cell((1, 0)).unwrap().unwrap().value_text(),
        "#DIV/0!"
    );

    let mut over_package = OpcPackage::from_bytes_with_limits(&fixture.bytes, generous).unwrap();
    let over_before = table_blob(&over_package);
    let over = "😀".repeat(32_768);
    assert_eq!(over.encode_utf16().count(), 65_536);
    let mut over_edit = edit_pivot_table_data(&mut over_package, "Pivot").unwrap();
    assert!(over_edit.set_text((0, 0), over.as_str()).is_err());
    assert!(over_edit.set_error((1, 0), "#REF!").is_err());
    assert!(!over_edit.is_changed());
    assert_eq!(table_blob(&over_package), over_before);
}

#[test]
fn borrowed_workbook_scalar_setters_preflight_limits_are_atomic_and_preserve_source() {
    let fixture = Fixture::build(FixtureOptions {
        table_exts: r#"<x15:opaqueData keep="yes"><x15:unknown/></x15:opaqueData>"#.to_owned(),
        ..FixtureOptions::default()
    });
    let attribute_limited = ReadLimits::builder()
        .max_xml_attribute_bytes(4_096)
        .unwrap()
        .build()
        .unwrap();
    let workbook = Workbook::from_slice_with_limits(&fixture.bytes, attribute_limited).unwrap();
    let original = workbook.to_bytes().unwrap();
    let too_wide = "z".repeat(5_000);
    let mut edit = workbook.edit_pivot_table_data("Pivot").unwrap();
    assert!(edit.set_text((0, 0), too_wide.as_str()).is_err());
    assert!(!edit.is_changed());
    assert!(edit.set_text((0, 0), "borrowed").unwrap());
    assert!(edit.set_error((1, 0), "#VALUE!").unwrap());
    let commit = edit.commit().unwrap();
    assert!(commit.changed());
    assert_eq!(workbook.to_bytes().unwrap(), original);
    let edited = commit.workbook();
    let view = edited.pivot_table_data("Pivot").unwrap().unwrap();
    assert_eq!(view.cell((0, 0)).unwrap().unwrap().value_text(), "borrowed");
    assert_eq!(view.cell((1, 0)).unwrap().unwrap().value_text(), "#VALUE!");
    let table = String::from_utf8(workbook_table_blob(edited)).unwrap();
    assert!(table.contains(r#"opaqueData keep="yes""#));
    assert!(table.contains(r#"in="+0""#));

    let part_limited = ReadLimits::builder()
        .max_part_bytes(fixture.table_xml.len() as u64)
        .unwrap()
        .build()
        .unwrap();
    let constrained =
        Workbook::from_bytes_with_limits(fixture.bytes.clone(), part_limited).unwrap();
    let before_failed_commit = constrained.to_bytes().unwrap();
    let mut failing = constrained.edit_pivot_table_data("Pivot").unwrap();
    let replacement = "q".repeat(4_096);
    assert!(failing.set_text((0, 0), replacement.as_str()).unwrap());
    assert!(failing.commit().is_err());
    assert_eq!(constrained.to_bytes().unwrap(), before_failed_commit);
}

#[test]
fn lowered_limits_charge_authored_rows_cells_and_dimensions_independently() {
    let workbook = Fixture::build(FixtureOptions::default()).workbook();
    for (scope, limits) in [
        (
            "pivotTableData authored rows",
            PivotTableDataLimits::new().with_max_authored_rows(1),
        ),
        (
            "pivotTableData authored cells",
            PivotTableDataLimits::new().with_max_authored_cells(3),
        ),
        (
            "pivotTableData rowCount",
            PivotTableDataLimits::new().with_max_row_count(1),
        ),
        (
            "pivotTableData columnCount",
            PivotTableDataLimits::new().with_max_column_count(1),
        ),
    ] {
        let error = workbook
            .pivot_table_data_with_limits("Pivot", &limits)
            .unwrap_err();
        assert_resource_limit(&error, scope);
    }

    let view = workbook
        .pivot_table_data_with_limits("Pivot", &PivotTableDataLimits::new())
        .unwrap()
        .unwrap();
    assert_eq!(view.rows().len(), 2);
    assert_eq!(view.rows()[0].cells().len(), 2);
}

#[test]
fn retained_byte_limit_has_exact_and_one_under_boundaries_and_failure_is_atomic() {
    let fixture = Fixture::build(FixtureOptions::default());
    let package = fixture.package();
    let required = minimum_retained_limit(&package);
    assert!(required > 0);

    let exact_limits = PivotTableDataLimits::new().with_max_retained_bytes(required);
    let snapshot = load_pivot_table_data_with_limits(&package, "Pivot", &exact_limits).unwrap();
    assert_eq!(snapshot.limits().max_retained_bytes(), required);
    let expected_data = snapshot.data().clone();
    let expected_source = snapshot.source_xml().to_vec();
    let expected_owner = snapshot.source_owner_range();
    let cloned = snapshot.clone();
    drop(snapshot);
    assert_eq!(cloned.data(), &expected_data);
    assert_eq!(cloned.source_xml(), expected_source.as_slice());
    assert_eq!(cloned.source_owner_range(), expected_owner);
    assert_eq!(cloned.limits(), exact_limits);

    let one_under = PivotTableDataLimits::new().with_max_retained_bytes(required - 1);
    let package_error =
        load_pivot_table_data_with_limits(&package, "Pivot", &one_under).unwrap_err();
    assert_resource_limit(&package_error, "pivotTableData");

    let workbook = fixture.workbook();
    let before = workbook.to_bytes().unwrap();
    let read_error = workbook
        .pivot_table_data_with_limits("Pivot", &one_under)
        .unwrap_err();
    assert_resource_limit(&read_error, "pivotTableData");
    assert_eq!(workbook.to_bytes().unwrap(), before);

    let transaction_error = match workbook.edit_pivot_table_data_with_limits("Pivot", &exact_limits)
    {
        Err(error) => error,
        Ok(_) => panic!("exact retained limit unexpectedly admitted staged transaction"),
    };
    assert_resource_limit(&transaction_error, "Workbook pivotTableData staged state");
    assert_eq!(workbook.to_bytes().unwrap(), before);
}

#[test]
fn exact_wire_part_cap_does_not_become_semantic_retained_cap() {
    let fixture = Fixture::build(FixtureOptions::default());
    let wire_limits = ReadLimits::builder()
        .max_part_bytes(fixture.table_xml.len() as u64)
        .unwrap()
        .build()
        .unwrap();
    assert_eq!(wire_limits.max_part_bytes(), fixture.table_xml.len() as u64);

    let workbook = Workbook::from_bytes_with_limits(fixture.bytes.clone(), wire_limits).unwrap();
    let view = workbook
        .pivot_table_data_with_limits("Pivot", &PivotTableDataLimits::new())
        .unwrap()
        .unwrap();
    assert_eq!(view.cell((0, 0)).unwrap().unwrap().value_text(), "alpha");

    let package = OpcPackage::from_bytes_with_limits(&fixture.bytes, wire_limits).unwrap();
    let snapshot =
        load_pivot_table_data_with_limits(&package, "Pivot", &PivotTableDataLimits::new()).unwrap();
    assert_eq!(
        snapshot.cell((0, 0)).unwrap().unwrap().value_text(),
        "alpha"
    );
}

#[test]
fn borrowed_workbook_limits_noop_preserves_source_and_snapshot_clone_is_observable() {
    let fixture = Fixture::build(FixtureOptions::default());
    let workbook = Workbook::from_slice_with_limits(&fixture.bytes, ReadLimits::default()).unwrap();
    let limits = PivotTableDataLimits::new();
    let before = workbook.to_bytes().unwrap();
    let mut edit = workbook
        .edit_pivot_table_data_with_limits("Pivot", &limits)
        .unwrap();
    assert!(!edit.set_text((0, 0), "alpha").unwrap());
    let commit = edit.commit().unwrap();
    assert!(!commit.changed());
    assert!(commit.patch().is_empty());
    assert_eq!(workbook.to_bytes().unwrap(), before);
    assert_eq!(commit.workbook().to_bytes().unwrap(), before);
    assert_eq!(
        workbook
            .pivot_table_data_with_limits("Pivot", &limits)
            .unwrap()
            .unwrap()
            .cell((0, 0))
            .unwrap()
            .unwrap()
            .value_text(),
        "alpha"
    );
}

#[test]
fn lowered_retained_budget_rejects_borrowed_setter_before_staging_mutation() {
    let fixture = Fixture::build(FixtureOptions::default());
    let replacement = "L".repeat(1_024);
    let transaction_required = minimum_transaction_limit(&fixture);
    let setter_required = minimum_setter_limit(&fixture, replacement.as_str());
    assert!(
        setter_required > transaction_required,
        "setter admission did not require retained bytes beyond transaction setup: setter={}, transaction={}",
        setter_required,
        transaction_required
    );

    let limits = PivotTableDataLimits::new().with_max_retained_bytes(setter_required - 1);
    let workbook = fixture.workbook();
    let before = workbook.to_bytes().unwrap();
    let mut edit = workbook
        .edit_pivot_table_data_with_limits("Pivot", &limits)
        .unwrap();
    let error = edit.set_text((0, 0), replacement.as_str()).unwrap_err();
    assert_resource_limit(&error, "pivotTableData");
    assert!(!edit.is_changed());

    let commit = edit.commit().unwrap();
    assert!(!commit.changed());
    assert!(commit.patch().is_empty());
    assert_eq!(workbook.to_bytes().unwrap(), before);
    assert_eq!(commit.workbook().to_bytes().unwrap(), before);
}

#[test]
fn low_level_lowered_retained_budget_rejects_borrowed_setter_atomically() {
    let fixture = Fixture::build(FixtureOptions::default());
    let replacement = "L".repeat(1_024);
    let transaction_required = minimum_low_level_transaction_limit(&fixture);
    let setter_required = minimum_low_level_setter_limit(&fixture, replacement.as_str());
    assert!(
        setter_required > transaction_required,
        "low-level setter admission did not require retained bytes beyond transaction setup: setter={}, transaction={}",
        setter_required,
        transaction_required
    );

    let limits = PivotTableDataLimits::new().with_max_retained_bytes(setter_required - 1);
    let mut package = fixture.package();
    let before = table_blob(&package);
    let mut edit = edit_pivot_table_data_with_limits(&mut package, "Pivot", &limits).unwrap();
    let error = edit.set_text((0, 0), replacement.as_str()).unwrap_err();
    assert_resource_limit(&error, "pivotTableData");
    assert!(!edit.is_changed());
    let commit = edit.commit().unwrap();
    assert!(!commit.changed());
    assert!(commit.patch().is_empty());
    assert_eq!(table_blob(&package), before);
}

#[test]
fn owned_scalar_capacity_is_charged_before_staging() {
    let fixture = Fixture::build(FixtureOptions::default());
    let logical_length = 32;
    let compact_capacity = logical_length;
    let inflated_capacity = 32 * 1024;
    let compact_required = minimum_owned_setter_limit(&fixture, logical_length, compact_capacity);
    let inflated_required = minimum_owned_setter_limit(&fixture, logical_length, inflated_capacity);
    assert!(
        inflated_required > compact_required,
        "owned String capacity was not admitted separately from text length: compact={}, inflated={}",
        compact_required,
        inflated_required
    );

    let limits = PivotTableDataLimits::new().with_max_retained_bytes(inflated_required - 1);
    {
        let workbook = fixture.workbook();
        let mut compact = workbook
            .edit_pivot_table_data_with_limits("Pivot", &limits)
            .unwrap();
        assert!(
            compact
                .set_value(
                    (0, 0),
                    owned_text_with_capacity(logical_length, compact_capacity)
                )
                .unwrap()
        );
    }

    let workbook = fixture.workbook();
    let before = workbook.to_bytes().unwrap();
    let mut inflated = workbook
        .edit_pivot_table_data_with_limits("Pivot", &limits)
        .unwrap();
    let error = match inflated.set_value(
        (0, 0),
        owned_text_with_capacity(logical_length, inflated_capacity),
    ) {
        Err(error) => error,
        Ok(_) => panic!("over-capacity owned String unexpectedly fit the retained budget"),
    };
    assert_resource_limit(&error, "pivotTableData");
    assert!(!inflated.is_changed());
    let commit = inflated.commit().unwrap();
    assert!(!commit.changed());
    assert!(commit.patch().is_empty());
    assert_eq!(workbook.to_bytes().unwrap(), before);
    assert_eq!(commit.workbook().to_bytes().unwrap(), before);
}

#[test]
fn empty_owned_scalar_replacement_is_charged_at_the_boundary() {
    let fixture = Fixture::build(FixtureOptions::default());
    let transaction_required = minimum_transaction_limit(&fixture);
    let empty_required = minimum_owned_setter_limit(&fixture, 0, 0);
    assert!(
        empty_required > transaction_required,
        "empty owned Arc replacement did not require retained admission beyond transaction setup: empty={}, transaction={}",
        empty_required,
        transaction_required
    );

    let limits = PivotTableDataLimits::new().with_max_retained_bytes(empty_required - 1);
    let workbook = fixture.workbook();
    let before = workbook.to_bytes().unwrap();
    let mut edit = workbook
        .edit_pivot_table_data_with_limits("Pivot", &limits)
        .unwrap();
    let error = match edit.set_value((0, 0), owned_text_with_capacity(0, 0)) {
        Err(error) => error,
        Ok(_) => panic!("empty owned replacement unexpectedly fit the retained budget"),
    };
    assert_resource_limit(&error, "pivotTableData");
    assert!(!edit.is_changed());
    let commit = edit.commit().unwrap();
    assert!(!commit.changed());
    assert!(commit.patch().is_empty());
    assert_eq!(workbook.to_bytes().unwrap(), before);
    assert_eq!(commit.workbook().to_bytes().unwrap(), before);
}

#[test]
fn repeated_scalar_replacements_use_net_retained_delta() {
    let fixture = Fixture::build(FixtureOptions::default());
    let first = "R".repeat(1_024);
    let second = "S".repeat(1_024);
    let oversized = "O".repeat(4_096);
    let setter_required = minimum_setter_limit(&fixture, first.as_str());
    let exact = minimum_repeated_replacement_limit(&fixture, first.as_str(), second.as_str());
    assert!(exact >= setter_required);
    assert!(exact > 0);
    assert!(!repeated_replacement_limit_admits(
        &fixture,
        first.as_str(),
        second.as_str(),
        exact - 1,
    ));
    assert!(
        !setter_limit_admits(&fixture, oversized.as_str(), exact),
        "oversized replacement unexpectedly fit the committed retained-byte budget"
    );
    let limits = PivotTableDataLimits::new().with_max_retained_bytes(exact);
    let workbook = fixture.workbook();
    let mut edit = workbook
        .edit_pivot_table_data_with_limits("Pivot", &limits)
        .unwrap();

    assert!(edit.set_text((0, 0), first.as_str()).unwrap());
    let error = edit.set_text((0, 0), oversized.as_str()).unwrap_err();
    assert_resource_limit(&error, "pivotTableData");
    assert!(edit.is_changed());
    assert!(edit.set_text((0, 0), "short").unwrap());
    assert!(edit.set_text((0, 0), second.as_str()).unwrap());
    for _ in 0..16 {
        assert!(edit.set_text((0, 0), "short").unwrap());
        assert!(edit.set_text((0, 0), second.as_str()).unwrap());
    }

    let commit = edit.commit().unwrap();
    assert!(commit.changed());
    let view = commit
        .workbook()
        .pivot_table_data("Pivot")
        .unwrap()
        .unwrap();
    assert_eq!(view.cell((0, 0)).unwrap().unwrap().value_text(), second);
}

#[test]
fn selected_part_relationship_closure_needs_semantic_retained_budget() {
    let baseline = Fixture::build(FixtureOptions::default());
    let related = fixture_with_large_selected_part_relationships();
    let baseline_required = minimum_retained_limit(&baseline.package());
    let related_package = related.package();
    let related_required = minimum_retained_limit(&related_package);
    assert!(
        related_required > baseline_required,
        "large selected-part relationship closure did not raise semantic budget: baseline={}, related={}",
        baseline_required,
        related_required
    );

    let exact = PivotTableDataLimits::new().with_max_retained_bytes(related_required);
    let snapshot = load_pivot_table_data_with_limits(&related_package, "Pivot", &exact).unwrap();
    assert_eq!(
        snapshot.cell((0, 0)).unwrap().unwrap().value_text(),
        "alpha"
    );

    let one_under = PivotTableDataLimits::new().with_max_retained_bytes(related_required - 1);
    let error =
        load_pivot_table_data_with_limits(&related_package, "Pivot", &one_under).unwrap_err();
    assert_resource_limit(&error, "pivotTableData");
}

#[test]
fn invalid_extra_edit_is_atomic_on_low_level_and_workbook_transactions() {
    let fixture = Fixture::build(FixtureOptions::default());
    let invalid_extra = || PivotValueCellExtraEdit {
        background_color: PivotValueAttributeEdit::Set("0000BEEF".to_owned()),
        foreground_color: PivotValueAttributeEdit::Set("not-hex".to_owned()),
        ..PivotValueCellExtraEdit::keep()
    };

    let mut package = fixture.package();
    let before_table = table_blob(&package);
    let mut low_level = edit_pivot_table_data(&mut package, "Pivot").unwrap();
    assert!(low_level.set_extra((0, 0), invalid_extra()).is_err());
    assert!(!low_level.is_changed());
    let low_commit = low_level.commit().unwrap();
    assert!(!low_commit.changed());
    assert!(low_commit.patch().is_empty());
    assert_eq!(table_blob(&package), before_table);

    let workbook = fixture.workbook();
    let before_workbook = workbook.to_bytes().unwrap();
    let mut workbook_edit = workbook.edit_pivot_table_data("Pivot").unwrap();
    assert!(workbook_edit.set_extra((0, 0), invalid_extra()).is_err());
    assert!(!workbook_edit.is_changed());
    let workbook_commit = workbook_edit.commit().unwrap();
    assert!(!workbook_commit.changed());
    assert!(workbook_commit.patch().is_empty());
    assert_eq!(workbook.to_bytes().unwrap(), before_workbook);
    assert_eq!(
        workbook_commit.workbook().to_bytes().unwrap(),
        before_workbook
    );
}

#[test]
fn low_level_package_ingress_refuses_an_encrypted_pivot_table_entry() {
    let fixture = Fixture::build(FixtureOptions::default());
    let encrypted = mark_zip_entry_encrypted(fixture.bytes, "xl/pivotTables/pivotTable1.xml");
    let result = OpcPackage::from_bytes(&encrypted);
    assert!(
        matches!(result, Err(litchi_opc::OpcError::ZipError(_))),
        "encrypted pivotTable Part was admitted by low-level OPC ingress"
    );
}

#[test]
fn caller_limits_reject_source_and_candidate_before_publication() {
    let fixture = Fixture::build(FixtureOptions::default());
    let event_limits = ReadLimits::builder()
        .max_xml_events(8)
        .unwrap()
        .build()
        .unwrap();
    assert!(Workbook::from_bytes_with_limits(fixture.bytes.clone(), event_limits).is_err());

    let attribute_limits = ReadLimits::builder()
        .max_xml_attribute_bytes(8)
        .unwrap()
        .max_relationship_target_bytes(8)
        .unwrap()
        .build()
        .unwrap();
    assert!(Workbook::from_bytes_with_limits(fixture.bytes.clone(), attribute_limits).is_err());

    let part_limits = ReadLimits::builder()
        .max_part_bytes(fixture.table_xml.len() as u64)
        .unwrap()
        .build()
        .unwrap();
    let workbook = Workbook::from_bytes_with_limits(fixture.bytes.clone(), part_limits).unwrap();
    let before = workbook.to_bytes().unwrap();
    let mut edit = workbook.edit_pivot_table_data("Pivot").unwrap();
    edit.set_value((0, 0), PivotCellValueEdit::text("z".repeat(4_096)))
        .unwrap();
    assert!(edit.commit().is_err());
    assert_eq!(workbook.to_bytes().unwrap(), before);
}
