//! The facade's `sheet::Workbook` in-source tests (`sheet/workbook.rs`),
//! ported onto `File`, `WorkbookView` and its `WorkbookTrait`.
//!
//! Each test keeps its facade semantics: the route a fixture takes, the
//! owner it ends in, and the result or typed refusal it reports. Facade
//! errors that were boxed or string-flattened are asserted through the typed
//! code and leaf instead.

use litchi_core::sheet::WorkbookTrait;

#[cfg(any(feature = "xlsx", feature = "xlsb", feature = "xls"))]
use super::test_data;
use crate::{Code, File};
#[cfg(feature = "xlsx")]
use litchi_core::sheet::Cell as _;
#[cfg(any(feature = "xlsx", feature = "xls"))]
use litchi_core::sheet::Worksheet as _;

#[cfg(any(feature = "xlsx", feature = "xls"))]
fn open(relative: &str) -> File {
    crate::open(test_data(relative)).unwrap_or_else(|error| panic!("open {relative}: {error}"))
}

/// Read a workbook through the generic trait only, as a formula evaluator
/// generic over `W: WorkbookTrait + ?Sized` does.
fn generic_names<W: WorkbookTrait + ?Sized>(workbook: &W) -> Vec<String> {
    assert_eq!(workbook.worksheet_names().len(), workbook.worksheet_count());
    workbook.worksheet_names().to_vec()
}

// ---- XLSX ---------------------------------------------------------------

#[cfg(feature = "xlsx")]
mod xlsx {
    use std::io::{Cursor, Write as _};

    use litchi_core::sheet::CellValue;

    use super::*;
    use crate::format::xlsx::Xlsx;
    use crate::{OpenOptions, SourceMode};

    const WORKSHEET: &str =
        "application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml";
    const WORKBOOK: &str =
        "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml";

    pub(super) fn package_with_second_kind(
        names: [&str; 2],
        second_sheet: &[u8],
        title: &str,
        second_content_type: &str,
        second_relationship_type: &str,
    ) -> Vec<u8> {
        let [first, second] = names;
        let mut writer = zip::ZipWriter::new(Cursor::new(Vec::new()));
        let options = zip::write::SimpleFileOptions::default()
            .compression_method(zip::CompressionMethod::Stored);
        writer.start_file("[Content_Types].xml", options).unwrap();
        writer
            .write_all(
                format!(
                    r#"<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="{WORKBOOK}"/><Override PartName="/xl/worksheets/sheet1.xml" ContentType="{WORKSHEET}"/><Override PartName="/xl/worksheets/sheet2.xml" ContentType="{second_content_type}"/><Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/></Types>"#
                )
                .as_bytes(),
            )
            .unwrap();
        writer.start_file("_rels/.rels", options).unwrap();
        writer
            .write_all(
                br#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/></Relationships>"#,
            )
            .unwrap();
        writer.start_file("docProps/core.xml", options).unwrap();
        writer
            .write_all(
                format!(
                    r#"<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/"><dc:title>{title}</dc:title></cp:coreProperties>"#
                )
                .as_bytes(),
            )
            .unwrap();
        writer.start_file("xl/workbook.xml", options).unwrap();
        writer
            .write_all(
                format!(
                    r#"<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><workbookPr date1904="1"/><sheets><sheet name="{first}" sheetId="1" r:id="rId1"/><sheet name="{second}" sheetId="2" state="hidden" r:id="rId2"/></sheets></workbook>"#
                )
                .as_bytes(),
            )
            .unwrap();
        writer
            .start_file("xl/_rels/workbook.xml.rels", options)
            .unwrap();
        writer
            .write_all(
                format!(
                    r#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/><Relationship Id="rId2" Type="{second_relationship_type}" Target="worksheets/sheet2.xml"/></Relationships>"#
                )
                .as_bytes(),
            )
            .unwrap();
        writer
            .start_file("xl/worksheets/sheet1.xml", options)
            .unwrap();
        writer
            .write_all(
                br#"<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData><row r="1"><c r="A1"><v>1</v></c></row></sheetData></worksheet>"#,
            )
            .unwrap();
        writer
            .start_file("xl/worksheets/sheet2.xml", options)
            .unwrap();
        writer.write_all(second_sheet).unwrap();
        writer.finish().unwrap().into_inner()
    }

    fn package(second_sheet: &[u8], title: &str) -> Vec<u8> {
        package_with_second_kind(
            ["First", "Second"],
            second_sheet,
            title,
            WORKSHEET,
            "http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet",
        )
    }

    const SECOND_SHEET: &[u8] = br#"<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData><row r="1"><c r="A1"><v>2</v></c></row></sheetData></worksheet>"#;
    const TRUNCATED_SHEET: &[u8] =
        br#"<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData><row>"#;

    fn valid_package(title: &str) -> Vec<u8> {
        package(SECOND_SHEET, title)
    }

    fn xlsx(file: &File) -> &Xlsx {
        match file {
            File::Xlsx(xlsx) => xlsx,
            other => panic!("expected XLSX, got {other:?}"),
        }
    }

    fn first_cell(file: &File) -> CellValue {
        let view = file.workbook().unwrap();
        let sheet = WorkbookTrait::worksheet_by_index(&view, 0).unwrap();
        sheet.cell_by_coordinate("A1").unwrap().value().clone()
    }

    /// The facade's `Workbook::sheet` existed for XLSX only and refused the
    /// other formats at run time; the selection lives on the XLSX handle.
    /// XLSX names match case-insensitively.
    #[test]
    fn xlsx_sheet_selection_is_case_insensitive() {
        let bytes = package_with_second_kind(
            ["Data", "Notes"],
            SECOND_SHEET,
            "selection",
            WORKSHEET,
            "http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet",
        );
        for file in [
            crate::open_bytes(bytes.clone()).unwrap(),
            crate::open_bytes_with(
                bytes.clone(),
                &OpenOptions::new().source_mode(SourceMode::Eager),
            )
            .unwrap(),
        ] {
            let File::Xlsx(handle) = &file else {
                panic!("expected XLSX, got {file:?}");
            };
            let data = handle
                .sheet("dAtA")
                .unwrap()
                .expect("case-insensitive match");
            assert_eq!(data.name(), "Data");
            assert_eq!(data.position(), 0);
            let notes = handle
                .sheet(1_usize)
                .unwrap()
                .expect("selection by position");
            assert_eq!(notes.name(), "Notes");
            assert!(handle.sheet("missing").unwrap().is_none());
            assert!(handle.sheet(2_usize).unwrap().is_none());
            // The generic trait matches the same way.
            let view = file.workbook().unwrap();
            assert!(WorkbookTrait::worksheet_by_name(&view, "DATA").is_ok());
        }
    }

    #[cfg(any(unix, windows))]
    #[test]
    fn filesystem_and_bytes_xlsx_use_source_owner_and_match_projection() {
        let bytes = valid_package("source title");
        // The file name is not consulted: an `.ods` name holds an XLSX.
        let temporary = super::super::temporary_with(&bytes, ".ods");
        let source = crate::open(temporary.path()).unwrap();
        let from_bytes = crate::open_bytes(bytes.clone()).unwrap();
        assert!(xlsx(&source).is_source_backed());
        assert!(xlsx(&from_bytes).is_source_backed());
        let (source_view, bytes_view) =
            (source.workbook().unwrap(), from_bytes.workbook().unwrap());
        assert_eq!(source_view.worksheet_names().unwrap(), ["First", "Second"]);
        assert_eq!(
            source_view.worksheet_names().unwrap(),
            bytes_view.worksheet_names().unwrap()
        );
        assert_eq!(
            source_view.worksheet_count().unwrap(),
            bytes_view.worksheet_count().unwrap()
        );
        assert_eq!(generic_names(&source_view), ["First", "Second"]);
        let eager = litchi_xlsx::Package::from_bytes(&bytes).unwrap();
        assert_eq!(
            eager
                .workbook_ref()
                .sheets()
                .map(|sheet| sheet.name().to_owned())
                .collect::<Vec<_>>(),
            ["First", "Second"]
        );
        assert_eq!(source.text().unwrap(), from_bytes.text().unwrap());
        assert_eq!(source.text().unwrap(), "1\n2\n");
        assert_eq!(
            source.metadata().unwrap().unwrap().title.as_deref(),
            Some("source title")
        );
        assert_eq!(
            source.metadata().unwrap().unwrap().title,
            from_bytes.metadata().unwrap().unwrap().title
        );
        assert!(WorkbookTrait::is_1904_date_system(&source_view));
        assert_eq!(first_cell(&source), CellValue::Int(1));
    }

    #[cfg(any(unix, windows))]
    #[test]
    fn filesystem_xlsx_defers_unselected_worksheet_payload() {
        let temporary = super::super::temporary_with(&package(TRUNCATED_SHEET, "deferred"), "");
        let file = crate::open(temporary.path()).expect("catalog-only source open");
        assert_eq!(
            file.workbook().unwrap().worksheet_names().unwrap(),
            ["First", "Second"]
        );
        assert_eq!(first_cell(&file), CellValue::Int(1));
        assert!(file.text().is_err());
    }

    #[test]
    fn bytes_xlsx_defers_unselected_worksheet_payload() {
        let file = crate::open_bytes(package(TRUNCATED_SHEET, "deferred bytes")).unwrap();
        assert!(xlsx(&file).is_source_backed());
        assert_eq!(
            file.workbook().unwrap().worksheet_names().unwrap(),
            ["First", "Second"]
        );
        assert_eq!(first_cell(&file), CellValue::Int(1));
        assert!(file.text().is_err());
    }

    #[cfg(any(unix, windows))]
    #[test]
    fn filesystem_xlsx_text_skips_non_grid_sheet_kinds() {
        let bytes = package_with_second_kind(
            ["First", "Second"],
            br#"<chartsheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"/>"#,
            "chart",
            "application/vnd.openxmlformats-officedocument.spreadsheetml.chartsheet+xml",
            "http://schemas.openxmlformats.org/officeDocument/2006/relationships/chartsheet",
        );
        let temporary = super::super::temporary_with(&bytes, "");
        let source = crate::open(temporary.path()).unwrap();
        let eager =
            crate::open_bytes_with(bytes, &OpenOptions::new().source_mode(SourceMode::Eager))
                .unwrap();
        assert!(!xlsx(&eager).is_source_backed());
        assert_eq!(
            source.workbook().unwrap().worksheet_names().unwrap(),
            ["First", "Second"]
        );
        assert_eq!(source.text().unwrap(), "1\n");
        assert_eq!(source.text().unwrap(), eager.text().unwrap());
    }

    #[cfg(any(unix, windows))]
    #[test]
    fn filesystem_xlsx_cached_queries_report_source_change() {
        let temporary = super::super::temporary_with(&valid_package("before"), "");
        let file = crate::open(temporary.path()).unwrap();
        std::fs::write(
            temporary.path(),
            valid_package("after with a different size"),
        )
        .unwrap();
        let view = file.workbook().unwrap();
        for error in [
            view.worksheet_names().map(drop).unwrap_err(),
            view.worksheet_count().map(drop).unwrap_err(),
            file.metadata().map(drop).unwrap_err(),
            file.text().map(drop).unwrap_err(),
        ] {
            assert_eq!(error.code(), Code::IoError, "{error:?}");
            assert!(super::super::source_change(&error).is_some(), "{error:?}");
        }
    }

    #[cfg(any(unix, windows))]
    #[test]
    fn filesystem_xlsx_worksheet_handles_preserve_typed_source_change_errors() {
        let temporary = super::super::temporary_with(&valid_package("before"), "");
        let file = crate::open(temporary.path()).unwrap();
        let view = file.workbook().unwrap();
        let first = WorkbookTrait::worksheet_by_index(&view, 0).unwrap();
        std::fs::write(
            temporary.path(),
            valid_package("after with a different size"),
        )
        .unwrap();

        let typed = |error: crate::Error| {
            assert_eq!(error.code(), Code::IoError, "{error:?}");
            assert!(super::super::source_change(&error).is_some(), "{error:?}");
        };
        typed(first.cell_by_coordinate("A1").map(drop).unwrap_err());
        typed(
            first
                .rows()
                .next()
                .expect("a stale row result")
                .unwrap_err(),
        );
        typed(
            first
                .cells()
                .next()
                .expect("a stale cell result")
                .map(drop)
                .unwrap_err(),
        );
    }

    #[cfg(any(unix, windows))]
    #[test]
    fn filesystem_xlsx_propagates_the_opc_input_limit_before_fallback() {
        let temporary = super::super::temporary_with(&valid_package("bounded"), "");
        let maximum = litchi_opc::SourceLimits::default().max_input_bytes();
        temporary.as_file().set_len(maximum + 1).unwrap();
        let error = crate::open(temporary.path()).expect_err("over the default OPC limit");
        assert_eq!(error.code(), Code::LimitExceeded);
        assert!(matches!(
            error.leaf(),
            crate::error::Source::Opc(litchi_opc::OpcError::Limit {
                kind: litchi_opc::SourceLimitKind::InputBytes,
                ..
            })
        ));
    }

    #[cfg(any(unix, windows))]
    #[test]
    fn filesystem_xlsx_preflights_tight_input_limit_before_reads() {
        let temporary = super::super::temporary_with(&valid_package("bounded sparse"), "");
        temporary.as_file().set_len(1024 * 1024 + 1).unwrap();
        let limits = litchi_opc::SourceLimits::builder()
            .max_input_bytes(1024)
            .unwrap()
            .build()
            .unwrap();
        let error = crate::open_with(temporary.path(), &OpenOptions::new().ooxml_limits(limits))
            .unwrap_err();
        assert!(matches!(
            error.leaf(),
            crate::error::Source::Opc(litchi_opc::OpcError::Limit {
                kind: litchi_opc::SourceLimitKind::InputBytes,
                observed: 1_048_577,
                maximum: 1024,
            })
        ));
    }

    #[test]
    fn bytes_xlsx_propagates_malformed_opc_probe_error() {
        let error = crate::open_bytes(b"PK\x03\x04malformed".to_vec()).unwrap_err();
        assert_eq!(error.code(), Code::CorruptFile, "{error:?}");
    }

    #[test]
    fn bytes_xlsx_matching_catalog_owner_error_stays_typed() {
        let mut writer = zip::ZipWriter::new(Cursor::new(Vec::new()));
        let options = zip::write::SimpleFileOptions::default()
            .compression_method(zip::CompressionMethod::Stored);
        writer.start_file("[Content_Types].xml", options).unwrap();
        writer
            .write_all(
                format!(
                    r#"<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Override PartName="/xl/workbook.xml" ContentType="{WORKBOOK}"/></Types>"#
                )
                .as_bytes(),
            )
            .unwrap();
        writer.start_file("_rels/.rels", options).unwrap();
        writer
            .write_all(
                br#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>"#,
            )
            .unwrap();
        writer.start_file("xl/workbook.xml", options).unwrap();
        writer.write_all(b"<not-a-workbook/>").unwrap();
        let error = crate::open_bytes(writer.finish().unwrap().into_inner()).unwrap_err();
        assert_eq!(error.code(), Code::CorruptFile, "{error:?}");
        assert_eq!(
            error.format(),
            Some(litchi_core::FileFormat::Xlsx),
            "{error:?}"
        );
    }

    #[test]
    fn test_workbook_xlsx_fixtures() {
        let path = test_data("ooxml/xlsx/DateFormatTests.xlsx");
        for file in [
            crate::open(&path).unwrap(),
            crate::open_bytes(std::fs::read(&path).unwrap()).unwrap(),
        ] {
            let view = file.workbook().unwrap();
            assert!(!view.worksheet_names().unwrap().is_empty());
            assert!(view.worksheet_count().unwrap() > 0);
            assert_eq!(generic_names(&view), view.worksheet_names().unwrap());
            file.text().unwrap();
            let _ = file
                .metadata()
                .unwrap()
                .map(|metadata| (metadata.title, metadata.author));
        }
        for relative in [
            "ooxml/xlsx/condFormat_cellis.xlsx",
            "ooxml/xlsx/autofilter.xlsx",
            "ooxml/xlsx/DataValidationEvaluations.xlsx",
        ] {
            assert!(
                open(relative)
                    .workbook()
                    .unwrap()
                    .worksheet_count()
                    .unwrap()
                    > 0,
                "{relative}"
            );
        }
    }

    #[test]
    fn test_workbook_open_nonexistent_and_invalid() {
        assert_eq!(
            crate::open(test_data("nonexistent_file.xlsx"))
                .unwrap_err()
                .code(),
            Code::FileNotFound
        );
        let error =
            crate::open_bytes(b"This is not a valid spreadsheet file".to_vec()).unwrap_err();
        assert_eq!(error.code(), Code::UnsupportedFormat);
        assert_eq!(error.detected(), None);
    }

    /// The facade's XLS source probe preserved XLSX precedence.
    #[cfg(feature = "xls")]
    #[test]
    fn xlsx_precedence_survives_the_xls_feature() {
        assert!(matches!(
            open("ooxml/xlsx/DateFormatTests.xlsx"),
            File::Xlsx(_)
        ));
    }
}

// ---- XLSB ---------------------------------------------------------------

#[cfg(feature = "xlsb")]
mod xlsb {
    use std::io::{Cursor, Write as _};

    use super::*;

    fn minimal_xlsx_marked_opc() -> Vec<u8> {
        let mut writer = zip::ZipWriter::new(Cursor::new(Vec::new()));
        let options = zip::write::SimpleFileOptions::default()
            .compression_method(zip::CompressionMethod::Stored);
        writer.start_file("[Content_Types].xml", options).unwrap();
        writer
            .write_all(
                br#"<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/></Types>"#,
            )
            .unwrap();
        writer.start_file("_rels/.rels", options).unwrap();
        writer
            .write_all(
                br#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>"#,
            )
            .unwrap();
        writer.start_file("xl/workbook.xml", options).unwrap();
        writer
            .write_all(
                br#"<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheets/></workbook>"#,
            )
            .unwrap();
        writer.finish().unwrap().into_inner()
    }

    /// `xlsb_dynamic_open_rejects_known_xlsx_without_reopening_path`: an
    /// XLSX package never opens as XLSB; content detection names it.
    #[test]
    fn an_xlsx_package_is_never_opened_as_xlsb() {
        let result = crate::open_bytes(minimal_xlsx_marked_opc());
        #[cfg(feature = "xlsx")]
        assert!(matches!(result, Ok(File::Xlsx(_))), "{result:?}");
        #[cfg(not(feature = "xlsx"))]
        {
            let error = result.expect_err("the XLSX owner is not compiled in");
            assert_eq!(error.code(), Code::UnsupportedFormat);
            assert_eq!(error.detected(), Some(litchi_core::FileFormat::Xlsx));
            assert!(error.suggestion().is_some_and(|text| text.contains("xlsx")));
        }
    }

    #[test]
    fn xlsb_reads_through_the_view_and_the_trait() {
        let path = test_data("ooxml/xlsb/Simple.xlsb");
        for file in [
            crate::open(&path).unwrap(),
            crate::open_bytes(std::fs::read(&path).unwrap()).unwrap(),
        ] {
            assert!(matches!(file, File::Xlsb(_)));
            let view = file.workbook().unwrap();
            let names = view.worksheet_names().unwrap();
            assert!(!names.is_empty());
            assert_eq!(generic_names(&view), names);
            assert!(WorkbookTrait::worksheet_by_index(&view, 0).is_ok());
            file.text().unwrap();
        }
    }
}

// ---- ODS ----------------------------------------------------------------

#[cfg(feature = "ods")]
mod ods {
    use super::*;

    const MIME: &str = "application/vnd.oasis.opendocument.spreadsheet";
    const OFFICE: &str = "urn:oasis:names:tc:opendocument:xmlns:office:1.0";
    const TABLE: &str = "urn:oasis:names:tc:opendocument:xmlns:table:1.0";
    const TEXT: &str = "urn:oasis:names:tc:opendocument:xmlns:text:1.0";

    pub(super) fn package(title: &str) -> Vec<u8> {
        let content = format!(
            r#"<office:document-content xmlns:office="{OFFICE}" xmlns:table="{TABLE}" xmlns:text="{TEXT}"><office:body><office:spreadsheet><table:table table:name="Sheet1"><table:table-row table:number-rows-repeated="2"><table:table-cell table:number-columns-repeated="2" office:value-type="string"><text:p>hello</text:p></table:table-cell><table:table-cell office:value-type="string"><text:p>world</text:p></table:table-cell></table:table-row></table:table><table:table table:name="Sheet2"/></office:spreadsheet></office:body></office:document-content>"#
        );
        let metadata = format!(
            r#"<office:document-meta xmlns:office="{OFFICE}" xmlns:dc="http://purl.org/dc/elements/1.1/"><office:meta><dc:title>{title}</dc:title></office:meta></office:document-meta>"#
        );
        let mut writer = litchi_odf_common::core::PackageWriter::new(
            litchi_odf_common::core::StreamingArchiveLimits::default(),
        );
        writer.set_mimetype(MIME).unwrap();
        writer.add_file("content.xml", content.as_bytes()).unwrap();
        writer.add_file("meta.xml", metadata.as_bytes()).unwrap();
        writer
            .add_file_with_media_type("Pictures/sample.bin", b"media", "application/octet-stream")
            .unwrap();
        writer.finish_to_bytes().unwrap()
    }

    pub(super) fn malformed_body() -> Vec<u8> {
        let mut writer = litchi_odf_common::core::PackageWriter::new(
            litchi_odf_common::core::StreamingArchiveLimits::default(),
        );
        writer.set_mimetype(MIME).unwrap();
        writer
            .add_file("content.xml", b"<not-an-ods-document/>")
            .unwrap();
        writer.finish_to_bytes().unwrap()
    }

    fn ods(file: &File) -> &crate::format::ods::Ods {
        match file {
            File::Ods(ods) => ods,
            other => panic!("expected ODS, got {other:?}"),
        }
    }

    #[cfg(any(unix, windows))]
    #[test]
    fn filesystem_ods_uses_source_owner_and_matches_byte_projection() {
        let bytes = package("source");
        // An OOXML name on an ODF package: the content decides.
        let temporary = super::super::temporary_with(&bytes, ".xlsx");
        let source = crate::open(temporary.path()).unwrap();
        assert!(ods(&source).is_source_backed());
        let eager = crate::open_bytes(bytes).unwrap();
        assert!(!ods(&eager).is_source_backed());
        let (source_view, eager_view) = (source.workbook().unwrap(), eager.workbook().unwrap());
        assert_eq!(
            source_view.worksheet_names().unwrap(),
            eager_view.worksheet_names().unwrap()
        );
        assert_eq!(
            source_view.worksheet_count().unwrap(),
            eager_view.worksheet_count().unwrap()
        );
        assert_eq!(source.text().unwrap(), eager.text().unwrap());
        assert_eq!(
            source.metadata().unwrap().unwrap().title.as_deref(),
            Some("source")
        );
        assert_eq!(
            eager.metadata().unwrap().unwrap().title.as_deref(),
            Some("source")
        );
    }

    #[cfg(any(unix, windows))]
    #[test]
    fn source_root_metadata_reports_source_change() {
        let temporary = super::super::temporary_with(&package("before"), "");
        let file = crate::open(temporary.path()).unwrap();
        assert_eq!(
            file.metadata().unwrap().unwrap().title.as_deref(),
            Some("before")
        );
        std::fs::write(temporary.path(), package("after with a different size")).unwrap();
        let error = file.metadata().unwrap_err();
        assert_eq!(error.code(), Code::IoError);
        assert!(super::super::source_change(&error).is_some(), "{error:?}");
    }

    #[test]
    fn a_malformed_ods_body_is_refused_as_ods() {
        let refusal = |file: crate::Result<File>| match file {
            Ok(file) => file.text().expect_err("a malformed body is refused"),
            Err(error) => error,
        };
        let error = refusal(crate::open_bytes(malformed_body()));
        assert_eq!(
            error.format(),
            Some(litchi_core::FileFormat::Ods),
            "{error:?}"
        );
        #[cfg(any(unix, windows))]
        {
            let temporary = super::super::temporary_with(&malformed_body(), "");
            let error = refusal(crate::open(temporary.path()));
            assert_eq!(
                error.format(),
                Some(litchi_core::FileFormat::Ods),
                "{error:?}"
            );
        }
    }

    #[cfg(any(unix, windows))]
    #[test]
    fn source_root_preserves_repeated_text_semantics() {
        let temporary = super::super::temporary_with(&package("media"), "");
        let file = crate::open(temporary.path()).unwrap();
        assert_eq!(
            file.text().unwrap(),
            "hello\thello\tworld\nhello\thello\tworld\n"
        );
    }

    /// ODS has no sheet adapter yet: the trait reports names and count and
    /// refuses worksheet access with a typed capability refusal.
    #[test]
    fn the_trait_reports_ods_names_and_refuses_worksheet_access() {
        let file = crate::open_bytes(package("trait")).unwrap();
        let view = file.workbook().unwrap();
        assert_eq!(generic_names(&view), ["Sheet1", "Sheet2"]);
        assert_eq!(WorkbookTrait::worksheet_count(&view), 2);
        let refused = |error: crate::Error| {
            assert_eq!(error.code(), Code::UnsupportedCapability);
            assert_eq!(error.format(), Some(litchi_core::FileFormat::Ods));
        };
        refused(
            WorkbookTrait::worksheet_by_index(&view, 0)
                .map(drop)
                .unwrap_err(),
        );
        refused(
            WorkbookTrait::worksheet_by_name(&view, "Sheet1")
                .map(drop)
                .unwrap_err(),
        );
        refused(
            WorkbookTrait::active_worksheet(&view)
                .map(drop)
                .unwrap_err(),
        );
        let mut worksheets = WorkbookTrait::worksheets(&view);
        refused(
            worksheets
                .next()
                .expect("one refusal")
                .map(drop)
                .unwrap_err(),
        );
        assert!(worksheets.next().is_none());
    }

    #[test]
    fn flat_ods_detection_and_open_agree() {
        const FLAT_ODS: &[u8] = br#"<?xml version="1.0" encoding="UTF-8"?>
<office:document
    xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0"
    xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0"
    xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0"
    office:version="1.3"
    office:mimetype="application/vnd.oasis.opendocument.spreadsheet">
  <office:body>
    <office:spreadsheet>
      <table:table table:name="Sheet1">
        <table:table-row>
          <table:table-cell office:value-type="string"><text:p>value</text:p></table:table-cell>
        </table:table-row>
      </table:table>
    </office:spreadsheet>
  </office:body>
</office:document>"#;
        assert_eq!(
            crate::detect::detect_file_format_from_bytes(
                FLAT_ODS,
                &crate::detect::SourceLimits::default(),
            ),
            Some(litchi_core::FileFormat::Ods)
        );
        // Flat ODS opens through the leaf's flat owner (record 0277); both
        // routes agree with detection, and an edit republishes into the
        // flat bytes (record 0292).
        let opened = crate::open_bytes(FLAT_ODS.to_vec()).unwrap();
        assert_eq!(opened.format(), Some(litchi_core::FileFormat::Ods));
        assert_eq!(opened.text().unwrap(), "value");
        let edited = opened
            .set("/Sheet1/A1", crate::props! { "value" => 1 })
            .unwrap();
        let published = edited.to_bytes().unwrap();
        assert!(published.starts_with(b"<?xml"));
        assert_eq!(crate::open_bytes(published).unwrap().text().unwrap(), "1");
        #[cfg(any(unix, windows))]
        {
            let temporary = super::super::temporary_with(FLAT_ODS, "");
            let opened = crate::open(temporary.path()).unwrap();
            assert_eq!(opened.text().unwrap(), "value");
            assert_eq!(opened.to_bytes().unwrap(), FLAT_ODS);
        }
    }
}

// ---- OOXML/ODF polyglots (`ooxml_odf_polyglot_tests`) --------------------

#[cfg(all(feature = "ods", any(feature = "xlsx", feature = "xlsb")))]
mod polyglot {
    use std::io::{Cursor, Write as _};

    use super::*;
    use crate::OpenOptions;

    fn ordinary_ods() -> Vec<u8> {
        let mut writer = litchi_odf_common::core::PackageWriter::new(
            litchi_odf_common::core::StreamingArchiveLimits::default(),
        );
        writer
            .set_mimetype(litchi_odf_common::constants::ODF_SPREADSHEET)
            .unwrap();
        writer
            .add_file(
                "content.xml",
                br#"<office:document-content xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0"><office:body><office:spreadsheet><table:table table:name="Sheet1"/></office:spreadsheet></office:body></office:document-content>"#,
            )
            .unwrap();
        writer.finish_to_bytes().unwrap()
    }

    fn dual_marker_xlsx() -> Vec<u8> {
        let mut writer = zip::ZipWriter::new(Cursor::new(Vec::new()));
        let options = zip::write::SimpleFileOptions::default()
            .compression_method(zip::CompressionMethod::Stored);
        // A valid ODF local `mimetype` first, as ODF requires; the remaining
        // entries form a valid minimal OPC/XLSX package.
        writer.start_file("mimetype", options).unwrap();
        writer
            .write_all(b"application/vnd.oasis.opendocument.spreadsheet")
            .unwrap();
        writer.start_file("[Content_Types].xml", options).unwrap();
        writer
            .write_all(
                br#"<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/></Types>"#,
            )
            .unwrap();
        writer.start_file("_rels/.rels", options).unwrap();
        writer
            .write_all(
                br#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>"#,
            )
            .unwrap();
        writer.start_file("xl/workbook.xml", options).unwrap();
        writer
            .write_all(
                br#"<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="Sheet1" sheetId="1" r:id="rId1"/></sheets></workbook>"#,
            )
            .unwrap();
        writer
            .start_file("xl/_rels/workbook.xml.rels", options)
            .unwrap();
        writer
            .write_all(
                br#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/></Relationships>"#,
            )
            .unwrap();
        writer
            .start_file("xl/worksheets/sheet1.xml", options)
            .unwrap();
        writer
            .write_all(
                br#"<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData/></worksheet>"#,
            )
            .unwrap();
        writer.finish().unwrap().into_inner()
    }

    fn tight() -> OpenOptions {
        OpenOptions::new().ooxml_limits(
            litchi_opc::SourceLimits::builder()
                .max_input_bytes(1)
                .unwrap()
                .build()
                .unwrap(),
        )
    }

    /// An ordinary ODS uses the neutral ceiling, not the OPC limit, even
    /// under an OOXML name.
    #[cfg(any(unix, windows))]
    #[test]
    fn ordinary_ods_path_uses_ods_owner_under_tight_ooxml_limit() {
        let temporary = super::super::temporary_with(&ordinary_ods(), ".xlsx");
        let file = crate::open_with(temporary.path(), &tight()).unwrap();
        assert!(matches!(file, File::Ods(_)));
    }

    /// An ODS/OOXML polyglot is an OPC candidate and stays bound by the
    /// caller's OPC limit.
    #[cfg(any(unix, windows))]
    #[test]
    fn ods_ooxml_polyglot_remains_caller_limited() {
        let bytes = dual_marker_xlsx();
        let temporary = super::super::temporary_with(&bytes, ".ods");
        for error in [
            crate::open_with(temporary.path(), &tight()).unwrap_err(),
            crate::open_bytes_with(bytes.clone(), &tight()).unwrap_err(),
        ] {
            assert!(
                matches!(
                    error.leaf(),
                    crate::error::Source::Opc(litchi_opc::OpcError::Limit {
                        kind: litchi_opc::SourceLimitKind::InputBytes,
                        observed: actual,
                        maximum: 1,
                    }) if *actual == bytes.len() as u64
                ),
                "{error:?}"
            );
        }
    }

    #[cfg(feature = "xlsx")]
    #[test]
    fn ooxml_first_precedence_survives_an_odf_local_mimetype_marker() {
        let bytes = dual_marker_xlsx();
        let mut files = vec![crate::open_bytes(bytes.clone()).unwrap()];
        #[cfg(any(unix, windows))]
        let temporary = super::super::temporary_with(&bytes, "");
        #[cfg(any(unix, windows))]
        files.push(crate::open(temporary.path()).unwrap());
        for file in files {
            assert!(matches!(file, File::Xlsx(_)));
            assert_eq!(
                file.workbook().unwrap().worksheet_names().unwrap(),
                ["Sheet1"]
            );
        }
    }

    /// Without the XLSX owner the polyglot is still classified as XLSX, and
    /// never falls through to ODS. The facade reported "not an Office file";
    /// `File` names the format and the feature it needs.
    #[cfg(not(feature = "xlsx"))]
    #[test]
    fn a_disabled_xlsx_owner_keeps_ooxml_precedence() {
        let bytes = dual_marker_xlsx();
        let mut errors = vec![crate::open_bytes(bytes.clone()).unwrap_err()];
        #[cfg(any(unix, windows))]
        let temporary = super::super::temporary_with(&bytes, "");
        #[cfg(any(unix, windows))]
        errors.push(crate::open(temporary.path()).unwrap_err());
        for error in errors {
            assert_eq!(error.code(), Code::UnsupportedFormat);
            assert_eq!(error.detected(), Some(litchi_core::FileFormat::Xlsx));
        }
    }

    #[test]
    fn invalid_odf_body_still_falls_back_to_typed_ods_validation() {
        let error = match crate::open_bytes(ods::malformed_body()) {
            Ok(file) => file.text().unwrap_err(),
            Err(error) => error,
        };
        assert_eq!(
            error.format(),
            Some(litchi_core::FileFormat::Ods),
            "{error:?}"
        );
    }

    #[test]
    fn valid_ods_wins_after_the_ooxml_probe_fails() {
        assert!(matches!(
            crate::open_bytes(ordinary_ods()).unwrap(),
            File::Ods(_)
        ));
    }
}

// ---- XLS ----------------------------------------------------------------

#[cfg(feature = "xls")]
mod xls {
    use std::io::Cursor;

    use litchi_cfb::OleWriter;

    use super::*;

    const FIXTURE: &str = "ole/xls/Simple.xls";

    fn xls(file: &File) -> &crate::format::xls::Xls {
        match file {
            File::Xls(xls) => xls,
            other => panic!("expected XLS, got {other:?}"),
        }
    }

    pub(super) fn with_streams(streams: &[(&str, &[u8])]) -> Vec<u8> {
        let mut writer = OleWriter::new();
        for (name, bytes) in streams {
            writer.create_stream(&[*name], bytes).unwrap();
        }
        let mut output = Cursor::new(Vec::new());
        writer.write_to(&mut output).unwrap();
        output.into_inner()
    }

    pub(super) fn workbook_stream() -> Vec<u8> {
        let mut ole = litchi_cfb::OleFile::open(
            Cursor::new(std::fs::read(test_data(FIXTURE)).unwrap()),
            litchi_cfb::OleFileLimits::default(),
        )
        .unwrap();
        ole.open_stream(&["Workbook"]).unwrap()
    }

    #[test]
    fn test_workbook_xls_fixtures() {
        let path = test_data(FIXTURE);
        for file in [
            crate::open(&path).unwrap(),
            crate::open_bytes(std::fs::read(&path).unwrap()).unwrap(),
        ] {
            let view = file.workbook().unwrap();
            let names = view.worksheet_names().unwrap();
            assert!(!names.is_empty());
            assert_eq!(view.worksheet_count().unwrap(), names.len());
            file.text().unwrap();
            let _ = file
                .metadata()
                .unwrap()
                .map(|metadata| (metadata.title, metadata.author));
        }
        for relative in [
            "ole/xls/ConditionalFormattingSamples.xls",
            "ole/xls/HyperlinksOnManySheets.xls",
        ] {
            assert!(
                !open(relative)
                    .workbook()
                    .unwrap()
                    .worksheet_names()
                    .unwrap()
                    .is_empty()
            );
        }
        open("ole/xls/FormulaEvalTestData.xls").text().unwrap();
    }

    /// XLS reads through the generic trait from its eager reader, which a
    /// path-opened workbook materializes from the retained source on first
    /// use.
    #[test]
    fn the_trait_reads_xls_on_both_routes() {
        let path = test_data(FIXTURE);
        let from_path = crate::open(&path).unwrap();
        let from_bytes = crate::open_bytes(std::fs::read(&path).unwrap()).unwrap();
        for file in [&from_path, &from_bytes] {
            let view = file.workbook().unwrap();
            assert_eq!(generic_names(&view), view.worksheet_names().unwrap());
            let sheet = WorkbookTrait::worksheet_by_index(&view, 0).unwrap();
            assert!(sheet.row_count() > 0);
        }
    }

    #[cfg(any(unix, windows))]
    #[test]
    fn filesystem_open_retains_xls_source() {
        let file = open(FIXTURE);
        assert!(xls(&file).is_source_backed());
        let view = file.workbook().unwrap();
        let names = view.worksheet_names().unwrap();
        assert!(!names.is_empty());
        assert_eq!(view.worksheet_count().unwrap(), names.len());
    }

    #[cfg(any(unix, windows))]
    #[test]
    fn source_xls_matches_eager_names_text_and_metadata() {
        let path = test_data(FIXTURE);
        let source = crate::open(&path).unwrap();
        let eager = crate::open_bytes(std::fs::read(&path).unwrap()).unwrap();
        assert!(xls(&source).is_source_backed());
        assert!(!xls(&eager).is_source_backed());
        let source_metadata = source.metadata().unwrap().unwrap();
        let eager_metadata = eager.metadata().unwrap().unwrap();
        assert!(source_metadata.application.is_some());
        assert_eq!(source_metadata.application, eager_metadata.application);
        assert_eq!(source_metadata.author, eager_metadata.author);
        let (source_view, eager_view) = (source.workbook().unwrap(), eager.workbook().unwrap());
        assert_eq!(
            source_view.worksheet_names().unwrap(),
            eager_view.worksheet_names().unwrap()
        );
        assert_eq!(
            source_view.worksheet_count().unwrap(),
            eager_view.worksheet_count().unwrap()
        );
        assert_eq!(source.text().unwrap(), eager.text().unwrap());
    }

    /// A BIFF5 workbook, which the positional owner does not read, takes
    /// the eager route from a path and matches the byte route.
    #[cfg(any(unix, windows))]
    #[test]
    fn legacy_biff_path_matches_bytes() {
        let mut stream = workbook_stream();
        stream[4..6].copy_from_slice(&0x0500_u16.to_le_bytes());
        let bytes = with_streams(&[("Workbook", &stream)]);
        let temporary = super::super::temporary_with(&bytes, "");
        let from_path = crate::open(temporary.path()).unwrap();
        let from_bytes = crate::open_bytes(bytes).unwrap();
        assert!(!xls(&from_path).is_source_backed());
        let (path_view, bytes_view) = (
            from_path.workbook().unwrap(),
            from_bytes.workbook().unwrap(),
        );
        assert_eq!(
            path_view.worksheet_names().unwrap(),
            bytes_view.worksheet_names().unwrap()
        );
        assert_eq!(
            path_view.worksheet_count().unwrap(),
            bytes_view.worksheet_count().unwrap()
        );
        assert_eq!(from_path.text().unwrap(), from_bytes.text().unwrap());
    }

    #[cfg(any(unix, windows))]
    #[test]
    fn source_xls_refuses_stale_path() {
        let temporary = super::super::temporary_copy(&test_data(FIXTURE), "");
        let file = crate::open(temporary.path()).unwrap();
        super::super::mutate(temporary.path());
        let error = file.workbook().unwrap().worksheet_names().unwrap_err();
        assert_eq!(error.code(), Code::IoError, "{error:?}");
    }

    /// An encrypted workbook is recognized and locked, never read eagerly.
    #[cfg(any(unix, windows))]
    #[test]
    fn source_xls_locks_an_encrypted_filepass() {
        let temporary = super::super::temporary_copy(
            &test_data("poi/test-data/spreadsheet/xor-encryption-abc.xls"),
            "",
        );
        let file = crate::open(temporary.path()).unwrap();
        let File::Locked(locked) = &file else {
            panic!("expected a locked workbook, got {file:?}");
        };
        assert_eq!(locked.format(), Some(litchi_core::FileFormat::Xls));
        assert_eq!(locked.encryption(), crate::EncryptionKind::XorObfuscation);
    }

    #[cfg(any(unix, windows))]
    #[test]
    fn source_xls_refuses_malformed_biff_without_eager_fallback() {
        let mut stream = 0x0809_u16.to_le_bytes().to_vec();
        stream.extend_from_slice(&16_u16.to_le_bytes());
        stream.extend_from_slice(&[0; 16]);
        let temporary = super::super::temporary_with(&with_streams(&[("Workbook", &stream)]), "");
        let error = crate::open(temporary.path()).expect_err("a malformed workbook");
        assert!(
            matches!(error.leaf(), crate::error::Source::XlsSource(_)),
            "{error:?}"
        );
    }

    #[cfg(unix)]
    #[test]
    fn source_xls_pins_the_open_file_across_atomic_path_replacement() {
        let original = super::super::temporary_copy(&test_data(FIXTURE), "");
        let file = crate::open(original.path()).unwrap();
        let names = file.workbook().unwrap().worksheet_names().unwrap();
        let replacement = super::super::temporary_with(b"not an XLS workbook", "");
        std::fs::rename(replacement.path(), original.path()).unwrap();
        // The open descriptor stays on the original inode.
        assert_eq!(file.workbook().unwrap().worksheet_names().unwrap(), names);
        assert!(crate::open(original.path()).is_err());
    }

    /// Host-stream precedence: `WordDocument`, then `PowerPoint Document` or
    /// `Current User`, then `Workbook` or `Book`. A host whose feature is off
    /// is named, not re-routed to XLS: the facade's workbook opener fell back
    /// to XLS in such builds; `File` reports the format detection found.
    #[cfg(any(unix, windows))]
    #[test]
    fn ole_host_stream_precedence_names_the_host() {
        use litchi_core::FileFormat;

        let workbook = workbook_stream();
        let cases: [(Vec<(&str, &[u8])>, Option<FileFormat>); 5] = [
            (
                vec![
                    ("WordDocument", b"word".as_slice()),
                    ("PowerPoint Document", b"ppt".as_slice()),
                    ("Workbook", workbook.as_slice()),
                ],
                Some(FileFormat::Doc),
            ),
            (
                vec![
                    ("PowerPoint Document", b"ppt".as_slice()),
                    ("Workbook", workbook.as_slice()),
                ],
                Some(FileFormat::Ppt),
            ),
            (
                vec![
                    ("Current User", b"ppt".as_slice()),
                    ("Workbook", workbook.as_slice()),
                ],
                Some(FileFormat::Ppt),
            ),
            (
                vec![
                    ("Book", b"xls".as_slice()),
                    ("Workbook", workbook.as_slice()),
                ],
                Some(FileFormat::Xls),
            ),
            (vec![("Other", b"ole".as_slice())], None),
        ];
        for (streams, expected) in cases {
            let bytes = with_streams(&streams);
            let temporary = super::super::temporary_with(&bytes, "");
            for result in [
                crate::open(temporary.path()),
                crate::open_bytes(bytes.clone()),
            ] {
                let named = match result {
                    Ok(file) => file.format(),
                    Err(error) => error.format().or(error.detected()),
                };
                assert_eq!(named, expected, "{streams:?}");
            }
        }
    }
}
