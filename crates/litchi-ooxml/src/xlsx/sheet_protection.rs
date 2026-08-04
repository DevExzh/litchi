//! Compatibility exports for the canonical XLSX worksheet-protection codec.

pub use litchi_xlsx::sheet_protection::*;

#[cfg(test)]
mod tests {
    use super::*;

    const CORE: &[u8] = b"http://schemas.openxmlformats.org/spreadsheetml/2006/main";
    const START: &str =
        r#"<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">"#;

    fn parse(body: &str) -> litchi_xlsx::Result<WorksheetProtectionMetadata> {
        parse_worksheet_protection(format!("{START}{body}</worksheet>").as_bytes())
    }

    #[test]
    fn reads_poi_libreoffice_and_synthetic_package_through_worksheet_accessor() {
        use crate::xlsx::{Workbook, Worksheet, WorksheetInfo};
        use std::fs;
        use std::path::Path;
        use std::time::{SystemTime, UNIX_EPOCH};

        fn inspect(path: &Path, relationship_id: &str, expected_ranges: usize, strong_sheet: bool) {
            let workbook = Workbook::open(path).unwrap();
            let mut worksheet = Worksheet::new(
                &workbook,
                WorksheetInfo {
                    name: "Sheet1".into(),
                    relationship_id: relationship_id.into(),
                    sheet_id: 1,
                    is_active: true,
                    print_area: None,
                    repeating_rows: None,
                    repeating_columns: None,
                },
            );
            worksheet.load_data().unwrap();
            let metadata = worksheet.protection_metadata();
            assert_eq!(metadata.protected_ranges().count(), expected_ranges);
            if strong_sheet {
                assert!(matches!(
                    metadata.sheet_protection().unwrap().verifier(),
                    Some(ProtectionPasswordVerifier::Strong(_))
                ));
            }
        }

        let root = Path::new(env!("CARGO_MANIFEST_DIR")).join("../..");
        inspect(
            &root.join(
                "test-data/poi/test-data/spreadsheet/workbookProtection-sheet_password-2013.xlsx",
            ),
            "rId1",
            0,
            true,
        );
        inspect(
            &root.join("test-data/libreoffice-core/sc/qa/unit/data/xlsx/enhanced-protection.xlsx"),
            "rId1",
            5,
            false,
        );
        inspect(&root.join("test-data/libreoffice-core/sc/qa/unit/data/xlsx/enhancedProtectionRangeShorthand.xlsx"), "rId2", 1, false);

        let metadata = parse(r#"<sheetData/><sheetProtection password="CC3D" sheet="1"/><protectedRanges><protectedRange name="Editable" sqref="A1:B2"/></protectedRanges>"#).unwrap();
        let fragment =
            write_worksheet_protection(&metadata, WorksheetProtectionConformance::Transitional)
                .unwrap();
        let sheet = format!(
            r#"<?xml version="1.0"?><worksheet xmlns="{}"><sheetData/>{fragment}</worksheet>"#,
            std::str::from_utf8(CORE).unwrap()
        );
        let package = make_package(&[
            (
                "[Content_Types].xml",
                r#"<?xml version="1.0"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/></Types>"#,
            ),
            (
                "_rels/.rels",
                r#"<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>"#,
            ),
            (
                "xl/workbook.xml",
                r#"<?xml version="1.0"?><workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="Sheet1" sheetId="1" r:id="rId1"/></sheets></workbook>"#,
            ),
            (
                "xl/_rels/workbook.xml.rels",
                r#"<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/></Relationships>"#,
            ),
            ("xl/worksheets/sheet1.xml", &sheet),
        ]);
        let path = std::env::temp_dir().join(format!(
            "litchi-sheet-protection-{}-{}.xlsx",
            std::process::id(),
            SystemTime::now()
                .duration_since(UNIX_EPOCH)
                .unwrap()
                .as_nanos()
        ));
        fs::write(&path, package).unwrap();
        inspect(&path, "rId1", 1, false);
        fs::remove_file(path).unwrap();
    }

    fn make_package(entries: &[(&str, &str)]) -> Vec<u8> {
        let mut bytes = Vec::new();
        let mut central = Vec::new();
        for (name, value) in entries {
            let offset = bytes.len() as u32;
            let data = value.as_bytes();
            let crc = crc32(data);
            push32(&mut bytes, 0x04034b50);
            push16(&mut bytes, 20);
            push16(&mut bytes, 0);
            push16(&mut bytes, 0);
            push16(&mut bytes, 0);
            push16(&mut bytes, 0);
            push32(&mut bytes, crc);
            push32(&mut bytes, data.len() as u32);
            push32(&mut bytes, data.len() as u32);
            push16(&mut bytes, name.len() as u16);
            push16(&mut bytes, 0);
            bytes.extend_from_slice(name.as_bytes());
            bytes.extend_from_slice(data);
            push32(&mut central, 0x02014b50);
            push16(&mut central, 20);
            push16(&mut central, 20);
            push16(&mut central, 0);
            push16(&mut central, 0);
            push16(&mut central, 0);
            push16(&mut central, 0);
            push32(&mut central, crc);
            push32(&mut central, data.len() as u32);
            push32(&mut central, data.len() as u32);
            push16(&mut central, name.len() as u16);
            push16(&mut central, 0);
            push16(&mut central, 0);
            push16(&mut central, 0);
            push16(&mut central, 0);
            push32(&mut central, 0);
            push32(&mut central, offset);
            central.extend_from_slice(name.as_bytes());
        }
        let offset = bytes.len() as u32;
        let size = central.len() as u32;
        bytes.extend_from_slice(&central);
        push32(&mut bytes, 0x06054b50);
        push16(&mut bytes, 0);
        push16(&mut bytes, 0);
        push16(&mut bytes, entries.len() as u16);
        push16(&mut bytes, entries.len() as u16);
        push32(&mut bytes, size);
        push32(&mut bytes, offset);
        push16(&mut bytes, 0);
        bytes
    }
    fn crc32(data: &[u8]) -> u32 {
        let mut crc = !0u32;
        for byte in data {
            crc ^= u32::from(*byte);
            for _ in 0..8 {
                crc = if crc & 1 != 0 {
                    (crc >> 1) ^ 0xedb88320
                } else {
                    crc >> 1
                };
            }
        }
        !crc
    }
    fn push16(out: &mut Vec<u8>, value: u16) {
        out.extend_from_slice(&value.to_le_bytes());
    }
    fn push32(out: &mut Vec<u8>, value: u32) {
        out.extend_from_slice(&value.to_le_bytes());
    }
}
