//! Compatibility re-exports for the XLSX-owned data-consolidation codec.

pub use litchi_xlsx::data_consolidation::*;

#[cfg(test)]
mod tests {
    use super::*;

    const T: &str = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
    const R: &str = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";

    // Provenance: LibreOffice's sc/source/filter/excel/ooxml-export-TODO.txt explicitly lists
    // dataConsolidate/dataRefs/dataRef, while POI exposes the same function enum only for pivot
    // tables. This deterministic package is therefore synthetic rather than mislabeled corpus data.
    #[test]
    fn immutable_worksheet_accessor_reads_synthetic_package() {
        use crate::xlsx::{Workbook, Worksheet, WorksheetInfo};
        use std::fs;
        use std::time::{SystemTime, UNIX_EPOCH};

        let sheet = format!(
            r#"<?xml version="1.0"?><worksheet xmlns="{T}"><sheetData/><dataConsolidate function="average" topLabels="1"><dataRefs count="1"><dataRef sheet="Input" ref="A1:C9"/></dataRefs></dataConsolidate></worksheet>"#
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
                &format!(
                    r#"<?xml version="1.0"?><workbook xmlns="{T}" xmlns:r="{R}"><sheets><sheet name="Result" sheetId="1" r:id="rId1"/></sheets></workbook>"#
                ),
            ),
            (
                "xl/_rels/workbook.xml.rels",
                r#"<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/></Relationships>"#,
            ),
            ("xl/worksheets/sheet1.xml", &sheet),
        ]);
        let path = std::env::temp_dir().join(format!(
            "litchi-data-consolidate-{}-{}.xlsx",
            std::process::id(),
            SystemTime::now()
                .duration_since(UNIX_EPOCH)
                .unwrap()
                .as_nanos()
        ));
        fs::write(&path, package).unwrap();
        let workbook = Workbook::open(&path).unwrap();
        let mut worksheet = Worksheet::new(
            &workbook,
            WorksheetInfo {
                name: "Result".into(),
                relationship_id: "rId1".into(),
                sheet_id: 1,
                is_active: true,
                print_area: None,
                repeating_rows: None,
                repeating_columns: None,
            },
        );
        worksheet.load_data().unwrap();
        let consolidation = worksheet.data_consolidation().unwrap();
        assert_eq!(consolidation.function(), Function::Average);
        assert!(consolidation.top_labels());
        assert_eq!(
            consolidation.data_references().unwrap().references().len(),
            1
        );
        fs::remove_file(path).unwrap();
    }

    fn make_package(entries: &[(&str, &str)]) -> Vec<u8> {
        let mut bytes = Vec::new();
        let mut central = Vec::new();
        for (name, value) in entries {
            let offset = bytes.len() as u32;
            let data = value.as_bytes();
            let crc = crc32(data);
            push_u32(&mut bytes, 0x04034b50);
            push_u16(&mut bytes, 20);
            push_u16(&mut bytes, 0);
            push_u16(&mut bytes, 0);
            push_u16(&mut bytes, 0);
            push_u16(&mut bytes, 0);
            push_u32(&mut bytes, crc);
            push_u32(&mut bytes, data.len() as u32);
            push_u32(&mut bytes, data.len() as u32);
            push_u16(&mut bytes, name.len() as u16);
            push_u16(&mut bytes, 0);
            bytes.extend_from_slice(name.as_bytes());
            bytes.extend_from_slice(data);
            push_u32(&mut central, 0x02014b50);
            push_u16(&mut central, 20);
            push_u16(&mut central, 20);
            push_u16(&mut central, 0);
            push_u16(&mut central, 0);
            push_u16(&mut central, 0);
            push_u16(&mut central, 0);
            push_u32(&mut central, crc);
            push_u32(&mut central, data.len() as u32);
            push_u32(&mut central, data.len() as u32);
            push_u16(&mut central, name.len() as u16);
            push_u16(&mut central, 0);
            push_u16(&mut central, 0);
            push_u16(&mut central, 0);
            push_u16(&mut central, 0);
            push_u32(&mut central, 0);
            push_u32(&mut central, offset);
            central.extend_from_slice(name.as_bytes());
        }
        let central_offset = bytes.len() as u32;
        let central_size = central.len() as u32;
        bytes.extend_from_slice(&central);
        push_u32(&mut bytes, 0x06054b50);
        push_u16(&mut bytes, 0);
        push_u16(&mut bytes, 0);
        push_u16(&mut bytes, entries.len() as u16);
        push_u16(&mut bytes, entries.len() as u16);
        push_u32(&mut bytes, central_size);
        push_u32(&mut bytes, central_offset);
        push_u16(&mut bytes, 0);
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
    fn push_u16(out: &mut Vec<u8>, value: u16) {
        out.extend_from_slice(&value.to_le_bytes());
    }
    fn push_u32(out: &mut Vec<u8>, value: u32) {
        out.extend_from_slice(&value.to_le_bytes());
    }
}
