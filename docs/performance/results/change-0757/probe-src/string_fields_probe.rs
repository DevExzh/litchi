//! Evidence probe for change 0757 (not part of the test suite): what the
//! writer at the base does with strings past each BIFF8 limit. Compiles
//! against both legs; it only reports outcomes.
#![allow(clippy::unwrap_used, clippy::expect_used, reason = "probe")]

use std::io::Cursor;

use litchi_core::sheet::{Cell as _, CellValue, WorkbookTrait as _};
use litchi_xls::writer::Writer;

fn write(writer: &mut Writer) -> Result<Vec<u8>, String> {
    let mut output = Cursor::new(Vec::new());
    writer.write_to(&mut output).map_err(|error| error.to_string())?;
    Ok(output.into_inner())
}

fn read(bytes: Vec<u8>) -> Result<litchi_xls::Workbook<Cursor<Vec<u8>>>, String> {
    litchi_xls::Workbook::new(Cursor::new(bytes)).map_err(|error| error.to_string())
}

fn report(case: &str, outcome: String) {
    println!("PROBE {case}: {outcome}");
}

fn cell_outcome(value: &str) -> String {
    let mut writer = Writer::new();
    let sheet = writer.add_worksheet("Probe").unwrap();
    if let Err(error) = writer.write_string(sheet, 0, 0, value) {
        return format!("write_string refused: {error}");
    }
    let bytes = match write(&mut writer) {
        Ok(bytes) => bytes,
        Err(error) => return format!("write_to refused: {error}"),
    };
    match read(bytes) {
        Err(error) => format!("written; reader refused the workbook: {error}"),
        Ok(workbook) => match workbook.xls_worksheet(0).unwrap().get_cell(0, 0).unwrap().value() {
            CellValue::String(text) => format!(
                "written; read back {} of {} UTF-16 units{}",
                text.encode_utf16().count(),
                value.encode_utf16().count(),
                if text == value { " (whole)" } else { " (TRUNCATED)" }
            ),
            other => format!("written; read back {other:?}"),
        },
    }
}

#[test]
fn probe_string_fields_past_their_limits() {
    report("cell 65,535 x b", cell_outcome(&"b".repeat(65_535)));
    report("cell 65,536 x b", cell_outcome(&"b".repeat(65_536)));
    report("cell 70,000 x e-acute", cell_outcome(&"é".repeat(70_000)));
    report("cell 65,534 x a + emoji", cell_outcome(&format!("{}😀", "a".repeat(65_534))));

    for (case, literal) in [
        ("formula literal 255 x a", "a".repeat(255)),
        ("formula literal 300 x a", "a".repeat(300)),
        ("formula literal 254 x a + emoji", format!("{}😀", "a".repeat(254))),
    ] {
        let mut writer = Writer::new();
        let sheet = writer.add_worksheet("Probe").unwrap();
        writer.write_formula(sheet, 0, 0, &format!("\"{literal}\"")).unwrap();
        let outcome = match write(&mut writer).and_then(read) {
            Err(error) => format!("refused: {error}"),
            Ok(workbook) => {
                let bytes = workbook
                    .xls_worksheet(0)
                    .unwrap()
                    .get_cell(0, 0)
                    .unwrap()
                    .formula_bytes()
                    .unwrap()
                    .to_vec();
                format!(
                    "written; PtgStr cch {} of {} UTF-16 units{}",
                    bytes[1],
                    literal.encode_utf16().count(),
                    if usize::from(bytes[1]) == literal.encode_utf16().count() {
                        " (whole)"
                    } else {
                        " (TRUNCATED)"
                    }
                )
            },
        };
        report(case, outcome);
    }

    for (case, pattern) in [
        ("number format 255 x 0", "0".repeat(255)),
        ("number format 256 x 0", "0".repeat(256)),
        ("number format 70,000 x 0", "0".repeat(70_000)),
    ] {
        let mut writer = Writer::new();
        writer.add_worksheet("Probe").unwrap();
        writer.register_number_format(&pattern);
        let outcome = match std::panic::catch_unwind(std::panic::AssertUnwindSafe(|| write(&mut writer))) {
            Err(_) => "write_to PANICKED".to_string(),
            Ok(Err(error)) => format!("write_to refused: {error}"),
            Ok(Ok(bytes)) => match read(bytes) {
                Err(error) => format!("written; reader refused the workbook: {error}"),
                Ok(_) => "written; reader accepted it".to_string(),
            },
        };
        report(case, outcome);
    }

    for (case, name) in [
        ("defined name Cafe (Latin-1)", "Café".to_string()),
        ("defined name with emoji", "Rocket😀Launch".to_string()),
        ("defined name 200 emoji (400 units)", "😀".repeat(200)),
    ] {
        let mut writer = Writer::new();
        writer.add_worksheet("Probe").unwrap();
        let outcome = match writer.define_name(&name, "A1:B2") {
            Err(error) => format!("define_name refused: {error}"),
            Ok(()) => match write(&mut writer).and_then(read) {
                Err(error) => format!("written; reader refused the workbook: {error}"),
                Ok(workbook) => {
                    let names: Vec<String> = workbook
                        .defined_names()
                        .iter()
                        .map(|defined| defined.name.clone())
                        .collect();
                    format!(
                        "written; read back {names:?}{}",
                        if names == [name.clone()] { " (whole)" } else { " (WRONG)" }
                    )
                },
            },
        };
        report(case, outcome);
    }

    for (case, name) in [
        ("sheet name 16 x e-acute (32 bytes)", "é".repeat(16)),
        ("sheet name 31 x e-acute", "é".repeat(31)),
        ("sheet name 32 x b", "b".repeat(32)),
    ] {
        let mut writer = Writer::new();
        let outcome = match writer.add_worksheet(&name) {
            Err(error) => format!("add_worksheet refused: {error}"),
            Ok(_) => match write(&mut writer).and_then(read) {
                Err(error) => format!("written; reader refused: {error}"),
                Ok(workbook) => format!("written; read back {:?}", workbook.worksheet_names()),
            },
        };
        report(case, outcome);
    }
}
