use std::io::Cursor;

use litchi_cfb::OleWriter;
use litchi_core::sheet::{Cell as _, CellValue};
use litchi_xls::Workbook;
use litchi_xls::formula_metadata::Range;
use litchi_xls::writer::Writer;

const BOF: u16 = 0x0809;
const BOUNDSHEET8: u16 = 0x0085;
const FONT: u16 = 0x0031;
const WINDOW1: u16 = 0x003d;
const XF: u16 = 0x00e0;
const RRTABID: u16 = 0x013d;
const DIMENSIONS: u16 = 0x0200;
const FORMULA: u16 = 0x0006;
const ARRAY: u16 = 0x0221;
const STRING: u16 = 0x0207;
const EOF: u16 = 0x000a;
const PTG_EXP: u8 = 0x01;
const PTG_INT: u8 = 0x1e;

fn push_record(stream: &mut Vec<u8>, kind: u16, payload: &[u8]) {
    stream.extend_from_slice(&kind.to_le_bytes());
    stream.extend_from_slice(
        &u16::try_from(payload.len())
            .expect("fixture record fits BIFF8")
            .to_le_bytes(),
    );
    stream.extend_from_slice(payload);
}

fn push_bof(stream: &mut Vec<u8>, substream: u16) {
    let mut payload = Vec::with_capacity(16);
    payload.extend_from_slice(&0x0600u16.to_le_bytes());
    payload.extend_from_slice(&substream.to_le_bytes());
    payload.extend_from_slice(&[0; 12]);
    push_record(stream, BOF, &payload);
}

fn ptg_exp(row: u16, col: u8) -> [u8; 5] {
    let row = row.to_le_bytes();
    [PTG_EXP, row[0], row[1], col, 0]
}

fn push_dimensions(
    stream: &mut Vec<u8>,
    first_row: u32,
    last_row_exclusive: u32,
    first_col: u16,
    last_col_exclusive: u16,
) {
    let mut payload = Vec::with_capacity(14);
    payload.extend_from_slice(&first_row.to_le_bytes());
    payload.extend_from_slice(&last_row_exclusive.to_le_bytes());
    payload.extend_from_slice(&first_col.to_le_bytes());
    payload.extend_from_slice(&last_col_exclusive.to_le_bytes());
    payload.extend_from_slice(&0u16.to_le_bytes());
    push_record(stream, DIMENSIONS, &payload);
}

fn formula(row: u16, col: u8, anchor: (u16, u8), cached: [u8; 8]) -> Vec<u8> {
    let tokens = ptg_exp(anchor.0, anchor.1);
    let mut payload = Vec::with_capacity(27);
    payload.extend_from_slice(&row.to_le_bytes());
    payload.extend_from_slice(&u16::from(col).to_le_bytes());
    payload.extend_from_slice(&0u16.to_le_bytes()); // ixfe
    payload.extend_from_slice(&cached);
    payload.extend_from_slice(&0u16.to_le_bytes()); // grbit
    payload.extend_from_slice(&0u32.to_le_bytes()); // chn
    payload.extend_from_slice(&(tokens.len() as u16).to_le_bytes());
    payload.extend_from_slice(&tokens);
    payload
}

fn scalar_formula(row: u16, col: u8, value: u16) -> Vec<u8> {
    let tokens = [PTG_INT, value as u8, (value >> 8) as u8];
    let mut payload = Vec::with_capacity(25);
    payload.extend_from_slice(&row.to_le_bytes());
    payload.extend_from_slice(&u16::from(col).to_le_bytes());
    payload.extend_from_slice(&0u16.to_le_bytes());
    payload.extend_from_slice(&f64::from(value).to_le_bytes());
    payload.extend_from_slice(&0u16.to_le_bytes());
    payload.extend_from_slice(&0u32.to_le_bytes());
    payload.extend_from_slice(&(tokens.len() as u16).to_le_bytes());
    payload.extend_from_slice(&tokens);
    payload
}

fn array(first: (u16, u8), last: (u16, u8), value: u16) -> Vec<u8> {
    let tokens = [PTG_INT, value as u8, (value >> 8) as u8];
    let mut payload = Vec::with_capacity(17);
    payload.extend_from_slice(&first.0.to_le_bytes());
    payload.extend_from_slice(&last.0.to_le_bytes());
    payload.push(first.1);
    payload.push(last.1);
    payload.extend_from_slice(&0u16.to_le_bytes()); // fAlwaysCalc and reserved bits
    payload.extend_from_slice(&0u32.to_le_bytes()); // unused
    payload.extend_from_slice(&(tokens.len() as u16).to_le_bytes());
    payload.extend_from_slice(&tokens);
    payload
}

fn workbook_with(worksheet_records: &[u8]) -> Vec<u8> {
    let mut workbook = Vec::new();
    push_bof(&mut workbook, 0x0005);
    push_record(&mut workbook, RRTABID, &1u16.to_le_bytes());
    let mut font = Vec::with_capacity(21);
    font.extend_from_slice(&200u16.to_le_bytes()); // dyHeight
    font.extend_from_slice(&0u16.to_le_bytes()); // grbit
    font.extend_from_slice(&0x7fffu16.to_le_bytes()); // icv automatic
    font.extend_from_slice(&400u16.to_le_bytes()); // bls normal
    font.extend_from_slice(&0u16.to_le_bytes()); // sss
    font.extend_from_slice(&[0, 0, 0, 0]); // underline, family, charset, reserved
    font.extend_from_slice(&[5, 0]); // compressed XLUnicodeString length/options
    font.extend_from_slice(b"Arial");
    for _ in 0..4 {
        push_record(&mut workbook, FONT, &font);
    }
    // BIFF8 requires fifteen built-in style XFs followed by a cell XF.
    for index in 0..16 {
        let mut xf = [0; 20];
        if index < 15 {
            xf[4] = 0x04; // fStyle
        }
        push_record(&mut workbook, XF, &xf);
    }
    let mut window = [0; 18];
    window[4..6].copy_from_slice(&1000u16.to_le_bytes());
    window[6..8].copy_from_slice(&800u16.to_le_bytes());
    window[14..16].copy_from_slice(&1u16.to_le_bytes()); // one selected tab
    push_record(&mut workbook, WINDOW1, &window);

    let boundsheet_offset = workbook.len() + 4;
    let mut boundsheet = vec![0; 8];
    boundsheet[6] = 1; // cch
    boundsheet.extend_from_slice(b"S");
    push_record(&mut workbook, BOUNDSHEET8, &boundsheet);
    push_record(&mut workbook, EOF, &[]);

    let worksheet_offset = u32::try_from(workbook.len()).expect("fixture offset fits BIFF8");
    workbook[boundsheet_offset..boundsheet_offset + 4]
        .copy_from_slice(&worksheet_offset.to_le_bytes());
    push_bof(&mut workbook, 0x0010);
    workbook.extend_from_slice(worksheet_records);
    push_record(&mut workbook, EOF, &[]);

    let mut cfb = OleWriter::new();
    cfb.create_stream(&["Workbook"], &workbook).unwrap();
    let mut output = Cursor::new(Vec::new());
    cfb.write_to(&mut output).unwrap();
    output.into_inner()
}

fn open(worksheet_records: &[u8]) -> litchi_xls::Result<Workbook<Cursor<Vec<u8>>>> {
    Workbook::new(Cursor::new(workbook_with(worksheet_records)))
}

#[test]
fn hand_authored_single_cell_array_and_intervening_string_are_publicly_bound() {
    // MS-XLS 2.1: Formula [Array] [String]. The Array must not consume or
    // detach the cached String that follows it.
    let mut records = Vec::new();
    push_dimensions(&mut records, 4, 5, 5, 6);
    push_record(
        &mut records,
        FORMULA,
        &formula(4, 5, (4, 5), [0, 0, 0, 0, 0, 0, 0xff, 0xff]),
    );
    push_record(&mut records, ARRAY, &array((4, 5), (4, 5), 7));
    push_record(
        &mut records,
        STRING,
        &[5, 0, 0, b'a', b'r', b'r', b'a', b'y'],
    );

    let workbook = open(&records).unwrap();
    let sheet = workbook.xls_worksheet(0).unwrap();
    let owner = sheet.array_formulas().next().unwrap();
    let cell = sheet.get_cell(4, 5).unwrap();

    assert_eq!(sheet.array_formulas().len(), 1);
    assert_eq!(owner.anchor().row(), 4);
    assert_eq!(owner.anchor().col(), 5);
    assert_eq!(owner.tokens(), &[PTG_INT, 7, 0]);
    assert_eq!(owner.cell_count(), 1);
    assert!(cell.is_array_formula());
    assert_eq!(cell.formula(), Some("=7"));
    assert!(matches!(cell.value(), CellValue::String(value) if value == "array"));
}

#[test]
fn hand_authored_two_by_two_array_binds_every_formula_cell() {
    let anchor = (2, 3);
    let mut records = Vec::new();
    push_dimensions(&mut records, 2, 4, 3, 5);
    push_record(
        &mut records,
        FORMULA,
        &formula(2, 3, anchor, 7.0f64.to_le_bytes()),
    );
    push_record(&mut records, ARRAY, &array(anchor, (3, 4), 7));
    for (row, col) in [(2, 4), (3, 3), (3, 4)] {
        push_record(
            &mut records,
            FORMULA,
            &formula(row, col, anchor, 7.0f64.to_le_bytes()),
        );
    }

    let workbook = open(&records).unwrap();
    let sheet = workbook.xls_worksheet(0).unwrap();
    let owner = sheet.array_formulas().next().unwrap();

    assert_eq!(owner.range().first().row(), 2);
    assert_eq!(owner.range().first().col(), 3);
    assert_eq!(owner.range().last().row(), 3);
    assert_eq!(owner.range().last().col(), 4);
    assert_eq!(owner.cells().count(), 4);
    for (row, col) in [(2u32, 3u32), (2, 4), (3, 3), (3, 4)] {
        let cell = sheet.get_cell(row, col).unwrap();
        assert!(cell.is_array_formula(), "missing binding at ({row}, {col})");
        assert_eq!(cell.array_formula().unwrap().range(), owner.range());
        assert_eq!(
            sheet.array_formula_at(row, col).unwrap().range(),
            owner.range()
        );
    }
}

#[test]
fn malformed_array_links_are_rejected_through_the_public_reader() {
    let anchor = (0, 2);

    let mut orphan = Vec::new();
    push_dimensions(&mut orphan, 0, 1, 2, 3);
    push_record(&mut orphan, FORMULA, &scalar_formula(0, 2, 7));
    push_record(&mut orphan, ARRAY, &array(anchor, anchor, 7));

    let mut incomplete = Vec::new();
    push_dimensions(&mut incomplete, 0, 2, 2, 4);
    push_record(
        &mut incomplete,
        FORMULA,
        &formula(0, 2, anchor, 0.0f64.to_le_bytes()),
    );
    push_record(&mut incomplete, ARRAY, &array(anchor, (1, 3), 7));
    for (row, col) in [(0, 3), (1, 2)] {
        push_record(
            &mut incomplete,
            FORMULA,
            &formula(row, col, anchor, 0.0f64.to_le_bytes()),
        );
    }

    let mut wrong_anchor = Vec::new();
    push_dimensions(&mut wrong_anchor, 0, 1, 2, 4);
    push_record(
        &mut wrong_anchor,
        FORMULA,
        &formula(0, 2, (0, 3), 0.0f64.to_le_bytes()),
    );
    push_record(&mut wrong_anchor, ARRAY, &array(anchor, anchor, 7));

    for (name, malformed) in [
        ("incomplete", &incomplete),
        ("wrong anchor", &wrong_anchor),
        ("orphan", &orphan),
    ] {
        match open(malformed) {
            Err(_) => {},
            Ok(workbook) => assert!(
                workbook.xls_worksheet(0).is_err(),
                "accepted malformed {name} link"
            ),
        }
    }
}

#[test]
fn public_writer_array_formula_roundtrips_through_the_reader() {
    let mut writer = Writer::new();
    let sheet = writer.add_worksheet("S").unwrap();
    let range = Range::try_new(6, 1, 7, 2).unwrap();
    writer.write_array_formula(sheet, range, "A1+7").unwrap();

    let mut output = Cursor::new(Vec::new());
    writer.write_to(&mut output).unwrap();
    let workbook = Workbook::new(Cursor::new(output.into_inner())).unwrap();
    let sheet = workbook.xls_worksheet(0).unwrap();

    assert_eq!(sheet.array_formulas().len(), 1);
    assert_eq!(sheet.array_formulas().next().unwrap().range(), range);
    for (row, col) in [(6, 1), (6, 2), (7, 1), (7, 2)] {
        assert!(sheet.get_cell(row, col).unwrap().is_array_formula());
    }
}
