use litchi_cfb::{OleWriter, SharedOleFile};
use litchi_core::OwnedSource;
use std::{io::Cursor, path::PathBuf, sync::Arc};

fn workbook_stream(bytes: Vec<u8>) -> Vec<u8> {
    SharedOleFile::open(Arc::new(OwnedSource::new(bytes)))
        .unwrap()
        .open_stream(&["Workbook"])
        .unwrap()
}

fn cfb_with_workbook(stream: &[u8]) -> Vec<u8> {
    let mut writer = OleWriter::new();
    writer.create_stream(&["Workbook"], stream).unwrap();
    let mut output = Vec::new();
    writer.write_to(&mut Cursor::new(&mut output)).unwrap();
    output
}

fn frame_bytes(kind: u16, payload: &[u8]) -> Vec<u8> {
    let mut output = Vec::with_capacity(4 + payload.len());
    output.extend_from_slice(&kind.to_le_bytes());
    output.extend_from_slice(&(u16::try_from(payload.len()).unwrap()).to_le_bytes());
    output.extend_from_slice(payload);
    output
}

fn number_frame(row: u16, column: u16, value: f64) -> Vec<u8> {
    let mut payload = Vec::with_capacity(14);
    payload.extend_from_slice(&row.to_le_bytes());
    payload.extend_from_slice(&column.to_le_bytes());
    payload.extend_from_slice(&0_u16.to_le_bytes());
    payload.extend_from_slice(&value.to_le_bytes());
    frame_bytes(0x0203, &payload)
}

fn first_sheet_offset(stream: &[u8]) -> usize {
    let mut offset = 0;
    while offset + 4 <= stream.len() {
        let kind = u16::from_le_bytes([stream[offset], stream[offset + 1]]);
        let length = usize::from(u16::from_le_bytes([stream[offset + 2], stream[offset + 3]]));
        if kind == 0x0085 {
            return usize::try_from(u32::from_le_bytes([
                stream[offset + 4],
                stream[offset + 5],
                stream[offset + 6],
                stream[offset + 7],
            ]))
            .unwrap();
        }
        offset += 4 + length;
    }
    panic!("workbook has no BoundSheet8");
}

fn worksheet_eof_offset(stream: &[u8], start: usize) -> usize {
    let mut offset = start;
    loop {
        assert!(offset + 4 <= stream.len());
        let kind = u16::from_le_bytes([stream[offset], stream[offset + 1]]);
        let length = usize::from(u16::from_le_bytes([stream[offset + 2], stream[offset + 3]]));
        if kind == 0x000A {
            return offset;
        }
        offset += 4 + length;
    }
}

fn patch_later_sheet_offsets(
    output: &mut [u8],
    first_sheet_start: usize,
    insertion_boundary: usize,
    extra_len: usize,
) {
    let delta = u32::try_from(extra_len).unwrap();
    let mut cursor = 0;
    while cursor < first_sheet_start {
        assert!(cursor + 4 <= output.len());
        let kind = u16::from_le_bytes([output[cursor], output[cursor + 1]]);
        let length = usize::from(u16::from_le_bytes([output[cursor + 2], output[cursor + 3]]));
        if kind == 0x0085 && cursor + 8 <= output.len() {
            let position = u32::from_le_bytes([
                output[cursor + 4],
                output[cursor + 5],
                output[cursor + 6],
                output[cursor + 7],
            ]);
            if usize::try_from(position).unwrap() > insertion_boundary {
                let shifted = position.checked_add(delta).unwrap();
                output[cursor + 4..cursor + 8].copy_from_slice(&shifted.to_le_bytes());
            }
        }
        cursor += 4 + length;
    }
}

fn insert_before_worksheet_eof(stream: &[u8], extra: &[u8]) -> Vec<u8> {
    let eof = worksheet_eof_offset(stream, first_sheet_offset(stream));
    let first_sheet_start = first_sheet_offset(stream);
    let mut output = Vec::with_capacity(stream.len() + extra.len());
    output.extend_from_slice(&stream[..eof]);
    output.extend_from_slice(extra);
    output.extend_from_slice(&stream[eof..]);
    patch_later_sheet_offsets(&mut output, first_sheet_start, eof, extra.len());
    output
}

fn main() -> Result<(), Box<dyn std::error::Error>> {
    let args: Vec<_> = std::env::args().skip(1).collect();
    if args.len() != 2 {
        return Err("usage: xls0686-generate SIMPLE_XLS OUTPUT_DIRECTORY".into());
    }
    let template = workbook_stream(std::fs::read(&args[0])?);
    let output = PathBuf::from(&args[1]);
    std::fs::create_dir_all(&output)?;
    for count in [70_000_u32, 100_000] {
        let mut original = template.clone();
        let start = first_sheet_offset(&original);
        let eof = worksheet_eof_offset(&original, start);
        let mut offset = start;
        while offset < eof {
            let kind = u16::from_le_bytes(original[offset..offset + 2].try_into()?);
            let length = usize::from(u16::from_le_bytes(
                original[offset + 2..offset + 4].try_into()?,
            ));
            if kind == 0x0200 && length == 14 {
                let last_row = count / 2 + 1;
                original[offset + 8..offset + 12].copy_from_slice(&last_row.to_le_bytes());
                original[offset + 14..offset + 16].copy_from_slice(&2_u16.to_le_bytes());
            }
            offset += 4 + length;
        }
        let mut extra = Vec::new();
        for index in 0..count {
            let row = u16::try_from(index / 2 + 1)?;
            let column = u16::try_from(index % 2)?;
            extra.extend_from_slice(&number_frame(row, column, f64::from(index + 1)));
        }
        let bytes = cfb_with_workbook(&insert_before_worksheet_eof(&original, &extra));
        let path = output.join(format!("numeric-{count}.xls"));
        std::fs::write(&path, &bytes)?;
        println!("{}\t{}\t{}", path.display(), count + 1, bytes.len());
    }
    Ok(())
}
