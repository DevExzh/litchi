use std::io::Cursor;

use crate::{OleFile, writer::OleWriter};

fn sample_file() -> Vec<u8> {
    let mut writer = OleWriter::new();
    writer.create_stream(&["Data"], b"payload").unwrap();
    let mut output = Cursor::new(Vec::new());
    writer.write_to(&mut output).unwrap();
    output.into_inner()
}

fn read_u32(bytes: &[u8], offset: usize) -> u32 {
    u32::from_le_bytes(bytes[offset..offset + 4].try_into().unwrap())
}

fn write_u32(bytes: &mut [u8], offset: usize, value: u32) {
    bytes[offset..offset + 4].copy_from_slice(&value.to_le_bytes());
}

#[test]
fn rejects_version_3_directory_sector_count() {
    let mut bytes = sample_file();
    write_u32(&mut bytes, 0x28, 1);
    assert!(OleFile::open(Cursor::new(bytes)).is_err());
}

#[test]
fn rejects_difat_start_count_disagreement() {
    let mut bytes = sample_file();
    write_u32(&mut bytes, 0x48, 1);
    write_u32(&mut bytes, 0x44, 0xFFFF_FFFE);
    assert!(OleFile::open(Cursor::new(bytes)).is_err());
}

#[test]
fn rejects_directory_and_fat_sector_overlap() {
    let mut bytes = sample_file();
    let first_fat_sector = read_u32(&bytes, 0x4C);
    write_u32(&mut bytes, 0x30, first_fat_sector);
    assert!(OleFile::open(Cursor::new(bytes)).is_err());
}

#[test]
fn rejects_fat_sector_without_fatsect_marker() {
    let mut bytes = sample_file();
    let sector_size = 1usize << u16::from_le_bytes([bytes[0x1E], bytes[0x1F]]);
    let first_fat_sector = read_u32(&bytes, 0x4C) as usize;
    let fat_offset = (first_fat_sector + 1) * sector_size + first_fat_sector * 4;
    write_u32(&mut bytes, fat_offset, 0xFFFF_FFFF);
    assert!(OleFile::open(Cursor::new(bytes)).is_err());
}

#[test]
fn rejects_incorrect_minifat_sector_count() {
    let mut bytes = sample_file();
    let count = read_u32(&bytes, 0x40);
    assert!(count > 0, "sample must contain a MiniFAT");
    write_u32(&mut bytes, 0x40, count + 1);
    assert!(OleFile::open(Cursor::new(bytes)).is_err());
}

#[test]
fn opens_real_world_word_compound_files() {
    for (name, bytes) in [
        (
            "FancyFoot.doc",
            include_bytes!("../../../test-data/ole/doc/FancyFoot.doc").as_slice(),
        ),
        (
            "Lists.doc",
            include_bytes!("../../../test-data/ole/doc/Lists.doc").as_slice(),
        ),
    ] {
        OleFile::open(Cursor::new(bytes))
            .unwrap_or_else(|error| panic!("failed to open {name}: {error}"));
    }
}

#[test]
fn opens_the_complete_legacy_office_compound_file_corpus() {
    fn collect(dir: &std::path::Path, files: &mut Vec<std::path::PathBuf>) {
        for entry in std::fs::read_dir(dir)
            .unwrap_or_else(|error| panic!("failed to read {}: {error}", dir.display()))
        {
            let path = entry.unwrap().path();
            if path.is_dir() {
                collect(&path, files);
            } else if path
                .extension()
                .and_then(|extension| extension.to_str())
                .is_some_and(|extension| {
                    extension.eq_ignore_ascii_case("doc")
                        || extension.eq_ignore_ascii_case("xls")
                        || extension.eq_ignore_ascii_case("ppt")
                })
            {
                files.push(path);
            }
        }
    }

    let root = std::path::Path::new(env!("CARGO_MANIFEST_DIR")).join("../../test-data/ole");
    let mut files = Vec::new();
    collect(&root, &mut files);
    files.sort();
    assert!(!files.is_empty(), "legacy Office fixture corpus is empty");

    for path in files {
        let bytes = std::fs::read(&path)
            .unwrap_or_else(|error| panic!("failed to read {}: {error}", path.display()));
        OleFile::open(Cursor::new(bytes))
            .unwrap_or_else(|error| panic!("failed to open {}: {error}", path.display()));
    }
}
