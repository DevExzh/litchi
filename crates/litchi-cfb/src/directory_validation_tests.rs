#![allow(
    clippy::unwrap_used,
    clippy::expect_used,
    reason = "test assertions panic on failure by design"
)]

use crate::consts::{DIRENTRY_SIZE, NOSTREAM};
use crate::file::OleFile;
use crate::writer::DirectoryBuilder;
use std::io::Cursor;

fn valid_directory(count: u32) -> Vec<u8> {
    let mut builder = DirectoryBuilder::new(NOSTREAM, 0);
    for index in 0..count {
        builder
            .add_stream(format!("stream-{index:03}"), NOSTREAM, 0)
            .unwrap();
    }
    builder.generate_directory_stream().unwrap()
}

fn validate(data: &[u8]) -> Result<(), crate::OleError> {
    OleFile::<Cursor<Vec<u8>>>::validate_directory(data, 512)
}

fn read_sid(data: &[u8], entry_sid: u32, offset: usize) -> u32 {
    let start = entry_sid as usize * DIRENTRY_SIZE + offset;
    u32::from_le_bytes(data[start..start + 4].try_into().unwrap())
}

#[test]
fn accepts_writer_generated_red_black_directories() {
    for count in 0..=64 {
        validate(&valid_directory(count)).unwrap();
    }
}

#[test]
fn accepts_interoperable_root_color_values() {
    for color in [0, 1] {
        let mut directory = valid_directory(3);
        directory[67] = color;
        validate(&directory).unwrap();
    }
}

#[test]
fn rejects_invalid_names_types_and_colors() {
    let mut invalid_name_length = valid_directory(1);
    invalid_name_length[DIRENTRY_SIZE + 64] = 3;
    assert!(validate(&invalid_name_length).is_err());

    let mut invalid_type = valid_directory(1);
    invalid_type[DIRENTRY_SIZE + 66] = 3;
    assert!(validate(&invalid_type).is_err());

    let mut invalid_color = valid_directory(1);
    invalid_color[DIRENTRY_SIZE + 67] = 2;
    assert!(validate(&invalid_color).is_err());
}

#[test]
fn accepts_non_black_tree_roots_but_rejects_invalid_ordering() {
    let mut red_root = valid_directory(3);
    let tree_root = read_sid(&red_root, 0, 76);
    red_root[tree_root as usize * DIRENTRY_SIZE + 67] = 0;
    assert!(validate(&red_root).is_ok());

    let mut reversed = valid_directory(3);
    let reversed_root = read_sid(&reversed, 0, 76);
    let left = read_sid(&reversed, reversed_root, 68);
    let right = read_sid(&reversed, reversed_root, 72);
    let root_offset = reversed_root as usize * DIRENTRY_SIZE;
    reversed[root_offset + 68..root_offset + 72].copy_from_slice(&right.to_le_bytes());
    reversed[root_offset + 72..root_offset + 76].copy_from_slice(&left.to_le_bytes());
    assert!(validate(&reversed).is_err());
}

#[test]
fn accepts_black_height_mismatch_but_rejects_stream_children() {
    let mut mismatched = valid_directory(3);
    let tree_root = read_sid(&mismatched, 0, 76);
    let left = read_sid(&mismatched, tree_root, 68);
    mismatched[left as usize * DIRENTRY_SIZE + 67] = 1;
    assert!(validate(&mismatched).is_ok());

    let mut stream_child = valid_directory(1);
    let stream_sid = read_sid(&stream_child, 0, 76);
    let offset = stream_sid as usize * DIRENTRY_SIZE + 76;
    stream_child[offset..offset + 4].copy_from_slice(&stream_sid.to_le_bytes());
    assert!(validate(&stream_child).is_err());
}
