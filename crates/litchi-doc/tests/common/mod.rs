//! Shared DOC fixture normalization for integration tests.

use litchi_doc::parts::fib::FileInformationBlock;
use litchi_doc::writer::Writer;
use litchi_ole_common::object::{Editor as ObjectEditor, Limits, Targets};
use std::io::Cursor;

const DOP_INDEX: usize = 31;
const FIB_POINTERS: usize = 154;
const WORD97_DOP_LENGTH: usize = 500;

fn extract_dop(bytes: Vec<u8>) -> (Vec<u8>, FileInformationBlock) {
    let package = ObjectEditor::open(bytes, Targets::default(), Limits::default())
        .expect("DOC fixture object editor");
    let word_path = ["WordDocument".to_string()];
    let word = package.stream(&word_path).expect("WordDocument stream");
    let fib = FileInformationBlock::parse(word).expect("DOC fixture FIB");
    let table_name = if fib.which_table_stream() {
        "1Table"
    } else {
        "0Table"
    };
    let table_path = [table_name.to_string()];
    let table = package
        .stream(&table_path)
        .expect("DOC fixture table stream");
    let (offset, length) = fib
        .get_table_pointer(DOP_INDEX)
        .expect("DOC fixture DOP pointer");
    let offset = usize::try_from(offset).expect("DOC fixture DOP offset");
    let length = usize::try_from(length).expect("DOC fixture DOP length");
    let end = offset.checked_add(length).expect("DOC fixture DOP end");
    (table[offset..end].to_vec(), fib)
}

/// Returns a canonical, valid Word 97 DOP payload from the checked-in writer.
pub(crate) fn valid_word97_dop() -> Vec<u8> {
    let mut writer = Writer::new();
    writer
        .add_paragraph("DOP fixture")
        .expect("DOP fixture paragraph");
    let mut output = Cursor::new(Vec::new());
    writer.write_to(&mut output).expect("DOP fixture DOC");
    let (mut dop, _) = extract_dop(output.into_inner());
    dop.truncate(WORD97_DOP_LENGTH);
    assert_eq!(dop.len(), WORD97_DOP_LENGTH);
    dop
}

fn valid_word2002_dop() -> Vec<u8> {
    let mut dop = valid_word97_dop();
    dop.resize(594, 0);
    dop
}

fn valid_modern_dop() -> Vec<u8> {
    let path = std::path::Path::new(env!("CARGO_MANIFEST_DIR"))
        .join("../../test-data/ole/doc/NoHeadFoot.doc");
    let (dop, _) = extract_dop(std::fs::read(path).expect("modern DOP fixture"));
    assert!(matches!(dop.len(), 674 | 690 | 694));
    dop
}

fn expected_pointer_count(nfib: u16) -> Option<usize> {
    match nfib {
        0x00C1 => Some(0x5D),
        0x00D9 => Some(0x6C),
        0x0101 => Some(0x88),
        0x010C => Some(0xA4),
        0x0112 => Some(0xB7),
        _ => None,
    }
}

fn effective_fib_version(word: &[u8], count: usize) -> Option<(u16, u16)> {
    let pointer_end = FIB_POINTERS.checked_add(count.checked_mul(8)?)?;
    let csw_new = u16::from_le_bytes(word.get(pointer_end..pointer_end + 2)?.try_into().ok()?);
    let nfib = if csw_new == 0 {
        u16::from_le_bytes(word.get(2..4)?.try_into().ok()?)
    } else {
        u16::from_le_bytes(
            word.get(pointer_end + 2..pointer_end + 4)?
                .try_into()
                .ok()?,
        )
    };
    Some((nfib, csw_new))
}

fn normalize_fib_shape(word: &mut [u8]) {
    let Some(count_bytes) = word.get(152..154) else {
        return;
    };
    let count = usize::from(u16::from_le_bytes([count_bytes[0], count_bytes[1]]));
    let Some((effective, csw_new)) = effective_fib_version(word, count) else {
        return;
    };
    let expected_csw_new = match effective {
        0x00C1 => Some(0),
        0x00D9 | 0x0101 | 0x010C => Some(2),
        0x0112 => Some(5),
        _ => None,
    };
    if expected_pointer_count(effective) == Some(count) && expected_csw_new == Some(csw_new) {
        return;
    }
    // Several checked-in producer fixtures carry an overlong legacy
    // FibRgFcLcb while retaining Word 97's base nFib. Keep the physical
    // bytes and all pointers used by the editor, but expose the valid modern
    // counted prefix consumed by this test's edit path.
    if count < 136 {
        return;
    }
    let pointer_end = FIB_POINTERS + 136 * 8;
    if word.len() < pointer_end + 4 {
        return;
    }
    word[152..154].copy_from_slice(&136u16.to_le_bytes());
    word[pointer_end..pointer_end + 2].copy_from_slice(&2u16.to_le_bytes());
    word[pointer_end + 2..pointer_end + 4].copy_from_slice(&0x0101u16.to_le_bytes());
}

/// Replaces a fixture's DOP with a valid payload while preserving every other
/// stream and byte range. This is used only by edit tests whose fixture intent
/// is unrelated to malformed producer-specific DOP lengths.
pub(crate) fn with_valid_word97_dop(bytes: Vec<u8>) -> Vec<u8> {
    let word_path = ["WordDocument".to_string()];
    let mut package = ObjectEditor::open(bytes, Targets::default(), Limits::default())
        .expect("DOC fixture object editor");
    let original_word = package.stream(&word_path).expect("WordDocument stream");
    let mut word = original_word.to_vec();
    normalize_fib_shape(&mut word);
    let fib = FileInformationBlock::parse(&word).expect("DOC fixture FIB");
    let table_name = if fib.which_table_stream() {
        "1Table"
    } else {
        "0Table"
    };
    let table_path = [table_name.to_string()];
    let mut table = package
        .stream(&table_path)
        .expect("DOC fixture table stream")
        .to_vec();
    let (offset, length) = fib
        .get_table_pointer(DOP_INDEX)
        .expect("DOC fixture DOP pointer");
    let offset = usize::try_from(offset).expect("DOC fixture DOP offset");
    let length = usize::try_from(length).expect("DOC fixture DOP length");
    let effective = effective_fib_version(
        &word,
        fib.table_pointer_count().expect("DOC fixture FIB count"),
    )
    .map(|(nfib, _)| nfib)
    .unwrap_or(0x0101);
    let replacement = if effective == 0x0112 {
        Some(valid_modern_dop())
    } else if effective == 0x0101 {
        Some(valid_word2002_dop())
    } else {
        Some(valid_word97_dop())
    };
    if let Some(dop) = replacement {
        if length < dop.len() {
            let insertion = offset
                .checked_add(length)
                .expect("DOC fixture DOP insertion");
            let extra = dop.len() - length;
            table.splice(insertion..insertion, std::iter::repeat_n(0, extra));
            let count = fib.table_pointer_count().expect("DOC fixture FIB count");
            for index in 0..count {
                let pointer = FIB_POINTERS + index * 8;
                let current = usize::try_from(u32::from_le_bytes(
                    word[pointer..pointer + 4]
                        .try_into()
                        .expect("DOC fixture FIB pointer"),
                ))
                .expect("DOC fixture FIB pointer");
                if current >= insertion {
                    let shifted =
                        u32::try_from(current + extra).expect("shifted DOC fixture FIB pointer");
                    word[pointer..pointer + 4].copy_from_slice(&shifted.to_le_bytes());
                }
            }
        }
        let end = offset.checked_add(dop.len()).expect("DOC fixture DOP end");
        assert!(table.len() >= end);
        table[offset..end].copy_from_slice(&dop);
        let pointer = FIB_POINTERS + DOP_INDEX * 8;
        word[pointer + 4..pointer + 8].copy_from_slice(&(dop.len() as u32).to_le_bytes());
    }
    package
        .put_stream(&word_path, word)
        .expect("normalized WordDocument stream");
    package
        .put_stream(&table_path, table)
        .expect("normalized table stream");
    package
        .finish()
        .expect("normalized DOC fixture publication")
}
