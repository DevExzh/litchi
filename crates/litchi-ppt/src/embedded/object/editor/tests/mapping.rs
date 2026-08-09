use super::super::{mapping, rewrite};
use super::ppt_record;
use crate::writer::{PersistPtrBuilder, UserEditAtom};

#[test]
fn merges_newest_incremental_mapping_over_prior_edit() {
    let object1 = ppt_record(0, 0x1111, b"one");
    let object2 = ppt_record(0, 0x2222, b"two");
    let mut document = object1.clone();
    document.extend_from_slice(&object2);
    let mut first_dir = PersistPtrBuilder::new();
    first_dir.set_offset(1, 0);
    let first_dir_offset = u32::try_from(document.len()).unwrap();
    document.extend_from_slice(&first_dir.generate_full_record());
    let first_edit_offset = u32::try_from(document.len()).unwrap();
    document
        .extend_from_slice(&UserEditAtom::new_minimal(first_dir_offset, 1, 1, 0).generate_record());
    let replacement_offset = u32::try_from(document.len()).unwrap();
    document.extend_from_slice(&object2);
    let mut second_dir = PersistPtrBuilder::new();
    second_dir.set_offset(1, replacement_offset);
    let second_dir_offset = u32::try_from(document.len()).unwrap();
    document.extend_from_slice(&second_dir.generate_incremental_record());
    let mut edit = UserEditAtom::new_minimal(second_dir_offset, 1, 1, 0);
    edit.offset_last_edit = first_edit_offset;
    let second_edit_offset = u32::try_from(document.len()).unwrap();
    document.extend_from_slice(&edit.generate_record());
    let (mapping, document_id) = mapping::read(&document, second_edit_offset).unwrap();
    assert_eq!(document_id, 1);
    assert_eq!(mapping.get(&1), Some(&replacement_offset));
    assert_eq!(rewrite::type_of(&document[..8]).unwrap(), 0x1111);
}
