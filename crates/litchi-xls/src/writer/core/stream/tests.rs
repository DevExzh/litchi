use std::collections::HashMap;

use crate::Error;

use super::semantic::lookup_shared_string_index;

#[test]
fn missing_shared_string_mapping_returns_invalid_data() {
    let shared_strings = vec!["present".to_string()];
    let string_map = HashMap::new();

    let result = lookup_shared_string_index(&shared_strings, &string_map, "present");

    assert!(matches!(
        result,
        Err(Error::InvalidData(message)) if message.contains("missing from the shared string table")
    ));
}

#[test]
fn shared_string_mapping_must_match_table_entry() {
    let shared_strings = vec!["first".to_string(), "second".to_string()];
    let mut string_map = HashMap::new();
    string_map.insert("second".to_string(), 0);

    let result = lookup_shared_string_index(&shared_strings, &string_map, "second");

    assert!(matches!(
        result,
        Err(Error::InvalidData(message)) if message.contains("does not match the shared string table")
    ));
}
