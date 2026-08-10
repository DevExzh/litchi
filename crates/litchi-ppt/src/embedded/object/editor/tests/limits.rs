use super::super::{Collection, Editor};
use crate::package::Error;
use std::collections::{BTreeMap, HashSet};
use std::sync::Arc;

#[test]
fn rejects_a_source_above_the_publication_limit_before_parsing() {
    let error = Editor::open_records_arc_with_limit(Arc::from(vec![0u8; 2]), 1)
        .err()
        .expect("source above limit must be rejected");
    assert!(matches!(
        error,
        Error::ResourceLimit(message) if message.contains("editor source")
    ));
}

#[test]
fn rejects_a_projected_incremental_stream_above_the_publication_limit() {
    let mut staged_storage = BTreeMap::new();
    staged_storage.insert(1, vec![0; 8]);
    let editor = Editor {
        original: Arc::from(vec![0u8; 8]),
        max_output_bytes: 15,
        streams: Vec::new(),
        document_path: vec!["PowerPoint Document".into()],
        current_user_path: vec!["Current User".into()],
        document: vec![0; 8],
        current_user: vec![0; 20],
        mappings: BTreeMap::from([(1, 0)]),
        current_edit_offset: 0,
        document_persist_id: 1,
        collection: Collection {
            id_seed: 1,
            objects: Vec::new(),
            unknown_records: Vec::new(),
        },
        staged_storage,
        removed_persist_ids: HashSet::new(),
        rewrite_object_list: false,
        changed: true,
    };

    let error = editor.finish().unwrap_err();
    assert!(matches!(
        error,
        Error::ResourceLimit(message) if message.contains("incremental document stream")
    ));
}
