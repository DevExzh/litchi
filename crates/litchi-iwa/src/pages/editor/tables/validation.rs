//! Validation boundary for reachable Pages table identities.

use super::*;

use litchi_iwa_protos::table_model_discovery_codec::{
    DecodeOptions, TableModelSnapshot, decode_table_model,
};

use crate::archive::RawMessage;

/// Strictly project the table-model facts needed by body-table discovery.
pub(super) fn decode_table_models<'a>(
    messages: impl Iterator<Item = &'a RawMessage>,
    model_id: u64,
) -> Result<Vec<TableModelSnapshot<'a>>> {
    messages
        .map(|message| {
            let source = message.data.as_slice();
            decode_table_model(source, DecodeOptions::for_source(source)).map_err(|error| {
                Error::InvalidFormat(format!(
                    "Pages table model {model_id} contains malformed table-model payload: {error}"
                ))
            })
        })
        .collect()
}

/// Resolve a model identifier only when it is a validated body-owned table.
pub(super) fn validate_body_table(
    editor: &PagesEditor,
    model_object_id: u64,
) -> Result<PagesTableGraph> {
    body_table_graphs(editor)?
        .into_iter()
        .find(|graph| graph.info.model_object_id == model_object_id)
        .ok_or_else(|| {
            Error::ParseError(format!(
                "Pages table model {model_object_id} is not attached to the body"
            ))
        })
}

impl PagesEditor {
    /// Require a reachable body table before any table operation.
    pub(super) fn require_body_table(&self, model_object_id: u64) -> Result<PagesTableGraph> {
        validate_body_table(self, model_object_id)
    }
}
