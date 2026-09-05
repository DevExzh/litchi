//! Validation boundary for reachable Pages table identities.

use super::*;

use litchi_iwa_protos::table_model_discovery_codec::{
    DecodeOptions, TableModelSnapshot, decode_table_model,
};

use crate::archive::RawMessage;
use litchi_iwa_common::{LimitKind, WireLimits};
use litchi_iwa_protos::table_info_codec;

/// Borrow only the model, parent and lock branches needed by discovery.
/// Unrelated drawable envelopes remain opaque source bytes.
pub(super) fn decode_table_info_ownership(
    source: &[u8],
) -> Result<table_info_codec::TableInfoSnapshot> {
    let options = table_info_codec::DecodeOptions::for_source(source);
    let options = options
        .with_max_message_bytes(options.max_message_bytes().min(WireLimits::MAX_INPUT_BYTES))
        .with_max_fields(options.max_fields().min(WireLimits::MAX_FIELDS))
        .with_max_work_bytes(options.max_work_bytes().min(WireLimits::MAX_REWRITE_WORK));
    table_info_codec::decode_table_info_with_parent(source, options).map_err(|error| {
        let limit = if let Some((observed, maximum)) = error.field_limit_values() {
            Some((LimitKind::Fields, observed, maximum))
        } else if let Some((observed, maximum)) = error.work_limit_values() {
            Some((LimitKind::RewriteWork, observed, maximum))
        } else {
            match error.wire_resource_limit() {
                Some(table_info_codec::WireResourceLimit::Bytes { observed, maximum }) => Some((
                    LimitKind::InputBytes,
                    observed.unwrap_or(usize::MAX),
                    maximum.unwrap_or(options.max_message_bytes()),
                )),
                Some(table_info_codec::WireResourceLimit::Nesting { observed, maximum }) => Some((
                    LimitKind::Nesting,
                    observed
                        .and_then(|value| usize::try_from(value).ok())
                        .unwrap_or(usize::MAX),
                    usize::try_from(maximum.unwrap_or(options.recursion_limit()))
                        .unwrap_or(usize::MAX),
                )),
                _ => None,
            }
        };
        if let Some((kind, observed, limit)) = limit {
            Error::IwaCommon(litchi_iwa_common::Error::LimitExceeded {
                kind,
                observed,
                limit,
            })
        } else {
            Error::InvalidFormat(format!(
                "Pages table-info ownership failed strict validation: {error}"
            ))
        }
    })
}

/// Strictly project the table-model facts needed by body-table discovery.
pub(super) fn decode_unique_table_model<'a>(
    messages: impl Iterator<Item = &'a RawMessage>,
    model_id: u64,
) -> Result<TableModelSnapshot<'a>> {
    let mut selected = None;
    for message in messages {
        let source = message.data.as_slice();
        let model =
            decode_table_model(source, DecodeOptions::for_source(source)).map_err(|error| {
                Error::InvalidFormat(format!(
                    "Pages table model {model_id} contains malformed table-model payload: {error}"
                ))
            })?;
        if selected.replace(model).is_some() {
            return Err(Error::InvalidFormat(format!(
                "Pages table model {model_id} must contain exactly one table-model payload"
            )));
        }
    }
    selected.ok_or_else(|| {
        Error::InvalidFormat(format!(
            "Pages table model {model_id} must contain exactly one table-model payload"
        ))
    })
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
