//! Byte-preserving protobuf mutations for embedded merge formula storage.

use prost::Message;

use super::*;

const TABLE_MODEL_MERGE_OWNER_FIELD: u32 = 47;
const TABLE_MODEL_MESSAGE_TYPES: &[u32] = &[6_000, 6_001];
const MERGE_OWNER_FORMULA_STORE_FIELD: u32 = 2;
const FORMULA_STORE_NEXT_INDEX_FIELD: u32 = 2;
const FORMULA_STORE_FORMULAS_FIELD: u32 = 3;

pub(super) fn append_formula(
    store_data: &[u8],
    next_formula_index: u32,
    pair_data: Vec<u8>,
) -> Result<Vec<u8>> {
    let current = repeated_length_delimited_payloads(store_data, FORMULA_STORE_FORMULAS_FIELD)?;
    let mut formulas = Vec::with_capacity(current.len() + 1);
    formulas.extend(current.into_iter().map(<[u8]>::to_vec));
    formulas.push(pair_data);
    let data = patch_varint_field(
        store_data,
        FORMULA_STORE_NEXT_INDEX_FIELD,
        true,
        Some(u64::from(next_formula_index)),
    )?;
    rewrite_repeated_length_delimited_fields(&data, FORMULA_STORE_FORMULAS_FIELD, &formulas)
}

pub(super) fn remove_formula(store_data: &[u8], remove_index: u32) -> Result<Vec<u8>> {
    let mut found = false;
    let payloads = repeated_length_delimited_payloads(store_data, FORMULA_STORE_FORMULAS_FIELD)?;
    let mut formulas = Vec::with_capacity(payloads.len().saturating_sub(1));
    for payload in payloads {
        let pair = tst::formula_store_archive::FormulaStorePair::decode(payload)?;
        if pair.formula_index == remove_index {
            if found {
                return Err(Error::InvalidFormat(format!(
                    "iWork merge formula {remove_index} occurs more than once in wire storage"
                )));
            }
            found = true;
        } else {
            formulas.push(payload.to_vec());
        }
    }
    if !found {
        return Err(Error::InvalidFormat(format!(
            "iWork merge formula {remove_index} is missing from wire storage"
        )));
    }
    rewrite_repeated_length_delimited_fields(store_data, FORMULA_STORE_FORMULAS_FIELD, &formulas)
}

pub(super) fn patch_table_model(
    package: &mut IWorkPackage,
    table_id: u64,
    patch: impl FnOnce(&[u8]) -> Result<Vec<u8>>,
) -> Result<()> {
    let locations = object_locations(package)?;
    let archive_name = locations.get(&table_id).ok_or_else(|| {
        Error::InvalidFormat(format!("iWork table model object {table_id} is missing"))
    })?;
    package.update_archive(archive_name, |archive| {
        let object = archive.object_mut(table_id).ok_or_else(|| {
            Error::InvalidFormat(format!("iWork table model object {table_id} is missing"))
        })?;
        let message_index = object
            .messages
            .iter()
            .position(|message| {
                TABLE_MODEL_MESSAGE_TYPES.contains(&message.type_)
                    && TableModelArchive::decode(message.data.as_slice()).is_ok()
            })
            .ok_or_else(|| {
                Error::InvalidFormat(format!(
                    "Object {table_id} has no iWork table-model payload"
                ))
            })?;
        let message_type = object.messages[message_index].type_;
        let data = patch(object.messages[message_index].data.as_slice())?;
        TableModelArchive::decode(data.as_slice())?;
        object.replace_message(
            message_index,
            RawMessage {
                type_: message_type,
                data,
            },
        )?;
        Ok(())
    })
}

pub(super) fn add_merge_owner(original: &[u8], owner: &tst::MergeOwnerArchive) -> Result<Vec<u8>> {
    patch_length_delimited_field(
        original,
        TABLE_MODEL_MERGE_OWNER_FIELD,
        false,
        Some(&owner.encode_to_vec()),
    )
}

pub(super) fn remove_merge_owner(original: &[u8]) -> Result<Vec<u8>> {
    patch_length_delimited_field(original, TABLE_MODEL_MERGE_OWNER_FIELD, true, None)
}

pub(super) fn transform_merge_owner(
    original: &[u8],
    transform: impl FnOnce(&[u8]) -> Result<Vec<u8>>,
) -> Result<Vec<u8>> {
    transform_length_delimited_field(original, TABLE_MODEL_MERGE_OWNER_FIELD, transform)
}

pub(super) fn add_formula_store(
    owner_data: &[u8],
    store: &tst::FormulaStoreArchive,
) -> Result<Vec<u8>> {
    patch_length_delimited_field(
        owner_data,
        MERGE_OWNER_FORMULA_STORE_FIELD,
        false,
        Some(&store.encode_to_vec()),
    )
}

pub(super) fn transform_formula_store(
    owner_data: &[u8],
    transform: impl FnOnce(&[u8]) -> Result<Vec<u8>>,
) -> Result<Vec<u8>> {
    transform_length_delimited_field(owner_data, MERGE_OWNER_FORMULA_STORE_FIELD, transform)
}
