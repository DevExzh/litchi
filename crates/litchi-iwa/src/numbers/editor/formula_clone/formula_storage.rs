//! Formula-list UUID rewrites for cloned table storage.

use super::*;

pub(in crate::numbers::editor) fn remap_cloned_formula_storage(
    object: &mut ArchiveObject,
    source_table_uuid: &str,
    new_table_uuid: &str,
) -> Result<()> {
    let source_uuid = parse_table_uuid(source_table_uuid)?;
    let new_uuid = parse_table_uuid(new_table_uuid)?;
    remap_formula_storage(object, &source_uuid, &new_uuid)
}

pub(in crate::numbers::editor) fn remap_cloned_formula_owner_storage(
    object: &mut ArchiveObject,
    source_owner_uuid: &tsp::Uuid,
    new_owner_uuid: &tsp::Uuid,
) -> Result<()> {
    remap_formula_storage(object, source_owner_uuid, new_owner_uuid)
}

fn remap_formula_storage(
    object: &mut ArchiveObject,
    source_uuid: &tsp::Uuid,
    new_uuid: &tsp::Uuid,
) -> Result<()> {
    for index in 0..object.messages.len() {
        let message = object.messages[index].clone();
        let list_type = TableDataList::decode(message.data.as_slice())
            .ok()
            .and_then(|list| tst::table_data_list::ListType::try_from(list.list_type).ok())
            .or_else(|| {
                TableDataListSegment::decode(message.data.as_slice())
                    .ok()
                    .and_then(|list| tst::table_data_list::ListType::try_from(list.list_type).ok())
            });
        if list_type != Some(tst::table_data_list::ListType::Formula) {
            continue;
        }
        let data = transform_length_delimited_fields_at_path(
            message.data.as_slice(),
            &[3, 5],
            |formula| remap_formula_wire(formula, source_uuid, new_uuid),
        )?;
        object.replace_message(
            index,
            RawMessage {
                type_: message.type_,
                data,
            },
        )?;
    }
    Ok(())
}

fn remap_formula_wire(
    original: &[u8],
    source: &tsp::Uuid,
    replacement: &tsp::Uuid,
) -> Result<Vec<u8>> {
    let previous = tsce::FormulaArchive::decode(original)?;
    let mut data = original.to_vec();
    if previous.host_table_uid.as_ref() == Some(source) {
        data = patch_uuid_at_path(&data, &[7], replacement)?;
    }
    data = transform_length_delimited_field(&data, 1, |ast| {
        remap_ast_array_wire(ast, source, replacement)
    })?;
    let verified = tsce::FormulaArchive::decode(data.as_slice())?;
    if previous.host_table_uid.as_ref() == Some(source)
        && verified.host_table_uid.as_ref() != Some(replacement)
    {
        return Err(Error::InvalidFormat(
            "Numbers formula host UUID clone failed validation".to_owned(),
        ));
    }
    Ok(data)
}

fn remap_ast_array_wire(
    data: &[u8],
    source: &tsp::Uuid,
    replacement: &tsp::Uuid,
) -> Result<Vec<u8>> {
    transform_length_delimited_fields_at_path(data, &[1], |node| {
        let mut node = transform_length_delimited_fields_at_path(node, &[16, 5], |uuid| {
            remap_cfuuid_wire(uuid, source, replacement)
        })?;
        node = transform_length_delimited_fields_at_path(&node, &[28, 1], |uuid| {
            remap_cfuuid_wire(uuid, source, replacement)
        })?;
        transform_length_delimited_fields_at_path(&node, &[14], |nested| {
            remap_ast_array_wire(nested, source, replacement)
        })
    })
}

fn remap_cfuuid_wire(data: &[u8], source: &tsp::Uuid, replacement: &tsp::Uuid) -> Result<Vec<u8>> {
    let decoded = tsp::CfuuidArchive::decode(data)?;
    if cfuuid_key(&decoded) != Some(uuid_key(source)) {
        return Ok(data.to_vec());
    }
    let mut data = data.to_vec();
    if decoded.uuid_bytes.is_some() {
        data = crate::wire::patch_length_delimited_field(
            &data,
            1,
            true,
            Some(&uuid_bytes(replacement)),
        )?;
    }
    for (field, present, value) in [
        (2, decoded.uuid_w0.is_some(), replacement.lower as u32),
        (
            3,
            decoded.uuid_w1.is_some(),
            (replacement.lower >> 32) as u32,
        ),
        (4, decoded.uuid_w2.is_some(), replacement.upper as u32),
        (
            5,
            decoded.uuid_w3.is_some(),
            (replacement.upper >> 32) as u32,
        ),
    ] {
        if present {
            data = patch_varint_field(&data, field, true, Some(u64::from(value)))?;
        }
    }
    Ok(data)
}

fn patch_uuid_at_path(data: &[u8], path: &[u32], uuid: &tsp::Uuid) -> Result<Vec<u8>> {
    transform_length_delimited_fields_at_path(data, path, |uuid_data| {
        patch_uuid_wire(uuid_data, uuid)
    })
}

fn patch_uuid_wire(data: &[u8], uuid: &tsp::Uuid) -> Result<Vec<u8>> {
    let data = patch_varint_field(data, 1, true, Some(uuid.lower))?;
    patch_varint_field(&data, 2, true, Some(uuid.upper))
}

fn cfuuid_key(uuid: &tsp::CfuuidArchive) -> Option<(u64, u64)> {
    Some((
        u64::from(uuid.uuid_w0?) | (u64::from(uuid.uuid_w1?) << 32),
        u64::from(uuid.uuid_w2?) | (u64::from(uuid.uuid_w3?) << 32),
    ))
}

fn uuid_bytes(uuid: &tsp::Uuid) -> [u8; 16] {
    let value = (u128::from(uuid.upper) << 64) | u128::from(uuid.lower);
    value.to_be_bytes()
}
