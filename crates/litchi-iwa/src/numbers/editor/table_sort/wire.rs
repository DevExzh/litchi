//! Lossless protobuf wire handling for Numbers table sort orders.

use super::*;

use litchi_iwa_protos::table_sort_order_codec as codec;

fn codec_options(source: &[u8], model: &TableModelArchive) -> codec::DecodeOptions {
    let columns = usize::try_from(model.number_of_columns).unwrap_or(usize::MAX);
    codec::DecodeOptions::for_source(source).with_max_columns(columns)
}

fn codec_error(error: codec::DecodeError) -> Error {
    Error::InvalidFormat(error.to_string())
}

pub(super) fn read_native_table_sort_order_wire(
    original: &[u8],
    model: &TableModelArchive,
) -> Result<Option<codec::SortOrderSnapshot>> {
    let (snapshot, has_sort_field) = codec::decode_table_model_sort_order_with_presence(
        original,
        codec_options(original, model),
    )
    .map_err(codec_error)?;
    if has_sort_field != model.sort_order.is_some() {
        return Err(Error::InvalidFormat(
            "Numbers table sort-order wire payload is missing or inconsistent".to_owned(),
        ));
    }
    if let Some(native) = model.sort_order.as_ref() {
        let consistent = match snapshot.as_ref() {
            Some(snapshot) => {
                native.r#type == snapshot.scope().native_value()
                    && native.rules.len() == snapshot.rules().len()
                    && native
                        .rules
                        .iter()
                        .zip(snapshot.rules())
                        .all(|(native, rule)| {
                            native.index == rule.column()
                                && native.direction == rule.direction().native_value()
                        })
            },
            None => native.rules.is_empty() && matches!(native.r#type, 0 | 1),
        };
        if !consistent {
            return Err(Error::InvalidFormat(
                "Numbers table sort-order wire payload is missing or inconsistent".to_owned(),
            ));
        }
    }
    Ok(snapshot)
}

pub(super) fn delete_table_sort_column_wire(
    original: &[u8],
    model: &TableModelArchive,
    column: u32,
    new_columns: u32,
) -> Result<Vec<u8>> {
    let Some(previous) = read_native_table_sort_order_wire(original, model)? else {
        return Ok(original.to_vec());
    };
    let rules = previous
        .rules()
        .iter()
        .copied()
        .filter(|rule| rule.column() != column && rule.column() < new_columns)
        .collect::<Vec<_>>();
    let expected = if rules.is_empty() {
        None
    } else {
        Some(codec::SortOrderSnapshot::new(previous.scope(), rules).map_err(codec_error)?)
    };
    if expected.as_ref() == Some(&previous) {
        return Ok(original.to_vec());
    }
    let data = codec::rewrite_table_model_sort_order(
        original,
        expected.clone(),
        codec_options(original, model),
    )
    .map_err(codec_error)?
    .into_bytes();
    let verified = TableModelArchive::decode(data.as_slice())?;
    if read_native_table_sort_order_wire(&data, &verified)? != expected {
        return Err(Error::InvalidFormat(
            "Numbers table sort-order column deletion failed validation".to_owned(),
        ));
    }
    Ok(data)
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::numbers::NumbersDocumentBuilder;
    use crate::protobuf::tst::TableModelArchive;
    use prost::Message;

    const ROOT_GROUP_BEFORE: [u8; 6] = [0xa3, 0x06, 0x08, 0x01, 0xa4, 0x06];
    const ROOT_GROUP_AFTER: [u8; 6] = [0xab, 0x06, 0x10, 0x02, 0xac, 0x06];

    fn table_model_source() -> Vec<u8> {
        let editor = NumbersDocumentBuilder::new()
            .table_dimensions(2, 3)
            .build()
            .expect("source-built Numbers table");
        let table_id = editor.tables().expect("table catalog")[0].object_id;
        let archive = editor
            .package
            .archive("Index/Document.iwa")
            .expect("document archive");
        let object = archive.object(table_id).expect("table model object");
        let message_index = find_table_model_message(object).expect("table model message");
        object.messages[message_index].data.clone()
    }

    fn table_model_source_with_root_groups(
        order: &codec::SortOrderSnapshot,
    ) -> (Vec<u8>, usize, Vec<u8>) {
        let mut source = table_model_source();
        let sort_payload = codec::canonical_table_sort_order(order).expect("sort payload");
        let mut sort_field = Vec::new();
        crate::wire::append_length_delimited_field(&mut sort_field, 44, &sort_payload)
            .expect("sort field");
        // Unknown balanced groups are valid protobuf wire fields.  Keep one
        // on each side of field 44 so the regression covers both traversal
        // order and exact retention through a rewrite.
        let tail_start = source.len();
        source.extend_from_slice(&ROOT_GROUP_BEFORE);
        source.extend_from_slice(&sort_field);
        source.extend_from_slice(&ROOT_GROUP_AFTER);
        (source, tail_start, sort_field)
    }

    #[test]
    fn strict_codec_reads_sort_with_unrelated_root_groups() {
        let order = codec::SortOrderSnapshot::new(
            codec::SortScope::EntireTable,
            [codec::SortRule::new(1, codec::SortDirection::Ascending)],
        )
        .expect("sort order");
        let (source, tail_start, sort_field) = table_model_source_with_root_groups(&order);
        let model = TableModelArchive::decode(source.as_slice()).expect("model with sort");

        assert_eq!(
            read_native_table_sort_order_wire(&source, &model).expect("strict sort read"),
            Some(order)
        );
        assert_eq!(
            &source[tail_start..tail_start + ROOT_GROUP_BEFORE.len()],
            ROOT_GROUP_BEFORE
        );
        let sort_start = tail_start + ROOT_GROUP_BEFORE.len();
        assert_eq!(
            &source[sort_start..sort_start + sort_field.len()],
            sort_field
        );
        assert_eq!(&source[sort_start + sort_field.len()..], ROOT_GROUP_AFTER);
    }

    #[test]
    fn column_delete_retains_root_groups_and_their_source_order() {
        let order = codec::SortOrderSnapshot::new(
            codec::SortScope::EntireTable,
            [
                codec::SortRule::new(1, codec::SortDirection::Ascending),
                codec::SortRule::new(2, codec::SortDirection::Descending),
            ],
        )
        .expect("sort order");
        let (source, tail_start, _sort_field) = table_model_source_with_root_groups(&order);
        let model = TableModelArchive::decode(source.as_slice()).expect("model with sort");

        let changed =
            delete_table_sort_column_wire(&source, &model, 2, 2).expect("column deletion rewrite");
        let changed_model =
            TableModelArchive::decode(changed.as_slice()).expect("rewritten model with groups");
        let expected = codec::SortOrderSnapshot::new(
            codec::SortScope::EntireTable,
            [codec::SortRule::new(1, codec::SortDirection::Ascending)],
        )
        .expect("remaining sort order");
        assert_eq!(
            read_native_table_sort_order_wire(&changed, &changed_model)
                .expect("strict rewritten sort read"),
            Some(expected.clone())
        );

        let mut expected_sort_field = Vec::new();
        let expected_payload = codec::canonical_table_sort_order(&expected).expect("sort payload");
        crate::wire::append_length_delimited_field(&mut expected_sort_field, 44, &expected_payload)
            .expect("sort field");
        let changed_after_start = changed.len() - ROOT_GROUP_AFTER.len();
        let changed_sort_start = changed_after_start - expected_sort_field.len();
        let changed_group_start = changed_sort_start - ROOT_GROUP_BEFORE.len();
        assert_eq!(changed_group_start, tail_start);
        assert_eq!(&changed[..tail_start], &source[..tail_start]);
        assert_eq!(
            &changed[changed_group_start..changed_sort_start],
            ROOT_GROUP_BEFORE
        );
        assert_eq!(
            &changed[changed_sort_start..changed_after_start],
            expected_sort_field
        );
        assert_eq!(&changed[changed_after_start..], ROOT_GROUP_AFTER);
    }

    #[test]
    fn strict_codec_reads_semantics_without_eager_sort_archive_decode() {
        let source = table_model_source();
        let order = codec::SortOrderSnapshot::new(
            codec::SortScope::EntireTable,
            [codec::SortRule::new(1, codec::SortDirection::Ascending)],
        )
        .expect("sort order");
        let sort_payload = codec::canonical_table_sort_order(&order).expect("sort payload");
        let source =
            crate::wire::append_repeated_length_delimited_field(&source, 44, &sort_payload)
                .expect("sort field");
        let model = TableModelArchive::decode(source.as_slice()).expect("model with sort");
        assert_eq!(model.number_of_columns, 3);
        assert_eq!(
            read_native_table_sort_order_wire(&source, &model).expect("strict sort read"),
            Some(order)
        );
    }

    #[test]
    fn strict_codec_rewrite_keeps_unknown_outer_sort_and_rule_fields() {
        let mut source = table_model_source();
        let order = codec::SortOrderSnapshot::new(
            codec::SortScope::EntireTable,
            [
                codec::SortRule::new(1, codec::SortDirection::Ascending),
                codec::SortRule::new(2, codec::SortDirection::Descending),
            ],
        )
        .expect("sort order");
        let mut sort_payload = codec::canonical_table_sort_order(&order).expect("sort payload");
        let sort_unknown = {
            let mut bytes = Vec::new();
            crate::wire::append_varint_field(&mut bytes, 98, 980).expect("sort unknown");
            bytes
        };
        sort_payload.extend_from_slice(&sort_unknown);
        let rule_unknown = {
            let mut bytes = Vec::new();
            crate::wire::append_varint_field(&mut bytes, 97, 970).expect("rule unknown");
            bytes
        };
        sort_payload =
            crate::wire::transform_length_delimited_fields_at_path(&sort_payload, &[2], |rule| {
                let mut rule = rule.to_vec();
                rule.extend_from_slice(&rule_unknown);
                Ok(rule)
            })
            .expect("rule unknown");
        crate::wire::append_varint_field(&mut source, 99, 990).expect("outer unknown");
        crate::wire::append_length_delimited_field(&mut source, 44, &sort_payload)
            .expect("sort field");
        let model = TableModelArchive::decode(source.as_slice()).expect("model with sort");

        let changed =
            delete_table_sort_column_wire(&source, &model, 2, 2).expect("column deletion rewrite");
        let changed_model = TableModelArchive::decode(changed.as_slice()).expect("rewritten model");
        let changed_sort = crate::wire::repeated_length_delimited_payloads(&changed, 44)
            .expect("sort payloads")
            .pop()
            .expect("sort payload");
        let changed_rules =
            crate::wire::repeated_length_delimited_payloads(changed_sort, 2).expect("sort rules");
        assert_eq!(changed_rules.len(), 1);
        assert!(changed_rules[0].ends_with(&rule_unknown));
        assert!(changed_sort.ends_with(&sort_unknown));
        assert!(
            crate::wire::parse_wire_fields(&changed)
                .expect("outer fields")
                .iter()
                .any(|field| field.number() == 99)
        );
        assert_eq!(
            read_native_table_sort_order_wire(&changed, &changed_model)
                .expect("strict rewritten sort read")
                .expect("remaining sort rule")
                .rules(),
            &[codec::SortRule::new(1, codec::SortDirection::Ascending)]
        );
    }
}
