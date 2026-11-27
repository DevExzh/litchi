use std::collections::HashMap;
use std::path::{Component as PathComponent, Path};

use litchi_iwa_common::{decode_varint_from_bytes, varint::encoded_len};

use super::{
    METADATA_COMPONENT, MediaClosureState, PACKAGE_METADATA_TYPE, Package, RawField, Selection,
    TransactionBudget, charge_message_info, physical_catalog, validate_selected_metadata,
};
use crate::soundtrack::Error;

pub(super) fn validate_media_closure(
    package: &Package,
    selection: &Selection<'_>,
    budget: &mut TransactionBudget,
) -> Result<(), Error> {
    let info = selection
        .soundtrack
        .archive_info
        .message_infos
        .get(selection.soundtrack_message_index)
        .ok_or(Error::InvalidSource)?;
    if info.data_references.is_empty() {
        return Ok(());
    }
    let mut states = HashMap::new();
    states
        .try_reserve(info.data_references.len())
        .map_err(|_error| Error::Allocation {
            amount: info.data_references.len(),
        })?;
    for identifier in &info.data_references {
        let state = states.entry(*identifier).or_insert(MediaClosureState {
            payload_occurrences: 0,
            component_declarations: 0,
            owner_occurrences: 0,
            owner_count: 0,
            data_declarations: 0,
            filename: None,
            materialized_length: None,
        });
        state.payload_occurrences = state
            .payload_occurrences
            .checked_add(1)
            .ok_or(Error::InvalidSource)?;
    }

    let catalog = physical_catalog(package)?;
    let metadata_component = catalog
        .components()
        .get(METADATA_COMPONENT)
        .ok_or(Error::InvalidSource)?;
    let mut selected_metadata_payload = None;
    for object in &metadata_component.archive().objects {
        if object.messages.len() != object.archive_info.message_infos.len() {
            return Err(Error::InvalidSource);
        }
        for (index, (message, message_info)) in object
            .messages
            .iter()
            .zip(&object.archive_info.message_infos)
            .enumerate()
        {
            budget.charge_work(message.data.len().saturating_add(1))?;
            if message.type_ != message_info.type_
                || usize::try_from(message_info.length).ok() != Some(message.data.len())
            {
                return Err(Error::InvalidSource);
            }
            if message.type_ == PACKAGE_METADATA_TYPE {
                if selected_metadata_payload
                    .replace(message.data.as_slice())
                    .is_some()
                {
                    return Err(Error::InvalidSource);
                }
                charge_message_info(object, index, budget)?;
                validate_selected_metadata(object, index)?;
            }
        }
    }
    let metadata_payload = selected_metadata_payload.ok_or(Error::InvalidSource)?;
    let locator = selection
        .soundtrack_component
        .strip_prefix("Index/")
        .and_then(|name| name.strip_suffix(".iwa"))
        .ok_or(Error::InvalidSource)?;
    let selected_components = stream_package_metadata(
        metadata_payload,
        locator.as_bytes(),
        selection.soundtrack_identifier,
        &mut states,
        budget,
    )?;
    if selected_components != 1 {
        return Err(Error::InvalidSource);
    }
    let mut materialized = HashMap::new();
    materialized
        .try_reserve(states.len())
        .map_err(|_error| Error::Allocation {
            amount: states.len(),
        })?;
    for state in states.values() {
        let filename = state.filename.ok_or(Error::InvalidSource)?;
        if materialized.insert(filename, (0usize, None)).is_some() {
            return Err(Error::InvalidSource);
        }
    }
    for entry in catalog.package().iter() {
        budget.charge_work(1)?;
        if let Some(filename) = entry.name().strip_prefix("Data/")
            && let Some((count, length)) = materialized.get_mut(filename.as_bytes())
        {
            *count = count.checked_add(1).ok_or(Error::InvalidSource)?;
            *length = Some(entry.data().len());
        }
    }
    for state in states.values() {
        let filename = state.filename.ok_or(Error::InvalidSource)?;
        if state.component_declarations != 1
            || state.owner_occurrences != 1
            || state.owner_count != state.payload_occurrences
            || state.data_declarations != 1
            || materialized.get(filename).copied() != Some((1, state.materialized_length))
        {
            return Err(Error::InvalidSource);
        }
    }
    Ok(())
}

fn stream_package_metadata<'a>(
    source: &'a [u8],
    locator: &[u8],
    soundtrack_identifier: u64,
    states: &mut HashMap<u64, MediaClosureState<'a>>,
    budget: &mut TransactionBudget,
) -> Result<usize, Error> {
    budget.require_depth(1)?;
    budget.charge_work(source.len())?;
    let mut locator_input = source;
    let mut selected_components = 0usize;
    while let Some(field) = next_raw_field(&mut locator_input, budget)? {
        match field.number {
            3 | 11 if field.wire == 2 => {
                let component = field.bytes.ok_or(Error::InvalidSource)?;
                if stream_component_info(component, locator, soundtrack_identifier, states, budget)?
                {
                    selected_components = selected_components
                        .checked_add(1)
                        .ok_or(Error::InvalidSource)?;
                }
            },
            4 if field.wire == 2 => {
                stream_data_info(field.bytes.ok_or(Error::InvalidSource)?, states, budget)?;
            },
            3 | 4 | 11 => return Err(Error::InvalidSource),
            _ => {},
        }
    }
    Ok(selected_components)
}

fn stream_component_info(
    source: &[u8],
    locator: &[u8],
    soundtrack_identifier: u64,
    states: &mut HashMap<u64, MediaClosureState<'_>>,
    budget: &mut TransactionBudget,
) -> Result<bool, Error> {
    budget.require_depth(2)?;
    budget.charge_work(source.len())?;
    let mut input = source;
    let mut preferred = None;
    let mut current = None;
    while let Some(field) = next_raw_field(&mut input, budget)? {
        match field.number {
            2 if field.wire == 2 && preferred.is_none() => preferred = field.bytes,
            3 if field.wire == 2 && current.is_none() => current = field.bytes,
            2 | 3 => return Err(Error::InvalidSource),
            _ => {},
        }
    }
    if current.or(preferred) != Some(locator) {
        return Ok(false);
    }
    budget.charge_work(source.len())?;
    let mut reference_input = source;
    while let Some(field) = next_raw_field(&mut reference_input, budget)? {
        if field.number == 7 {
            if field.wire != 2 {
                return Err(Error::InvalidSource);
            }
            stream_component_data_reference(
                field.bytes.ok_or(Error::InvalidSource)?,
                soundtrack_identifier,
                states,
                budget,
            )?;
        }
    }
    Ok(true)
}

fn stream_component_data_reference(
    source: &[u8],
    soundtrack_identifier: u64,
    states: &mut HashMap<u64, MediaClosureState<'_>>,
    budget: &mut TransactionBudget,
) -> Result<(), Error> {
    budget.require_depth(3)?;
    budget.charge_work(source.len())?;
    let mut identifier_input = source;
    let mut data_identifier = None;
    while let Some(field) = next_raw_field(&mut identifier_input, budget)? {
        match field.number {
            1 if field.wire == 0 && data_identifier.is_none() => data_identifier = field.varint,
            2 if field.wire == 2 => {},
            1 | 2 => return Err(Error::InvalidSource),
            _ => {},
        }
    }
    let Some(state) = data_identifier.and_then(|identifier| states.get_mut(&identifier)) else {
        return Ok(());
    };
    state.component_declarations = state
        .component_declarations
        .checked_add(1)
        .ok_or(Error::InvalidSource)?;
    budget.charge_work(source.len())?;
    let mut owner_list_input = source;
    while let Some(owner_field) = next_raw_field(&mut owner_list_input, budget)? {
        if owner_field.number != 2 {
            continue;
        }
        let owner = owner_field.bytes.ok_or(Error::InvalidSource)?;
        budget.require_depth(4)?;
        budget.charge_work(owner.len())?;
        let mut owner_input = owner;
        let mut object = None;
        let mut count = None;
        while let Some(owner_value) = next_raw_field(&mut owner_input, budget)? {
            match owner_value.number {
                1 if owner_value.wire == 0 && object.is_none() => object = owner_value.varint,
                2 if owner_value.wire == 0 && count.is_none() => count = owner_value.varint,
                1 | 2 => return Err(Error::InvalidSource),
                _ => {},
            }
        }
        budget.charge_references(1)?;
        if object == Some(soundtrack_identifier) {
            state.owner_occurrences = state
                .owner_occurrences
                .checked_add(1)
                .ok_or(Error::InvalidSource)?;
            state.owner_count = usize::try_from(count.ok_or(Error::InvalidSource)?)
                .map_err(|_error| Error::InvalidSource)?;
        }
    }
    Ok(())
}

fn stream_data_info<'a>(
    source: &'a [u8],
    states: &mut HashMap<u64, MediaClosureState<'a>>,
    budget: &mut TransactionBudget,
) -> Result<(), Error> {
    budget.require_depth(2)?;
    budget.charge_work(source.len())?;
    let mut input = source;
    let mut identifier = None;
    let mut digest = None;
    let mut preferred = None;
    let mut current = None;
    let mut materialized_length = None;
    while let Some(field) = next_raw_field(&mut input, budget)? {
        match field.number {
            1 if field.wire == 0 && identifier.is_none() => identifier = field.varint,
            2 if field.wire == 2 && digest.is_none() => digest = field.bytes,
            3 if field.wire == 2 && preferred.is_none() => preferred = field.bytes,
            4 if field.wire == 2 && current.is_none() => current = field.bytes,
            18 if field.wire == 0 && materialized_length.is_none() => {
                materialized_length = field.varint;
            },
            1 | 2 | 3 | 4 | 18 => return Err(Error::InvalidSource),
            _ => {},
        }
    }
    let Some(state) = identifier.and_then(|data_identifier| states.get_mut(&data_identifier))
    else {
        return Ok(());
    };
    state.data_declarations = state
        .data_declarations
        .checked_add(1)
        .ok_or(Error::InvalidSource)?;
    if digest.is_none_or(|digest_bytes| digest_bytes.len() != 20) {
        return Err(Error::InvalidSource);
    }
    state.materialized_length = Some(
        usize::try_from(materialized_length.ok_or(Error::InvalidSource)?)
            .map_err(|_error| Error::InvalidSource)?,
    );
    let filename_bytes = current.or(preferred).ok_or(Error::InvalidSource)?;
    let filename = std::str::from_utf8(filename_bytes).map_err(|_error| Error::InvalidSource)?;
    if filename.is_empty()
        || filename.contains(['\0', '\\'])
        || Path::new(filename).is_absolute()
        || Path::new(filename)
            .components()
            .any(|component| !matches!(component, PathComponent::Normal(_)))
    {
        return Err(Error::InvalidSource);
    }
    state.filename = Some(filename.as_bytes());
    Ok(())
}

pub(super) fn next_raw_field<'a>(
    input: &mut &'a [u8],
    budget: &mut TransactionBudget,
) -> Result<Option<RawField<'a>>, Error> {
    if input.is_empty() {
        return Ok(None);
    }
    let source = *input;
    budget.charge_fields(1)?;
    let tag = take_canonical_varint(input)?;
    let number = u32::try_from(tag >> 3).map_err(|_error| Error::InvalidSource)?;
    let wire = u8::try_from(tag & 7).map_err(|_error| Error::InvalidSource)?;
    if number == 0 || number > 0x1fff_ffff {
        return Err(Error::InvalidSource);
    }
    let mut field = RawField {
        number,
        wire,
        varint: None,
        bytes: None,
        raw: &[],
    };
    match wire {
        0 => field.varint = Some(take_canonical_varint(input)?),
        1 => {
            field.bytes = Some(take_bytes(input, 8)?);
        },
        2 => {
            let length = usize::try_from(take_canonical_varint(input)?)
                .map_err(|_error| Error::InvalidSource)?;
            field.bytes = Some(take_bytes(input, length)?);
        },
        5 => {
            field.bytes = Some(take_bytes(input, 4)?);
        },
        _ => return Err(Error::InvalidSource),
    }
    let consumed = source
        .len()
        .checked_sub(input.len())
        .ok_or(Error::InvalidSource)?;
    field.raw = source.get(..consumed).ok_or(Error::InvalidSource)?;
    Ok(Some(field))
}

fn take_canonical_varint(input: &mut &[u8]) -> Result<u64, Error> {
    let (value, consumed) =
        decode_varint_from_bytes(input).map_err(|_error| Error::InvalidSource)?;
    if consumed != encoded_len(value) {
        return Err(Error::InvalidSource);
    }
    *input = input.get(consumed..).ok_or(Error::InvalidSource)?;
    Ok(value)
}

fn take_bytes<'a>(input: &mut &'a [u8], count: usize) -> Result<&'a [u8], Error> {
    let value = input.get(..count).ok_or(Error::InvalidSource)?;
    *input = input.get(count..).ok_or(Error::InvalidSource)?;
    Ok(value)
}

#[cfg(test)]
mod tests {
    use std::collections::HashMap;

    use litchi_iwa_common::{WireLimits, varint::encode_varint_into};

    use super::{MediaClosureState, TransactionBudget, stream_package_metadata};

    const SOUNDTRACK_IDENTIFIER: u64 = 10;
    const LOCATOR: &[u8] = b"Soundtrack";

    #[derive(Clone, Copy)]
    struct Counts {
        fields: usize,
        work: usize,
        references: usize,
    }

    #[test]
    fn streaming_media_metadata_counts_scale_linearly_from_4096_to_8192() {
        let small = stream_fixture(4_096);
        let large = stream_fixture(8_192);

        assert_eq!(small.references, 4_096);
        assert_eq!(large.references, 8_192);
        assert_at_most_2_3x(small.fields, large.fields);
        assert_at_most_2_3x(small.work, large.work);
        assert_at_most_2_3x(small.references, large.references);
    }

    fn stream_fixture(media_count: usize) -> Counts {
        let source = package_metadata(media_count);
        let mut states = HashMap::with_capacity(media_count);
        for offset in 0..media_count {
            let identifier = data_identifier(offset);
            assert!(
                states
                    .insert(
                        identifier,
                        MediaClosureState {
                            payload_occurrences: 1,
                            component_declarations: 0,
                            owner_occurrences: 0,
                            owner_count: 0,
                            data_declarations: 0,
                            filename: None,
                            materialized_length: None,
                        },
                    )
                    .is_none()
            );
        }
        let mut budget = TransactionBudget {
            fields: 0,
            work: 0,
            references: 0,
            max_fields: WireLimits::MAX_FIELDS,
            max_work: WireLimits::MAX_REWRITE_WORK,
            max_references: 1_000_000,
            max_nesting: WireLimits::MAX_NESTING,
        };

        assert_eq!(
            stream_package_metadata(
                &source,
                LOCATOR,
                SOUNDTRACK_IDENTIFIER,
                &mut states,
                &mut budget,
            )
            .expect("synthetic metadata must be valid"),
            1
        );
        assert_eq!(states.len(), media_count);
        assert!(states.values().all(|state| {
            state.payload_occurrences == 1
                && state.component_declarations == 1
                && state.owner_occurrences == 1
                && state.owner_count == 1
                && state.data_declarations == 1
                && state.filename.is_some()
                && state.materialized_length == Some(1)
        }));

        Counts {
            fields: budget.fields,
            work: budget.work,
            references: budget.references,
        }
    }

    fn package_metadata(media_count: usize) -> Vec<u8> {
        let mut component = Vec::new();
        push_bytes_field(&mut component, 3, LOCATOR);
        for offset in 0..media_count {
            let mut owner = Vec::new();
            push_varint_field(&mut owner, 1, SOUNDTRACK_IDENTIFIER);
            push_varint_field(&mut owner, 2, 1);

            let mut reference = Vec::new();
            push_varint_field(&mut reference, 1, data_identifier(offset));
            push_bytes_field(&mut reference, 2, &owner);
            push_bytes_field(&mut component, 7, &reference);
        }

        let mut metadata = Vec::new();
        push_bytes_field(&mut metadata, 3, &component);
        for offset in 0..media_count {
            let identifier = data_identifier(offset);
            let filename = format!("m{identifier:04x}");
            let mut data = Vec::new();
            push_varint_field(&mut data, 1, identifier);
            push_bytes_field(&mut data, 2, &[0; 20]);
            push_bytes_field(&mut data, 4, filename.as_bytes());
            push_varint_field(&mut data, 18, 1);
            push_bytes_field(&mut metadata, 4, &data);
        }
        metadata
    }

    fn data_identifier(offset: usize) -> u64 {
        4_096u64
            .checked_add(u64::try_from(offset).expect("test count fits u64"))
            .expect("test identifier fits u64")
    }

    fn push_varint_field(output: &mut Vec<u8>, number: u32, value: u64) {
        encode_varint_into(output, u64::from(number) << 3);
        encode_varint_into(output, value);
    }

    fn push_bytes_field(output: &mut Vec<u8>, number: u32, value: &[u8]) {
        encode_varint_into(output, (u64::from(number) << 3) | 2);
        encode_varint_into(
            output,
            u64::try_from(value.len()).expect("test field length fits u64"),
        );
        output.extend_from_slice(value);
    }

    fn assert_at_most_2_3x(small: usize, large: usize) {
        assert!(
            large.saturating_mul(10) <= small.saturating_mul(23),
            "expected at most 2.3x growth: small={small}, large={large}"
        );
    }
}
