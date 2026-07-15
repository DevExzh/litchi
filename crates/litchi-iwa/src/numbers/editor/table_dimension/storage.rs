//! Wire-preserving table dimension header storage.

use super::*;
use crate::wire::patch_fixed32_field;

const HEADER_ENTRIES_FIELD: u32 = 2;
const HEADER_SIZE_FIELD: u32 = 2;

pub(super) fn read_dimension_size(
    package: &IWorkPackage,
    locations: &HashMap<u64, String>,
    model: &TableModelArchive,
    dimension: NumbersTableDimension,
) -> Result<Option<f32>> {
    let identifier = header_bucket_identifier(model, dimension)?;
    if identifier == 0 {
        return Ok(None);
    }
    let archive_name = locations.get(&identifier).ok_or_else(|| {
        Error::InvalidFormat(format!(
            "Numbers header bucket object {identifier} is missing"
        ))
    })?;
    let archive = package.archive(archive_name)?;
    let object = archive.object(identifier).ok_or_else(|| {
        Error::InvalidFormat(format!(
            "Numbers header bucket object {identifier} is missing"
        ))
    })?;
    let bucket = object
        .messages
        .iter()
        .find_map(|message| tst::HeaderStorageBucket::decode(message.data.as_slice()).ok())
        .ok_or_else(|| {
            Error::InvalidFormat(format!(
                "Object {identifier} has no Numbers header bucket payload"
            ))
        })?;
    let index = u32::try_from(dimension.index())
        .map_err(|_| Error::ParseError("Numbers table dimension exceeds u32".to_owned()))?;
    let mut matches = bucket.headers.iter().filter(|header| header.index == index);
    let size = matches.next().map(|header| header.size);
    if matches.next().is_some() {
        return Err(Error::InvalidFormat(format!(
            "Numbers header bucket repeats dimension index {index}"
        )));
    }
    Ok(size.filter(|size| *size != DEFAULT_DIMENSION_POINTS))
}

pub(super) fn write_dimension_size(
    package: &mut IWorkPackage,
    locations: &HashMap<u64, String>,
    model: &TableModelArchive,
    dimension: NumbersTableDimension,
    points: f32,
) -> Result<()> {
    let identifier = header_bucket_identifier(model, dimension)?;
    if identifier == 0 {
        return Err(Error::InvalidFormat(
            "Numbers table has no header storage for dimension sizing".to_owned(),
        ));
    }
    let archive_name = locations.get(&identifier).ok_or_else(|| {
        Error::InvalidFormat(format!(
            "Numbers header bucket object {identifier} is missing"
        ))
    })?;
    package.update_archive(archive_name, |archive| {
        let object = archive.object_mut(identifier).ok_or_else(|| {
            Error::InvalidFormat(format!(
                "Numbers header bucket object {identifier} is missing"
            ))
        })?;
        let message_index = object
            .messages
            .iter()
            .position(|message| tst::HeaderStorageBucket::decode(message.data.as_slice()).is_ok())
            .ok_or_else(|| {
                Error::InvalidFormat(format!(
                    "Object {identifier} has no Numbers header bucket payload"
                ))
            })?;
        let original = object.messages[message_index].data.clone();
        let previous = tst::HeaderStorageBucket::decode(original.as_slice())?;
        let raw_headers = repeated_length_delimited_payloads(&original, HEADER_ENTRIES_FIELD)?;
        if raw_headers.len() != previous.headers.len() {
            return Err(Error::InvalidFormat(
                "Numbers header bucket wire count is inconsistent".to_owned(),
            ));
        }
        let index = u32::try_from(dimension.index())
            .map_err(|_| Error::ParseError("Numbers table dimension exceeds u32".to_owned()))?;
        let positions = previous
            .headers
            .iter()
            .enumerate()
            .filter_map(|(position, header)| (header.index == index).then_some(position))
            .collect::<Vec<_>>();
        if positions.len() > 1 {
            return Err(Error::InvalidFormat(format!(
                "Numbers header bucket repeats dimension index {index}"
            )));
        }
        let mut current = previous.clone();
        match positions.first().copied() {
            Some(position) if current.headers[position].size.to_bits() == points.to_bits() => {
                return Ok(());
            },
            Some(position) if points == DEFAULT_DIMENSION_POINTS => {
                let header = &current.headers[position];
                let canonical_wire = header.encode_to_vec() == raw_headers[position];
                let removable = header.number_of_cells == 0
                    && header.hiding_state == 0
                    && header.cell_style.is_none()
                    && header.text_style.is_none()
                    && canonical_wire;
                if removable {
                    current.headers.remove(position);
                } else {
                    current.headers[position].size = DEFAULT_DIMENSION_POINTS;
                }
            },
            Some(position) => current.headers[position].size = points,
            None if points != DEFAULT_DIMENSION_POINTS => {
                current.headers.push(tst::header_storage_bucket::Header {
                    index,
                    size: points,
                    hiding_state: 0,
                    number_of_cells: 0,
                    cell_style: None,
                    text_style: None,
                });
                current.headers.sort_by_key(|header| header.index);
            },
            None => return Ok(()),
        }
        let data = rewrite_dimension_size_wire(&original, &previous, &current)?;
        let message_type = object.messages[message_index].type_;
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

fn header_bucket_identifier(
    model: &TableModelArchive,
    dimension: NumbersTableDimension,
) -> Result<u64> {
    match dimension {
        NumbersTableDimension::Column(_) => Ok(model.base_data_store.column_headers.identifier),
        NumbersTableDimension::Row(row) => model
            .base_data_store
            .row_headers
            .buckets
            .get(row / HEADER_BUCKET_ROWS)
            .map(|reference| reference.identifier)
            .ok_or_else(|| {
                Error::InvalidFormat(format!(
                    "Numbers table {:?} has no row-header bucket for row {row}",
                    model.table_name
                ))
            }),
    }
}

fn rewrite_dimension_size_wire(
    original: &[u8],
    previous: &tst::HeaderStorageBucket,
    current: &tst::HeaderStorageBucket,
) -> Result<Vec<u8>> {
    if previous.bucket_hash_function != current.bucket_hash_function {
        return Err(Error::InvalidFormat(
            "Numbers header-bucket hash function changed during size mutation".to_owned(),
        ));
    }
    let raw_headers = repeated_length_delimited_payloads(original, HEADER_ENTRIES_FIELD)?;
    if raw_headers.len() != previous.headers.len() {
        return Err(Error::InvalidFormat(
            "Numbers header bucket wire count is inconsistent".to_owned(),
        ));
    }
    let mut existing = HashMap::with_capacity(previous.headers.len());
    for (header, raw) in previous.headers.iter().zip(raw_headers) {
        if tst::header_storage_bucket::Header::decode(raw)? != *header
            || existing.insert(header.index, (header, raw)).is_some()
        {
            return Err(Error::InvalidFormat(
                "Numbers header bucket has inconsistent or duplicate entries".to_owned(),
            ));
        }
    }
    let mut seen = HashSet::with_capacity(current.headers.len());
    let replacements = current
        .headers
        .iter()
        .map(|header| {
            if !seen.insert(header.index) {
                return Err(Error::InvalidFormat(
                    "Numbers header bucket would contain duplicate entries".to_owned(),
                ));
            }
            let Some((previous, raw)) = existing.get(&header.index) else {
                return Ok(header.encode_to_vec());
            };
            let mut immutable = *header;
            immutable.size = previous.size;
            if immutable != **previous {
                return Err(Error::InvalidFormat(format!(
                    "Numbers header {} changed outside its dimension size",
                    header.index
                )));
            }
            patch_fixed32_field(raw, HEADER_SIZE_FIELD, true, Some(header.size.to_bits()))
        })
        .collect::<Result<Vec<_>>>()?;
    let data =
        rewrite_repeated_length_delimited_fields(original, HEADER_ENTRIES_FIELD, &replacements)?;
    if tst::HeaderStorageBucket::decode(data.as_slice())? != *current {
        return Err(Error::InvalidFormat(
            "Numbers header dimension size failed wire validation".to_owned(),
        ));
    }
    Ok(data)
}
