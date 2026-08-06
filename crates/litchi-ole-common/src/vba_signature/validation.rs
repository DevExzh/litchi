//! Structural and resource validation for VBA signature blobs.

use std::ops::Range;

use super::codec::{Header, INFO_HEADER_SIZE, Layout, Outer};
use super::model::{Error, Kind, Limits};

pub(crate) fn blob_size(size: usize, limits: Limits) -> Result<(), Error> {
    if size > limits.max_blob_bytes {
        return Err(Error::Limit("blob byte"));
    }
    Ok(())
}

pub(crate) fn layout(
    source: &[u8],
    kind: Kind,
    outer: Outer,
    header: Header,
    limits: Limits,
) -> Result<Layout, Error> {
    if outer.info.len() < INFO_HEADER_SIZE {
        return Err(Error::Truncated("information header"));
    }
    if header.project_name_size != 0 {
        return Err(Error::invalid("reserved project-name size must be zero"));
    }
    if header.timestamp_url_size != 0 {
        return Err(Error::invalid("reserved timestamp-URL size must be zero"));
    }

    let signature_size = size(
        header.signature_size,
        "signature",
        limits.max_signature_bytes,
    )?;
    let certificate_store_size = size(
        header.certificate_store_size,
        "certificate-store",
        limits.max_certificate_store_bytes,
    )?;
    let header_end = outer.info.start + INFO_HEADER_SIZE;
    let signature = field_range(
        outer.base,
        header.signature_offset,
        signature_size,
        header_end,
        outer.info.end,
        "signature",
    )?;
    let certificate_store = field_range(
        outer.base,
        header.certificate_store_offset,
        certificate_store_size,
        header_end,
        outer.info.end,
        "certificate store",
    )?;
    let project_name = field_range(
        outer.base,
        header.project_name_offset,
        2,
        header_end,
        outer.info.end,
        "reserved project name",
    )?;
    let timestamp_url = field_range(
        outer.base,
        header.timestamp_url_offset,
        2,
        header_end,
        outer.info.end,
        "reserved timestamp URL",
    )?;

    let fields = [
        (&signature, "signature"),
        (&certificate_store, "certificate store"),
        (&project_name, "reserved project name"),
        (&timestamp_url, "reserved timestamp URL"),
    ];
    for (index, (left, left_name)) in fields.iter().enumerate() {
        for (right, right_name) in fields.iter().skip(index + 1) {
            if overlaps(left, right) {
                return Err(Error::invalid(format!("{left_name} overlaps {right_name}")));
            }
        }
    }
    if source[project_name.clone()] != [0, 0] {
        return Err(Error::invalid(
            "reserved project name must be one null UTF-16 code unit",
        ));
    }
    if source[timestamp_url.clone()] != [0, 0] {
        return Err(Error::invalid(
            "reserved timestamp URL must be one null UTF-16 code unit",
        ));
    }

    let content_end = fields
        .iter()
        .map(|(range, _)| range.end)
        .max()
        .unwrap_or(header_end);
    let (info, padding) = match kind {
        Kind::Property => {
            let expected = padding_for(content_end, 4);
            if content_end.checked_add(expected) != Some(outer.total_end) {
                return Err(Error::invalid(
                    "DigSigBlob padding does not align signature info to four bytes",
                ));
            }
            (outer.info.start..content_end, content_end..outer.total_end)
        },
        Kind::Word => {
            let expected_total = outer
                .info
                .end
                .checked_add(padding_for(outer.info.end, 2))
                .ok_or_else(|| Error::invalid("WordSigBlob padding overflows usize"))?;
            if expected_total != outer.total_end {
                return Err(Error::invalid(
                    "WordSigBlob count does not match its info and padding",
                ));
            }
            (outer.info.clone(), outer.info.end..outer.total_end)
        },
    };

    Ok(Layout {
        info,
        signature,
        certificate_store,
        project_name,
        timestamp_url,
        timestamp_marker: header.timestamp_marker,
        padding,
    })
}

fn size(raw: u32, field: &'static str, maximum: usize) -> Result<usize, Error> {
    let value =
        usize::try_from(raw).map_err(|_| Error::invalid(format!("{field} size overflows")))?;
    if value > maximum {
        return Err(Error::Limit(field));
    }
    Ok(value)
}

fn field_range(
    base: usize,
    raw_offset: u32,
    size: usize,
    minimum: usize,
    maximum: usize,
    field: &'static str,
) -> Result<Range<usize>, Error> {
    let offset = usize::try_from(raw_offset)
        .map_err(|_| Error::invalid(format!("{field} offset overflows usize")))?;
    let start = base
        .checked_add(offset)
        .ok_or_else(|| Error::invalid(format!("{field} offset overflows usize")))?;
    let end = start
        .checked_add(size)
        .ok_or_else(|| Error::invalid(format!("{field} size overflows usize")))?;
    if start < minimum {
        return Err(Error::invalid(format!(
            "{field} starts inside the information header"
        )));
    }
    if end > maximum {
        return Err(Error::Truncated(field));
    }
    Ok(start..end)
}

fn overlaps(left: &Range<usize>, right: &Range<usize>) -> bool {
    !left.is_empty() && !right.is_empty() && left.start < right.end && right.start < left.end
}

fn padding_for(offset: usize, alignment: usize) -> usize {
    (alignment - (offset % alignment)) % alignment
}
