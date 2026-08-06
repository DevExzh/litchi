//! Wire decoding for `[MS-OSHARED]` VBA signature containers.

use std::ops::Range;

use super::model::{Error, Kind, Limits};
use super::validation;

pub(crate) const INFO_HEADER_SIZE: usize = 36;

#[derive(Debug, Clone, PartialEq, Eq)]
pub(crate) struct Layout {
    pub(crate) info: Range<usize>,
    pub(crate) signature: Range<usize>,
    pub(crate) certificate_store: Range<usize>,
    pub(crate) project_name: Range<usize>,
    pub(crate) timestamp_url: Range<usize>,
    pub(crate) timestamp_marker: u32,
    pub(crate) padding: Range<usize>,
}

#[derive(Debug, Clone, Copy)]
pub(crate) struct Header {
    pub(crate) signature_size: u32,
    pub(crate) signature_offset: u32,
    pub(crate) certificate_store_size: u32,
    pub(crate) certificate_store_offset: u32,
    pub(crate) project_name_size: u32,
    pub(crate) project_name_offset: u32,
    pub(crate) timestamp_marker: u32,
    pub(crate) timestamp_url_size: u32,
    pub(crate) timestamp_url_offset: u32,
}

pub(crate) fn parse(source: &[u8], kind: Kind, limits: Limits) -> Result<Layout, Error> {
    validation::blob_size(source.len(), limits)?;

    let outer = match kind {
        Kind::Property => property_outer(source)?,
        Kind::Word => word_outer(source)?,
    };
    let header = header(source, outer.info.start)?;
    validation::layout(source, kind, outer, header, limits)
}

#[derive(Debug, Clone)]
pub(crate) struct Outer {
    pub(crate) base: usize,
    pub(crate) info: Range<usize>,
    pub(crate) total_end: usize,
}

fn property_outer(source: &[u8]) -> Result<Outer, Error> {
    let declared = usize::try_from(read_u32(source, 0, "DigSigBlob size")?)
        .map_err(|_| Error::invalid("DigSigBlob size overflows usize"))?;
    let pointer = read_u32(source, 4, "DigSigBlob serialized pointer")?;
    if pointer != 8 {
        return Err(Error::invalid("DigSigBlob serialized pointer must equal 8"));
    }
    let total_end = 8usize
        .checked_add(declared)
        .ok_or_else(|| Error::invalid("DigSigBlob size overflows usize"))?;
    if total_end > source.len() {
        return Err(Error::Truncated("DigSigBlob payload"));
    }
    if total_end != source.len() {
        return Err(Error::invalid("DigSigBlob has trailing bytes"));
    }
    Ok(Outer {
        base: 0,
        info: 8..total_end,
        total_end,
    })
}

fn word_outer(source: &[u8]) -> Result<Outer, Error> {
    let code_units = usize::from(read_u16(source, 0, "WordSigBlob character count")?);
    let info_size = usize::try_from(read_u32(source, 2, "WordSigBlob info size")?)
        .map_err(|_| Error::invalid("WordSigBlob info size overflows usize"))?;
    let pointer = read_u32(source, 6, "WordSigBlob serialized pointer")?;
    if pointer != 8 {
        return Err(Error::invalid(
            "WordSigBlob serialized pointer must equal 8",
        ));
    }
    let total_end = 2usize
        .checked_add(
            code_units
                .checked_mul(2)
                .ok_or_else(|| Error::invalid("WordSigBlob size overflows usize"))?,
        )
        .ok_or_else(|| Error::invalid("WordSigBlob size overflows usize"))?;
    if total_end > source.len() {
        return Err(Error::Truncated("WordSigBlob payload"));
    }
    if total_end != source.len() {
        return Err(Error::invalid("WordSigBlob has trailing bytes"));
    }
    let info_end = 10usize
        .checked_add(info_size)
        .ok_or_else(|| Error::invalid("WordSigBlob info size overflows usize"))?;
    if info_end > total_end {
        return Err(Error::Truncated("WordSigBlob signature info"));
    }
    Ok(Outer {
        base: 2,
        info: 10..info_end,
        total_end,
    })
}

fn header(source: &[u8], offset: usize) -> Result<Header, Error> {
    Ok(Header {
        signature_size: read_u32(source, offset, "signature size")?,
        signature_offset: read_u32(source, offset + 4, "signature offset")?,
        certificate_store_size: read_u32(source, offset + 8, "certificate-store size")?,
        certificate_store_offset: read_u32(source, offset + 12, "certificate-store offset")?,
        project_name_size: read_u32(source, offset + 16, "project-name size")?,
        project_name_offset: read_u32(source, offset + 20, "project-name offset")?,
        timestamp_marker: read_u32(source, offset + 24, "timestamp marker")?,
        timestamp_url_size: read_u32(source, offset + 28, "timestamp-URL size")?,
        timestamp_url_offset: read_u32(source, offset + 32, "timestamp-URL offset")?,
    })
}

fn read_u16(source: &[u8], offset: usize, field: &'static str) -> Result<u16, Error> {
    let bytes = source
        .get(offset..offset + 2)
        .ok_or(Error::Truncated(field))?;
    Ok(u16::from_le_bytes([bytes[0], bytes[1]]))
}

fn read_u32(source: &[u8], offset: usize, field: &'static str) -> Result<u32, Error> {
    let bytes = source
        .get(offset..offset + 4)
        .ok_or(Error::Truncated(field))?;
    Ok(u32::from_le_bytes([bytes[0], bytes[1], bytes[2], bytes[3]]))
}
