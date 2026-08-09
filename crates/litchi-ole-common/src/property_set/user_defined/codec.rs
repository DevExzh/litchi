use super::model::{Hyperlink, Hyperlinks, Limits, LinkBase};
use litchi_cfb::OleError;

const VT_I4: u16 = 0x0003;
const VT_LPWSTR: u16 = 0x001f;

struct Reader<'a> {
    data: &'a [u8],
    offset: usize,
}

impl<'a> Reader<'a> {
    const fn new(data: &'a [u8]) -> Self {
        Self { data, offset: 0 }
    }

    fn remaining(&self) -> usize {
        self.data.len() - self.offset
    }

    fn take(&mut self, count: usize, field: &str) -> Result<&'a [u8], OleError> {
        let end = self
            .offset
            .checked_add(count)
            .ok_or_else(|| invalid(format!("{field} offset overflow")))?;
        let bytes = self
            .data
            .get(self.offset..end)
            .ok_or_else(|| invalid(format!("{field} is truncated")))?;
        self.offset = end;
        Ok(bytes)
    }

    fn u16(&mut self, field: &str) -> Result<u16, OleError> {
        let bytes = self.take(2, field)?;
        Ok(u16::from_le_bytes([bytes[0], bytes[1]]))
    }

    fn u32(&mut self, field: &str) -> Result<u32, OleError> {
        let bytes = self.take(4, field)?;
        Ok(u32::from_le_bytes([bytes[0], bytes[1], bytes[2], bytes[3]]))
    }

    fn i4(&mut self, field: &str) -> Result<i32, OleError> {
        if self.u16(&format!("{field} type"))? != VT_I4 {
            return Err(invalid(format!("{field} must be VT_I4")));
        }
        // MS-OLEPS requires zero on write but does not make a nonzero
        // retained padding word a read failure.
        self.u16(&format!("{field} padding"))?;
        Ok(i32::from_ne_bytes(
            self.u32(&format!("{field} value"))?.to_ne_bytes(),
        ))
    }

    fn string(
        &mut self,
        limits: Limits,
        aggregate: &mut usize,
        field: &str,
    ) -> Result<String, OleError> {
        if self.u16(&format!("{field} type"))? != VT_LPWSTR {
            return Err(invalid(format!("{field} must be VT_LPWSTR")));
        }
        self.u16(&format!("{field} padding"))?;
        let units = usize::try_from(self.u32(&format!("{field} UTF-16 unit count"))?).map_err(
            |_conversion_error| invalid(format!("{field} UTF-16 unit count is too large")),
        )?;
        check_units(units, limits, aggregate, field)?;
        let byte_count = units
            .checked_mul(2)
            .ok_or_else(|| invalid(format!("{field} UTF-16 byte length overflow")))?;
        let bytes = self.take(byte_count, field)?;
        let mut decoded = String::new();
        decoded
            .try_reserve(units)
            .map_err(|source| allocation("user-defined hyperlink string", source))?;
        let mut values = Vec::new();
        values
            .try_reserve_exact(units)
            .map_err(|source| allocation("user-defined hyperlink UTF-16", source))?;
        for code_unit_bytes in bytes.chunks_exact(2) {
            values.push(u16::from_le_bytes([code_unit_bytes[0], code_unit_bytes[1]]));
        }
        if values.last().copied() != Some(0) {
            return Err(invalid(format!("{field} must be NUL terminated")));
        }
        if values[..values.len() - 1].contains(&0) {
            return Err(invalid(format!("{field} contains an interior NUL")));
        }
        for character in char::decode_utf16(values[..values.len() - 1].iter().copied()) {
            decoded.push(
                character
                    .map_err(|_utf16_error| invalid(format!("{field} contains invalid UTF-16")))?,
            );
        }
        let padding = (4 - (byte_count % 4)) % 4;
        self.take(padding, &format!("{field} padding"))?;
        Ok(decoded)
    }

    fn finish(self, field: &str) -> Result<(), OleError> {
        if self.remaining() == 0 {
            Ok(())
        } else {
            Err(invalid(format!("{field} has trailing bytes")))
        }
    }
}

fn invalid(message: impl Into<String>) -> OleError {
    OleError::InvalidFormat(message.into())
}

fn allocation(resource: &'static str, source: std::collections::TryReserveError) -> OleError {
    OleError::Allocation { resource, source }
}

pub(super) fn decode_link_base(data: &[u8], limits: Limits) -> Result<LinkBase, OleError> {
    check_blob(data, limits, "_PID_LINKBASE")?;
    if data.is_empty() {
        return LinkBase::new("");
    }
    if !data.len().is_multiple_of(2) {
        return Err(invalid(
            "_PID_LINKBASE UTF-16 payload must have an even byte length",
        ));
    }
    let mut aggregate = 0;
    check_units(data.len() / 2, limits, &mut aggregate, "_PID_LINKBASE")?;
    let value = decode_utf16(data, "_PID_LINKBASE")?;
    LinkBase::new(value)
}

pub(super) fn encode_link_base(value: &LinkBase, limits: Limits) -> Result<Vec<u8>, OleError> {
    let mut units = 0;
    let mut output = Vec::new();
    let count = value
        .value()
        .encode_utf16()
        .count()
        .checked_add(1)
        .ok_or_else(|| invalid("_PID_LINKBASE UTF-16 length overflow"))?;
    check_units(count, limits, &mut units, "_PID_LINKBASE")?;
    let bytes = count
        .checked_mul(2)
        .ok_or_else(|| invalid("_PID_LINKBASE UTF-16 byte length overflow"))?;
    check_blob_len(bytes, limits, "_PID_LINKBASE")?;
    output
        .try_reserve_exact(bytes)
        .map_err(|source| allocation("user-defined link base", source))?;
    for unit in value.value().encode_utf16() {
        output.extend_from_slice(&unit.to_le_bytes());
    }
    output.extend_from_slice(&0u16.to_le_bytes());
    check_blob(&output, limits, "_PID_LINKBASE")?;
    Ok(output)
}

pub(super) fn decode_hyperlinks(data: &[u8], limits: Limits) -> Result<Hyperlinks, OleError> {
    check_blob(data, limits, "_PID_HLINKS")?;
    let mut reader = Reader::new(data);
    let declared_bytes_raw = reader.u32("_PID_HLINKS cbData")?;
    let declared_bytes = usize::try_from(declared_bytes_raw)
        .map_err(|_conversion_error| invalid("_PID_HLINKS cbData is too large"))?;
    if declared_bytes != reader.remaining() {
        return Err(invalid(
            "_PID_HLINKS cbData does not exactly bound VecVtHyperlink",
        ));
    }
    let elements = usize::try_from(reader.u32("_PID_HLINKS cElements")?)
        .map_err(|_conversion_error| invalid("_PID_HLINKS cElements is too large"))?;
    if elements % 6 != 0 {
        return Err(invalid("_PID_HLINKS cElements must be divisible by six"));
    }
    let count = elements / 6;
    if count > limits.max_links() {
        return Err(invalid(format!(
            "_PID_HLINKS link count {count} exceeds configured limit {}",
            limits.max_links()
        )));
    }
    let mut links = Vec::new();
    links
        .try_reserve_exact(count)
        .map_err(|source| allocation("user-defined hyperlinks", source))?;
    let mut units = 0;
    for _ in 0..count {
        let hash = u32::from_ne_bytes(reader.i4("hyperlink hash")?.to_ne_bytes());
        let app = reader.i4("hyperlink app")?;
        let office_art = reader.i4("hyperlink OfficeArt")?;
        let info = reader.i4("hyperlink info")?;
        let target = reader.string(limits, &mut units, "hyperlink target")?;
        let location = reader.string(limits, &mut units, "hyperlink location")?;
        links.push(Hyperlink::from_wire(
            hash, app, office_art, info, target, location,
        ));
    }
    reader.finish("_PID_HLINKS")?;
    Ok(Hyperlinks::new(links))
}

pub(super) fn encode_hyperlinks(value: &Hyperlinks, limits: Limits) -> Result<Vec<u8>, OleError> {
    if value.len() > limits.max_links() {
        return Err(invalid(format!(
            "_PID_HLINKS link count {} exceeds configured limit {}",
            value.len(),
            limits.max_links()
        )));
    }
    let element_count_usize = value
        .len()
        .checked_mul(6)
        .ok_or_else(|| invalid("_PID_HLINKS element count overflow"))?;
    let element_count = u32::try_from(element_count_usize)
        .map_err(|_conversion_error| invalid("_PID_HLINKS element count exceeds u32"))?;
    let planned = encoded_hyperlinks_size(value, limits)?;
    let mut output = Vec::new();
    output
        .try_reserve_exact(planned)
        .map_err(|source| allocation("user-defined hyperlinks", source))?;
    output.extend_from_slice(&0u32.to_le_bytes());
    output.extend_from_slice(&element_count.to_le_bytes());
    let mut units = 0;
    for link in value.links() {
        append_i4(
            &mut output,
            i32::from_ne_bytes(link.stored_hash().to_ne_bytes()),
        )?;
        append_i4(&mut output, link.app())?;
        append_i4(&mut output, link.office_art())?;
        append_i4(&mut output, link.info())?;
        append_string(
            &mut output,
            link.target(),
            limits,
            &mut units,
            "hyperlink target",
        )?;
        append_string(
            &mut output,
            link.location(),
            limits,
            &mut units,
            "hyperlink location",
        )?;
    }
    let payload_size = output
        .len()
        .checked_sub(4)
        .ok_or_else(|| invalid("_PID_HLINKS payload underflow"))?;
    let data_size = u32::try_from(payload_size)
        .map_err(|_conversion_error| invalid("_PID_HLINKS cbData exceeds u32"))?;
    if output.len() != planned {
        return Err(invalid(
            "_PID_HLINKS encoded size does not match its checked plan",
        ));
    }
    output[..4].copy_from_slice(&data_size.to_le_bytes());
    check_blob(&output, limits, "_PID_HLINKS")?;
    Ok(output)
}

fn check_blob(data: &[u8], limits: Limits, field: &str) -> Result<(), OleError> {
    check_blob_len(data.len(), limits, field)
}

fn check_blob_len(length: usize, limits: Limits, field: &str) -> Result<(), OleError> {
    if length > limits.max_blob_bytes() {
        return Err(invalid(format!(
            "{field} BLOB size {length} exceeds configured limit {}",
            limits.max_blob_bytes()
        )));
    }
    Ok(())
}

fn encoded_hyperlinks_size(value: &Hyperlinks, limits: Limits) -> Result<usize, OleError> {
    let mut total = 8usize;
    let mut units = 0usize;
    for link in value.links() {
        total = total
            .checked_add(32)
            .ok_or_else(|| invalid("_PID_HLINKS fixed record size overflow"))?;
        total = total
            .checked_add(encoded_string_size(
                link.target(),
                limits,
                &mut units,
                "hyperlink target",
            )?)
            .ok_or_else(|| invalid("_PID_HLINKS encoded size overflow"))?;
        total = total
            .checked_add(encoded_string_size(
                link.location(),
                limits,
                &mut units,
                "hyperlink location",
            )?)
            .ok_or_else(|| invalid("_PID_HLINKS encoded size overflow"))?;
        check_blob_len(total, limits, "_PID_HLINKS")?;
    }
    check_blob_len(total, limits, "_PID_HLINKS")?;
    Ok(total)
}

fn encoded_string_size(
    value: &str,
    limits: Limits,
    aggregate: &mut usize,
    field: &str,
) -> Result<usize, OleError> {
    if value.contains('\0') {
        return Err(invalid(format!("{field} must not contain NUL")));
    }
    let units = value
        .encode_utf16()
        .count()
        .checked_add(1)
        .ok_or_else(|| invalid(format!("{field} UTF-16 length overflow")))?;
    check_units(units, limits, aggregate, field)?;
    let byte_count = units
        .checked_mul(2)
        .ok_or_else(|| invalid(format!("{field} UTF-16 byte length overflow")))?;
    let padding = (4 - (byte_count % 4)) % 4;
    8usize
        .checked_add(byte_count)
        .and_then(|size| size.checked_add(padding))
        .ok_or_else(|| invalid(format!("{field} encoded size overflow")))
}

fn append_i4(output: &mut Vec<u8>, value: i32) -> Result<(), OleError> {
    output
        .try_reserve(8)
        .map_err(|source| allocation("user-defined hyperlink value", source))?;
    output.extend_from_slice(&VT_I4.to_le_bytes());
    output.extend_from_slice(&0u16.to_le_bytes());
    output.extend_from_slice(&value.to_le_bytes());
    Ok(())
}

fn append_string(
    output: &mut Vec<u8>,
    value: &str,
    limits: Limits,
    aggregate: &mut usize,
    field: &str,
) -> Result<(), OleError> {
    if value.contains('\0') {
        return Err(invalid(format!("{field} must not contain NUL")));
    }
    let units = value
        .encode_utf16()
        .count()
        .checked_add(1)
        .ok_or_else(|| invalid(format!("{field} UTF-16 length overflow")))?;
    check_units(units, limits, aggregate, field)?;
    let byte_count = units
        .checked_mul(2)
        .ok_or_else(|| invalid(format!("{field} UTF-16 byte length overflow")))?;
    let padding = (4 - (byte_count % 4)) % 4;
    let required = 8usize
        .checked_add(byte_count)
        .and_then(|size| size.checked_add(padding))
        .ok_or_else(|| invalid(format!("{field} encoded size overflow")))?;
    output
        .try_reserve(required)
        .map_err(|source| allocation("user-defined hyperlink string", source))?;
    output.extend_from_slice(&VT_LPWSTR.to_le_bytes());
    output.extend_from_slice(&0u16.to_le_bytes());
    output.extend_from_slice(
        &u32::try_from(units)
            .map_err(|_conversion_error| invalid(format!("{field} UTF-16 length exceeds u32")))?
            .to_le_bytes(),
    );
    for unit in value.encode_utf16() {
        output.extend_from_slice(&unit.to_le_bytes());
    }
    output.extend_from_slice(&0u16.to_le_bytes());
    output.resize(output.len() + padding, 0);
    Ok(())
}

fn check_units(
    units: usize,
    limits: Limits,
    aggregate: &mut usize,
    field: &str,
) -> Result<(), OleError> {
    if units == 0 || units > limits.max_string_units() {
        return Err(invalid(format!(
            "{field} UTF-16 unit count {units} exceeds configured per-string limit {}",
            limits.max_string_units()
        )));
    }
    *aggregate = aggregate
        .checked_add(units)
        .ok_or_else(|| invalid("user-defined hyperlink aggregate UTF-16 unit count overflow"))?;
    if *aggregate > limits.max_total_utf16_units() {
        return Err(invalid(format!(
            "user-defined hyperlink aggregate UTF-16 unit count {} exceeds configured limit {}",
            *aggregate,
            limits.max_total_utf16_units()
        )));
    }
    Ok(())
}

fn decode_utf16(bytes: &[u8], field: &str) -> Result<String, OleError> {
    if !bytes.len().is_multiple_of(2) {
        return Err(invalid(format!(
            "{field} UTF-16 payload has an odd byte length"
        )));
    }
    if bytes.is_empty() {
        return Err(invalid(format!("{field} must include a terminating NUL")));
    }
    let mut values = Vec::new();
    values
        .try_reserve_exact(bytes.len() / 2)
        .map_err(|source| allocation("user-defined hyperlink UTF-16", source))?;
    for pair in bytes.chunks_exact(2) {
        values.push(u16::from_le_bytes([pair[0], pair[1]]));
    }
    if values.last().copied() != Some(0) {
        return Err(invalid(format!("{field} must be NUL terminated")));
    }
    if values[..values.len() - 1].contains(&0) {
        return Err(invalid(format!("{field} contains an interior NUL")));
    }
    let mut decoded = String::new();
    decoded
        .try_reserve(values.len())
        .map_err(|source| allocation("user-defined hyperlink string", source))?;
    for character in char::decode_utf16(values[..values.len() - 1].iter().copied()) {
        decoded.push(
            character
                .map_err(|_utf16_error| invalid(format!("{field} contains invalid UTF-16")))?,
        );
    }
    Ok(decoded)
}
