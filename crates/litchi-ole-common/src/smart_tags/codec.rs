use std::collections::HashSet;

use litchi_codepage::Ansi;

use super::model::{
    Error, Limits, Property, PropertyBag, PropertyBagStore, PropertyBagString,
    PropertyBagStringEncoding, Type,
};

impl PropertyBagStore {
    /// Parse a store prefix and return the number of bytes consumed.
    pub fn parse_prefix(data: &[u8], ansi: Ansi, limits: Limits) -> Result<(Self, usize), Error> {
        let mut cursor = Cursor::new(data);
        let type_count = cursor.u32("smart-tag type count")?;
        let type_count = bounded_count(
            type_count,
            cursor.remaining(),
            14,
            limits.max_types,
            "smart-tag type count",
        )?;
        let mut types = Vec::with_capacity(type_count);
        let mut type_ids = HashSet::with_capacity(type_count);
        for _ in 0..type_count {
            let size = usize::try_from(cursor.u32("FactoidType size")?)
                .map_err(|_| Error::new("FactoidType size overflows usize"))?;
            let end = cursor
                .offset
                .checked_add(size)
                .ok_or_else(|| Error::new("FactoidType size overflows"))?;
            if end > data.len() {
                return Err(Error::new("FactoidType is truncated"));
            }
            let id = u16::try_from(cursor.u32("FactoidType id")?)
                .map_err(|_| Error::new("FactoidType id exceeds 0xFFFF"))?;
            let namespace_uri = cursor.pb_string(ansi)?;
            let tag_name = cursor.pb_string(ansi)?;
            let download_url = cursor.pb_string(ansi)?;
            if cursor.offset != end {
                return Err(Error::new(
                    "FactoidType byte count does not match its contents",
                ));
            }
            if !type_ids.insert(id) {
                return Err(Error::new(
                    "PropertyBagStore has duplicate smart-tag type ids",
                ));
            }
            types.push(Type {
                id,
                namespace_uri,
                tag_name,
                download_url,
            });
        }

        if cursor.u16("property-bag header size")? != 0x000c
            || cursor.u16("property-bag version")? != 0x0100
        {
            return Err(Error::new(
                "PropertyBagStore has an invalid header size or version",
            ));
        }
        let reserved_factoid_count = cursor.u32("property-bag reserved value")?;
        let string_count = cursor.u32("smart-tag string count")?;
        let string_count = bounded_count(
            string_count,
            cursor.remaining(),
            2,
            limits.max_strings,
            "smart-tag string count",
        )?;
        let mut strings = Vec::with_capacity(string_count);
        for _ in 0..string_count {
            strings.push(cursor.pb_string(ansi)?);
        }

        Ok((
            Self {
                ansi,
                reserved_factoid_count,
                types,
                strings,
            },
            cursor.offset,
        ))
    }

    /// Resolve a declared type by its stable 16-bit identifier.
    pub fn tag_type(&self, id: u16) -> Option<&Type> {
        self.types.iter().find(|kind| kind.id == id)
    }

    /// Resolve a string-table index.
    pub fn string(&self, index: u32) -> Option<&str> {
        self.strings
            .get(usize::try_from(index).ok()?)
            .map(|value| value.value.as_str())
    }

    /// Resolve both strings referenced by a property.
    pub fn resolve_property(&self, property: Property) -> Option<(&str, &str)> {
        Some((
            self.string(property.key_index)?,
            self.string(property.value_index)?,
        ))
    }

    /// Serialize the shared store without any format-specific property bags.
    pub fn to_bytes(&self) -> Result<Vec<u8>, Error> {
        self.to_bytes_with_bags(&[])
    }

    /// Serialize the shared store followed by the supplied property bags.
    ///
    /// ANSI strings are encoded with [`Self::ansi`]; values that are not
    /// representable are rejected instead of being replaced.
    pub fn to_bytes_with_bags(&self, bags: &[PropertyBag]) -> Result<Vec<u8>, Error> {
        let type_count = u32::try_from(self.types.len())
            .map_err(|_| Error::new("smart-tag type count exceeds u32"))?;
        let string_count = u32::try_from(self.strings.len())
            .map_err(|_| Error::new("smart-tag string count exceeds u32"))?;
        let mut type_ids = HashSet::with_capacity(self.types.len());
        let mut output = Vec::new();
        output.extend_from_slice(&type_count.to_le_bytes());
        for kind in &self.types {
            if !type_ids.insert(kind.id) {
                return Err(Error::new(
                    "PropertyBagStore has duplicate smart-tag type ids",
                ));
            }
            let mut payload = u32::from(kind.id).to_le_bytes().to_vec();
            append_pb_string(&mut payload, &kind.namespace_uri, self.ansi)?;
            append_pb_string(&mut payload, &kind.tag_name, self.ansi)?;
            append_pb_string(&mut payload, &kind.download_url, self.ansi)?;
            let payload_len = u32::try_from(payload.len())
                .map_err(|_| Error::new("FactoidType payload exceeds u32"))?;
            output.extend_from_slice(&payload_len.to_le_bytes());
            output.extend_from_slice(&payload);
        }
        output.extend_from_slice(&0x000cu16.to_le_bytes());
        output.extend_from_slice(&0x0100u16.to_le_bytes());
        output.extend_from_slice(&self.reserved_factoid_count.to_le_bytes());
        output.extend_from_slice(&string_count.to_le_bytes());
        for value in &self.strings {
            append_pb_string(&mut output, value, self.ansi)?;
        }
        for bag in bags {
            if !type_ids.contains(&bag.type_id) {
                return Err(Error::new(
                    "PropertyBag references an unknown smart-tag type",
                ));
            }
            let property_count = u16::try_from(bag.properties.len())
                .map_err(|_| Error::new("PropertyBag contains more than 65535 properties"))?;
            output.extend_from_slice(&bag.type_id.to_le_bytes());
            output.extend_from_slice(&property_count.to_le_bytes());
            output.extend_from_slice(&0u16.to_le_bytes());
            for property in &bag.properties {
                if self.resolve_property(*property).is_none() {
                    return Err(Error::new(
                        "smart-tag property string index is out of range",
                    ));
                }
                output.extend_from_slice(&property.key_index.to_le_bytes());
                output.extend_from_slice(&property.value_index.to_le_bytes());
            }
        }
        Ok(output)
    }

    /// Parse an exact number of property bags from `data`.
    pub fn parse_bags(
        &self,
        data: &[u8],
        count: usize,
        limits: Limits,
    ) -> Result<Vec<PropertyBag>, Error> {
        if count > limits.max_bags || count > data.len() / 6 {
            return Err(Error::new(
                "smart-tag bag count exceeds the configured or encoded limit",
            ));
        }
        let mut cursor = Cursor::new(data);
        let mut total_properties = 0usize;
        let mut bags = Vec::with_capacity(count);
        for _ in 0..count {
            bags.push(self.parse_bag(&mut cursor, &mut total_properties, limits)?);
        }
        if cursor.remaining() != 0 {
            return Err(Error::new("smart-tag property bags contain trailing bytes"));
        }
        Ok(bags)
    }

    /// Parse property bags until `data` is exhausted, as used by MS-DOC.
    pub fn parse_bags_to_end(
        &self,
        data: &[u8],
        limits: Limits,
    ) -> Result<Vec<PropertyBag>, Error> {
        let mut cursor = Cursor::new(data);
        let mut total_properties = 0usize;
        let mut bags = Vec::new();
        while cursor.remaining() != 0 {
            if bags.len() >= limits.max_bags {
                return Err(Error::new(
                    "smart-tag bag count exceeds the configured limit",
                ));
            }
            bags.push(self.parse_bag(&mut cursor, &mut total_properties, limits)?);
        }
        Ok(bags)
    }

    fn parse_bag(
        &self,
        cursor: &mut Cursor<'_>,
        total_properties: &mut usize,
        limits: Limits,
    ) -> Result<PropertyBag, Error> {
        let type_id = cursor.u16("smart-tag type id")?;
        let property_count = usize::from(cursor.u16("smart-tag property count")?);
        if cursor.u16("smart-tag reserved value")? != 0 {
            return Err(Error::new("PropertyBag has a nonzero reserved field"));
        }
        if self.tag_type(type_id).is_none() {
            return Err(Error::new(
                "PropertyBag references an unknown smart-tag type",
            ));
        }
        *total_properties = total_properties
            .checked_add(property_count)
            .ok_or_else(|| Error::new("smart-tag property count overflows"))?;
        if *total_properties > limits.max_properties || property_count > cursor.remaining() / 8 {
            return Err(Error::new(
                "smart-tag property count exceeds the configured or encoded limit",
            ));
        }
        let mut properties = Vec::with_capacity(property_count);
        for _ in 0..property_count {
            let property = Property {
                key_index: cursor.u32("smart-tag property key index")?,
                value_index: cursor.u32("smart-tag property value index")?,
            };
            if self.resolve_property(property).is_none() {
                return Err(Error::new(
                    "smart-tag property string index is out of range",
                ));
            }
            properties.push(property);
        }
        Ok(PropertyBag {
            type_id,
            properties,
        })
    }
}

fn append_pb_string(
    output: &mut Vec<u8>,
    value: &PropertyBagString,
    ansi: Ansi,
) -> Result<(), Error> {
    if value.value.contains('\0') {
        return Err(Error::new(
            "PBString values cannot contain embedded NUL characters",
        ));
    }
    match value.encoding {
        PropertyBagStringEncoding::Ansi => {
            let encoded = ansi
                .encode(&value.value)
                .map_err(|error| Error::new(format!("PBString ANSI encoding failed: {error}")))?;
            let count = u16::try_from(encoded.len())
                .ok()
                .filter(|count| *count <= 0x7fff)
                .ok_or_else(|| Error::new("ANSI PBString exceeds 32767 bytes"))?;
            output.extend_from_slice(&(count | 0x8000).to_le_bytes());
            output.extend_from_slice(&encoded);
        },
        PropertyBagStringEncoding::Utf16 => {
            let units = value.value.encode_utf16().collect::<Vec<_>>();
            let count = u16::try_from(units.len())
                .ok()
                .filter(|count| *count <= 0x7fff)
                .ok_or_else(|| Error::new("UTF-16 PBString exceeds 32767 code units"))?;
            output.extend_from_slice(&count.to_le_bytes());
            output.extend(units.into_iter().flat_map(u16::to_le_bytes));
        },
    }
    Ok(())
}

fn bounded_count(
    value: u32,
    remaining: usize,
    item_minimum: usize,
    configured_maximum: usize,
    name: &str,
) -> Result<usize, Error> {
    let value =
        usize::try_from(value).map_err(|_| Error::new(format!("{name} overflows usize")))?;
    if value > configured_maximum || value > remaining / item_minimum {
        return Err(Error::new(format!(
            "{name} exceeds the configured or encoded limit"
        )));
    }
    Ok(value)
}

struct Cursor<'a> {
    data: &'a [u8],
    offset: usize,
}

impl<'a> Cursor<'a> {
    fn new(data: &'a [u8]) -> Self {
        Self { data, offset: 0 }
    }

    fn remaining(&self) -> usize {
        self.data.len() - self.offset
    }

    fn bytes(&mut self, count: usize, name: &str) -> Result<&'a [u8], Error> {
        let end = self
            .offset
            .checked_add(count)
            .ok_or_else(|| Error::new(format!("{name} offset overflows")))?;
        let bytes = self
            .data
            .get(self.offset..end)
            .ok_or_else(|| Error::new(format!("{name} is truncated")))?;
        self.offset = end;
        Ok(bytes)
    }

    fn u16(&mut self, name: &str) -> Result<u16, Error> {
        let bytes = self.bytes(2, name)?;
        Ok(u16::from_le_bytes([bytes[0], bytes[1]]))
    }

    fn u32(&mut self, name: &str) -> Result<u32, Error> {
        let bytes = self.bytes(4, name)?;
        Ok(u32::from_le_bytes([bytes[0], bytes[1], bytes[2], bytes[3]]))
    }

    fn pb_string(&mut self, ansi: Ansi) -> Result<PropertyBagString, Error> {
        let header = self.u16("PBString header")?;
        let count = usize::from(header & 0x7fff);
        let encoding = if header & 0x8000 != 0 {
            PropertyBagStringEncoding::Ansi
        } else {
            PropertyBagStringEncoding::Utf16
        };
        let byte_count = match encoding {
            PropertyBagStringEncoding::Ansi => count,
            PropertyBagStringEncoding::Utf16 => count
                .checked_mul(2)
                .ok_or_else(|| Error::new("PBString size overflows"))?,
        };
        let bytes = self.bytes(byte_count, "PBString")?;
        let value = match encoding {
            PropertyBagStringEncoding::Ansi => ansi
                .decode(bytes)
                .map_err(|error| Error::new(format!("PBString ANSI decoding failed: {error}")))?
                .into_owned(),
            PropertyBagStringEncoding::Utf16 => {
                let units = bytes
                    .chunks_exact(2)
                    .map(|pair| u16::from_le_bytes([pair[0], pair[1]]))
                    .collect::<Vec<_>>();
                String::from_utf16(&units)
                    .map_err(|_| Error::new("PBString contains invalid UTF-16"))?
            },
        };
        if value.contains('\0') {
            return Err(Error::new("PBString contains an embedded NUL character"));
        }
        Ok(PropertyBagString { value, encoding })
    }
}
