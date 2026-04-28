use crate::{Error, ErrorKind, Header, utils::le_u16};
use std::io::Write;

/// A numeric identifier for an extra field in a Zip archive.
///
/// Constants defined here correspond to the IDs defined in the Zip specification.
///
/// See sections 4.5 and 4.6 of the Zip spec.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub struct ExtraFieldId(u16);

impl ExtraFieldId {
    pub const ZIP64: Self = Self(0x0001);
    pub const AV_INFO: Self = Self(0x0007);
    pub const EXTENDED_LANGUAGE_ENCODING: Self = Self(0x0008);
    pub const OS2: Self = Self(0x0009);
    pub const NTFS: Self = Self(0x000a);
    pub const OPENVMS: Self = Self(0x000c);
    pub const UNIX: Self = Self(0x000d);
    pub const FILE_STREAM_AND_FORK_DESCRIPTORS: Self = Self(0x000e);
    pub const PATCH_DESCRIPTOR: Self = Self(0x000f);
    pub const PKCS7_STORE: Self = Self(0x0014);
    pub const X509_CERT_ID_AND_SIG: Self = Self(0x0015);
    pub const X509_CERT_ID_CENTRAL_DIR: Self = Self(0x0016);
    pub const STRONG_ENCRYPTION_HEADER: Self = Self(0x0017);
    pub const RECORD_MANAGEMENT_CONTROLS: Self = Self(0x0018);
    pub const PKCS7_ENCRYPTION_RECIPIENT_CERT_LIST: Self = Self(0x0019);
    pub const TIMESTAMP_RECORD: Self = Self(0x0020);
    pub const POLICY_DECRYPTION_KEY_RECORD: Self = Self(0x0021);
    pub const SMARTCRYPT_KEY_PROVIDER: Self = Self(0x0022);
    pub const SMARTCRYPT_POLICY_KEY_DATA: Self = Self(0x0023);
    pub const IBM_S390_AS400_UNCOMPRESSED: Self = Self(0x0065);
    pub const IBM_S390_AS400_COMPRESSED: Self = Self(0x0066);
    pub const POSZIP_4690: Self = Self(0x4690);
    pub const EXTENDED_TIMESTAMP: Self = Self(0x5455);
    pub const INFO_ZIP_UNIX_ORIGINAL: Self = Self(0x5855);
    pub const INFO_ZIP_UNIX: Self = Self(0x7855);
    pub const INFO_ZIP_UNIX_UID_GID: Self = Self(0x7875);
    pub const JAVA_JAR: Self = Self(0xCAFE);
    pub const ANDROID_ZIP_ALIGNMENT: Self = Self(0xD935);
    pub const MACINTOSH: Self = Self(0x07c8);
    pub const ACORN_SPARKFS: Self = Self(0x4341);
    pub const WINDOWS_NT_SECURITY_DESCRIPTOR: Self = Self(0x4653);
    pub const AOS_VS_ACL: Self = Self(0x5356);
    pub const INFO_ZIP_UNICODE_COMMENT: Self = Self(0x6375);
    pub const INFO_ZIP_UNICODE_PATH: Self = Self(0x7075);
    pub const DATA_STREAM_ALIGNMENT: Self = Self(0xa11e);
    pub const MICROSOFT_OPEN_PACKAGING_GROWTH_HINT: Self = Self(0xa220);

    /// Returns the raw `u16` value of the extra field ID.
    #[inline]
    pub const fn new(id: u16) -> Self {
        Self(id)
    }

    /// Returns the raw `u16` value of the extra field ID.
    #[inline]
    pub const fn as_u16(self) -> u16 {
        self.0
    }
}

/// An iterator over extra field entries in a Zip archive.
///
/// This follows zip spec section 4.5 defines extensible data fields:
///
/// - Header ID - 2 bytes
/// - Data Size - 2 bytes
/// - Data - variable length
///
/// If the iterator encounters malformed or truncated data, it will stop
/// yielding entries. You can check [`ExtraFields::remaining_bytes()`] after
/// iteration to detect if any data was left unparsed.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub struct ExtraFields<'a> {
    data: &'a [u8],
}

impl<'a> ExtraFields<'a> {
    /// Creates a new iterator over the extra fields in the provided data slice.
    #[inline]
    pub fn new(data: &'a [u8]) -> Self {
        Self { data }
    }

    /// Returns the remaining unparsed bytes in the extra field data.
    #[inline]
    pub fn remaining_bytes(&self) -> &'a [u8] {
        self.data
    }

    #[inline]
    fn next_data(&mut self) -> Option<&'a [u8]> {
        let scratch = self.data;
        if scratch.len() < 4 {
            return None;
        }

        let size = le_u16(&scratch[2..4]) as usize;
        let total_field_len = size + 4;
        if scratch.len() < total_field_len {
            return None;
        }

        let (body, rest) = scratch.split_at(total_field_len);

        // Only advance once we have the entire entry
        self.data = rest;
        Some(body)
    }
}

impl<'a> Iterator for ExtraFields<'a> {
    type Item = (ExtraFieldId, &'a [u8]);

    #[inline]
    fn next(&mut self) -> Option<Self::Item> {
        let next_chunk = self.next_data()?;
        let kind = le_u16(&next_chunk[0..2]);
        let body = &next_chunk[4..];
        Some((ExtraFieldId(kind), body))
    }
}

/// Container for extra fields with a shared data buffer and cached sizes.
#[derive(Debug, Clone)]
pub(crate) struct ExtraFieldsContainer {
    entries: StackVec<Header, 5>,
    data_buffer: StackVec<u8, 15>,
    pub(crate) local_size: u16,
    pub(crate) central_size: u16,
}

impl ExtraFieldsContainer {
    pub fn new() -> Self {
        Self {
            entries: StackVec::new(Header::new(0)),
            data_buffer: StackVec::new(0u8),
            local_size: 0,
            central_size: 0,
        }
    }

    pub fn add_field(
        &mut self,
        id: ExtraFieldId,
        data: &[u8],
        location: Header,
    ) -> Result<(), Error> {
        let size_delta = 4 + data.len();
        let mut current_size = 0;
        if location.includes_local() {
            current_size = self.local_size;
        }
        if location.includes_central() {
            current_size = std::cmp::max(self.central_size, current_size);
        }

        if size_delta + (current_size as usize) > u16::MAX as usize {
            return Err(Error::from(ErrorKind::InvalidInput {
                msg: "extra field data too large".to_string(),
            }));
        }

        let mut buffer = [0u8; 4];
        buffer[0..2].copy_from_slice(&id.as_u16().to_le_bytes());
        buffer[2..4].copy_from_slice(&(data.len() as u16).to_le_bytes());
        self.data_buffer.extend_from_slice(&buffer);
        self.data_buffer.extend_from_slice(data);
        if location.includes_local() {
            self.local_size += size_delta as u16;
        }
        if location.includes_central() {
            self.central_size += size_delta as u16;
        }

        self.entries.push(location);
        Ok(())
    }

    fn write_extra_fields_iter(
        &self,
        writer: &mut impl Write,
        filter: Header,
    ) -> Result<(), Error> {
        let fields = self.data_buffer.as_slice();
        let mut extra_fields = ExtraFields::new(fields);
        let entries = self.entries.as_slice();
        for entry in entries {
            let extra_field = extra_fields.next_data().expect("Entry should have data");
            let write = entry.intersects(filter);
            if write {
                writer.write_all(extra_field)?;
            }
        }
        Ok(())
    }

    #[inline]
    pub fn write_extra_fields(&self, writer: &mut impl Write, filter: Header) -> Result<(), Error> {
        if filter == Header::LOCAL && self.local_size == 0 {
            // No local fields to write
            Ok(())
        } else if filter == Header::CENTRAL && self.central_size == 0 {
            // No central fields to write
            Ok(())
        } else if self.local_size == 0 || self.central_size == 0 {
            // If everything is one sided, we can dump everything
            writer.write_all(self.data_buffer.as_slice())?;
            Ok(())
        } else {
            self.write_extra_fields_iter(writer, filter)
        }
    }
}

/// A stack-first vector that avoids heap allocation for small amounts of data.
///
/// A poor man's `smallvec` as we aren't able to store as many elements inline
/// (by one byte), but it's still an extremely effective no dependency, no
/// unsafe solution, as benchmarks showed a 33% throughput improvement when
/// writing out files with timestamps.
#[derive(Debug, Clone)]
pub(crate) enum StackVec<T, const N: usize>
where
    T: Copy + Clone,
{
    /// Inline storage for up to N elements
    Small { data: [T; N], len: u8 },
    /// Heap storage for more elements
    Large(Vec<T>),
}

impl<T, const N: usize> StackVec<T, N>
where
    T: Copy + Clone,
{
    pub fn new(default_val: T) -> Self {
        Self::Small {
            data: [default_val; N],
            len: 0,
        }
    }

    pub fn push(&mut self, item: T) {
        match self {
            Self::Small { data, len } => {
                if (*len as usize) < N {
                    // Still fits in small storage
                    data[*len as usize] = item;
                    *len += 1;
                } else {
                    // Need to promote to large storage
                    let mut vec = Vec::with_capacity(N + 1);
                    vec.extend_from_slice(&data[..N]);
                    vec.push(item);
                    *self = Self::Large(vec);
                }
            },
            Self::Large(vec) => {
                vec.push(item);
            },
        }
    }

    pub fn as_slice(&self) -> &[T] {
        match self {
            Self::Small { data, len } => &data[..*len as usize],
            Self::Large(vec) => vec.as_slice(),
        }
    }
}

// Specialized methods for StackVec<u8, N> (byte buffers)
impl<const N: usize> StackVec<u8, N> {
    pub fn extend_from_slice(&mut self, slice: &[u8]) {
        match self {
            Self::Small { data, len } => {
                let current_len = *len as usize;
                let end = current_len + slice.len();
                if end <= N {
                    data[current_len..current_len + slice.len()].copy_from_slice(slice);
                    *len += slice.len() as u8;
                } else {
                    // Need to promote to large buffer
                    let mut vec = Vec::with_capacity(current_len + slice.len());
                    vec.extend_from_slice(&data[..current_len]);
                    vec.extend_from_slice(slice);
                    *self = Self::Large(vec);
                }
            },
            Self::Large(vec) => {
                vec.extend_from_slice(slice);
            },
        }
    }
}

#[derive(Debug)]
pub enum StackVecIter<'a, T, const N: usize>
where
    T: Copy + Clone,
{
    Small {
        data: &'a [T; N],
        len: u8,
        index: u8,
    },
    Large(std::slice::Iter<'a, T>),
}

impl<'a, T, const N: usize> Iterator for StackVecIter<'a, T, N>
where
    T: Copy + Clone,
{
    type Item = &'a T;

    fn next(&mut self) -> Option<Self::Item> {
        match self {
            Self::Small { data, len, index } => {
                if *index < *len {
                    let result = &data[*index as usize];
                    *index += 1;
                    Some(result)
                } else {
                    None
                }
            },
            Self::Large(iter) => iter.next(),
        }
    }
}

#[cfg(test)]
mod tests {
    use std::io::Cursor;

    use super::*;

    #[test]
    fn test_partial_parsing_with_remaining_bytes() {
        let data = [0x55, 0x54, 0x01, 0x00, 0xFF, 0x01, 0x00, 0x05];
        let mut iter = ExtraFields::new(&data);
        assert_eq!(iter.remaining_bytes(), &data);

        let (id, body) = iter.next().unwrap();
        assert_eq!(id, ExtraFieldId::EXTENDED_TIMESTAMP);
        assert_eq!(body, &[0xFF]);

        assert_eq!(iter.next(), None);
        assert_eq!(iter.remaining_bytes(), &[0x01, 0x00, 0x05]);
    }

    #[test]
    fn test_unknown_field_id() {
        let data = [0xFF, 0xFF, 0x02, 0x00, 0xDE, 0xAD];
        let mut iter = ExtraFields::new(&data);

        let (id, body) = iter.next().unwrap();
        assert_eq!(id, ExtraFieldId(0xFFFF));
        assert_eq!(body, &[0xDE, 0xAD]);

        assert_eq!(iter.next(), None);
    }

    #[test]
    fn test_stack_vec_u8_inline_operations() {
        let mut buf = StackVec::<u8, 4>::new(0);
        assert_eq!(buf.as_slice(), &[]);

        buf.push(1);
        assert_eq!(buf.as_slice(), &[1]);

        buf.extend_from_slice(&[2, 3]);
        assert_eq!(buf.as_slice(), &[1, 2, 3]);
    }

    #[test]
    fn test_stack_vec_u8_promote_to_heap() {
        let mut buf = StackVec::<u8, 2>::new(0);

        // Fill inline capacity
        buf.extend_from_slice(&[1, 2]);
        assert_eq!(buf.as_slice(), &[1, 2]);

        // Force promotion to heap
        buf.extend_from_slice(&[3, 4, 5]);
        assert_eq!(buf.as_slice(), &[1, 2, 3, 4, 5]);

        buf.push(6);
        assert_eq!(buf.as_slice(), &[1, 2, 3, 4, 5, 6]);
    }

    #[test]
    fn test_stack_vec_size_constraints() {
        // Test that StackVec for bytes is same size as Vec
        assert!(
            std::mem::size_of::<StackVec<u8, 15>>() <= 24,
            "StackVec should not exceed Vec size on 64 bits"
        );
    }

    #[test]
    fn test_stack_vec_clone() {
        let mut buf = StackVec::<u8, 2>::new(0);
        buf.extend_from_slice(&[1, 2, 3]); // Force heap promotion

        let cloned = buf.clone();
        assert_eq!(buf.as_slice(), cloned.as_slice());
    }

    fn round_trip_extra_fields(fields: &[(ExtraFieldId, &[u8], Header)]) {
        let mut container = ExtraFieldsContainer::new();

        for (id, data, location) in fields {
            container.add_field(*id, data, *location).unwrap();
        }

        for location in [Header::LOCAL, Header::CENTRAL] {
            let mut cursor = Cursor::new(Vec::new());
            container.write_extra_fields(&mut cursor, location).unwrap();

            let written_fields = fields
                .iter()
                .filter(|&&(_, _, loc)| loc == location)
                .map(|&(id, data, _)| (id, data))
                .collect::<Vec<_>>();
            let read_fields = ExtraFields::new(cursor.get_ref()).collect::<Vec<_>>();

            assert_eq!(written_fields, read_fields);
        }
    }

    #[test]
    fn test_extra_fields() {
        // Only local extra fields
        round_trip_extra_fields(&[
            (ExtraFieldId::new(0), &[0u8; 16], Header::LOCAL),
            (ExtraFieldId::new(1), &[1u8; 16], Header::LOCAL),
        ]);

        // Only central extra fields
        round_trip_extra_fields(&[
            (ExtraFieldId::new(0), &[0u8; 16], Header::CENTRAL),
            (ExtraFieldId::new(1), &[1u8; 16], Header::CENTRAL),
        ]);

        // Mixed extra fields where the local and central sizes are the same
        round_trip_extra_fields(&[
            (ExtraFieldId::new(0), &[0u8; 16], Header::CENTRAL),
            (ExtraFieldId::new(1), &[1u8; 16], Header::LOCAL),
        ]);
    }
}
