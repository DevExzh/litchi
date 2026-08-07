use super::codec::encode_record;
use super::{FontCollections, PackageLimits, PackageOptions};
use crate::package::{Error, Package, Result};
use crate::records::Record;
use std::io::Cursor;
use std::sync::Arc;

/// Immutable whole-CFB font snapshot bound to the exact live document owner.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Snapshot {
    pub(crate) bytes: Arc<[u8]>,
    pub(crate) document: Arc<[u8]>,
    pub(crate) document_record: Arc<Record>,
    pub(crate) fonts: FontCollections,
    pub(crate) document_persist_id: u32,
    pub(crate) limits: PackageLimits,
}

impl Snapshot {
    pub fn parse(bytes: &[u8]) -> Result<Self> {
        Self::parse_with_limits(bytes, PackageLimits::default())
    }

    /// Copy a borrowed package only after enforcing the caller's source limit.
    pub fn parse_with_limits(bytes: &[u8], limits: PackageLimits) -> Result<Self> {
        validate_source_len(bytes.len(), limits)?;
        let mut owned = Vec::new();
        owned
            .try_reserve_exact(bytes.len())
            .map_err(|_| Error::AllocationFailed("PowerPoint font snapshot source"))?;
        owned.extend_from_slice(bytes);
        Self::from_arc(Arc::from(owned), limits)
    }

    pub fn from_bytes(bytes: Vec<u8>) -> Result<Self> {
        Self::from_bytes_with_limits(bytes, PackageLimits::default())
    }

    pub fn from_bytes_with_limits(bytes: Vec<u8>, limits: PackageLimits) -> Result<Self> {
        Self::from_bytes_with_options(
            bytes,
            PackageOptions {
                password: None,
                limits,
            },
        )
    }

    pub fn from_bytes_with_options(bytes: Vec<u8>, options: PackageOptions<'_>) -> Result<Self> {
        validate_source_len(bytes.len(), options.limits)?;
        Self::from_arc_with_options(Arc::from(bytes), options)
    }

    pub(crate) fn from_arc(bytes: Arc<[u8]>, limits: PackageLimits) -> Result<Self> {
        Self::from_arc_with_options(
            bytes,
            PackageOptions {
                password: None,
                limits,
            },
        )
    }

    fn from_arc_with_options(bytes: Arc<[u8]>, options: PackageOptions<'_>) -> Result<Self> {
        validate_source_len(bytes.len(), options.limits)?;
        let mut ingress_limits = options.limits.fonts.records;
        ingress_limits.max_package_bytes = ingress_limits
            .max_package_bytes
            .min(options.limits.max_source_bytes);
        let mut package =
            Package::from_reader_with_limits(Cursor::new(bytes.clone()), ingress_limits)?;
        #[cfg(feature = "encryption")]
        let presentation = package.presentation_with_options_and_limits(
            crate::OpenOptions {
                password: options.password,
            },
            options.limits.fonts.records,
        )?;
        #[cfg(not(feature = "encryption"))]
        let presentation = {
            if options.password.is_some() {
                return Err(Error::InvalidFormat(
                    "password opening requires the encryption feature".into(),
                ));
            }
            package.presentation_with_limits(options.limits.fonts.records)?
        };
        let (document_persist_id, source) =
            match crate::embedded::object::Editor::inspect_live_document(&bytes) {
                Ok(value) => value,
                Err(_error) if options.password.is_some() => {
                    let offset = presentation.slide_directory().document_offset();
                    let stream = presentation.document_stream();
                    let (_, consumed) = Record::parse_strict_with_limits(
                        stream,
                        offset,
                        options.limits.fonts.records,
                    )?;
                    let end = offset
                        .checked_add(consumed)
                        .ok_or_else(|| Error::Corrupted("live document range overflow".into()))?;
                    let source = stream.get(offset..end).ok_or_else(|| {
                        Error::Corrupted("live document exceeds its stream".into())
                    })?;
                    (
                        presentation.slide_directory().document_persist_id(),
                        source.to_vec(),
                    )
                },
                Err(error) => return Err(error),
            };
        let (mut record, consumed) =
            Record::parse_strict_with_limits(&source, 0, options.limits.fonts.records)?;
        if record.record_type != crate::RecordType::Document {
            return Err(Error::Corrupted(
                "live document persist owner is not a DocumentContainer".into(),
            ));
        }
        if consumed != source.len() {
            return Err(Error::Corrupted(
                "live Document persist record has trailing bytes".into(),
            ));
        }
        let encoded = encode_record(&record, options.limits.fonts.records)?;
        if encoded != source.as_slice() {
            return Err(Error::InvalidFormat(
                "live font owner is not losslessly representable".into(),
            ));
        }
        let fonts = FontCollections::take_from_document(&mut record, options.limits.fonts)?;
        Ok(Self {
            bytes,
            document: Arc::from(source.into_boxed_slice()),
            document_record: Arc::new(record),
            fonts,
            document_persist_id,
            limits: options.limits,
        })
    }

    pub fn bytes(&self) -> &[u8] {
        &self.bytes
    }
    pub fn document_bytes(&self) -> &[u8] {
        &self.document
    }
    pub const fn fonts(&self) -> &FontCollections {
        &self.fonts
    }
    pub const fn document_persist_id(&self) -> u32 {
        self.document_persist_id
    }
    pub const fn limits(&self) -> PackageLimits {
        self.limits
    }
    pub fn revision(&self) -> super::Revision {
        super::Revision::from_bytes(&self.bytes)
    }

    pub fn edit(&self) -> Result<super::Transaction> {
        super::Transaction::new(self.clone())
    }
}

fn validate_source_len(len: usize, limits: PackageLimits) -> Result<()> {
    if limits.max_source_bytes == 0 || len > limits.max_source_bytes {
        return Err(Error::ResourceLimit(format!(
            "PowerPoint font snapshot source size {len} exceeds limit {}",
            limits.max_source_bytes
        )));
    }
    Ok(())
}
