use litchi_cfb::OleError;

/// Error type for DOC writing
#[derive(Debug)]
pub enum WriteError {
    /// I/O error
    Io(std::io::Error),
    /// Invalid data
    InvalidData(String),
    /// OLE error
    Ole(OleError),
    /// MS-OVBA project authoring error
    Vba(litchi_vba::Error),
    /// The staged document assembled, but its package could not be parsed.
    Package(crate::package::Error),
    /// The metadata given to [`super::Builder::set_metadata`] cannot be
    /// stored in the OLE property sets.
    Metadata(litchi_ole_common::property_set::summary_properties::MetadataError),
    /// Document properties were combined with `CryptoAPI` RC4 encryption.
    ///
    /// Office keeps the properties of such a document in an encrypted summary
    /// stream ([MS-OFFCRYPTO] 2.3.5.4), which this writer does not produce;
    /// it refuses rather than store them in plaintext beside the encrypted
    /// document.
    EncryptedMetadata,
}

impl From<std::io::Error> for WriteError {
    fn from(err: std::io::Error) -> Self {
        WriteError::Io(err)
    }
}

impl From<OleError> for WriteError {
    fn from(err: OleError) -> Self {
        WriteError::Ole(err)
    }
}

impl From<litchi_ole_common::property_set::summary_properties::MetadataError> for WriteError {
    fn from(err: litchi_ole_common::property_set::summary_properties::MetadataError) -> Self {
        WriteError::Metadata(err)
    }
}

impl From<litchi_vba::Error> for WriteError {
    fn from(err: litchi_vba::Error) -> Self {
        WriteError::Vba(err)
    }
}

impl std::fmt::Display for WriteError {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        match self {
            WriteError::Io(e) => write!(f, "I/O error: {e}"),
            WriteError::InvalidData(s) => write!(f, "Invalid data: {s}"),
            WriteError::Ole(e) => write!(f, "OLE error: {e}"),
            WriteError::Vba(e) => write!(f, "VBA project error: {e}"),
            WriteError::Package(e) => write!(f, "DOC package error: {e}"),
            WriteError::Metadata(e) => write!(f, "DOC metadata error: {e}"),
            WriteError::EncryptedMetadata => f.write_str(
                "DOC document properties cannot be written with CryptoAPI RC4 encryption",
            ),
        }
    }
}
impl std::error::Error for WriteError {
    fn source(&self) -> Option<&(dyn std::error::Error + 'static)> {
        match self {
            Self::Io(error) => Some(error),
            Self::Ole(error) => Some(error),
            Self::Vba(error) => Some(error),
            Self::Package(error) => Some(error),
            Self::InvalidData(_) => None,
            Self::Metadata(error) => Some(error),
            Self::EncryptedMetadata => None,
        }
    }
}
