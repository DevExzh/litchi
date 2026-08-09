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
        }
    }
}
impl std::error::Error for WriteError {
    fn source(&self) -> Option<&(dyn std::error::Error + 'static)> {
        match self {
            Self::Io(error) => Some(error),
            Self::Ole(error) => Some(error),
            Self::Vba(error) => Some(error),
            Self::InvalidData(_) => None,
        }
    }
}
