use super::super::{OleError, OleFile};
/// Package implementation for legacy Word documents (.doc).
use super::document::Document;
use std::fs::File;
use std::io::{self, Read, Seek};
use std::path::Path;

/// Options controlling how a legacy Word document is opened.
#[derive(Debug, Clone, Copy, Default)]
pub struct DocOpenOptions<'a> {
    /// Password used for password-to-open encryption.
    pub password: Option<&'a str>,
    /// How non-structural stylesheet defects are treated.
    ///
    /// Defaults to [`crate::doc::DocLeniency::Strict`], which is the historical behaviour.
    pub leniency: super::leniency::DocLeniency,
}

/// Password-to-open encryption schemes identified in a DOC file.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum DocEncryptionKind {
    /// Legacy Word XOR obfuscation.
    XorObfuscation,
    /// Office CryptoAPI encryption.
    CryptoApi,
    /// An encryption version not recognized by this implementation.
    Unknown {
        /// Encryption header major version.
        major: u16,
        /// Encryption header minor version.
        minor: u16,
    },
}

/// Error types for DOC file parsing.
#[derive(Debug)]
pub enum DocError {
    /// IO error
    Io(io::Error),
    /// OLE file error
    Ole(OleError),
    /// Invalid DOC format
    InvalidFormat(String),
    /// Stream not found
    StreamNotFound(String),
    /// Corrupted file
    Corrupted(String),
    /// The document is encrypted and no password was supplied.
    PasswordRequired,
    /// The supplied password did not validate.
    InvalidPassword,
    /// The document uses a recognized but unsupported encryption scheme.
    UnsupportedEncryption(DocEncryptionKind),
    /// The clear encryption header is malformed.
    MalformedEncryptionHeader(String),
    /// The file predates the Word 97 binary format this reader implements.
    ///
    /// Word 6.0 and Word 95 documents keep the structures that MS-DOC places
    /// in a table stream inside `WordDocument` instead, so they cannot be read
    /// as MS-DOC no matter how tolerant the parser is.
    UnsupportedVersion {
        /// The `nFib` value found in the FIB.
        nfib: u16,
        /// Human-readable name of the originating Word release.
        name: &'static str,
    },
}

impl From<io::Error> for DocError {
    fn from(err: io::Error) -> Self {
        DocError::Io(err)
    }
}

impl From<OleError> for DocError {
    fn from(err: OleError) -> Self {
        DocError::Ole(err)
    }
}

impl std::fmt::Display for DocError {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        match self {
            DocError::Io(e) => write!(f, "IO error: {}", e),
            DocError::Ole(e) => write!(f, "OLE error: {}", e),
            DocError::InvalidFormat(s) => write!(f, "Invalid format: {}", s),
            DocError::StreamNotFound(s) => write!(f, "Stream not found: {}", s),
            DocError::Corrupted(s) => write!(f, "Corrupted file: {}", s),
            DocError::PasswordRequired => write!(f, "a password is required to open this document"),
            DocError::InvalidPassword => write!(f, "the document password is invalid"),
            DocError::UnsupportedEncryption(kind) => {
                write!(f, "unsupported DOC encryption: {kind:?}")
            },
            DocError::MalformedEncryptionHeader(s) => {
                write!(f, "malformed DOC encryption header: {s}")
            },
            DocError::UnsupportedVersion { nfib, name } => write!(
                f,
                "{name} documents (nFib {nfib:#06x}) predate the Word 97 binary format and are not supported"
            ),
        }
    }
}

impl std::error::Error for DocError {}

/// Result type for DOC operations.
pub type Result<T> = std::result::Result<T, DocError>;

/// A Word (.doc) package.
///
/// This is the main entry point for working with legacy Word documents.
/// It wraps an OLE file and provides Word-specific functionality.
///
/// # Examples
///
/// ```rust,no_run
/// use litchi_ole::doc::Package;
///
/// // Open an existing document
/// let mut pkg = Package::open("document.doc")?;
///
/// // Get the main document
/// let doc = pkg.document()?;
///
/// // Extract text
/// let text = doc.text()?;
/// println!("{}", text);
/// # Ok::<(), Box<dyn std::error::Error>>(())
/// ```
pub struct Package<R: Read + Seek = File> {
    /// The underlying OLE file
    ole: OleFile<R>,
}

impl Package<File> {
    /// Open a .doc package from a file path.
    ///
    /// # Arguments
    ///
    /// * `path` - Path to the .doc file
    ///
    /// # Examples
    ///
    /// ```rust,no_run
    /// use litchi_ole::doc::Package;
    ///
    /// let pkg = Package::open("document.doc")?;
    /// # Ok::<(), Box<dyn std::error::Error>>(())
    /// ```
    pub fn open<P: AsRef<Path>>(path: P) -> Result<Self> {
        let file = File::open(path)?;
        Package::from_reader(file)
    }
}

impl<R: Read + Seek> Package<R> {
    /// Create a Package from any reader that implements Read + Seek.
    ///
    /// # Arguments
    ///
    /// * `reader` - A reader containing the .doc file data
    ///
    /// # Examples
    ///
    /// ```rust,no_run
    /// use std::fs::File;
    /// use litchi_ole::doc::Package;
    ///
    /// let file = File::open("document.doc")?;
    /// let pkg = Package::from_reader(file)?;
    /// # Ok::<(), Box<dyn std::error::Error>>(())
    /// ```
    pub fn from_reader(reader: R) -> Result<Self> {
        let ole = OleFile::open(reader)?;

        // Verify it's a Word document by checking for the WordDocument stream
        if !ole.exists(&["WordDocument"]) {
            return Err(DocError::InvalidFormat(
                "Not a valid Word document: WordDocument stream not found".to_string(),
            ));
        }

        Ok(Self { ole })
    }

    /// Create a Package from an already-parsed OLE file.
    ///
    /// This is used for single-pass parsing where the OLE file has already
    /// been parsed during format detection. It avoids double-parsing.
    ///
    /// # Arguments
    ///
    /// * `ole` - An already-parsed OLE file
    ///
    /// # Examples
    ///
    /// ```rust,no_run
    /// use litchi_ole::{OleFile, doc::Package};
    /// use std::io::Cursor;
    ///
    /// let bytes = std::fs::read("document.doc")?;
    /// let ole = OleFile::open(Cursor::new(bytes))?;
    /// let pkg = Package::from_ole_file(ole)?;
    /// # Ok::<(), Box<dyn std::error::Error>>(())
    /// ```
    pub fn from_ole_file(ole: OleFile<R>) -> Result<Self> {
        // Verify it's a Word document by checking for the WordDocument stream
        if !ole.exists(&["WordDocument"]) {
            return Err(DocError::InvalidFormat(
                "Not a valid Word document: WordDocument stream not found".to_string(),
            ));
        }

        Ok(Self { ole })
    }

    /// Get the main document.
    ///
    /// Returns the `Document` object which provides access to the document's
    /// content, formatting, tables, and other features.
    ///
    /// # Examples
    ///
    /// ```rust,no_run
    /// use litchi_ole::doc::Package;
    ///
    /// let mut pkg = Package::open("document.doc")?;
    /// let doc = pkg.document()?;
    /// # Ok::<(), Box<dyn std::error::Error>>(())
    /// ```
    pub fn document(&mut self) -> Result<Document> {
        Document::from_ole(&mut self.ole)
    }

    /// Get the main document using explicit password-to-open options.
    pub fn document_with_options(&mut self, options: DocOpenOptions<'_>) -> Result<Document> {
        Document::from_ole_with_options(&mut self.ole, options)
    }

    /// Get the underlying OLE file.
    ///
    /// This provides access to lower-level OLE operations and streams.
    #[inline]
    pub fn ole_file(&mut self) -> &mut OleFile<R> {
        &mut self.ole
    }

    /// Discover the optional MS-DOC `Macros` project storage without opening streams.
    ///
    /// This inspects CFB directory names only. It never opens, decompresses,
    /// parses, or executes the `PROJECT`, `dir`, `_VBA_PROJECT`, `__SRP_*`,
    /// or candidate module streams.
    pub fn vba_project_storages(&self) -> Vec<super::vba::VbaProjectStorage> {
        super::vba::discover_vba_project_storages(&self.ole.list_streams())
    }

    /// Discover the optional MS-DOC `Macros` project storage.
    pub fn vba_project_storage(&self) -> Option<super::vba::VbaProjectStorage> {
        self.vba_project_storages().into_iter().next()
    }

    /// Parse the optional MS-DOC VBA project with safe default limits.
    ///
    /// Source is decompressed and decoded according to the project code page,
    /// but is never compiled, interpreted, or executed.
    pub fn vba(
        &mut self,
    ) -> std::result::Result<Option<litchi_vba::project::Project>, litchi_vba::Error> {
        self.vba_with(&litchi_vba::Limits::default())
    }

    /// Parse the optional MS-DOC VBA project with explicit resource limits.
    pub fn vba_with(
        &mut self,
        limits: &litchi_vba::Limits,
    ) -> std::result::Result<Option<litchi_vba::project::Project>, litchi_vba::Error> {
        let Some(storage) = self.vba_project_storage() else {
            return Ok(None);
        };
        if !storage.is_structurally_complete() {
            return Ok(None);
        }
        let path: Vec<&str> = storage
            .project_root_path()
            .iter()
            .map(String::as_str)
            .collect();
        litchi_vba::project::Project::open(&mut self.ole, &path, limits).map(Some)
    }

    /// Read the legacy Custom XML Data Storage without resolving schema URIs.
    pub fn custom_xml_data_store(
        &mut self,
    ) -> litchi_ole_common::custom_xml_data::Result<
        Option<litchi_ole_common::custom_xml_data::MsoDataStore>,
    > {
        litchi_ole_common::custom_xml_data::inspect_mso_data_store(&mut self.ole)
    }

    pub fn summary_information(&mut self) -> Result<Option<litchi_cfb::PropertySetStream>> {
        match self
            .ole
            .property_set_stream(&["\u{0005}SummaryInformation"])
        {
            Ok(value) => Ok(Some(value)),
            Err(litchi_cfb::OleError::StreamNotFound) => Ok(None),
            Err(error) => Err(error.into()),
        }
    }

    /// Verify document XML signatures without evaluating certificate trust or
    /// opening any VBA project stream.
    pub fn verify_digital_signatures(
        &mut self,
        policy: &crate::signature::SignatureVerificationPolicy,
    ) -> crate::signature::Result<Vec<crate::signature::BinaryOfficeSignatureVerification>> {
        crate::signature::verify_binary_office_signatures(
            &mut self.ole,
            crate::signature::BinaryOfficeFormat::Doc,
            policy,
        )
    }

    pub fn document_summary_information(
        &mut self,
    ) -> Result<Option<litchi_cfb::PropertySetStream>> {
        match self
            .ole
            .property_set_stream(&["\u{0005}DocumentSummaryInformation"])
        {
            Ok(value) => Ok(Some(value)),
            Err(litchi_cfb::OleError::StreamNotFound) => Ok(None),
            Err(error) => Err(error.into()),
        }
    }

    pub fn user_defined_properties(&mut self) -> Result<Option<litchi_cfb::PropertySet>> {
        Ok(self.document_summary_information()?.and_then(|stream| {
            stream
                .section(litchi_cfb::USER_DEFINED_PROPERTIES_FMTID)
                .cloned()
        }))
    }
}

// `From<DocError> for litchi_core::Error` lives here (not in the umbrella) so
// the orphan rule is satisfied — both source and target crates are external
// to the umbrella.
impl From<DocError> for litchi_core::Error {
    fn from(err: DocError) -> Self {
        match err {
            DocError::Io(e) => litchi_core::Error::Io(e),
            DocError::Ole(ole_err) => litchi_core::Error::from(ole_err),
            DocError::InvalidFormat(s) => litchi_core::Error::InvalidFormat(s),
            DocError::StreamNotFound(s) => litchi_core::Error::ComponentNotFound(s),
            DocError::Corrupted(s) => litchi_core::Error::CorruptedFile(s),
            DocError::PasswordRequired => {
                litchi_core::Error::InvalidFormat("DOC password required".to_string())
            },
            DocError::InvalidPassword => {
                litchi_core::Error::InvalidFormat("invalid DOC password".to_string())
            },
            DocError::UnsupportedEncryption(kind) => {
                litchi_core::Error::InvalidFormat(format!("unsupported DOC encryption: {kind:?}"))
            },
            DocError::MalformedEncryptionHeader(s) => litchi_core::Error::CorruptedFile(s),
            DocError::UnsupportedVersion { nfib, name } => litchi_core::Error::InvalidFormat(
                format!("{name} documents (nFib {nfib:#06x}) are not supported"),
            ),
        }
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    #[ignore] // Requires test file
    fn test_open_package() {
        let result = Package::open("test.doc");
        assert!(result.is_ok());
    }

    #[test]
    #[ignore] // Requires test file
    fn test_invalid_file() {
        // Create a non-DOC file
        std::fs::write("test_invalid.tmp", b"Not a DOC file").unwrap();
        let result = Package::open("test_invalid.tmp");
        assert!(result.is_err());
        std::fs::remove_file("test_invalid.tmp").ok();
    }

    fn poi_fixture(name: &str) -> std::path::PathBuf {
        std::path::Path::new(env!("CARGO_MANIFEST_DIR"))
            .join("../../test-data/poi/test-data/document")
            .join(name)
    }

    #[test]
    fn opens_apache_poi_binary_rc4_document() {
        let path = poi_fixture("password_tika_binaryrc4.doc");

        let mut package = Package::open(&path).unwrap();
        assert!(matches!(
            package.document(),
            Err(DocError::PasswordRequired)
        ));

        let mut package = Package::open(&path).unwrap();
        assert!(matches!(
            package.document_with_options(DocOpenOptions {
                password: Some("wrong"),
                ..Default::default()
            }),
            Err(DocError::InvalidPassword)
        ));

        let mut package = Package::open(path).unwrap();
        let document = package
            .document_with_options(DocOpenOptions {
                password: Some("tika"),
                ..Default::default()
            })
            .unwrap();
        assert!(!document.text().unwrap().trim().is_empty());
    }

    #[test]
    fn opens_apache_poi_cryptoapi_document() {
        let path = poi_fixture("password_password_cryptoapi.doc");

        let mut package = Package::open(&path).unwrap();
        assert!(matches!(
            package.document(),
            Err(DocError::PasswordRequired)
        ));

        let mut package = Package::open(&path).unwrap();
        assert!(matches!(
            package.document_with_options(DocOpenOptions {
                password: Some("wrong"),
                ..Default::default()
            }),
            Err(DocError::InvalidPassword)
        ));

        let mut package = Package::open(path).unwrap();
        let document = package
            .document_with_options(DocOpenOptions {
                password: Some("password"),
                ..Default::default()
            })
            .unwrap();
        assert!(!document.text().unwrap().trim().is_empty());
    }
}
