//! Fresh Pages documents: [`Package::new`] and the option-bearing
//! [`Builder`] (ADR 0034 §4 and §9).
//!
//! [`Package::new`] returns the minimal valid Pages document: one empty body
//! section in US Letter geometry. It is deterministic: every call publishes
//! the same bytes, carrying the fixed identity of the built-in template.
//!
//! [`Builder`] authors a document with a caller-chosen [`Identity`] and
//! optional initial body text. The identity is a required, explicit input:
//! [`Identity::generate`] gives a native-like fresh identity (the only place
//! this crate consults randomness, and only when called), and
//! [`Identity::from_bytes`] a fixed one for deterministic pipelines and
//! goldens. [`Builder::build`] without an identity is
//! [`PackageError::MissingIdentity`].
//!
//! ```
//! use litchi_pages::builder::{Builder, Identity};
//! use litchi_pages::Package;
//!
//! # fn main() -> Result<(), Box<dyn std::error::Error>> {
//! let blank = Package::new()?;
//! assert_eq!(blank.sections().len(), 1);
//!
//! let identity = Identity::from_bytes([7; 16]);
//! let first = Builder::new().identity(identity).body_text("Hello").build()?;
//! let second = Builder::new().identity(identity).body_text("Hello").build()?;
//! assert_eq!(first.to_bytes()?, second.to_bytes()?);
//! assert_eq!(first.text()?, "Hello");
//! assert!(matches!(
//!     Builder::new().build(),
//!     Err(litchi_pages::PackageError::MissingIdentity)
//! ));
//! # Ok(())
//! # }
//! ```
//!
//! The built-in template was authored once by the legacy
//! `litchi_iwa::pages::Builder` (record 0299 lists its provenance and digest); the
//! crate embeds the bytes, so building performs no protobuf authoring.

use std::fmt;

use crate::{Package, PackageError, PackageResult};

/// The built-in minimal Pages document.
pub(crate) const TEMPLATE: &[u8] = include_bytes!("../assets/blank.pages");

/// The identity of a fresh Pages document: its document UUID (also the
/// `DocumentIdentifier`, `stableDocumentUUID` and `shareUUID`), its version
/// UUID (the properties' `versionUUID` and `revision` and the package
/// metadata's revision) and its private UUID.
#[derive(Clone, Copy, PartialEq, Eq, Hash)]
pub struct Identity {
    document: [u8; 16],
    version: [u8; 16],
    private: [u8; 16],
}

impl fmt::Debug for Identity {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter
            .debug_struct("Identity")
            .field("document", &uuid(self.document))
            .finish_non_exhaustive()
    }
}

/// Force the RFC 4122 version-4 and variant bits.
const fn version_four(mut bytes: [u8; 16]) -> [u8; 16] {
    bytes[6] = (bytes[6] & 0x0f) | 0x40;
    bytes[8] = (bytes[8] & 0x3f) | 0x80;
    bytes
}

impl Identity {
    /// A deterministic seed for the identities a build derives.
    fn seed(&self) -> u64 {
        let mut seed = 0x6c69_7463_6869_0420_u64;
        for chunk in self.document.chunks(8).chain(self.version.chunks(8)) {
            let mut word = [0_u8; 8];
            word.copy_from_slice(chunk);
            seed = litchi_numbers_wire::table_create::mix(seed ^ u64::from_le_bytes(word));
        }
        seed
    }

    /// A fresh identity from the operating system's randomness, as Pages
    /// creates for a new document.
    #[must_use]
    pub fn generate() -> Self {
        Self {
            document: litchi_core::id::generate_guid_bytes(),
            version: litchi_core::id::generate_guid_bytes(),
            private: litchi_core::id::generate_guid_bytes(),
        }
    }

    /// A fixed identity: the document UUID is `bytes` (with the version-4
    /// bits forced) and the version and private UUIDs are derived from it,
    /// so equal inputs give byte-identical documents.
    #[must_use]
    pub const fn from_bytes(bytes: [u8; 16]) -> Self {
        let mut private = bytes;
        let mut version = bytes;
        let mut index = 0;
        while index < 16 {
            private[index] = bytes[15 - index] ^ 0x5a;
            version[index] = bytes[(index + 7) % 16] ^ 0xc3;
            index += 1;
        }
        Self {
            document: version_four(bytes),
            version: version_four(version),
            private: version_four(private),
        }
    }

    /// A fixed identity from both UUIDs (the version-4 bits forced), for a
    /// document that must carry an identity authored elsewhere.
    #[must_use]
    pub const fn from_parts(document: [u8; 16], version: [u8; 16], private: [u8; 16]) -> Self {
        Self {
            document: version_four(document),
            version: version_four(version),
            private: version_four(private),
        }
    }

    /// The document UUID in Pages' uppercase hyphenated form.
    #[must_use]
    pub fn document_uuid(&self) -> String {
        uuid(self.document)
    }

    /// The version UUID in Pages' uppercase hyphenated form.
    #[must_use]
    pub fn version_uuid(&self) -> String {
        uuid(self.version)
    }

    /// The private UUID in Pages' uppercase hyphenated form.
    #[must_use]
    pub fn private_uuid(&self) -> String {
        uuid(self.private)
    }
}

fn uuid(bytes: [u8; 16]) -> String {
    let mut out = String::with_capacity(36);
    for (index, byte) in bytes.iter().enumerate() {
        if matches!(index, 4 | 6 | 8 | 10) {
            out.push('-');
        }
        out.push_str(&format!("{byte:02X}"));
    }
    out
}

/// Option-bearing authoring of a fresh Pages document.
#[derive(Debug, Clone, Default)]
#[must_use]
pub struct Builder {
    identity: Option<Identity>,
    body_text: String,
    language: Option<String>,
    locale: Option<String>,
    tables: Vec<(String, u32, u32)>,
}

impl Builder {
    /// Start a builder with no identity and an empty body.
    pub fn new() -> Self {
        Self::default()
    }

    /// Select the document identity (required).
    pub const fn identity(mut self, identity: Identity) -> Self {
        self.identity = Some(identity);
        self
    }

    /// Set the initial body text; `\n` separates paragraphs.
    pub fn body_text(mut self, text: impl Into<String>) -> Self {
        self.body_text = text.into();
        self
    }

    /// Set the document language (a BCP 47 tag such as `"fr"`), written to
    /// the document, its settings and the body's language, as Pages does.
    /// Unset, the template's `"en"` stays.
    pub fn language(mut self, language: impl Into<String>) -> Self {
        self.language = Some(language.into());
        self
    }

    /// Set the document locale (for example `"fr_FR"`), written to the
    /// document's current and creation locale and the settings' creation
    /// and last locale. Unset, the template's `"en_US"` stays.
    pub fn locale(mut self, locale: impl Into<String>) -> Self {
        self.locale = Some(locale.into());
        self
    }

    /// Add an empty body table `name` of `rows` × `columns` cells (one
    /// header row and one header column included), anchored after the body
    /// text and the tables added before it, as the host builder's
    /// `body_table` places it. The first table is the host builder's
    /// from-scratch table graph (its own stylesheet, calculation engine and
    /// formula owner); later ones are cloned from it with
    /// [`crate::table::creation`].
    pub fn body_table(mut self, name: impl Into<String>, rows: u32, columns: u32) -> Self {
        self.tables.push((name.into(), rows, columns));
        self
    }

    /// Build the document.
    ///
    /// # Errors
    ///
    /// [`PackageError::MissingIdentity`] without an identity;
    /// [`PackageError::InvalidLanguage`] and [`PackageError::InvalidLocale`]
    /// for an invalid language tag or locale;
    /// [`PackageError::InvalidTable`] for a body table with an empty or NUL
    /// name or a size a table cannot take; a typed text error
    /// ([`PackageError::InvalidFormat`]) when the body text holds a reserved
    /// Pages marker character.
    pub fn build(self) -> PackageResult<Package> {
        let identity = self.identity.ok_or(PackageError::MissingIdentity)?;
        if let Some(language) = self.language.as_deref() {
            litchi_iwa_text::TextLanguageTag::new(language)
                .map_err(PackageError::InvalidLanguage)?;
        }
        if self
            .locale
            .as_deref()
            .is_some_and(|locale| locale.trim().is_empty() || locale.contains('\0'))
        {
            return Err(PackageError::InvalidLocale);
        }
        for (name, rows, columns) in &self.tables {
            if name.is_empty()
                || name.contains('\0')
                || !litchi_numbers_wire::table_create::admits_size(*rows, *columns)
            {
                return Err(PackageError::InvalidTable {
                    rows: *rows,
                    columns: *columns,
                });
            }
        }
        let template = Package::new()?;
        let body_text = if self.body_text.is_empty() {
            None
        } else {
            // The body-text transaction's validation (reserved markers,
            // limits); its staged edit is not committed: the text goes out
            // in the one identity publication.
            template
                .edit_body_text()
                .and_then(|mut edit| edit.set(&self.body_text).map(|_| ()))
                .map_err(|error| PackageError::InvalidFormat(error.to_string()))?;
            Some(self.body_text.as_str())
        };
        let package = template.with_identity(
            &identity,
            self.language.as_deref(),
            self.locale.as_deref(),
            body_text,
        )?;
        if let Some(text) = body_text
            && package.text()? != text
        {
            return Err(PackageError::InvalidFormat(
                "the built Pages body text failed verification".to_owned(),
            ));
        }
        let mut tables = self.tables.iter();
        let Some((name, rows, columns)) = tables.next() else {
            return Ok(package);
        };
        let mut package = super::package::builder_table::install(
            &package,
            (name, *rows, *columns),
            self.locale.as_deref().unwrap_or("en_US"),
            identity.seed(),
            None,
        )?;
        for (name, rows, columns) in tables {
            let anchor = package.text()?.encode_utf16().count();
            package = package
                .edit_body_table_creation()
                .insert(
                    litchi_core::Position::new(anchor),
                    crate::table::creation::NewTable::new(*rows, *columns).with_name(name.clone()),
                )
                .and_then(crate::table::creation::Edit::commit)
                .map_err(|error| PackageError::InvalidFormat(error.to_string()))?
                .into_package();
        }
        Ok(package)
    }
}

impl Package {
    /// The minimal valid Pages document: one empty body section.
    ///
    /// Deterministic: every call gives the same bytes and the built-in
    /// template's fixed identity. Use [`Builder`] with
    /// [`Identity::generate`] for a fresh identity.
    ///
    /// # Errors
    ///
    /// Only when the built-in template fails to open, which its tests rule
    /// out.
    pub fn new() -> PackageResult<Self> {
        Self::from_bytes(TEMPLATE)
    }

    /// The same document with `identity` written into its identity
    /// sidecars (`Metadata/DocumentIdentifier` and the UUID keys and
    /// revision of `Metadata/Properties.plist`) and the package metadata's
    /// revision identifier, and with `language`, `locale` and the body text
    /// written where the host builder writes them, in one publication.
    fn with_identity(
        &self,
        identity: &Identity,
        language: Option<&str>,
        locale: Option<&str>,
        body_text: Option<&str>,
    ) -> PackageResult<Self> {
        let document = identity.document_uuid();
        let version = identity.version_uuid();
        let private = identity.private_uuid();
        let properties = super::package::sidecars::properties(self)?;
        let mut dictionary = properties.unwrap_or_default();
        for key in ["documentUUID", "stableDocumentUUID", "shareUUID"] {
            dictionary.insert(key.to_owned(), plist::Value::String(document.clone()));
        }
        dictionary.insert("privateUUID".to_owned(), plist::Value::String(private));
        dictionary.insert(
            "versionUUID".to_owned(),
            plist::Value::String(version.clone()),
        );
        let sequence = dictionary
            .get("revision")
            .and_then(plist::Value::as_string)
            .and_then(|revision| revision.split_once("::"))
            .map_or_else(|| "0".to_owned(), |(sequence, _)| sequence.to_owned());
        dictionary.insert(
            "revision".to_owned(),
            plist::Value::String(format!("{sequence}::{version}")),
        );
        let encoded = super::package::sidecars::encode(&dictionary)?;
        let mut rewrites = self.language_and_locale_rewrites(language, locale)?;
        rewrites.extend(self.revision_rewrite(&version)?);
        if let Some(text) = body_text {
            self.body_text_rewrite(&mut rewrites, text)?;
        }
        super::package::body_drawable_graph::publish_with_replacements(
            self,
            rewrites,
            &[
                (super::package::sidecars::PROPERTIES, &encoded),
                (super::package::sidecars::IDENTIFIER, document.as_bytes()),
            ],
        )
        .map_err(|error| PackageError::InvalidFormat(format!("{error:?}")))
    }
}

#[cfg(test)]
mod tests {
    use super::{Builder, Identity, TEMPLATE};
    use crate::{Package, PackageError};

    #[test]
    fn the_template_opens_and_is_deterministic() {
        let first = Package::new().unwrap_or_else(|error| panic!("template: {error}"));
        let second = Package::new().unwrap_or_else(|error| panic!("template: {error}"));
        assert_eq!(first.sections().len(), 1);
        assert_eq!(first.text().ok().as_deref(), Some(""));
        assert_eq!(first.to_bytes().ok().as_deref(), Some(TEMPLATE));
        assert_eq!(first.to_bytes().ok(), second.to_bytes().ok());
    }

    #[test]
    fn fixed_identities_are_byte_identical_and_distinct_from_each_other() {
        let build = |seed: u8| {
            Builder::new()
                .identity(Identity::from_bytes([seed; 16]))
                .body_text("One\nTwo")
                .build()
                .unwrap_or_else(|error| panic!("build: {error}"))
        };
        let (first, again, other) = (build(1), build(1), build(2));
        assert_eq!(first.to_bytes().ok(), again.to_bytes().ok());
        assert_ne!(first.to_bytes().ok(), other.to_bytes().ok());
        assert_eq!(first.text().ok().as_deref(), Some("One\nTwo"));
        let identifier = first.metadata().identifier;
        assert_eq!(
            identifier.as_deref(),
            Some(Identity::from_bytes([1; 16]).document_uuid().as_str())
        );
    }

    #[test]
    fn a_build_without_identity_is_refused() {
        assert!(matches!(
            Builder::new().body_text("x").build(),
            Err(PackageError::MissingIdentity)
        ));
    }

    #[test]
    fn invalid_languages_and_locales_are_refused() {
        let builder = || Builder::new().identity(Identity::from_bytes([3; 16]));
        assert!(matches!(
            builder().language(" ").build(),
            Err(PackageError::InvalidLanguage(_))
        ));
        assert!(matches!(
            builder().locale("").build(),
            Err(PackageError::InvalidLocale)
        ));
        let built = builder()
            .language("fr")
            .locale("fr_FR")
            .body_text("Bonjour")
            .build()
            .unwrap_or_else(|error| panic!("build: {error}"));
        assert_eq!(built.text().ok().as_deref(), Some("Bonjour"));
    }

    #[test]
    fn body_tables_are_built_from_scratch() {
        let build = |seed: u8| {
            Builder::new()
                .identity(Identity::from_bytes([seed; 16]))
                .body_text("Budget")
                .body_table("Q1", 4, 3)
                .body_table("Q2", 2, 2)
                .build()
                .unwrap_or_else(|error| panic!("build: {error}"))
        };
        let (first, again) = (build(5), build(5));
        assert_eq!(first.to_bytes().ok(), again.to_bytes().ok());
        let tables = first
            .body_tables()
            .unwrap_or_else(|error| panic!("tables: {error}"));
        let shapes: Vec<_> = tables
            .iter()
            .map(|table| (table.name().to_owned(), table.rows(), table.columns()))
            .collect();
        assert_eq!(shapes, [("Q1".to_owned(), 4, 3), ("Q2".to_owned(), 2, 2)]);
        assert_eq!(first.text().ok().as_deref(), Some("Budget\u{fffc}\u{fffc}"));
        let builder = || Builder::new().identity(Identity::from_bytes([3; 16]));
        assert!(matches!(
            builder().body_table("", 2, 2).build(),
            Err(PackageError::InvalidTable { .. })
        ));
        assert!(matches!(
            builder().body_table("T", 0, 2).build(),
            Err(PackageError::InvalidTable {
                rows: 0,
                columns: 2
            })
        ));
    }

    #[test]
    fn uuids_use_the_pages_spelling() {
        let identity = Identity::from_bytes([0xab; 16]);
        let text = identity.document_uuid();
        assert_eq!(text.len(), 36);
        assert_eq!(&text[14..15], "4");
        assert!(text.chars().all(|c| c == '-' || c.is_ascii_hexdigit()));
        assert_ne!(identity.document_uuid(), identity.private_uuid());
    }
}
