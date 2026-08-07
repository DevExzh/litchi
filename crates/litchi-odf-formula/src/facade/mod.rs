//! Ergonomic Formula document facade.

use litchi_core::Result;
use std::io::Read;
use std::path::Path;

use crate::codec;
use crate::model::Element;

/// An immutable-or-transactionally-edited `OpenDocument` Formula package.
pub struct Formula {
    package: crate::package::Package,
    root: Element,
}

impl Formula {
    /// Create a standard `.odf` package from validated `MathML`.
    ///
    /// # Errors
    ///
    /// Returns an error when the `MathML` fails validation or the package
    /// cannot be created.
    pub fn create(mathml: impl AsRef<str>) -> Result<Self> {
        Self::create_with_flavor(mathml.as_ref(), false)
    }

    /// Create a Formula template `.otf` package from validated `MathML`.
    ///
    /// # Errors
    ///
    /// Returns an error when the `MathML` fails validation or the package
    /// cannot be created.
    pub fn create_template(mathml: impl AsRef<str>) -> Result<Self> {
        Self::create_with_flavor(mathml.as_ref(), true)
    }

    fn create_with_flavor(mathml: &str, template: bool) -> Result<Self> {
        let root = codec::parse(mathml)?;
        let package = crate::package::Package::create(mathml.as_bytes(), template)?;
        Ok(Self { package, root })
    }

    /// Open a Formula package from a path.
    ///
    /// # Errors
    ///
    /// Returns an error when the file cannot be opened or its contents are
    /// not a valid Formula package.
    pub fn open(path: impl AsRef<Path>) -> Result<Self> {
        let file = std::fs::File::open(path)?;
        Self::from_reader(file)
    }

    /// Read a Formula package from a stream.
    ///
    /// # Errors
    ///
    /// Returns an error when reading fails or the data is not a valid Formula
    /// package.
    pub fn from_reader(reader: impl Read) -> Result<Self> {
        Self::from_bytes(read_all(reader)?)
    }

    /// Read and validate a Formula package from owned bytes.
    ///
    /// # Errors
    ///
    /// Returns an error when the bytes are not a valid Formula package.
    pub fn from_bytes(bytes: Vec<u8>) -> Result<Self> {
        let package = crate::package::Package::from_bytes(bytes)?;
        let root = codec::parse(&package.content_xml()?)?;
        Ok(Self { package, root })
    }

    /// Return the complete `MathML` `math` root.
    #[must_use]
    pub fn root(&self) -> &Element {
        &self.root
    }

    /// Replace the `MathML` root and atomically rebuild the package.
    ///
    /// # Errors
    ///
    /// Returns an error when the serialized root fails re-validation or the
    /// rebuilt package cannot be assembled.
    pub fn set_root(&mut self, root: Element) -> Result<()> {
        let xml = codec::serialize(&root);
        drop(root);
        let validated = codec::parse(&xml)?;
        let package = self.package.replace_content(xml.as_bytes())?;
        self.package = package;
        self.root = validated;
        Ok(())
    }

    /// Return every `MathML` annotation in document order.
    #[must_use]
    pub fn annotations(&self) -> Vec<&Element> {
        let mut annotations = Vec::new();
        self.root.collect_annotations(&mut annotations);
        annotations
    }

    /// Return the first `StarMath` annotation source, when present.
    #[must_use]
    pub fn starmath_source(&self) -> Option<String> {
        self.annotations().into_iter().find_map(|annotation| {
            annotation
                .attribute(None, "encoding")
                .is_some_and(|encoding| encoding.eq_ignore_ascii_case("StarMath 5.0"))
                .then(|| annotation.all_text())
        })
    }

    /// Return the exact package MIME type.
    #[must_use]
    pub fn mimetype(&self) -> &'static str {
        self.package.mimetype()
    }

    /// Whether this package uses the Formula template MIME type.
    #[must_use]
    pub const fn is_template(&self) -> bool {
        self.package.is_template()
    }

    /// Read the exact package `content.xml` as UTF-8.
    ///
    /// # Errors
    ///
    /// Returns an error when `content.xml` is missing or not valid UTF-8.
    pub fn content_xml(&self) -> Result<String> {
        self.package.content_xml()
    }

    /// List package members in archive order.
    ///
    /// # Errors
    ///
    /// Returns an error when the archive member list cannot be read.
    pub fn files(&self) -> Result<Vec<String>> {
        self.package.files()
    }

    /// Return the exact original package bytes.
    #[must_use]
    pub fn as_bytes(&self) -> &[u8] {
        self.package.as_bytes()
    }

    /// Clone the exact package bytes.
    #[must_use]
    pub fn to_bytes(&self) -> Vec<u8> {
        self.as_bytes().to_vec()
    }

    /// Consume the facade and return the exact package bytes.
    #[must_use]
    pub fn into_bytes(self) -> Vec<u8> {
        self.package.into_bytes()
    }

    /// Save the exact package bytes without rebuilding the archive.
    ///
    /// # Errors
    ///
    /// Returns an error when the file cannot be written.
    pub fn save(&self, path: impl AsRef<Path>) -> Result<()> {
        std::fs::write(path, self.as_bytes())?;
        Ok(())
    }
}

fn read_all(mut reader: impl Read) -> Result<Vec<u8>> {
    let mut bytes = Vec::new();
    reader.read_to_end(&mut bytes)?;
    Ok(bytes)
}
