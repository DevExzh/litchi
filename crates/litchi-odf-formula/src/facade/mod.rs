//! Ergonomic Formula document facade.

use litchi_core::Result;
use std::io::Read;
use std::path::Path;

use crate::codec;
use crate::model::Element;

/// An immutable-or-transactionally-edited OpenDocument Formula package.
pub struct Formula {
    package: crate::package::Package,
    root: Element,
}

impl Formula {
    /// Create a standard `.odf` package from validated MathML.
    pub fn create(mathml: impl AsRef<str>) -> Result<Self> {
        Self::create_with_flavor(mathml.as_ref(), false)
    }

    /// Create a Formula template `.otf` package from validated MathML.
    pub fn create_template(mathml: impl AsRef<str>) -> Result<Self> {
        Self::create_with_flavor(mathml.as_ref(), true)
    }

    fn create_with_flavor(mathml: &str, template: bool) -> Result<Self> {
        let root = codec::parse(mathml)?;
        let package = crate::package::Package::create(mathml.as_bytes(), template)?;
        Ok(Self { package, root })
    }

    /// Open a Formula package from a path.
    pub fn open(path: impl AsRef<Path>) -> Result<Self> {
        let file = std::fs::File::open(path)?;
        Self::from_reader(file)
    }

    /// Read a Formula package from a stream.
    pub fn from_reader(reader: impl Read) -> Result<Self> {
        Self::from_bytes(read_all(reader)?)
    }

    /// Read and validate a Formula package from owned bytes.
    pub fn from_bytes(bytes: Vec<u8>) -> Result<Self> {
        let package = crate::package::Package::from_bytes(bytes)?;
        let root = codec::parse(&package.content_xml()?)?;
        Ok(Self { package, root })
    }

    /// Return the complete MathML `math` root.
    pub fn root(&self) -> &Element {
        &self.root
    }

    /// Replace the MathML root and atomically rebuild the package.
    pub fn set_root(&mut self, root: Element) -> Result<()> {
        let xml = codec::serialize(&root);
        let validated = codec::parse(&xml)?;
        let package = self.package.replace_content(xml.as_bytes())?;
        self.package = package;
        self.root = validated;
        Ok(())
    }

    /// Return every MathML annotation in document order.
    pub fn annotations(&self) -> Vec<&Element> {
        let mut annotations = Vec::new();
        self.root.collect_annotations(&mut annotations);
        annotations
    }

    /// Return the first StarMath annotation source, when present.
    pub fn starmath_source(&self) -> Option<String> {
        self.annotations().into_iter().find_map(|annotation| {
            annotation
                .attribute(None, "encoding")
                .is_some_and(|encoding| encoding.eq_ignore_ascii_case("StarMath 5.0"))
                .then(|| annotation.all_text())
        })
    }

    /// Return the exact package MIME type.
    pub fn mimetype(&self) -> &'static str {
        self.package.mimetype()
    }

    /// Whether this package uses the Formula template MIME type.
    pub const fn is_template(&self) -> bool {
        self.package.is_template()
    }

    /// Read the exact package `content.xml` as UTF-8.
    pub fn content_xml(&self) -> Result<String> {
        self.package.content_xml()
    }

    /// List package members in archive order.
    pub fn files(&self) -> Result<Vec<String>> {
        self.package.files()
    }

    /// Return the exact original package bytes.
    pub fn as_bytes(&self) -> &[u8] {
        self.package.as_bytes()
    }

    /// Clone the exact package bytes.
    pub fn to_bytes(&self) -> Vec<u8> {
        self.as_bytes().to_vec()
    }

    /// Consume the facade and return the exact package bytes.
    pub fn into_bytes(self) -> Vec<u8> {
        self.package.into_bytes()
    }

    /// Save the exact package bytes without rebuilding the archive.
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
