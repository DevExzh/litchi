//! Unified creation of new package owners.
//!
//! Creation is deliberately separate from opening: a new package is built by
//! the format leaf and adopted by its unified owner directly.  In particular,
//! DOCX creation does not serialize a package and reopen it through the source
//! detector.  This keeps the leaf's typed package state and makes the owner
//! used by the returned [`File`] the same owner the creator produced.

use std::fmt;

use litchi_core::{FileFormat, Metadata};

use crate::File;
use crate::error::{Code, Error, Result};

/// Whether [`create_with`] makes an ordinary document or a template.
///
/// The content flavor is explicit.  It is never inferred from a destination
/// filename, because creation has no destination and a filename cannot alter
/// the package's content type.
#[derive(Clone, Copy, PartialEq, Eq, Hash, Default)]
#[non_exhaustive]
pub enum Flavor {
    /// An ordinary document.
    #[default]
    Document,
    /// A format-native template.
    Template,
}

impl fmt::Debug for Flavor {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter.write_str(match self {
            Self::Document => "Flavor::Document",
            Self::Template => "Flavor::Template",
        })
    }
}

/// The OOXML conformance requested for a newly created package.
#[derive(Clone, Copy, PartialEq, Eq, Hash, Default)]
#[non_exhaustive]
pub enum Conformance {
    /// ECMA-376 Transitional, the default emitted by the DOCX leaf creator.
    #[default]
    Transitional,
    /// ECMA-376 Strict.
    Strict,
}

impl fmt::Debug for Conformance {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter.write_str(match self {
            Self::Transitional => "Conformance::Transitional",
            Self::Strict => "Conformance::Strict",
        })
    }
}

/// Options for creating one new package.
///
/// The options are validated against the selected leaf creator.  A leaf that
/// cannot honor a flavor, conformance, or metadata field returns a typed
/// refusal rather than silently changing or dropping the request.
///
/// `Debug` reports only whether metadata was supplied.  Metadata is authored
/// content and is therefore not copied into diagnostic output.
#[derive(Clone, Default)]
#[non_exhaustive]
pub struct CreateOptions {
    pub(crate) flavor: Flavor,
    pub(crate) conformance: Conformance,
    pub(crate) metadata: Option<Metadata>,
}

impl fmt::Debug for CreateOptions {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter
            .debug_struct("CreateOptions")
            .field("flavor", &self.flavor)
            .field("conformance", &self.conformance)
            .field("metadata", &self.metadata.as_ref().map(|_| "[REDACTED]"))
            .finish()
    }
}

impl CreateOptions {
    /// The default options: a Transitional ordinary document without initial
    /// metadata.
    #[must_use]
    pub const fn new() -> Self {
        Self {
            flavor: Flavor::Document,
            conformance: Conformance::Transitional,
            metadata: None,
        }
    }

    /// Select an ordinary document or a template.
    #[must_use]
    pub const fn flavor(mut self, flavor: Flavor) -> Self {
        self.flavor = flavor;
        self
    }

    /// Select the OOXML conformance class.
    #[must_use]
    pub const fn conformance(mut self, conformance: Conformance) -> Self {
        self.conformance = conformance;
        self
    }

    /// Set initial unified metadata.
    #[must_use]
    pub fn metadata(mut self, metadata: Metadata) -> Self {
        self.metadata = Some(metadata);
        self
    }

    /// The requested flavor.
    #[must_use]
    pub const fn flavor_value(&self) -> Flavor {
        self.flavor
    }

    /// The requested OOXML conformance.
    #[must_use]
    pub const fn conformance_value(&self) -> Conformance {
        self.conformance
    }

    /// The initial metadata, if supplied.
    #[must_use]
    pub fn metadata_value(&self) -> Option<&Metadata> {
        self.metadata.as_ref()
    }
}

/// Create a new package with the default options.
///
/// DOC, DOCX, PPT, XLS, XLSX, XLSB and PPTX have native-backed unified creators.
/// The DOC creator emits one ordinary blank document paragraph; template,
/// Strict and initial metadata options are refused explicitly. Other formats
/// return a typed capability or format refusal according to the features and
/// creators present in this build.
///
/// # Errors
///
/// Returns [`Code::UnsupportedCapability`] when the selected format is
/// present but its creator has not landed, [`Code::UnsupportedFormat`] when
/// its feature is unavailable, and the typed leaf error when creation fails.
pub fn create(format: FileFormat) -> Result<File> {
    create_with(format, &CreateOptions::new())
}

/// Create a new package using `options`.
///
/// The selected creator validates every option before publication. DOC
/// supports an ordinary Transitional document without initial metadata.
/// DOCX supports ordinary documents, templates, Transitional conformance, and
/// the core metadata fields its leaf stores. PPT, XLS, XLSX, XLSB and PPTX
/// support ordinary Transitional documents; template, Strict and initial
/// metadata requests are refused explicitly by those creators.
///
/// # Errors
///
/// See [`create`].
pub fn create_with(format: FileFormat, options: &CreateOptions) -> Result<File> {
    #[cfg(feature = "xls")]
    if format == FileFormat::Xls {
        return crate::format::xls::Xls::create(options).map(File::Xls);
    }

    #[cfg(feature = "doc")]
    if format == FileFormat::Doc {
        return crate::format::doc::Doc::create(options).map(File::Doc);
    }

    #[cfg(feature = "docx")]
    if format == FileFormat::Docx {
        return crate::format::docx::create::create(options).map(File::Docx);
    }

    #[cfg(feature = "xlsx")]
    if format == FileFormat::Xlsx {
        return crate::format::xlsx::create::create(options).map(File::Xlsx);
    }

    #[cfg(feature = "pptx")]
    if format == FileFormat::Pptx {
        return crate::format::pptx::create::create(options).map(File::Pptx);
    }

    #[cfg(feature = "ppt")]
    if format == FileFormat::Ppt {
        return crate::format::ppt::Ppt::create(options).map(File::Ppt);
    }

    #[cfg(feature = "xlsb")]
    if format == FileFormat::Xlsb {
        return crate::format::xlsb::create::create(options).map(File::Xlsb);
    }

    let _ = options;
    if format_feature_enabled(format) {
        Err(unsupported_capability(format))
    } else {
        Err(unsupported_format(format, feature_suggestion(format)))
    }
}

/// Whether the selected format's leaf dependency is available in this build.
///
/// This is deliberately kept beside creation's refusal mapping rather than
/// inferred from `File`: a format can have its leaf enabled while its unified
/// creator is still a pending capability.
fn format_feature_enabled(format: FileFormat) -> bool {
    match format {
        FileFormat::Doc => cfg!(feature = "doc"),
        FileFormat::Docx => cfg!(feature = "docx"),
        FileFormat::Ppt => cfg!(feature = "ppt"),
        FileFormat::Pptx => cfg!(feature = "pptx"),
        FileFormat::Xls => cfg!(feature = "xls"),
        FileFormat::Xlsx => cfg!(feature = "xlsx"),
        FileFormat::Xlsb => cfg!(feature = "xlsb"),
        FileFormat::Rtf => cfg!(feature = "rtf"),
        FileFormat::Odt => cfg!(feature = "odt"),
        FileFormat::Ods => cfg!(feature = "ods"),
        FileFormat::Odp => cfg!(feature = "odp"),
        FileFormat::Pages
        | FileFormat::Keynote
        | FileFormat::Numbers
        | FileFormat::Odg
        | FileFormat::Odc
        | FileFormat::Odf
        | FileFormat::Odi
        | FileFormat::Odm
        | FileFormat::Oth
        | FileFormat::Odb
        | _ => false,
    }
}

fn feature_suggestion(format: FileFormat) -> &'static str {
    match format {
        FileFormat::Doc => "enable the `doc` cargo feature",
        FileFormat::Docx => "enable the `docx` cargo feature",
        FileFormat::Ppt => "enable the `ppt` cargo feature",
        FileFormat::Pptx => "enable the `pptx` cargo feature",
        FileFormat::Xls => "enable the `xls` cargo feature",
        FileFormat::Xlsx => "enable the `xlsx` cargo feature",
        FileFormat::Xlsb => "enable the `xlsb` cargo feature",
        FileFormat::Rtf => "enable the `rtf` cargo feature",
        FileFormat::Odt => "enable the `odt` cargo feature",
        FileFormat::Ods => "enable the `ods` cargo feature",
        FileFormat::Odp => "enable the `odp` cargo feature",
        _ => "the format has no unified creator in this build",
    }
}

fn unsupported_capability(format: FileFormat) -> Error {
    Error::new(
        Code::UnsupportedCapability,
        "creation is not integrated for this format",
    )
    .with_format(format)
    .with_suggestion("use a format creator listed by the capability report")
}

fn unsupported_format(format: FileFormat, suggestion: &'static str) -> Error {
    Error::new(
        Code::UnsupportedFormat,
        "the requested format is not enabled",
    )
    .with_format(format)
    .with_suggestion(suggestion)
}

#[cfg(feature = "docx")]
pub(crate) fn unsupported_docx_option(message: &'static str, suggestion: &'static str) -> Error {
    Error::new(Code::UnsupportedCapability, message)
        .with_format(FileFormat::Docx)
        .with_suggestion(suggestion)
}

#[cfg(feature = "docx")]
pub(crate) fn invalid_created_metadata() -> Error {
    unsupported_docx_option(
        "the requested initial metadata is not representable by DOCX core properties",
        "supply only metadata fields supported by the DOCX core-properties part",
    )
}
