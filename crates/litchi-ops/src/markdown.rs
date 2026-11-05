//! Markdown output: `ToMarkdown` for [`File`], its family views and the
//! [`compat`](crate::compat) element wrappers (M0a step S7). The crate
//! documentation's "Markdown" section describes what converts and how each
//! refusal is typed.
//!
//! The adapter moved here from the `litchi` facade, where it served the
//! facade's format-family types (deleted in M0a step S9). `ToMarkdown` and
//! `File` are both foreign to the facade, so the implementation must live
//! next to `File`. The rendering is unchanged: the facade's Markdown corpus
//! goldens, recorded from the deleted types, pin it. What changed is the
//! error, which is the typed [`Error`] of this crate.

#[cfg(any(feature = "doc", feature = "docx", feature = "rtf", feature = "odt"))]
mod document;
#[cfg(feature = "docx")]
mod docx;
#[cfg(any(feature = "pages", feature = "keynote"))]
mod iwork;
#[cfg(any(feature = "ppt", feature = "pptx", feature = "odp"))]
mod presentation;
#[cfg(any(feature = "doc", feature = "docx", feature = "rtf", feature = "odt"))]
mod writer;

#[cfg(test)]
mod tests;

use litchi_core::FileFormat;
use litchi_markdown::{MarkdownOptions, ToMarkdown};

#[cfg(any(feature = "doc", feature = "docx", feature = "rtf", feature = "odt"))]
use crate::DocumentView;
#[cfg(any(feature = "ppt", feature = "pptx", feature = "odp"))]
use crate::PresentationView;
use crate::{Code, Error, File, Result};

impl ToMarkdown for File {
    type Error = Error;

    /// Convert a document or presentation to Markdown (a Pages document or a
    /// Keynote presentation included, record 0306); a [`File::Signed`]
    /// converts the file it wraps.
    ///
    /// # Errors
    ///
    /// Returns [`Code::UnsupportedCapability`] for a workbook,
    /// [`Code::DocumentProtected`] for a [`File::Locked`], and otherwise the
    /// errors the [crate documentation](crate#markdown) lists.
    #[cfg_attr(
        not(any(
            feature = "doc",
            feature = "docx",
            feature = "rtf",
            feature = "odt",
            feature = "ppt",
            feature = "pptx",
            feature = "odp",
            feature = "pages",
            feature = "keynote"
        )),
        allow(
            clippy::only_used_in_recursion,
            reason = "without a document or presentation format only a signed file passes the options on"
        )
    )]
    fn to_markdown_with_options(&self, options: &MarkdownOptions) -> Result<String> {
        match self {
            #[cfg(feature = "docx")]
            Self::Docx(handle) => DocumentView::docx(handle).to_markdown_with_options(options),
            #[cfg(feature = "doc")]
            Self::Doc(handle) => DocumentView::doc(handle).to_markdown_with_options(options),
            #[cfg(feature = "rtf")]
            Self::Rtf(handle) => DocumentView::rtf(handle).to_markdown_with_options(options),
            #[cfg(feature = "odt")]
            Self::Odt(handle) | Self::Oth(handle) | Self::Odm(handle) => {
                DocumentView::odt(handle).to_markdown_with_options(options)
            },
            #[cfg(feature = "pptx")]
            Self::Pptx(handle) => PresentationView::pptx(handle).to_markdown_with_options(options),
            #[cfg(feature = "ppt")]
            Self::Ppt(handle) => PresentationView::ppt(handle).to_markdown_with_options(options),
            #[cfg(feature = "odp")]
            Self::Odp(handle) => PresentationView::odp(handle).to_markdown_with_options(options),
            #[cfg(feature = "pages")]
            Self::Pages(handle) => iwork::convert(self, handle, options)
                .map_err(|error| in_format(error, FileFormat::Pages)),
            #[cfg(feature = "keynote")]
            Self::Keynote(handle) => iwork::convert(self, handle, options)
                .map_err(|error| in_format(error, FileFormat::Keynote)),
            Self::Locked(locked) => Err(locked.read_error()),
            Self::Signed(signed) => signed.file().to_markdown_with_options(options),
            #[allow(
                unreachable_patterns,
                reason = "only workbooks remain, and a build without a workbook format has none"
            )]
            other => Err(workbook_unsupported(other.format())),
        }
    }
}

/// The refusal of Markdown output for a workbook.
fn workbook_unsupported(format: Option<FileFormat>) -> Error {
    let error = Error::new(
        Code::UnsupportedCapability,
        "Markdown output is available for documents and presentations, not for workbooks",
    );
    match format {
        Some(format) => error.with_format(format),
        None => error,
    }
}

#[cfg(any(feature = "doc", feature = "docx", feature = "rtf", feature = "odt"))]
impl ToMarkdown for DocumentView<'_> {
    type Error = Error;

    /// Convert the document: optional YAML front matter, then every
    /// paragraph and table in document order, then DOCX note definitions.
    ///
    /// # Errors
    ///
    /// As the [crate documentation](crate#markdown) lists.
    fn to_markdown_with_options(&self, options: &MarkdownOptions) -> Result<String> {
        document::convert(self, options)
    }
}

#[cfg(any(feature = "ppt", feature = "pptx", feature = "odp"))]
impl ToMarkdown for PresentationView<'_> {
    type Error = Error;

    /// Convert the presentation: optional YAML front matter, then one
    /// `# Slide N` section per slide.
    ///
    /// # Errors
    ///
    /// As the [crate documentation](crate#markdown) lists.
    fn to_markdown_with_options(&self, options: &MarkdownOptions) -> Result<String> {
        presentation::convert(self, options)
    }
}

#[cfg(any(feature = "doc", feature = "docx", feature = "rtf", feature = "odt"))]
impl ToMarkdown for crate::compat::word::Paragraph {
    type Error = Error;

    /// Convert one paragraph without its package context. A DOCX paragraph
    /// whose hyperlinks, images or notes need the package is refused;
    /// convert the whole document for those.
    fn to_markdown_with_options(&self, options: &MarkdownOptions) -> Result<String> {
        let mut writer = writer::MarkdownWriter::new(*options);
        writer.write_paragraph(self)?;
        Ok(writer.finish().trim_end().to_string())
    }
}

#[cfg(any(feature = "doc", feature = "docx", feature = "rtf", feature = "odt"))]
impl ToMarkdown for crate::compat::word::Run {
    type Error = Error;

    /// Convert one run with its direct formatting.
    fn to_markdown_with_options(&self, options: &MarkdownOptions) -> Result<String> {
        let mut writer = writer::MarkdownWriter::new(*options);
        writer.write_run(self)?;
        Ok(writer.finish())
    }
}

#[cfg(any(feature = "doc", feature = "docx", feature = "rtf", feature = "odt"))]
impl ToMarkdown for crate::compat::word::Table {
    type Error = Error;

    /// Convert one table: a Markdown table, or HTML when cells are merged or
    /// the options ask for it.
    fn to_markdown_with_options(&self, options: &MarkdownOptions) -> Result<String> {
        let mut writer = writer::MarkdownWriter::new(*options);
        writer.write_table(self)?;
        Ok(writer.finish().trim_end().to_string())
    }
}

#[cfg(any(feature = "ppt", feature = "pptx", feature = "odp"))]
impl ToMarkdown for crate::compat::slide::Slide {
    type Error = Error;

    /// The slide's text, escaped; presentations format whole slides.
    fn to_markdown_with_options(&self, _options: &MarkdownOptions) -> Result<String> {
        Ok(litchi_markdown::escape::text(&self.text()?).into_owned())
    }
}

/// Refuse content Markdown cannot represent faithfully.
#[cfg(any(feature = "doc", feature = "docx", feature = "rtf", feature = "odt"))]
fn unrepresentable(message: impl Into<std::borrow::Cow<'static, str>>) -> Error {
    Error::new(Code::UnsupportedElement, message)
}

/// A failed internal invariant of the adapter.
#[cfg(any(feature = "doc", feature = "docx", feature = "rtf", feature = "odt"))]
fn invariant(message: &'static str) -> Error {
    Error::new(Code::InternalError, message)
}

/// A failed allocation of a bounded collection the adapter builds.
#[cfg(any(feature = "doc", feature = "docx", feature = "rtf", feature = "odt"))]
fn allocation(resource: &'static str, source: std::collections::TryReserveError) -> Error {
    Error::from(litchi_core::Error::Allocation { resource, source })
}

/// Output that needs a cargo feature this build lacks.
#[cfg(any(
    all(feature = "docx", not(feature = "formula")),
    all(
        not(feature = "yaml"),
        any(
            feature = "doc",
            feature = "docx",
            feature = "rtf",
            feature = "odt",
            feature = "ppt",
            feature = "pptx",
            feature = "odp",
            feature = "pages",
            feature = "keynote"
        )
    )
))]
fn feature_required(what: &str, feature: &'static str) -> Error {
    Error::new(
        Code::UnsupportedCapability,
        format!("{what} needs the `{feature}` cargo feature"),
    )
    .with_suggestion(format!("enable the `{feature}` cargo feature"))
}

/// The refusal of front matter in a build without `yaml`.
#[cfg(all(
    not(feature = "yaml"),
    any(
        feature = "doc",
        feature = "docx",
        feature = "rtf",
        feature = "odt",
        feature = "ppt",
        feature = "pptx",
        feature = "odp",
        feature = "pages",
        feature = "keynote"
    )
))]
fn yaml_required() -> Error {
    feature_required("Markdown front matter", "yaml")
}

/// The YAML front matter of `metadata`, or nothing when it has no data.
#[cfg(all(
    feature = "yaml",
    any(
        feature = "doc",
        feature = "docx",
        feature = "rtf",
        feature = "odt",
        feature = "ppt",
        feature = "pptx",
        feature = "odp",
        feature = "pages",
        feature = "keynote"
    )
))]
fn yaml_front_matter(metadata: &litchi_core::Metadata) -> Result<String> {
    if !metadata.has_data() {
        return Ok(String::new());
    }
    let yaml = serde_saphyr::to_string(metadata).map_err(|error| {
        Error::new(
            Code::InternalError,
            format!("the document properties cannot be serialized as YAML: {error}"),
        )
    })?;
    Ok(format!("---\n{yaml}---\n\n"))
}

/// Type an ODT read failure: the ODT leaf reports `litchi_core::Error`.
#[cfg(feature = "odt")]
fn odt_error(error: litchi_core::Error) -> Error {
    Error::from(error).with_format(FileFormat::Odt)
}

/// Name `format` on an error that does not name one yet.
#[cfg(any(
    feature = "doc",
    feature = "docx",
    feature = "rtf",
    feature = "odt",
    feature = "ppt",
    feature = "pptx",
    feature = "odp",
    feature = "pages",
    feature = "keynote"
))]
fn in_format(error: Error, format: FileFormat) -> Error {
    if error.format().is_some() {
        error
    } else {
        error.with_format(format)
    }
}
