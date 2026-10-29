//! The typed leaf errors of the iWork owners, kept whole behind one
//! [`Source`](crate::error::Source) variant (record 0296).

use std::fmt;

use crate::error::Code;

/// A Pages, Keynote or Numbers leaf error, as the owner raised it.
///
/// [`Error::leaf`](crate::Error::leaf) returns it inside
/// [`Source::Iwork`](crate::error::Source::Iwork); the standard error source
/// chain reaches the leaf error itself.
#[derive(Debug)]
#[non_exhaustive]
pub enum LeafError {
    /// iWork package detection or preparation refused the input.
    Detection(litchi_iwa_detect::Error),
    /// The Pages package could not be read.
    #[cfg(feature = "pages")]
    Pages(litchi_pages::PackageError),
    /// A Pages section-text transaction failed.
    #[cfg(feature = "pages")]
    PagesText(litchi_pages::section::text::Error),
    /// Writing the Pages package to a sink failed.
    #[cfg(feature = "pages")]
    PagesWrite(litchi_pages::WriteError),
    /// A Pages paragraph transaction failed.
    #[cfg(feature = "pages")]
    PagesParagraphs(litchi_pages::section::paragraphs::Error),
    /// A Pages section-name transaction failed.
    #[cfg(feature = "pages")]
    PagesName(litchi_pages::section::name::Error),
    /// A Pages character-style transaction failed (record 0309).
    #[cfg(feature = "pages")]
    PagesTextStyle(litchi_pages::section::text_style::Error),
    /// A Pages body-chart point-size transaction failed (D7-A).
    #[cfg(feature = "pages")]
    PagesChartFont(litchi_pages::chart::font::Error),
    /// Reading a Pages body-chart inventory failed while resolving a
    /// canonical chart path.
    #[cfg(feature = "pages")]
    PagesChartArrangement(litchi_pages::chart::arrangement::Error),
    /// A Pages body-table value transaction failed (record 0309).
    #[cfg(feature = "pages")]
    PagesTableValues(litchi_pages::table::values::Error),
    /// A Pages body-table cell comment transaction (D7).
    #[cfg(feature = "pages")]
    PagesTableComments(litchi_pages::table::comments::Error),
    /// A Pages header or footer text transaction failed.
    #[cfg(feature = "pages")]
    PagesHeaderFooter(litchi_pages::header_footer::text::Error),
    /// A Pages footnote text transaction failed.
    #[cfg(feature = "pages")]
    PagesFootnote(litchi_pages::footnote::text::Error),
    /// A Pages body-footnote lifecycle transaction failed.
    #[cfg(feature = "pages")]
    PagesFootnotes(litchi_pages::footnote::body::transaction::Error),
    /// A Pages document-properties transaction failed.
    #[cfg(feature = "pages")]
    PagesMetadata(litchi_pages::metadata::Error),
    /// Saving a Pages package directory failed (record 0301).
    #[cfg(feature = "pages")]
    PagesSave(litchi_pages::SaveError),
    /// A Pages body-table deletion failed.
    #[cfg(feature = "pages")]
    PagesTableDeletion(litchi_pages::table::deletion::Error),
    /// A Pages body-table creation failed (record 0329).
    #[cfg(feature = "pages")]
    PagesTableCreation(litchi_pages::table::creation::Error),
    /// A Pages section add or remove failed (record 0312).
    #[cfg(feature = "pages")]
    PagesSectionLifecycle(litchi_pages::section::lifecycle::Error),
    /// The Keynote package could not be read.
    #[cfg(feature = "keynote")]
    Keynote(litchi_keynote::ReadError),
    /// A Keynote slide title or body transaction failed.
    #[cfg(feature = "keynote")]
    KeynoteText(litchi_keynote::slide::text::Error),
    /// A Keynote slide skip transaction failed.
    #[cfg(feature = "keynote")]
    KeynoteEdit(litchi_keynote::edit::Error),
    /// A Keynote slide order transaction failed.
    #[cfg(feature = "keynote")]
    KeynoteOrder(litchi_keynote::slide::order::Error),
    /// A Keynote slide deletion failed.
    #[cfg(feature = "keynote")]
    KeynoteDeletion(litchi_keynote::slide::delete::Error),
    /// Writing the Keynote package to a sink failed.
    #[cfg(feature = "keynote")]
    KeynoteWrite(litchi_keynote::WriteError),
    /// A Keynote speaker-notes transaction failed.
    #[cfg(feature = "keynote")]
    KeynoteNotes(litchi_keynote::slide::notes::Error),
    /// A Keynote slide-drawable transaction failed.
    #[cfg(feature = "keynote")]
    KeynoteDrawables(litchi_keynote::slide::drawable::Error),
    /// A Keynote slide-chart point-size transaction failed (D7-A).
    #[cfg(feature = "keynote")]
    KeynoteChartFont(litchi_keynote::chart::font::size::Error),
    /// A Keynote slide-table cell comment transaction (D7).
    #[cfg(feature = "keynote")]
    KeynoteTableComments(litchi_keynote::slide::table::comments::Error),
    /// A Keynote movie or audio removal failed (record 0309).
    #[cfg(feature = "keynote")]
    KeynoteMedia(litchi_keynote::slide::media::lifecycle::Error),
    /// A Keynote placeholder visibility change failed (record 0309).
    #[cfg(feature = "keynote")]
    KeynotePlaceholder(litchi_keynote::slide::placeholder::Error),
    /// A Keynote slide-creation transaction failed.
    #[cfg(feature = "keynote")]
    KeynoteCreation(litchi_keynote::slide::creation::Error),
    /// Saving a Keynote package directory failed (record 0306).
    #[cfg(feature = "keynote")]
    KeynoteSave(litchi_keynote::SaveError),
    /// The Numbers package could not be read.
    #[cfg(feature = "numbers")]
    Numbers(litchi_numbers::Error),
    /// The Numbers workbook text could not be rendered.
    #[cfg(feature = "numbers")]
    NumbersText(litchi_numbers::workbook::ReadError),
    /// A Numbers table-cell transaction failed.
    #[cfg(feature = "numbers")]
    NumbersCells(litchi_numbers::table::cells::Error),
    /// Writing the Numbers package to a sink failed.
    #[cfg(feature = "numbers")]
    NumbersWrite(litchi_numbers::WriteError),
    /// A Numbers transaction other than the cell transaction failed: sheet
    /// and table names, sheet order, table relocation, merges, comments,
    /// display formats, formula text, or a fresh workbook build.
    #[cfg(feature = "numbers")]
    NumbersEdit(super::numbers::EditError),
    /// A Numbers sheet-chart point-size transaction failed (D7-A).
    #[cfg(feature = "numbers")]
    NumbersChartFont(litchi_numbers::sheet::chart::font::Error),
    /// Reading a Numbers sheet's drawable inventory failed while resolving a
    /// canonical chart path.
    #[cfg(feature = "numbers")]
    NumbersDrawables(litchi_numbers::sheet::drawable::Error),
    /// Saving a Numbers package directory failed (record 0306).
    #[cfg(feature = "numbers")]
    NumbersSave(litchi_numbers::SaveError),
}

impl fmt::Display for LeafError {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        match self {
            Self::Detection(error) => error.fmt(formatter),
            #[cfg(feature = "pages")]
            Self::Pages(error) => error.fmt(formatter),
            #[cfg(feature = "pages")]
            Self::PagesText(error) => error.fmt(formatter),
            #[cfg(feature = "pages")]
            Self::PagesWrite(error) => error.fmt(formatter),
            #[cfg(feature = "pages")]
            Self::PagesParagraphs(error) => error.fmt(formatter),
            #[cfg(feature = "pages")]
            Self::PagesName(error) => error.fmt(formatter),
            #[cfg(feature = "pages")]
            Self::PagesTextStyle(error) => error.fmt(formatter),
            #[cfg(feature = "pages")]
            Self::PagesChartFont(error) => error.fmt(formatter),
            #[cfg(feature = "pages")]
            Self::PagesChartArrangement(error) => error.fmt(formatter),
            #[cfg(feature = "pages")]
            Self::PagesTableValues(error) => error.fmt(formatter),
            #[cfg(feature = "pages")]
            Self::PagesTableComments(error) => error.fmt(formatter),
            #[cfg(feature = "pages")]
            Self::PagesHeaderFooter(error) => error.fmt(formatter),
            #[cfg(feature = "pages")]
            Self::PagesFootnote(error) => error.fmt(formatter),
            #[cfg(feature = "pages")]
            Self::PagesFootnotes(error) => error.fmt(formatter),
            #[cfg(feature = "pages")]
            Self::PagesMetadata(error) => error.fmt(formatter),
            #[cfg(feature = "pages")]
            Self::PagesSave(error) => error.fmt(formatter),
            #[cfg(feature = "pages")]
            Self::PagesTableDeletion(error) => error.fmt(formatter),
            #[cfg(feature = "pages")]
            Self::PagesTableCreation(error) => error.fmt(formatter),
            #[cfg(feature = "pages")]
            Self::PagesSectionLifecycle(error) => error.fmt(formatter),
            #[cfg(feature = "keynote")]
            Self::Keynote(error) => error.fmt(formatter),
            #[cfg(feature = "keynote")]
            Self::KeynoteText(error) => error.fmt(formatter),
            #[cfg(feature = "keynote")]
            Self::KeynoteEdit(error) => error.fmt(formatter),
            #[cfg(feature = "keynote")]
            Self::KeynoteOrder(error) => error.fmt(formatter),
            #[cfg(feature = "keynote")]
            Self::KeynoteDeletion(error) => error.fmt(formatter),
            #[cfg(feature = "keynote")]
            Self::KeynoteWrite(error) => error.fmt(formatter),
            #[cfg(feature = "keynote")]
            Self::KeynoteNotes(error) => error.fmt(formatter),
            #[cfg(feature = "keynote")]
            Self::KeynoteDrawables(error) => error.fmt(formatter),
            #[cfg(feature = "keynote")]
            Self::KeynoteChartFont(error) => error.fmt(formatter),
            #[cfg(feature = "keynote")]
            Self::KeynoteTableComments(error) => error.fmt(formatter),
            #[cfg(feature = "keynote")]
            Self::KeynoteMedia(error) => error.fmt(formatter),
            #[cfg(feature = "keynote")]
            Self::KeynotePlaceholder(error) => error.fmt(formatter),
            #[cfg(feature = "keynote")]
            Self::KeynoteCreation(error) => error.fmt(formatter),
            #[cfg(feature = "keynote")]
            Self::KeynoteSave(error) => error.fmt(formatter),
            #[cfg(feature = "numbers")]
            Self::Numbers(error) => error.fmt(formatter),
            #[cfg(feature = "numbers")]
            Self::NumbersText(error) => error.fmt(formatter),
            #[cfg(feature = "numbers")]
            Self::NumbersCells(error) => error.fmt(formatter),
            #[cfg(feature = "numbers")]
            Self::NumbersWrite(error) => error.fmt(formatter),
            #[cfg(feature = "numbers")]
            Self::NumbersEdit(error) => error.fmt(formatter),
            #[cfg(feature = "numbers")]
            Self::NumbersChartFont(error) => error.fmt(formatter),
            #[cfg(feature = "numbers")]
            Self::NumbersDrawables(error) => error.fmt(formatter),
            #[cfg(feature = "numbers")]
            Self::NumbersSave(error) => error.fmt(formatter),
        }
    }
}

impl std::error::Error for LeafError {
    fn source(&self) -> Option<&(dyn std::error::Error + 'static)> {
        match self {
            Self::Detection(error) => Some(error),
            #[cfg(feature = "pages")]
            Self::Pages(error) => Some(error),
            #[cfg(feature = "pages")]
            Self::PagesText(error) => Some(error),
            #[cfg(feature = "pages")]
            Self::PagesWrite(error) => Some(error),
            #[cfg(feature = "pages")]
            Self::PagesParagraphs(error) => Some(error),
            #[cfg(feature = "pages")]
            Self::PagesName(error) => Some(error),
            #[cfg(feature = "pages")]
            Self::PagesTextStyle(error) => Some(error),
            #[cfg(feature = "pages")]
            Self::PagesChartFont(error) => Some(error),
            #[cfg(feature = "pages")]
            Self::PagesChartArrangement(error) => Some(error),
            #[cfg(feature = "pages")]
            Self::PagesTableValues(error) => Some(error),
            #[cfg(feature = "pages")]
            Self::PagesTableComments(error) => Some(error),
            #[cfg(feature = "pages")]
            Self::PagesHeaderFooter(error) => Some(error),
            #[cfg(feature = "pages")]
            Self::PagesFootnote(error) => Some(error),
            #[cfg(feature = "pages")]
            Self::PagesFootnotes(error) => Some(error),
            #[cfg(feature = "pages")]
            Self::PagesMetadata(error) => Some(error),
            #[cfg(feature = "pages")]
            Self::PagesSave(error) => Some(error),
            #[cfg(feature = "pages")]
            Self::PagesTableDeletion(error) => Some(error),
            #[cfg(feature = "pages")]
            Self::PagesTableCreation(error) => Some(error),
            #[cfg(feature = "pages")]
            Self::PagesSectionLifecycle(error) => Some(error),
            #[cfg(feature = "keynote")]
            Self::Keynote(error) => Some(error),
            #[cfg(feature = "keynote")]
            Self::KeynoteText(error) => Some(error),
            #[cfg(feature = "keynote")]
            Self::KeynoteEdit(error) => Some(error),
            #[cfg(feature = "keynote")]
            Self::KeynoteOrder(error) => Some(error),
            #[cfg(feature = "keynote")]
            Self::KeynoteDeletion(error) => Some(error),
            #[cfg(feature = "keynote")]
            Self::KeynoteWrite(error) => Some(error),
            #[cfg(feature = "keynote")]
            Self::KeynoteNotes(error) => Some(error),
            #[cfg(feature = "keynote")]
            Self::KeynoteDrawables(error) => Some(error),
            #[cfg(feature = "keynote")]
            Self::KeynoteChartFont(error) => Some(error),
            #[cfg(feature = "keynote")]
            Self::KeynoteTableComments(error) => Some(error),
            #[cfg(feature = "keynote")]
            Self::KeynoteMedia(error) => Some(error),
            #[cfg(feature = "keynote")]
            Self::KeynotePlaceholder(error) => Some(error),
            #[cfg(feature = "keynote")]
            Self::KeynoteCreation(error) => Some(error),
            #[cfg(feature = "keynote")]
            Self::KeynoteSave(error) => Some(error),
            #[cfg(feature = "numbers")]
            Self::Numbers(error) => Some(error),
            #[cfg(feature = "numbers")]
            Self::NumbersText(error) => Some(error),
            #[cfg(feature = "numbers")]
            Self::NumbersCells(error) => Some(error),
            #[cfg(feature = "numbers")]
            Self::NumbersWrite(error) => Some(error),
            #[cfg(feature = "numbers")]
            Self::NumbersEdit(error) => error.source_error(),
            #[cfg(feature = "numbers")]
            Self::NumbersChartFont(error) => Some(error),
            #[cfg(feature = "numbers")]
            Self::NumbersDrawables(error) => Some(error),
            #[cfg(feature = "numbers")]
            Self::NumbersSave(error) => Some(error),
        }
    }
}

impl LeafError {
    /// The unified code of the leaf error.
    ///
    /// Read failures of a package detection accepted are corrupt files,
    /// except typed resource ceilings (`limit_exceeded`) and I/O. Edit
    /// failures are the request's fault: a selector that matches nothing is
    /// `path_not_found`, a refused value `invalid_value`, a native shape the
    /// leaf does not rewrite yet `unsupported_capability`, a protected table
    /// `document_protected`, and a failed readback `internal_error`.
    #[must_use]
    pub(crate) fn code(&self) -> Code {
        match self {
            Self::Detection(error) => detection_code(error),
            #[cfg(feature = "pages")]
            Self::Pages(error) => pages_code(error),
            #[cfg(feature = "pages")]
            Self::PagesText(error) => pages_text_code(error),
            #[cfg(feature = "pages")]
            Self::PagesParagraphs(error) => pages_paragraphs_code(error),
            #[cfg(feature = "pages")]
            Self::PagesName(error) => pages_name_code(error),
            #[cfg(feature = "pages")]
            Self::PagesTextStyle(error) => pages_text_style_code(error),
            #[cfg(feature = "pages")]
            Self::PagesChartFont(error) => pages_chart_font_code(error),
            #[cfg(feature = "pages")]
            Self::PagesChartArrangement(error) => pages_chart_arrangement_code(error),
            #[cfg(feature = "pages")]
            Self::PagesTableValues(error) => pages_table_values_code(error),
            #[cfg(feature = "pages")]
            Self::PagesTableComments(error) => pages_table_comments_code(error),
            #[cfg(feature = "pages")]
            Self::PagesHeaderFooter(error) => pages_region_code(error),
            #[cfg(feature = "pages")]
            Self::PagesFootnote(error) => pages_footnote_code(error),
            #[cfg(feature = "pages")]
            Self::PagesFootnotes(error) => pages_footnotes_code(error),
            #[cfg(feature = "pages")]
            Self::PagesMetadata(error) => pages_metadata_code(error),
            #[cfg(feature = "pages")]
            Self::PagesSave(error) => match error.write_error() {
                Some(_) => Code::IoError,
                None if matches!(error, litchi_pages::SaveError::Bundle(_)) => Code::InvalidInput,
                None => Code::IoError,
            },
            #[cfg(feature = "pages")]
            Self::PagesTableDeletion(error) => pages_deletion_code(error),
            #[cfg(feature = "pages")]
            Self::PagesTableCreation(error) => pages_creation_code(error),
            #[cfg(feature = "pages")]
            Self::PagesSectionLifecycle(error) => pages_section_lifecycle_code(error),
            #[cfg(feature = "keynote")]
            Self::Keynote(error) => keynote_code(error),
            #[cfg(feature = "keynote")]
            Self::KeynoteText(error) => keynote_text_code(error),
            #[cfg(feature = "keynote")]
            Self::KeynoteEdit(error) => keynote_edit_code(error),
            #[cfg(feature = "keynote")]
            Self::KeynoteOrder(error) => keynote_order_code(error),
            #[cfg(feature = "keynote")]
            Self::KeynoteDeletion(error) => keynote_deletion_code(error),
            #[cfg(feature = "keynote")]
            Self::KeynoteNotes(error) => keynote_notes_code(error),
            #[cfg(feature = "keynote")]
            Self::KeynoteDrawables(error) => keynote_drawables_code(error),
            #[cfg(feature = "keynote")]
            Self::KeynoteChartFont(error) => keynote_chart_font_code(error),
            #[cfg(feature = "keynote")]
            Self::KeynoteTableComments(error) => keynote_table_comments_code(error),
            #[cfg(feature = "keynote")]
            Self::KeynoteMedia(error) => keynote_media_code(error),
            #[cfg(feature = "keynote")]
            Self::KeynotePlaceholder(error) => keynote_placeholder_code(error),
            #[cfg(feature = "keynote")]
            Self::KeynoteCreation(error) => keynote_creation_code(error),
            #[cfg(feature = "numbers")]
            Self::Numbers(error) => numbers_code(error),
            #[cfg(feature = "numbers")]
            Self::NumbersText(error) => match error {
                litchi_numbers::workbook::ReadError::Limit { .. } => Code::LimitExceeded,
                _ => Code::InternalError,
            },
            #[cfg(feature = "numbers")]
            Self::NumbersCells(error) => numbers_cells_code(error),
            #[cfg(feature = "pages")]
            Self::PagesWrite(_) => Code::IoError,
            #[cfg(feature = "keynote")]
            Self::KeynoteWrite(_) => Code::IoError,
            #[cfg(feature = "numbers")]
            Self::NumbersWrite(_) => Code::IoError,
            #[cfg(feature = "numbers")]
            Self::NumbersEdit(error) => error.code(),
            #[cfg(feature = "numbers")]
            Self::NumbersChartFont(error) => numbers_chart_font_code(error),
            #[cfg(feature = "numbers")]
            Self::NumbersDrawables(error) => numbers_drawables_code(error),
            #[cfg(feature = "keynote")]
            Self::KeynoteSave(error) => match error {
                litchi_keynote::SaveError::Bundle(_) => Code::InvalidInput,
                _ => Code::IoError,
            },
            #[cfg(feature = "numbers")]
            Self::NumbersSave(error) => match error {
                litchi_numbers::SaveError::Bundle(_) => Code::InvalidInput,
                _ => Code::IoError,
            },
        }
    }
}

/// The code of a refused leaf edit: its kind alone decides (record 0351
/// §2.4, record 0361), as for every other format.
fn refused(kind: litchi_core::edit::RefusalKind) -> Code {
    crate::error::refusal_code(kind)
}

/// The code of a physical iWork archive failure: a resource ceiling is
/// `limit_exceeded` (record 0306: source limits and package-directory
/// captures report through it), I/O is `io_error`, the rest is a malformed
/// package.
fn archive_code(error: &litchi_iwa_archive::Error) -> Code {
    use litchi_iwa_archive::Error as E;
    match error {
        E::Limit { .. } => Code::LimitExceeded,
        E::Io(_) => Code::IoError,
        E::SourceChanged { .. } | E::DirectoryChanged { .. } => Code::IoError,
        E::Allocation { .. } => Code::InternalError,
        _ => Code::CorruptFile,
    }
}

fn detection_code(error: &litchi_iwa_detect::Error) -> Code {
    match error {
        litchi_iwa_detect::Error::LimitExceeded { .. } => Code::LimitExceeded,
        litchi_iwa_detect::Error::Io(_) => Code::IoError,
        litchi_iwa_detect::Error::SourceChanged => Code::IoError,
        _ => Code::CorruptFile,
    }
}

#[cfg(feature = "pages")]
fn pages_code(error: &litchi_pages::PackageError) -> Code {
    use litchi_pages::PackageError as E;
    match error {
        E::Io(_) => Code::IoError,
        E::Detection(error) => detection_code(error),
        E::Archive(error) => archive_code(error),
        E::UnexpectedFormat { .. } | E::NotOfficeFile => Code::UnsupportedFormat,
        E::Limit { .. }
        | E::Semantic(
            litchi_pages::Error::TooManySections { .. }
            | litchi_pages::Error::TextTooLarge { .. }
            | litchi_pages::Error::TooManyBodyStorages { .. },
        )
        | E::SectionNamesTooLarge { .. }
        | E::PayloadLimit { .. }
        | E::ObjectLimit { .. } => Code::LimitExceeded,
        E::Allocation { .. } => Code::InternalError,
        _ => Code::CorruptFile,
    }
}

#[cfg(feature = "pages")]
fn pages_text_code(error: &litchi_pages::section::text::Error) -> Code {
    use litchi_pages::section::text::Error as E;
    match error {
        E::Refused(refusal) => refused(refusal.kind()),
        E::Limit { .. } => Code::LimitExceeded,
        E::InvalidSource | E::PatchConflict => Code::InvalidInput,
        _ => Code::InternalError,
    }
}

#[cfg(feature = "pages")]
fn pages_paragraphs_code(error: &litchi_pages::section::paragraphs::Error) -> Code {
    use litchi_pages::section::paragraphs::Error as E;
    match error {
        // Removing the final slide, section or paragraph is a caller
        // argument error, as for PPTX `SlideRemovalRefusal::FinalSlide`.
        E::Refused(litchi_pages::section::paragraphs::Refusal::LastParagraph) => {
            Code::InvalidArgument
        },
        E::Refused(refusal) => refused(refusal.kind()),
        E::Text(error) => pages_text_code(error),
        E::Limit { .. } => Code::LimitExceeded,
        E::PatchConflict => Code::InvalidInput,
        _ => Code::InternalError,
    }
}

#[cfg(feature = "pages")]
fn pages_name_code(error: &litchi_pages::section::name::Error) -> Code {
    use litchi_pages::section::name::Error as E;
    match error {
        E::Refused(refusal) => refused(refusal.kind()),
        E::Limit { .. } => Code::LimitExceeded,
        E::InvalidSource | E::PatchConflict => Code::InvalidInput,
        _ => Code::InternalError,
    }
}

#[cfg(feature = "pages")]
fn pages_text_style_code(error: &litchi_pages::section::text_style::Error) -> Code {
    use litchi_pages::section::text_style::Error as E;
    match error {
        E::Refused(refusal) => refused(refusal.kind()),
        E::Limit { .. } => Code::LimitExceeded,
        E::InvalidSource | E::PatchConflict => Code::InvalidInput,
        _ => Code::InternalError,
    }
}

#[cfg(feature = "pages")]
fn pages_chart_font_code(error: &litchi_pages::chart::font::Error) -> Code {
    use litchi_pages::chart::font::Error as E;
    match error {
        E::Refused(refusal) => refused(refusal.kind()),
        E::Limit { .. } => Code::LimitExceeded,
        E::InvalidSource | E::PatchConflict => Code::InvalidInput,
        E::Allocation { .. } | E::Verification => Code::InternalError,
        _ => Code::InternalError,
    }
}

#[cfg(feature = "pages")]
fn pages_chart_arrangement_code(error: &litchi_pages::chart::arrangement::Error) -> Code {
    use litchi_pages::chart::arrangement::Error as E;
    match error {
        E::Refused(refusal) => refused(refusal.kind()),
        E::Limit { .. } => Code::LimitExceeded,
        E::InvalidSource | E::PatchConflict => Code::InvalidInput,
        E::Allocation { .. } | E::Verification => Code::InternalError,
        _ => Code::InternalError,
    }
}

#[cfg(feature = "keynote")]
fn keynote_placeholder_code(error: &litchi_keynote::slide::placeholder::Error) -> Code {
    use litchi_keynote::slide::placeholder::Error as E;
    match error {
        E::Refused(refusal) => refused(refusal.kind()),
        E::Limit { .. } => Code::LimitExceeded,
        E::InvalidSource | E::PatchConflict => Code::InvalidInput,
        _ => Code::InternalError,
    }
}

#[cfg(feature = "keynote")]
fn keynote_media_code(error: &litchi_keynote::slide::media::lifecycle::Error) -> Code {
    use litchi_keynote::slide::media::lifecycle::Error as E;
    match error {
        E::Refused(refusal) => refused(refusal.kind()),
        E::Limit { .. } => Code::LimitExceeded,
        E::InvalidSource | E::PatchConflict => Code::InvalidInput,
        _ => Code::InternalError,
    }
}

#[cfg(feature = "pages")]
fn pages_table_values_code(error: &litchi_pages::table::values::Error) -> Code {
    use litchi_pages::table::values::Error as E;
    match error {
        E::Refused(refusal) => refused(refusal.kind()),
        E::Limit { .. } => Code::LimitExceeded,
        E::InvalidSource | E::PatchConflict => Code::InvalidInput,
        _ => Code::InternalError,
    }
}

#[cfg(feature = "pages")]
fn pages_table_comments_code(error: &litchi_pages::table::comments::Error) -> Code {
    use litchi_pages::table::comments::Error as E;
    match error {
        E::TableNotFound => Code::NotFound,
        E::AmbiguousTableName | E::AmbiguousSelector => Code::InvalidValue,
        E::UnsupportedSource | E::UnsupportedDependency => Code::UnsupportedCapability,
        E::Limit { .. } => Code::LimitExceeded,
        E::Allocation { .. } => Code::InternalError,
        E::InvalidSource | E::PatchConflict => Code::InvalidInput,
        E::Verification | E::Read | E::CommentNotFound => Code::InternalError,
        _ => Code::InternalError,
    }
}

#[cfg(feature = "pages")]
fn pages_region_code(error: &litchi_pages::header_footer::text::Error) -> Code {
    use litchi_pages::header_footer::text::Error as E;
    match error {
        E::Refused(refusal) => refused(refusal.kind()),
        E::Limit { .. } => Code::LimitExceeded,
        E::InvalidSource | E::PatchConflict => Code::InvalidInput,
        _ => Code::InternalError,
    }
}

#[cfg(feature = "pages")]
fn pages_footnote_code(error: &litchi_pages::footnote::text::Error) -> Code {
    use litchi_pages::footnote::text::Error as E;
    match error {
        E::Refused(refusal) => refused(refusal.kind()),
        E::Limit { .. } => Code::LimitExceeded,
        E::InvalidSource | E::PatchConflict => Code::InvalidInput,
        _ => Code::InternalError,
    }
}

#[cfg(feature = "pages")]
fn pages_footnotes_code(error: &litchi_pages::footnote::body::transaction::Error) -> Code {
    use litchi_pages::footnote::body::transaction::Error as E;
    match error {
        E::Refused(refusal) => refused(refusal.kind()),
        E::Limit { .. } => Code::LimitExceeded,
        E::InvalidSource | E::PatchConflict => Code::InvalidInput,
        _ => Code::InternalError,
    }
}

#[cfg(feature = "pages")]
fn pages_metadata_code(error: &litchi_pages::metadata::Error) -> Code {
    use litchi_pages::metadata::Error as E;
    match error {
        E::Refused(refusal) => refused(refusal.kind()),
        E::Limit { .. } => Code::LimitExceeded,
        E::InvalidSource | E::PatchConflict => Code::InvalidInput,
        _ => Code::InternalError,
    }
}

#[cfg(feature = "pages")]
fn pages_deletion_code(error: &litchi_pages::table::deletion::Error) -> Code {
    use litchi_pages::table::deletion::Error as E;
    match error {
        E::Refused(refusal) => refused(refusal.kind()),
        E::Limit { .. } => Code::LimitExceeded,
        E::PatchConflict => Code::InvalidInput,
        _ => Code::InternalError,
    }
}

#[cfg(feature = "pages")]
fn pages_creation_code(error: &litchi_pages::table::creation::Error) -> Code {
    use litchi_pages::table::creation::Error as E;
    match error {
        E::Refused(refusal) => refused(refusal.kind()),
        E::Values(error) => pages_table_values_code(error),
        E::Limit { .. } => Code::LimitExceeded,
        E::InvalidSource | E::PatchConflict => Code::InvalidInput,
        _ => Code::InternalError,
    }
}

#[cfg(feature = "pages")]
fn pages_section_lifecycle_code(error: &litchi_pages::section::lifecycle::Error) -> Code {
    use litchi_pages::section::lifecycle::Error as E;
    match error {
        // Removing the final slide, section or paragraph is a caller
        // argument error, as for PPTX `SlideRemovalRefusal::FinalSlide`.
        E::Refused(litchi_pages::section::lifecycle::Refusal::InitialSection) => {
            Code::InvalidArgument
        },
        E::Refused(refusal) => refused(refusal.kind()),
        E::Limit { .. } => Code::LimitExceeded,
        E::PatchConflict => Code::InvalidInput,
        _ => Code::InternalError,
    }
}

#[cfg(feature = "keynote")]
fn keynote_code(error: &litchi_keynote::ReadError) -> Code {
    use litchi_keynote::ReadError as E;
    match error {
        E::Io(_) => Code::IoError,
        E::Detection(error) => detection_code(error),
        E::Archive(error) => archive_code(error),
        E::UnexpectedFormat { .. } => Code::UnsupportedFormat,
        E::SemanticLimit { .. } | E::PayloadLimit { .. } => Code::LimitExceeded,
        E::Allocation { .. } => Code::InternalError,
        _ => Code::CorruptFile,
    }
}

#[cfg(feature = "keynote")]
fn keynote_text_code(error: &litchi_keynote::slide::text::Error) -> Code {
    use litchi_keynote::slide::text::Error as E;
    match error {
        E::Refused(refusal) => refused(refusal.kind()),
        E::Limit { .. } => Code::LimitExceeded,
        E::InvalidSource | E::PatchConflict => Code::InvalidInput,
        _ => Code::InternalError,
    }
}

#[cfg(feature = "keynote")]
fn keynote_edit_code(error: &litchi_keynote::edit::Error) -> Code {
    use litchi_keynote::edit::Error as E;
    match error {
        E::Refused(refusal) => refused(refusal.kind()),
        _ => Code::InternalError,
    }
}

#[cfg(feature = "keynote")]
fn keynote_order_code(error: &litchi_keynote::slide::order::Error) -> Code {
    use litchi_keynote::slide::order::Error as E;
    match error {
        E::Refused(refusal) => refused(refusal.kind()),
        E::Limit { .. } => Code::LimitExceeded,
        _ => Code::InternalError,
    }
}

#[cfg(feature = "keynote")]
fn keynote_deletion_code(error: &litchi_keynote::slide::delete::Error) -> Code {
    use litchi_keynote::slide::delete::Error as E;
    match error {
        // Removing the final slide, section or paragraph is a caller
        // argument error, as for PPTX `SlideRemovalRefusal::FinalSlide`.
        E::Refused(litchi_keynote::slide::delete::Refusal::CannotDeleteFinalSlide) => {
            Code::InvalidArgument
        },
        E::Refused(refusal) => refused(refusal.kind()),
        E::Limit { .. } => Code::LimitExceeded,
        _ => Code::InternalError,
    }
}

#[cfg(feature = "keynote")]
fn keynote_notes_code(error: &litchi_keynote::slide::notes::Error) -> Code {
    use litchi_keynote::slide::notes::Error as E;
    match error {
        E::Refused(refusal) => refused(refusal.kind()),
        E::Limit { .. } => Code::LimitExceeded,
        E::InvalidSource | E::PatchConflict => Code::InvalidInput,
        _ => Code::InternalError,
    }
}

#[cfg(feature = "keynote")]
fn keynote_drawables_code(error: &litchi_keynote::slide::drawable::Error) -> Code {
    use litchi_keynote::slide::drawable::Error as E;
    match error {
        E::Refused(refusal) => refused(refusal.kind()),
        E::Limit { .. } => Code::LimitExceeded,
        E::InvalidSource | E::PatchConflict => Code::InvalidInput,
        _ => Code::InternalError,
    }
}

#[cfg(feature = "keynote")]
fn keynote_chart_font_code(error: &litchi_keynote::chart::font::size::Error) -> Code {
    use litchi_keynote::chart::font::size::Error as E;
    match error {
        E::Refused(refusal) => refused(refusal.kind()),
        E::Limit { .. } => Code::LimitExceeded,
        E::InvalidSource | E::PatchConflict => Code::InvalidInput,
        E::Allocation { .. } | E::Verification => Code::InternalError,
        _ => Code::InternalError,
    }
}

#[cfg(feature = "keynote")]
fn keynote_table_comments_code(error: &litchi_keynote::slide::table::comments::Error) -> Code {
    use litchi_keynote::slide::table::comments::Error as E;
    match error {
        E::TableNotFound => Code::NotFound,
        E::AmbiguousTableName | E::AmbiguousSelector => Code::InvalidValue,
        E::UnsupportedSource | E::UnsupportedDependency => Code::UnsupportedCapability,
        E::TableLocked => Code::DocumentProtected,
        E::Limit { .. } => Code::LimitExceeded,
        E::Allocation { .. } => Code::InternalError,
        E::InvalidSource | E::PatchConflict => Code::InvalidInput,
        E::Verification | E::Read | E::CommentNotFound => Code::InternalError,
        _ => Code::InternalError,
    }
}

#[cfg(feature = "keynote")]
fn keynote_creation_code(error: &litchi_keynote::slide::creation::Error) -> Code {
    use litchi_keynote::slide::creation::Error as E;
    match error {
        E::Refused(refusal) => refused(refusal.kind()),
        E::Limit { .. } => Code::LimitExceeded,
        E::InvalidSource | E::PatchConflict => Code::InvalidInput,
        _ => Code::InternalError,
    }
}

#[cfg(feature = "numbers")]
fn numbers_code(error: &litchi_numbers::Error) -> Code {
    use litchi_numbers::Error as E;
    match error {
        E::Io(_) => Code::IoError,
        E::Detection(error) => detection_code(error),
        E::Archive(error) => archive_code(error),
        E::UnexpectedFormat { .. } => Code::UnsupportedFormat,
        E::SemanticLimit { .. } | E::InputTooLarge { .. } => Code::LimitExceeded,
        _ => Code::CorruptFile,
    }
}

#[cfg(feature = "numbers")]
fn numbers_cells_code(error: &litchi_numbers::table::cells::Error) -> Code {
    use litchi_numbers::table::cells::Error as E;
    match error {
        E::Refused(refusal) => refused(refusal.kind()),
        E::Limit { .. } => Code::LimitExceeded,
        E::InvalidSource { .. } | E::PatchConflict => Code::InvalidInput,
        _ => Code::InternalError,
    }
}

#[cfg(feature = "numbers")]
fn numbers_chart_font_code(error: &litchi_numbers::sheet::chart::font::Error) -> Code {
    use litchi_numbers::sheet::chart::font::Error as E;
    match error {
        E::Refused(refusal) => refused(refusal.kind()),
        E::Limit { .. } => Code::LimitExceeded,
        E::InvalidSource | E::PatchConflict => Code::InvalidInput,
        E::Allocation { .. } | E::Verification => Code::InternalError,
        _ => Code::InternalError,
    }
}

#[cfg(feature = "numbers")]
fn numbers_drawables_code(error: &litchi_numbers::sheet::drawable::Error) -> Code {
    use litchi_numbers::sheet::drawable::Error as E;
    match error {
        E::Refused(refusal) => refused(refusal.kind()),
        E::Limit { .. } => Code::LimitExceeded,
        E::InvalidSource | E::PatchConflict => Code::InvalidInput,
        _ => Code::InternalError,
    }
}

#[cfg(test)]
mod tests {
    use super::LeafError;
    use crate::error::Code;

    #[test]
    fn table_comment_selection_errors_keep_missing_and_ambiguous_codes() {
        #[cfg(feature = "pages")]
        {
            use litchi_pages::table::comments::Error;
            assert_eq!(
                LeafError::PagesTableComments(Error::TableNotFound).code(),
                Code::NotFound
            );
            assert_eq!(
                LeafError::PagesTableComments(Error::AmbiguousTableName).code(),
                Code::InvalidValue
            );
            assert_eq!(
                LeafError::PagesTableComments(Error::AmbiguousSelector).code(),
                Code::InvalidValue
            );
        }
        #[cfg(feature = "keynote")]
        {
            use litchi_keynote::slide::table::comments::Error;
            assert_eq!(
                LeafError::KeynoteTableComments(Error::TableNotFound).code(),
                Code::NotFound
            );
            assert_eq!(
                LeafError::KeynoteTableComments(Error::AmbiguousTableName).code(),
                Code::InvalidValue
            );
            assert_eq!(
                LeafError::KeynoteTableComments(Error::AmbiguousSelector).code(),
                Code::InvalidValue
            );
        }
    }

    #[cfg(feature = "pages")]
    #[test]
    fn pages_registry_resource_failures_keep_the_limit_code() {
        use litchi_pages::section::text_style::{Error, LimitKind};
        let error = LeafError::PagesTextStyle(Error::Limit {
            kind: LimitKind::RegistryFields,
            observed: 2,
            maximum: 1,
        });
        assert_eq!(error.code(), Code::LimitExceeded);
        assert_eq!(
            LeafError::PagesTextStyle(Error::InvalidSource).code(),
            Code::InvalidInput
        );
    }

    /// Record 0361: an iWork refusal reports its kind's code; a final-element
    /// removal keeps `invalid_argument` (the PPTX precedent); a patch for
    /// another base is `invalid_input`.
    #[test]
    fn iwork_refusals_classify_by_kind() {
        #[cfg(feature = "pages")]
        {
            use litchi_pages::table::deletion::{Error, Refusal};
            let missing = LeafError::PagesTableDeletion(Error::Refused(Refusal::TableNotFound));
            assert_eq!(missing.code(), Code::NotFound);
            let ambiguous =
                LeafError::PagesTableDeletion(Error::Refused(Refusal::AmbiguousSelector));
            assert_eq!(ambiguous.code(), Code::InvalidValue);
            let conflict = LeafError::PagesTableDeletion(Error::PatchConflict);
            assert_eq!(conflict.code(), Code::InvalidInput);
            let first =
                LeafError::PagesSectionLifecycle(litchi_pages::section::lifecycle::Error::Refused(
                    litchi_pages::section::lifecycle::Refusal::InitialSection,
                ));
            assert_eq!(first.code(), Code::InvalidArgument);
        }
        #[cfg(feature = "keynote")]
        {
            use litchi_keynote::slide::delete::{Error, Refusal};
            let last = LeafError::KeynoteDeletion(Error::Refused(Refusal::CannotDeleteFinalSlide));
            assert_eq!(last.code(), Code::InvalidArgument);
            let unsupported =
                LeafError::KeynoteDeletion(Error::Refused(Refusal::UnsupportedSource));
            assert_eq!(unsupported.code(), Code::UnsupportedCapability);
        }
        #[cfg(feature = "numbers")]
        {
            use litchi_numbers::table::cells::{Error, Refusal};
            let locked = LeafError::NumbersCells(Error::Refused(Refusal::TableLocked {
                path: litchi_numbers::table::cells::Path::Package,
            }));
            assert_eq!(locked.code(), Code::DocumentProtected);
        }
    }

    #[test]
    fn iwork_source_freshness_errors_are_io_errors() {
        use super::{archive_code, detection_code};
        use litchi_core::SourceVersion;
        use litchi_iwa_archive::Error as ArchiveError;

        assert_eq!(
            archive_code(&ArchiveError::SourceChanged {
                expected: SourceVersion::new(1, 1),
                observed: SourceVersion::new(1, 2),
            }),
            Code::IoError
        );
        assert_eq!(
            archive_code(&ArchiveError::DirectoryChanged {
                context: "test directory capture",
            }),
            Code::IoError
        );
        assert_eq!(
            detection_code(&litchi_iwa_detect::Error::SourceChanged),
            Code::IoError
        );
    }

    #[test]
    fn iwork_package_wrong_kinds_are_unsupported_formats() {
        use litchi_core::FileFormat;

        #[cfg(feature = "pages")]
        {
            assert_eq!(
                super::pages_code(&litchi_pages::PackageError::UnexpectedFormat {
                    detected: FileFormat::Keynote,
                }),
                Code::UnsupportedFormat
            );
            assert_eq!(
                super::pages_code(&litchi_pages::PackageError::NotOfficeFile),
                Code::UnsupportedFormat
            );
        }
        #[cfg(feature = "keynote")]
        assert_eq!(
            super::keynote_code(&litchi_keynote::ReadError::UnexpectedFormat {
                detected: FileFormat::Pages,
            }),
            Code::UnsupportedFormat
        );
        #[cfg(feature = "numbers")]
        assert_eq!(
            super::numbers_code(&litchi_numbers::Error::UnexpectedFormat {
                detected: FileFormat::Pages,
            }),
            Code::UnsupportedFormat
        );
    }
}
