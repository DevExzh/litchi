//! The word-processing family view, ported from the facade's
//! `document::Document` reads.

use litchi_core::FileFormat;

#[cfg(any(feature = "docx", feature = "odt"))]
use crate::Error;
use crate::Result;
use crate::compat::word::{DocumentElement, Paragraph, Table};
#[cfg(feature = "doc")]
use crate::format::doc::Doc;
#[cfg(feature = "docx")]
use crate::format::docx::{Docx, DocxOwner};
#[cfg(feature = "odt")]
use crate::format::odt::{Odt, OdtOwner};
#[cfg(feature = "rtf")]
use crate::format::rtf::Rtf;

/// The word-processing view of a [`File`](crate::File): paragraphs, tables
/// and the ordered elements of DOC, DOCX, RTF and ODT documents.
///
/// [`File::document`](crate::File::document) returns it. Reads keep the
/// facade's per-format behavior: DOCX and ODT select single paragraphs
/// through their owners, so a source-backed document reads only what it is
/// asked for and reports a source changed after open as
/// [`Code::IoError`](crate::Code::IoError); DOC and RTF project all
/// paragraphs, as their readers do.
///
/// ```no_run
/// # fn main() -> litchi_ops::Result<()> {
/// let file = litchi_ops::open("report.docx")?;
/// let document = file.document()?;
/// println!("{} paragraphs", document.paragraph_count()?);
/// for element in document.elements()? {
///     if let Some(paragraph) = element.as_paragraph() {
///         println!("{}", paragraph.text()?);
///     }
/// }
/// # Ok(())
/// # }
/// ```
#[derive(Clone, Copy)]
pub struct DocumentView<'a>(Inner<'a>);

/// The handle a [`DocumentView`] reads, for the crate's adapters built on
/// the view.
#[derive(Clone, Copy)]
pub(crate) enum Inner<'a> {
    #[cfg(feature = "docx")]
    Docx(&'a Docx),
    #[cfg(feature = "doc")]
    Doc(&'a Doc),
    #[cfg(feature = "rtf")]
    Rtf(&'a Rtf),
    #[cfg(feature = "odt")]
    Odt(&'a Odt),
}

/// Type an ODT element read: the ODT leaf reports `litchi_core::Error`.
#[cfg(feature = "odt")]
fn odt_error(error: litchi_core::Error) -> Error {
    Error::from(error).with_format(FileFormat::Odt)
}

/// Type an allocation failure of a projected collection.
#[cfg(any(feature = "docx", feature = "odt"))]
fn reserve<T>(capacity: usize, resource: &'static str, format: FileFormat) -> Result<Vec<T>> {
    let mut projected = Vec::new();
    projected.try_reserve_exact(capacity).map_err(|source| {
        Error::from(litchi_core::Error::Allocation { resource, source }).with_format(format)
    })?;
    Ok(projected)
}

/// Section properties affect pagination, not the linear body stream: they
/// appear as an opaque body child in DOCX's ordered-block view and must not
/// make an otherwise textual document unreadable.
#[cfg(feature = "docx")]
pub(crate) fn docx_unknown_is_section_properties(block: &litchi_docx::OpaqueBlock) -> bool {
    let bytes = block.xml_bytes();
    let Some(open) = bytes.iter().position(|byte| *byte == b'<') else {
        return false;
    };
    let name_start = open.saturating_add(1);
    let Some(name_end) = bytes[name_start..]
        .iter()
        .position(|byte| matches!(*byte, b' ' | b'\t' | b'\r' | b'\n' | b'/' | b'>'))
        .map(|offset| name_start.saturating_add(offset))
    else {
        return false;
    };
    let name = &bytes[name_start..name_end];
    name == b"sectPr" || name.strip_prefix(b"w:") == Some(b"sectPr")
}

/// Project DOCX ordered elements; an active unmodeled body block is refused
/// rather than silently dropped.
#[cfg(feature = "docx")]
fn docx_elements(raw: Vec<litchi_docx::Element>) -> Result<Vec<DocumentElement>> {
    let mut elements = reserve(
        raw.len(),
        "unified DOCX document elements",
        FileFormat::Docx,
    )?;
    for element in raw {
        match element {
            litchi_docx::Element::Paragraph(paragraph) => {
                elements.push(DocumentElement::Paragraph(Box::new(Paragraph::Docx(
                    *paragraph,
                ))));
            },
            litchi_docx::Element::Table(table) => {
                elements.push(DocumentElement::Table(Box::new(Table::Docx(table))));
            },
            litchi_docx::Element::Unknown(block) => {
                if !docx_unknown_is_section_properties(&block) {
                    return Err(Error::new(
                        crate::Code::UnsupportedElement,
                        "the document view cannot represent an active unmodeled DOCX body block",
                    )
                    .with_format(FileFormat::Docx));
                }
            },
        }
    }
    Ok(elements)
}

/// Detach RTF paragraphs from the source buffer.
#[cfg(feature = "rtf")]
fn owned_rtf_paragraph(
    paragraph: litchi_rtf::ParagraphContent<'_>,
) -> litchi_rtf::ParagraphContent<'static> {
    litchi_rtf::ParagraphContent::new(
        paragraph.properties,
        paragraph
            .runs
            .into_iter()
            .map(|run| {
                litchi_rtf::Run::new(
                    std::borrow::Cow::Owned(run.text.into_owned()),
                    run.formatting,
                )
            })
            .collect(),
    )
}

impl<'a> DocumentView<'a> {
    #[cfg(feature = "docx")]
    pub(crate) fn docx(handle: &'a Docx) -> Self {
        Self(Inner::Docx(handle))
    }

    #[cfg(feature = "doc")]
    pub(crate) fn doc(handle: &'a Doc) -> Self {
        Self(Inner::Doc(handle))
    }

    #[cfg(feature = "rtf")]
    pub(crate) fn rtf(handle: &'a Rtf) -> Self {
        Self(Inner::Rtf(handle))
    }

    #[cfg(feature = "odt")]
    pub(crate) fn odt(handle: &'a Odt) -> Self {
        Self(Inner::Odt(handle))
    }

    /// The handle this view reads.
    #[cfg_attr(
        not(feature = "markdown"),
        expect(dead_code, reason = "the Markdown adapter is the only reader")
    )]
    pub(crate) fn inner(&self) -> Inner<'a> {
        self.0
    }

    /// The format of the document.
    #[must_use]
    pub fn format(&self) -> FileFormat {
        match self.0 {
            #[cfg(feature = "docx")]
            Inner::Docx(_) => FileFormat::Docx,
            #[cfg(feature = "doc")]
            Inner::Doc(_) => FileFormat::Doc,
            #[cfg(feature = "rtf")]
            Inner::Rtf(_) => FileFormat::Rtf,
            #[cfg(feature = "odt")]
            Inner::Odt(odt) => odt.format(),
        }
    }

    /// The number of paragraphs.
    ///
    /// # Errors
    ///
    /// Returns [`Code::IoError`](crate::Code::IoError) when a source-backed
    /// document changed after open, and the format's typed error when its
    /// body cannot be parsed.
    pub fn paragraph_count(&self) -> Result<usize> {
        match self.0 {
            #[cfg(feature = "docx")]
            Inner::Docx(docx) => match docx.owner() {
                DocxOwner::Eager(package) => Ok(package.document()?.paragraph_count()?),
                owner @ DocxOwner::Source(package) => owner.finish(
                    package
                        .document()
                        .and_then(|document| document.paragraph_count())
                        .map_err(Error::from),
                ),
            },
            #[cfg(feature = "doc")]
            Inner::Doc(doc) => {
                let result = (|| Ok(doc.reader()?.paragraph_count()?))();
                doc.check_source()?;
                result
            },
            #[cfg(feature = "rtf")]
            Inner::Rtf(rtf) => Ok(rtf.document().paragraph_count()),
            #[cfg(feature = "odt")]
            Inner::Odt(odt) => odt.paragraph_count(),
        }
    }

    /// The visible text of the paragraph at zero-based `index`, or `None`
    /// past the last paragraph.
    ///
    /// DOCX and ODT select the paragraph through their owners; DOC and RTF
    /// use the same projection as [`Self::paragraphs`].
    ///
    /// # Errors
    ///
    /// As [`Self::paragraph_count`].
    pub fn paragraph_text(
        &self,
        index: impl Into<litchi_core::Position>,
    ) -> Result<Option<String>> {
        let index = index.into().get();
        match self.0 {
            #[cfg(feature = "docx")]
            Inner::Docx(docx) => match docx.owner() {
                DocxOwner::Eager(package) => package
                    .document()?
                    .paragraph(index)?
                    .map(|paragraph| Ok(paragraph.text()?))
                    .transpose(),
                owner @ DocxOwner::Source(package) => owner.finish(
                    package
                        .document()
                        .and_then(|document| document.paragraph_text(index))
                        .map_err(Error::from),
                ),
            },
            #[cfg(feature = "doc")]
            Inner::Doc(doc) => {
                let result = (|| {
                    doc.reader()?
                        .paragraphs()?
                        .into_iter()
                        .nth(index)
                        .map(|paragraph| Ok(paragraph.text()?.to_owned()))
                        .transpose()
                })();
                doc.check_source()?;
                result
            },
            #[cfg(feature = "rtf")]
            Inner::Rtf(rtf) => Ok(rtf
                .document()
                .paragraphs_with_content()
                .into_iter()
                .nth(index)
                .map(|paragraph| paragraph.text())),
            #[cfg(feature = "odt")]
            Inner::Odt(odt) => match odt.owner() {
                OdtOwner::Eager(document) => document
                    .paragraph(index)
                    .map_err(odt_error)?
                    .map(|paragraph| paragraph.text().map_err(odt_error))
                    .transpose(),
                #[cfg(any(unix, windows))]
                owner @ OdtOwner::Source(document) => {
                    owner.finish(document.paragraph(index).and_then(|paragraph| {
                        paragraph.map(|paragraph| paragraph.text()).transpose()
                    }))
                },
            },
        }
    }

    /// The paragraph at zero-based `index`, or `None` past the last
    /// paragraph.
    ///
    /// # Errors
    ///
    /// As [`Self::paragraph_count`].
    pub fn paragraph(&self, index: impl Into<litchi_core::Position>) -> Result<Option<Paragraph>> {
        let index = index.into().get();
        match self.0 {
            #[cfg(feature = "docx")]
            Inner::Docx(docx) => match docx.owner() {
                DocxOwner::Eager(package) => {
                    Ok(package.document()?.paragraph(index)?.map(Paragraph::Docx))
                },
                owner @ DocxOwner::Source(package) => owner.finish(
                    package
                        .document()
                        .and_then(|document| document.paragraph(index))
                        .map(|paragraph| paragraph.map(Paragraph::Docx))
                        .map_err(Error::from),
                ),
            },
            #[cfg(feature = "doc")]
            Inner::Doc(doc) => {
                let result = (|| {
                    Ok(doc
                        .reader()?
                        .paragraphs()?
                        .into_iter()
                        .nth(index)
                        .map(Paragraph::Doc))
                })();
                doc.check_source()?;
                result
            },
            #[cfg(feature = "rtf")]
            Inner::Rtf(_) => Ok(self.paragraphs()?.into_iter().nth(index)),
            #[cfg(feature = "odt")]
            Inner::Odt(odt) => match odt.owner() {
                OdtOwner::Eager(document) => Ok(document
                    .paragraph(index)
                    .map_err(odt_error)?
                    .map(Paragraph::Odt)),
                #[cfg(any(unix, windows))]
                owner @ OdtOwner::Source(document) => owner.finish(
                    document
                        .paragraph(index)
                        .map(|paragraph| paragraph.map(Paragraph::Odt)),
                ),
            },
        }
    }

    /// Every paragraph, in document order.
    ///
    /// # Errors
    ///
    /// As [`Self::paragraph_count`].
    pub fn paragraphs(&self) -> Result<Vec<Paragraph>> {
        match self.0 {
            #[cfg(feature = "docx")]
            Inner::Docx(docx) => match docx.owner() {
                DocxOwner::Eager(package) => Ok(package
                    .document()?
                    .paragraphs()?
                    .into_iter()
                    .map(Paragraph::Docx)
                    .collect()),
                owner @ DocxOwner::Source(package) => owner.finish((|| {
                    let paragraphs = package.document()?.paragraphs()?;
                    let mut projected = reserve(
                        paragraphs.len(),
                        "unified source-backed DOCX paragraphs",
                        FileFormat::Docx,
                    )?;
                    projected.extend(paragraphs.into_iter().map(Paragraph::Docx));
                    Ok(projected)
                })()),
            },
            #[cfg(feature = "doc")]
            Inner::Doc(doc) => {
                let result = (|| {
                    Ok(doc
                        .reader()?
                        .paragraphs()?
                        .into_iter()
                        .map(Paragraph::Doc)
                        .collect())
                })();
                doc.check_source()?;
                result
            },
            #[cfg(feature = "rtf")]
            Inner::Rtf(rtf) => Ok(rtf
                .document()
                .paragraphs_with_content()
                .into_iter()
                .map(|paragraph| Paragraph::Rtf(owned_rtf_paragraph(paragraph)))
                .collect()),
            #[cfg(feature = "odt")]
            Inner::Odt(odt) => match odt.owner() {
                OdtOwner::Eager(document) => Ok(document
                    .paragraphs()
                    .map_err(odt_error)?
                    .into_iter()
                    .map(Paragraph::Odt)
                    .collect()),
                #[cfg(any(unix, windows))]
                owner @ OdtOwner::Source(document) => {
                    let paragraphs = owner.finish(document.paragraphs())?;
                    let mut projected = reserve(
                        paragraphs.len(),
                        "unified source-backed ODT paragraphs",
                        FileFormat::Odt,
                    )?;
                    projected.extend(paragraphs.into_iter().map(Paragraph::Odt));
                    Ok(projected)
                },
            },
        }
    }

    /// Every table, in document order.
    ///
    /// # Errors
    ///
    /// As [`Self::paragraph_count`].
    pub fn tables(&self) -> Result<Vec<Table>> {
        match self.0 {
            #[cfg(feature = "docx")]
            Inner::Docx(docx) => match docx.owner() {
                DocxOwner::Eager(package) => Ok(package
                    .document()?
                    .tables()?
                    .into_iter()
                    .map(|table| Table::Docx(Box::new(table)))
                    .collect()),
                owner @ DocxOwner::Source(package) => owner.finish((|| {
                    let tables = package.document()?.tables()?;
                    let mut projected = reserve(
                        tables.len(),
                        "unified source-backed DOCX tables",
                        FileFormat::Docx,
                    )?;
                    projected.extend(tables.into_iter().map(|table| Table::Docx(Box::new(table))));
                    Ok(projected)
                })()),
            },
            #[cfg(feature = "doc")]
            Inner::Doc(doc) => {
                let result = (|| {
                    Ok(doc
                        .reader()?
                        .tables()?
                        .into_iter()
                        .map(|table| Table::Doc(Box::new(table)))
                        .collect())
                })();
                doc.check_source()?;
                result
            },
            // Detach each table from the source buffer without flattening it:
            // `into_owned` keeps merge roles, borders, widths, nested tables
            // and drawings that a text-only rebuild would discard.
            #[cfg(feature = "rtf")]
            Inner::Rtf(rtf) => Ok(rtf
                .document()
                .tables()
                .iter()
                .map(|table| Table::Rtf(Box::new(table.clone().into_owned())))
                .collect()),
            #[cfg(feature = "odt")]
            Inner::Odt(odt) => match odt.owner() {
                OdtOwner::Eager(document) => Ok(document
                    .tables()
                    .map_err(odt_error)?
                    .into_iter()
                    .map(|table| Table::Odt(Box::new(table)))
                    .collect()),
                #[cfg(any(unix, windows))]
                owner @ OdtOwner::Source(document) => {
                    let tables = owner.finish(document.tables())?;
                    let mut projected = reserve(
                        tables.len(),
                        "unified source-backed ODT tables",
                        FileFormat::Odt,
                    )?;
                    projected.extend(tables.into_iter().map(|table| Table::Odt(Box::new(table))));
                    Ok(projected)
                },
            },
        }
    }

    /// Every paragraph and table, in document order, for sequential
    /// processing such as Markdown conversion.
    ///
    /// DOCX section properties are skipped; any other unmodeled DOCX body
    /// block is refused rather than dropped. ODT headings become paragraphs;
    /// ODT lists are skipped, as the facade skipped them.
    ///
    /// # Errors
    ///
    /// Returns [`Code::UnsupportedElement`](crate::Code::UnsupportedElement)
    /// for an active unmodeled DOCX body block, and otherwise as
    /// [`Self::paragraph_count`].
    pub fn elements(&self) -> Result<Vec<DocumentElement>> {
        match self.0 {
            #[cfg(feature = "docx")]
            Inner::Docx(docx) => match docx.owner() {
                DocxOwner::Eager(package) => docx_elements(package.document()?.elements()?),
                owner @ DocxOwner::Source(package) => {
                    owner.finish((|| docx_elements(package.document()?.elements()?))())
                },
            },
            #[cfg(feature = "doc")]
            Inner::Doc(doc) => {
                let result = (|| {
                    Ok(doc
                        .reader()?
                        .elements()?
                        .into_iter()
                        .map(|element| match element {
                            litchi_doc::Element::Paragraph(paragraph) => {
                                DocumentElement::Paragraph(Box::new(Paragraph::Doc(*paragraph)))
                            },
                            litchi_doc::Element::Table(table) => {
                                DocumentElement::Table(Box::new(Table::Doc(table)))
                            },
                        })
                        .collect())
                })();
                doc.check_source()?;
                result
            },
            #[cfg(feature = "rtf")]
            Inner::Rtf(rtf) => Ok(rtf
                .document()
                .elements()
                .into_iter()
                .map(|element| match element {
                    litchi_rtf::DocumentElement::Paragraph(paragraph) => {
                        DocumentElement::Paragraph(Box::new(Paragraph::Rtf(owned_rtf_paragraph(
                            paragraph,
                        ))))
                    },
                    // Detach without flattening; see `tables`.
                    litchi_rtf::DocumentElement::Table(table) => {
                        DocumentElement::Table(Box::new(Table::Rtf(Box::new(table.into_owned()))))
                    },
                })
                .collect()),
            #[cfg(feature = "odt")]
            Inner::Odt(odt) => odt_elements(odt),
        }
    }
}

/// Project ODT ordered elements as the facade did for each owner.
#[cfg(feature = "odt")]
fn odt_elements(odt: &Odt) -> Result<Vec<DocumentElement>> {
    use litchi_odt::elements::parser::OrderElement;
    use litchi_odt::elements::text::Paragraph as ElementParagraph;

    match odt.owner() {
        OdtOwner::Eager(document) => {
            let mut elements = Vec::new();
            for element in document.elements().map_err(odt_error)? {
                match element {
                    OrderElement::Paragraph(paragraph) => {
                        elements.push(DocumentElement::Paragraph(Box::new(Paragraph::Odt(
                            paragraph,
                        ))));
                    },
                    OrderElement::NumberedParagraph(paragraph) => {
                        elements.push(DocumentElement::Paragraph(Box::new(Paragraph::Odt(
                            paragraph.into_paragraph(),
                        ))));
                    },
                    OrderElement::Heading(heading) => {
                        // A heading becomes a paragraph with its text and
                        // style; an unreadable heading is skipped, as the
                        // facade skipped it.
                        if let Ok(text) = heading.text() {
                            let mut paragraph = ElementParagraph::new();
                            paragraph.set_text(&text);
                            if let Some(style) = heading.style_name() {
                                paragraph.set_style_name(style);
                            }
                            elements.push(DocumentElement::Paragraph(Box::new(Paragraph::Odt(
                                paragraph,
                            ))));
                        }
                    },
                    OrderElement::Table(table) => {
                        elements.push(DocumentElement::Table(Box::new(Table::Odt(Box::new(
                            table,
                        )))));
                    },
                    OrderElement::List(_) => {},
                }
            }
            Ok(elements)
        },
        #[cfg(any(unix, windows))]
        owner @ OdtOwner::Source(document) => owner.finish((|| {
            let raw = document.elements()?;
            let mut elements = Vec::new();
            elements.try_reserve_exact(raw.len()).map_err(|source| {
                litchi_core::Error::Allocation {
                    resource: "unified source-backed ODT document elements",
                    source,
                }
            })?;
            for element in raw {
                match element {
                    OrderElement::Paragraph(paragraph) => {
                        elements.push(DocumentElement::Paragraph(Box::new(Paragraph::Odt(
                            paragraph,
                        ))));
                    },
                    OrderElement::NumberedParagraph(paragraph) => {
                        elements.push(DocumentElement::Paragraph(Box::new(Paragraph::Odt(
                            paragraph.into_paragraph(),
                        ))));
                    },
                    OrderElement::Heading(heading) => {
                        elements.push(DocumentElement::Paragraph(Box::new(Paragraph::Odt(
                            heading.try_into_paragraph()?,
                        ))));
                    },
                    OrderElement::Table(table) => {
                        elements.push(DocumentElement::Table(Box::new(Table::Odt(Box::new(
                            table,
                        )))));
                    },
                    OrderElement::List(_) => {},
                }
            }
            Ok(elements)
        })()),
    }
}

impl std::fmt::Debug for DocumentView<'_> {
    fn fmt(&self, formatter: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        formatter
            .debug_struct("DocumentView")
            .field("format", &self.format())
            .finish_non_exhaustive()
    }
}
