//! The raw escape layer (design §12; record 0015 W10/D19; record 0336):
//! package parts read as bytes, a part added with its content type and
//! relationship, an audited XPath-addressed XML edit, and a read-only
//! structural validation.
//!
//! These operations are format-specific escape hatches, not part of the node
//! view. They address *package parts*: an OPC part name (`/word/document.xml`)
//! or one of OfficeCLI's part aliases (`/document`, `/styles`, `/slide[2]`,
//! `/Sheet1`, `/Sheet1/drawing`, `/Sheet1/chart[1]`, ...), or an ODF member
//! (`/content.xml`, or the aliases `/content`, `/styles`, `/meta`,
//! `/settings`, `/manifest`).
//!
//! | Operation | Formats | Result |
//! |---|---|---|
//! | [`File::parts`], [`File::raw`] | DOCX, XLSX, PPTX, XLSB; packaged ODF | [`PartInfo`], [`Part`] |
//! | [`File::add_part`] | DOCX, XLSX, PPTX, XLSB | [`AddedPart`]: a new handle, the part name and the relationship id |
//! | [`File::raw_set`] | DOCX, XLSX, PPTX, XLSB (XML parts) | [`Edited`]: a new handle and the match count |
//! | [`File::validate`] | every format but iWork | [`Validation`] |
//!
//! Every operation works on the package the handle publishes
//! ([`File::to_bytes`]), so a raw read sees what a save would write, edits
//! included. The writes never mutate their input: they publish the package
//! with the change and open the result as a new handle. Every XML payload a
//! write introduces passes the package's publication audit
//! (`litchi_opc::OpcPackage::try_replace_owned_xml_part_bytes` and
//! `try_add_owned_xml_part_bytes`, records 0313 and 0317), and every member
//! the write does not touch keeps its bytes.
//!
//! ```
//! # #[cfg(feature = "docx")]
//! # {
//! use litchi_core::FileFormat;
//! use litchi_ops::raw::Action;
//!
//! let file = litchi_ops::create(FileFormat::Docx)?;
//! let document = file.raw("/document")?;
//! assert_eq!(document.name(), "/word/document.xml");
//! assert!(document.text()?.contains("<w:body"));
//!
//! let edited = file.raw_set(
//!     "/document",
//!     "//w:body",
//!     Action::Prepend,
//!     Some("<w:p><w:r><w:t>Hello</w:t></w:r></w:p>"),
//! )?;
//! assert_eq!(edited.matched, 1);
//! assert_eq!(edited.file.text()?.trim(), "Hello");
//! assert!(edited.file.validate()?.is_valid());
//! # }
//! # Ok::<(), litchi_ops::Error>(())
//! ```

#![cfg_attr(
    not(any(feature = "docx", feature = "xlsx", feature = "pptx", feature = "xlsb")),
    allow(
        dead_code,
        reason = "the part payloads are read only by the OOXML writes"
    )
)]

use std::fmt;
use std::str::FromStr;

use litchi_core::{FileFormat, ValidateReport};
use serde::ser::{Serialize, SerializeMap, Serializer};

use crate::{Code, Error as OpsError, File, Result};

#[cfg(any(
    feature = "docx",
    feature = "xlsx",
    feature = "pptx",
    feature = "xlsb",
    feature = "odt",
    feature = "ods",
    feature = "odp",
    odf_objects
))]
mod package;
#[cfg(any(feature = "docx", feature = "xlsx", feature = "pptx", feature = "xlsb"))]
mod parts;
mod validate;
#[cfg(any(feature = "docx", feature = "xlsx", feature = "pptx", feature = "xlsb"))]
mod xpath;

#[cfg(test)]
mod tests;

pub use validate::Error;

/// One package part read by [`File::raw`]: its name, content type and
/// payload bytes.
///
/// `Debug` shows the name, content type and size, never the payload.
#[derive(Clone, PartialEq, Eq)]
pub struct Part {
    name: String,
    content_type: String,
    bytes: Vec<u8>,
}

impl Part {
    /// The part name, for example `/word/document.xml` (ODF members are
    /// named with a leading `/` too: `/content.xml`).
    #[must_use]
    pub fn name(&self) -> &str {
        &self.name
    }

    /// The part's content type (OPC) or manifest media type (ODF); empty
    /// when the package declares none.
    #[must_use]
    pub fn content_type(&self) -> &str {
        &self.content_type
    }

    /// The payload bytes, decompressed.
    #[must_use]
    pub fn bytes(&self) -> &[u8] {
        &self.bytes
    }

    /// The payload bytes, without a copy.
    #[must_use]
    pub fn into_bytes(self) -> Vec<u8> {
        self.bytes
    }

    /// Whether the part is XML, by its name or its content type.
    #[must_use]
    pub fn is_xml(&self) -> bool {
        is_xml(&self.name, &self.content_type)
    }

    /// The payload as UTF-8 text, for XML parts.
    ///
    /// # Errors
    /// Returns `invalid_input` when the payload is not UTF-8 (a binary part,
    /// or XML in another encoding).
    pub fn text(&self) -> Result<&str> {
        std::str::from_utf8(&self.bytes).map_err(|_| {
            OpsError::new(Code::InvalidInput, "the part is not UTF-8 text")
                .with_suggestion("read binary parts with `Part::bytes`")
        })
    }
}

impl fmt::Debug for Part {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter
            .debug_struct("Part")
            .field("name", &self.name)
            .field("content_type", &self.content_type)
            .field("bytes", &self.bytes.len())
            .finish()
    }
}

/// One entry of [`File::parts`]: a part's name and content type.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct PartInfo {
    name: String,
    content_type: String,
}

impl PartInfo {
    /// The part name.
    #[must_use]
    pub fn name(&self) -> &str {
        &self.name
    }

    /// The part's content type; empty when the package declares none.
    #[must_use]
    pub fn content_type(&self) -> &str {
        &self.content_type
    }
}

impl Serialize for PartInfo {
    fn serialize<S: Serializer>(&self, serializer: S) -> std::result::Result<S::Ok, S::Error> {
        let mut map = serializer.serialize_map(Some(2))?;
        map.serialize_entry("name", &self.name)?;
        map.serialize_entry("contentType", &self.content_type)?;
        map.end()
    }
}

/// A part for [`File::add_part`]: its content type, the relationship type
/// that links it to its parent, its payload, and optionally its name.
///
/// [`NewPart::of_type`] builds OfficeCLI's `add-part --type` parts (an empty
/// chart, header or footer); [`NewPart::new`] adds any payload.
#[derive(Clone)]
pub struct NewPart {
    kind: Kind,
    name: Option<String>,
}

#[derive(Clone)]
enum Kind {
    Chart,
    Header,
    Footer,
    Custom {
        content_type: String,
        relationship_type: String,
        bytes: Vec<u8>,
    },
}

/// The DrawingML chart content type.
const CHART_CONTENT_TYPE: &str =
    "application/vnd.openxmlformats-officedocument.drawingml.chart+xml";
/// The chart relationship type.
const CHART_RELATIONSHIP: &str =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart";
const HEADER_CONTENT_TYPE: &str =
    "application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml";
const HEADER_RELATIONSHIP: &str =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/header";
const FOOTER_CONTENT_TYPE: &str =
    "application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml";
const FOOTER_RELATIONSHIP: &str =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer";

impl NewPart {
    /// An OfficeCLI `add-part --type` part: `chart` (an empty chart space,
    /// for DOCX, XLSX and PPTX), `header` or `footer` (an empty paragraph,
    /// DOCX only). Matching is ASCII case-insensitive.
    ///
    /// # Errors
    /// Returns `unsupported_type` for any other type.
    pub fn of_type(kind: &str) -> Result<Self> {
        let kind = if kind.eq_ignore_ascii_case("chart") {
            Kind::Chart
        } else if kind.eq_ignore_ascii_case("header") {
            Kind::Header
        } else if kind.eq_ignore_ascii_case("footer") {
            Kind::Footer
        } else {
            return Err(OpsError::new(Code::UnsupportedType, "unknown part type")
                .with_suggestion("use chart, header or footer, or NewPart::new for any payload"));
        };
        Ok(Self { kind, name: None })
    }

    /// Any part: its content type, the relationship type its parent links
    /// it with, and its payload. Give it a name with [`Self::named`].
    #[must_use]
    pub fn new(
        content_type: impl Into<String>,
        relationship_type: impl Into<String>,
        bytes: impl Into<Vec<u8>>,
    ) -> Self {
        Self {
            kind: Kind::Custom {
                content_type: content_type.into(),
                relationship_type: relationship_type.into(),
                bytes: bytes.into(),
            },
            name: None,
        }
    }

    /// The same part under this part name (for example
    /// `/customXml/item9.xml`). Typed parts are named after the next free
    /// numbered name when unnamed; a custom part needs a name.
    #[must_use]
    pub fn named(mut self, name: impl Into<String>) -> Self {
        self.name = Some(name.into());
        self
    }
}

impl fmt::Debug for NewPart {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter.debug_struct("NewPart").finish_non_exhaustive()
    }
}

/// The result of [`File::add_part`].
#[derive(Clone)]
#[non_exhaustive]
pub struct AddedPart {
    /// The new handle, holding the part.
    pub file: File,
    /// The new part's name, for example `/xl/charts/chart3.xml`.
    pub part: String,
    /// The id of the relationship from the parent part, for example `rId4`.
    pub relationship_id: String,
}

impl fmt::Debug for AddedPart {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter
            .debug_struct("AddedPart")
            .field("part", &self.part)
            .field("relationship_id", &self.relationship_id)
            .finish_non_exhaustive()
    }
}

/// The result of [`File::raw_set`].
#[derive(Clone)]
#[non_exhaustive]
pub struct Edited {
    /// The new handle.
    pub file: File,
    /// How many elements the XPath matched (each was edited).
    pub matched: usize,
}

impl fmt::Debug for Edited {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter
            .debug_struct("Edited")
            .field("matched", &self.matched)
            .finish_non_exhaustive()
    }
}

/// What [`File::raw_set`] does at each element the XPath matches
/// (OfficeCLI `raw-set --action`).
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
#[non_exhaustive]
pub enum Action {
    /// Insert the XML as the element's last children.
    Append,
    /// Insert the XML as the element's first children.
    Prepend,
    /// Insert the XML immediately before the element.
    InsertBefore,
    /// Insert the XML immediately after the element.
    InsertAfter,
    /// Replace the element with the XML.
    Replace,
    /// Remove the element.
    Remove,
    /// Set one attribute: the XML is `name=value` (`w:val=center`).
    SetAttr,
}

impl Action {
    /// The OfficeCLI spelling.
    #[must_use]
    pub const fn as_str(self) -> &'static str {
        match self {
            Self::Append => "append",
            Self::Prepend => "prepend",
            Self::InsertBefore => "insertbefore",
            Self::InsertAfter => "insertafter",
            Self::Replace => "replace",
            Self::Remove => "remove",
            Self::SetAttr => "setattr",
        }
    }
}

impl FromStr for Action {
    type Err = OpsError;

    /// OfficeCLI's action names, ASCII case-insensitively, including its
    /// aliases `before`, `after` and `delete`.
    fn from_str(value: &str) -> Result<Self> {
        const ACTIONS: [(&str, Action); 10] = [
            ("append", Action::Append),
            ("prepend", Action::Prepend),
            ("insertbefore", Action::InsertBefore),
            ("before", Action::InsertBefore),
            ("insertafter", Action::InsertAfter),
            ("after", Action::InsertAfter),
            ("replace", Action::Replace),
            ("remove", Action::Remove),
            ("delete", Action::Remove),
            ("setattr", Action::SetAttr),
        ];
        ACTIONS
            .iter()
            .find(|(name, _)| value.eq_ignore_ascii_case(name))
            .map(|(_, action)| *action)
            .ok_or_else(|| {
                OpsError::new(Code::InvalidArgument, "unknown raw-set action").with_suggestion(
                    "use append, prepend, insertbefore, insertafter, replace, remove or setattr",
                )
            })
    }
}

/// The result of [`File::validate`]: the format and a bounded, typed
/// [`ValidateReport`] of checks and findings.
///
/// Validation is structural: ZIP and package ingress, content types,
/// relationships and their targets, XML well-formedness, and the format
/// owner's own checks (DOCX and PPTX main-part semantics, XLS/PPT/DOC
/// compound-file and record structure, ODF manifest and package rules, RTF
/// syntax). It is not OOXML schema validation (OfficeCLI runs the Open XML
/// SDK's schema validator). It never mutates and never fetches external
/// targets.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Validation {
    format: FileFormat,
    report: ValidateReport,
}

impl Validation {
    /// The validated format.
    #[must_use]
    pub const fn format(&self) -> FileFormat {
        self.format
    }

    /// The checks and findings.
    #[must_use]
    pub const fn report(&self) -> &ValidateReport {
        &self.report
    }

    /// Whether no finding is an error or fatal. Informational findings
    /// (signature or macro presence, external targets) do not count.
    #[must_use]
    pub fn is_valid(&self) -> bool {
        !self.report.has_errors()
    }

    /// The findings of error or fatal severity, in report order.
    pub fn errors(&self) -> impl Iterator<Item = &litchi_core::ValidationIssue> {
        self.report
            .issues()
            .iter()
            .filter(|issue| issue.severity() >= litchi_core::IssueSeverity::Error)
    }
}

impl Serialize for Validation {
    fn serialize<S: Serializer>(&self, serializer: S) -> std::result::Result<S::Ok, S::Error> {
        let mut map = serializer.serialize_map(Some(5))?;
        map.serialize_entry("format", crate::schema::format_name(self.format))?;
        map.serialize_entry("valid", &self.is_valid())?;
        map.serialize_entry("complete", &self.report.is_complete())?;
        map.serialize_entry("checks", self.report.checks())?;
        map.serialize_entry("issues", self.report.issues())?;
        map.end()
    }
}

impl File {
    /// The package parts, in name order (design §12).
    ///
    /// # Errors
    /// Returns `unsupported_capability`, naming the format, for formats
    /// without package parts (legacy binary, RTF, flat ODF, iWork); the
    /// publication's errors otherwise.
    ///
    /// ```
    /// # #[cfg(feature = "xlsx")]
    /// # {
    /// let file = litchi_ops::create(litchi_core::FileFormat::Xlsx)?;
    /// let parts = file.parts()?;
    /// assert!(parts.iter().any(|part| part.name() == "/xl/workbook.xml"));
    /// # }
    /// # Ok::<(), litchi_ops::Error>(())
    /// ```
    pub fn parts(&self) -> Result<Vec<PartInfo>> {
        #[cfg(any(
            feature = "docx",
            feature = "xlsx",
            feature = "pptx",
            feature = "xlsb",
            feature = "odt",
            feature = "ods",
            feature = "odp",
            odf_objects
        ))]
        {
            package::Container::of(self)?.parts()
        }
        #[cfg(not(any(
            feature = "docx",
            feature = "xlsx",
            feature = "pptx",
            feature = "xlsb",
            feature = "odt",
            feature = "ods",
            feature = "odp",
            odf_objects
        )))]
        {
            Err(no_parts(self))
        }
    }

    /// Read one package part (design §12; OfficeCLI `raw`).
    ///
    /// `part` is a part name or an alias (see the [module](crate::raw)
    /// documentation). The bytes are those the handle would publish.
    ///
    /// # Errors
    /// Returns `path_not_found` for an unknown part or alias,
    /// `invalid_path` for a name that is not absolute, and the errors of
    /// [`Self::parts`].
    pub fn raw(&self, part: &str) -> Result<Part> {
        #[cfg(any(
            feature = "docx",
            feature = "xlsx",
            feature = "pptx",
            feature = "xlsb",
            feature = "odt",
            feature = "ods",
            feature = "odp",
            odf_objects
        ))]
        {
            package::Container::of(self)?.read(part)
        }
        #[cfg(not(any(
            feature = "docx",
            feature = "xlsx",
            feature = "pptx",
            feature = "xlsb",
            feature = "odt",
            feature = "ods",
            feature = "odp",
            odf_objects
        )))]
        {
            let _ = part;
            Err(no_parts(self))
        }
    }

    /// Add a part under `parent`, with its content type override and a
    /// relationship from the parent part (design §12; OfficeCLI
    /// `add-part`). `parent` is a part name or alias; `/` relates the part
    /// from the package itself. For an XLSX sheet parent a chart is related
    /// from the sheet's drawing, as OfficeCLI does.
    ///
    /// The handle is unchanged; the result holds a new one. An XML payload
    /// passes the package's publication audit first.
    ///
    /// # Errors
    /// Returns `document_protected` for signed or locked handles,
    /// `unsupported_capability` naming the format outside OOXML,
    /// `path_not_found` for an unknown parent, `file_exists` for a taken
    /// part name, `missing_argument` for an unnamed custom part, and
    /// `corrupt_file` for an XML payload that fails the audit.
    ///
    /// ```
    /// # #[cfg(feature = "docx")]
    /// # {
    /// use litchi_ops::raw::NewPart;
    ///
    /// let file = litchi_ops::create(litchi_core::FileFormat::Docx)?;
    /// let added = file.add_part("/document", NewPart::of_type("header")?)?;
    /// assert_eq!(added.part, "/word/header1.xml");
    /// assert!(added.relationship_id.starts_with("rId"));
    /// assert!(added.file.raw("/header[1]").is_ok());
    /// # }
    /// # Ok::<(), litchi_ops::Error>(())
    /// ```
    pub fn add_part(&self, parent: &str, part: NewPart) -> Result<AddedPart> {
        #[cfg(any(feature = "docx", feature = "xlsx", feature = "pptx", feature = "xlsb"))]
        {
            parts::add_part(self, parent, part)
        }
        #[cfg(not(any(feature = "docx", feature = "xlsx", feature = "pptx", feature = "xlsb")))]
        {
            let _ = (parent, part);
            Err(no_raw_write(self, "adding a part to"))
        }
    }

    /// Edit the XML of one part at every element `xpath` matches (design
    /// §12; OfficeCLI `raw-set`).
    ///
    /// The XPath is a location path: `/` and `//` steps of prefixed names
    /// or `*`, each with optional `[n]`, `[last()]` and `[@attr='value']`
    /// predicates. OfficeCLI's prefixes (`w`, `r`, `a`, `p`, `x`, `wp`,
    /// `mc`, `c`, `xdr`, `wps`, `wp14`, `v`) and the part root's own
    /// declarations are known; fragment prefixes resolve the same way and
    /// get declarations where the target scope lacks them. Inter-element
    /// whitespace in the fragment is dropped, as OfficeCLI does.
    ///
    /// The edited part is spliced, so every byte outside the matched
    /// elements is kept, and the whole result must pass the package's
    /// publication audit before the new handle is opened.
    ///
    /// # Errors
    /// Returns `invalid_xpath` for an XPath outside the subset or an
    /// unknown prefix, `zero_matches` when nothing matches,
    /// `invalid_argument` for a missing or malformed fragment or matches
    /// nested in one another, `unsupported_capability` for non-XML parts
    /// and formats outside OOXML, and the errors of [`Self::add_part`].
    pub fn raw_set(
        &self,
        part: &str,
        xpath: &str,
        action: Action,
        xml: Option<&str>,
    ) -> Result<Edited> {
        #[cfg(any(feature = "docx", feature = "xlsx", feature = "pptx", feature = "xlsb"))]
        {
            parts::raw_set(self, part, xpath, action, xml)
        }
        #[cfg(not(any(feature = "docx", feature = "xlsx", feature = "pptx", feature = "xlsb")))]
        {
            let _ = (part, xpath, action, xml);
            Err(no_raw_write(self, "raw XML edits of"))
        }
    }

    /// Validate the package structure this handle publishes, read-only
    /// (design §12; OfficeCLI `validate`). See [`Validation`] for what is
    /// checked.
    ///
    /// # Errors
    /// Returns `unsupported_capability` naming the format for iWork
    /// packages, `document_protected` for locked handles, and the
    /// publication's errors. Findings are not errors: they are in the
    /// report.
    pub fn validate(&self) -> Result<Validation> {
        validate::validate(self)
    }
}

pub(crate) fn is_xml(name: &str, content_type: &str) -> bool {
    let name = name.to_ascii_lowercase();
    let media = content_type.split(';').next().unwrap_or("").trim();
    name.ends_with(".xml")
        || name.ends_with(".rels")
        || name.ends_with(".rdf")
        || media.eq_ignore_ascii_case("application/xml")
        || media.eq_ignore_ascii_case("text/xml")
        || media
            .rsplit_once('+')
            .is_some_and(|(_, suffix)| suffix.eq_ignore_ascii_case("xml"))
}

/// `unsupported_capability` for a format without package parts.
pub(crate) fn no_parts(file: &File) -> OpsError {
    with_format(
        file,
        OpsError::new(
            Code::UnsupportedCapability,
            "this format has no package parts to read",
        )
        .with_suggestion(
            "the raw part layer covers OOXML (DOCX, XLSX, PPTX, XLSB) and packaged ODF; use get, query and view",
        ),
    )
}

/// `unsupported_capability` for a raw write outside OOXML.
#[cfg_attr(
    any(feature = "docx", feature = "xlsx", feature = "pptx", feature = "xlsb"),
    allow(dead_code, reason = "used by the non-OOXML dispatch arms")
)]
pub(crate) fn no_raw_write(file: &File, operation: &str) -> OpsError {
    with_format(
        file,
        OpsError::new(
            Code::UnsupportedCapability,
            format!("{operation} this format is not supported"),
        )
        .with_suggestion(
            "raw part writes cover OOXML (DOCX, XLSX, PPTX, XLSB); record 0336 follow-up for ODF",
        ),
    )
}

pub(crate) fn with_format(file: &File, error: OpsError) -> OpsError {
    match file.format() {
        Some(format) => error.with_format(format),
        None => error,
    }
}
