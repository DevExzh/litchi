//! The shared ODF metadata transaction (ADR 0034 §7): document properties
//! of `meta.xml`, edited as exact element splices.
//!
//! Every ODF leaf exposes the family as `Package::{edit_metadata,
//! apply_metadata}` and a `metadata` module whose `Edit`, `Commit` and
//! `Patch` are this module's generic types bound to the leaf's `Package`
//! (record 0288). [`Edit::set`] stages one [`Field`]; [`Edit::commit`]
//! replaces only the touched `office:meta` children, copies every other
//! byte and member of the source, and adopts the result through the leaf's
//! own ingress. Generator, dates and statistics are never rewritten, so a
//! commit consults no clock. An edit that changes nothing returns the source
//! package (an exact no-op); a changed edit of a signed or encrypted package
//! is refused.
//!
//! ```
//! use litchi_odf_common::metadata::{Edit, Field, Owner};
//! use litchi_odf_common::{Error, OwnedPackage};
//!
//! /// A minimal owner: the physical package itself.
//! #[derive(Clone)]
//! struct Raw(OwnedPackage);
//!
//! impl Owner for Raw {
//!     fn metadata_source(&self) -> Result<OwnedPackage, Error> {
//!         Ok(self.0.clone())
//!     }
//!     fn adopt_metadata(&self, bytes: Vec<u8>) -> Result<Self, Error> {
//!         Ok(Self(OwnedPackage::from_bytes(bytes)?))
//!     }
//!     fn metadata_bytes(&self) -> &[u8] {
//!         self.0.as_bytes()
//!     }
//! }
//!
//! # fn main() -> Result<(), Error> {
//! # let bytes = {
//! #     let mut writer = litchi_odf_common::core::PackageWriter::new(litchi_odf_common::core::StreamingArchiveLimits::default());
//! #     writer.set_mimetype("application/vnd.oasis.opendocument.text")?;
//! #     writer.add_file("content.xml", br#"<office:document-content xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" office:version="1.3"><office:body><office:text/></office:body></office:document-content>"#)?;
//! #     writer.add_file("meta.xml", br#"<office:document-meta xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:dc="http://purl.org/dc/elements/1.1/" office:version="1.3"><office:meta/></office:document-meta>"#)?;
//! #     writer.finish_to_bytes()?
//! # };
//! let source = Raw(OwnedPackage::from_bytes(bytes)?);
//! let mut edit = Edit::new(&source);
//! edit.set(Field::Title, Some("Quarterly report"))?;
//! let commit = edit.commit()?;
//! assert!(commit.diagnostics().changed());
//! let meta = commit.package().0.get_file("meta.xml")?;
//! assert!(String::from_utf8_lossy(&meta).contains("<dc:title>Quarterly report</dc:title>"));
//! # Ok(())
//! # }
//! ```

use std::collections::BTreeMap;
use std::ops::Range;

use litchi_core::Error as CoreError;
use quick_xml::NsReader;
use quick_xml::events::Event;
use quick_xml::name::{Namespace, PrefixDeclaration, ResolveResult};

use crate::core::{
    AuthoredXmlFragment, XmlSourcePart, XmlSplicePublication, rebuild_package_with_xml_splices,
};
use crate::{Error, OwnedPackage};

pub use crate::LimitKind;

const OFFICE: &[u8] = b"urn:oasis:names:tc:opendocument:xmlns:office:1.0";
const META: &[u8] = b"urn:oasis:names:tc:opendocument:xmlns:meta:1.0";
const DC: &[u8] = b"http://purl.org/dc/elements/1.1/";
/// The largest property value one assignment accepts.
const MAX_VALUE_BYTES: usize = 1024 * 1024;
/// The bound on the published package.
const MAX_PACKAGE_BYTES: usize = 2 * 1024 * 1024 * 1024 - 1;

/// A document property the metadata transaction writes.
#[derive(Clone, Copy, Debug, PartialEq, Eq, Hash, PartialOrd, Ord)]
#[non_exhaustive]
pub enum Field {
    /// `dc:title`.
    Title,
    /// `dc:subject`.
    Subject,
    /// `dc:description`.
    Description,
    /// `meta:initial-creator`, the original author.
    Author,
    /// `dc:creator`, who saved the document last.
    LastModifiedBy,
    /// `dc:language`.
    Language,
    /// The `meta:keyword` elements, one per keyword: [`Edit::set_keywords`]
    /// takes the list; [`Edit::set`] splits its text at `;` and `,`
    /// (`litchi_core::metadata::keywords::split`).
    Keywords,
}

impl Field {
    /// The namespace URI and local name of the field's element.
    const fn qualified(self) -> (&'static [u8], &'static str) {
        match self {
            Self::Title => (DC, "title"),
            Self::Subject => (DC, "subject"),
            Self::Description => (DC, "description"),
            Self::Author => (META, "initial-creator"),
            Self::LastModifiedBy => (DC, "creator"),
            Self::Language => (DC, "language"),
            Self::Keywords => (META, "keyword"),
        }
    }
}

/// A leaf package the metadata transaction edits.
///
/// Implemented by every ODF leaf `Package`; the leaf's `metadata` module
/// binds [`Edit`], [`Commit`] and [`Patch`] to it.
pub trait Owner: Clone {
    /// The physical package whose `meta.xml` an edit splices.
    ///
    /// # Errors
    ///
    /// The leaf's ingress errors when it must index its bytes again.
    fn metadata_source(&self) -> Result<OwnedPackage, Error>;

    /// Adopt published package bytes as a package of the same flavor.
    ///
    /// # Errors
    ///
    /// The leaf's ingress errors (a candidate the leaf refuses).
    fn adopt_metadata(&self, bytes: Vec<u8>) -> Result<Self, Error>;

    /// The package's exact bytes: a patch applies only to its source.
    fn metadata_bytes(&self) -> &[u8];
}

/// Staged field assignments, shared by the generic [`Edit`] and leaves that
/// publish metadata beside their own splices (ODT body edits).
#[derive(Clone, Debug, Default)]
pub struct Assignments {
    /// Each staged field's element texts (one per `meta:keyword`, else
    /// one), or `None` for a removal.
    fields: BTreeMap<Field, Option<Vec<String>>>,
}

impl Assignments {
    /// Stage setting (`Some`) or removing (`None`) one property. A later
    /// assignment of the same field replaces an earlier one.
    ///
    /// # Errors
    ///
    /// Returns [`Error::Core`] for text XML cannot carry or a value over
    /// 1 MiB.
    pub fn set(&mut self, field: Field, value: Option<&str>) -> Result<(), Error> {
        if let Some(value) = value {
            check_text(value)?;
        }
        let values = value.map(|value| {
            if field == Field::Keywords {
                litchi_core::metadata::keywords::split(value)
            } else {
                vec![value.to_owned()]
            }
        });
        self.fields.insert(field, values);
        Ok(())
    }

    /// Stage the keyword list, one `meta:keyword` element per keyword; an
    /// empty list removes every keyword. Empty keywords are dropped.
    ///
    /// # Errors
    ///
    /// Returns [`Error::Core`] for text XML cannot carry or a keyword over
    /// 1 MiB.
    pub fn set_keywords(&mut self, keywords: &[String]) -> Result<(), Error> {
        let mut list = Vec::with_capacity(keywords.len());
        for keyword in keywords.iter().filter(|keyword| !keyword.is_empty()) {
            check_text(keyword)?;
            list.push(keyword.clone());
        }
        self.fields
            .insert(Field::Keywords, (!list.is_empty()).then_some(list));
        Ok(())
    }

    /// Whether nothing is staged.
    #[must_use]
    pub fn is_empty(&self) -> bool {
        self.fields.is_empty()
    }

    /// The `meta.xml` splice publication of the staged assignments over
    /// `package`, or `None` when every assignment matches the source.
    ///
    /// # Errors
    ///
    /// Returns [`Error::Core`] when the package has no `meta.xml` or no
    /// `office:meta` element, or the part is malformed.
    pub fn publication(
        &self,
        package: &OwnedPackage,
    ) -> Result<Option<XmlSplicePublication>, Error> {
        match self.plan_publication(package)? {
            Publication::Unchanged => Ok(None),
            Publication::Splices(publication) => Ok(Some(publication)),
            Publication::MissingPart => Err(invalid(
                "the ODF package has no meta.xml for metadata edits",
            )),
        }
    }

    /// What the staged assignments publish over `package`, with one
    /// `meta.xml` lookup: splices into the part, nothing, or
    /// [`Publication::MissingPart`] when the package has none (then
    /// [`Assignments::created_package`] adds it).
    ///
    /// # Errors
    ///
    /// Returns [`Error::Core`] when `meta.xml` has no `office:meta` element
    /// or is malformed.
    pub fn plan_publication(&self, package: &OwnedPackage) -> Result<Publication, Error> {
        if self.fields.is_empty() {
            return Ok(Publication::Unchanged);
        }
        if !package.has_file("meta.xml")? {
            return Ok(Publication::MissingPart);
        }
        let part = XmlSourcePart::load(package, "meta.xml")?;
        let splices = plan(part.bytes(), &self.fields)?;
        if splices.is_empty() {
            return Ok(Publication::Unchanged);
        }
        let mut publication = XmlSplicePublication::new(part.clone());
        for (range, bytes) in splices {
            let proof = part.checked_range(range.clone(), &part.bytes()[range])?;
            let fragment = if bytes.is_empty() {
                AuthoredXmlFragment::deletion()
            } else {
                AuthoredXmlFragment::markup(bytes.into_bytes())?
            };
            publication.replace(proof, fragment)?;
        }
        Ok(Publication::Splices(publication))
    }

    /// The package with a new `meta.xml` holding the staged values, for a
    /// package that has none (OTH and ODM packages from the leaf builders);
    /// `None` when nothing would be written (only removals were staged). The
    /// part declares the `office:version` of `content.xml`; every other
    /// member is copied.
    ///
    /// # Errors
    ///
    /// Returns [`Error::Core`] for a signed or encrypted package, or when the
    /// package cannot be read or rewritten.
    pub fn created_package(&self, package: &OwnedPackage) -> Result<Option<Vec<u8>>, Error> {
        let mut children = String::new();
        for (field, value) in &self.fields {
            let Some(values) = value else { continue };
            let (uri, local) = field.qualified();
            let prefix = if uri == META { "meta" } else { "dc" };
            for value in values {
                render(&mut children, prefix, local, None, value);
            }
        }
        if children.is_empty() {
            return Ok(None);
        }
        refuse_envelope(package)?;
        let version = content_version(package)?;
        let mut xml = String::from(concat!(
            "<?xml version=\"1.0\" encoding=\"UTF-8\"?>",
            "<office:document-meta xmlns:office=\"urn:oasis:names:tc:opendocument:xmlns:office:1.0\" ",
            "xmlns:meta=\"urn:oasis:names:tc:opendocument:xmlns:meta:1.0\" ",
            "xmlns:dc=\"http://purl.org/dc/elements/1.1/\" office:version=\""
        ));
        xml.push_str(&version);
        xml.push_str("\"><office:meta>");
        xml.push_str(&children);
        xml.push_str("</office:meta></office:document-meta>");
        let mut writer = crate::core::PackageWriter::new_bounded(
            MAX_PACKAGE_BYTES,
            crate::core::StreamingArchiveLimits::default(),
        );
        writer.set_mimetype(&package.mimetype()?)?;
        writer.add_file("meta.xml", xml.as_bytes())?;
        writer.copy_source_files_from_except(package, &std::collections::HashSet::new())?;
        Ok(Some(writer.finish_to_bounded_bytes()?))
    }
}

/// The `office:version` of the package's `content.xml` root, so a created
/// `meta.xml` declares the same version (`1.3` when absent).
fn content_version(package: &OwnedPackage) -> Result<String, Error> {
    const FALLBACK: &str = "1.3";
    if !package.has_file("content.xml")? {
        return Ok(FALLBACK.to_owned());
    }
    let content = package.get_file("content.xml")?;
    let mut reader = NsReader::from_reader(content.as_slice());
    let mut buffer = Vec::new();
    loop {
        let event = reader
            .read_event_into(&mut buffer)
            .map_err(|_| invalid("ODF content.xml is malformed"))?;
        match event {
            Event::Start(ref tag) | Event::Empty(ref tag) => {
                for attribute in tag.attributes().with_checks(false).flatten() {
                    let (resolved, local) = reader.resolver().resolve_attribute(attribute.key);
                    if local.as_ref() == b"version"
                        && matches!(resolved, ResolveResult::Bound(Namespace(uri)) if uri == OFFICE)
                    {
                        let value = std::str::from_utf8(&attribute.value)
                            .map_err(|_| invalid("ODF office:version is not UTF-8"))?;
                        if !value.is_empty()
                            && value
                                .bytes()
                                .all(|byte| byte.is_ascii_digit() || byte == b'.')
                        {
                            return Ok(value.to_owned());
                        }
                    }
                }
                return Ok(FALLBACK.to_owned());
            },
            Event::Eof => return Ok(FALLBACK.to_owned()),
            _ => buffer.clear(),
        }
    }
}

/// What [`Assignments::plan_publication`] publishes over one package.
#[non_exhaustive]
pub enum Publication {
    /// Nothing is staged, or every assignment matches the source.
    Unchanged,
    /// Element splices into the package's `meta.xml`.
    Splices(XmlSplicePublication),
    /// The package has no `meta.xml`.
    MissingPart,
}

/// A detached metadata edit of one source package. Dropping it discards
/// the staged assignments.
pub struct Edit<P> {
    source: P,
    assignments: Assignments,
    count: usize,
}

impl<P: Owner> Edit<P> {
    /// Start an edit of `source`; leaves expose it as
    /// `Package::edit_metadata`.
    #[must_use]
    pub fn new(source: &P) -> Self {
        Self {
            source: source.clone(),
            assignments: Assignments::default(),
            count: 0,
        }
    }

    /// Stage setting (`Some`) or removing (`None`) one property. A later
    /// assignment of the same field replaces an earlier one.
    ///
    /// # Errors
    ///
    /// Returns [`Error::Core`] for text XML cannot carry or a value over
    /// 1 MiB.
    pub fn set(&mut self, field: Field, value: Option<&str>) -> Result<&mut Self, Error> {
        self.assignments.set(field, value)?;
        self.count += 1;
        Ok(self)
    }

    /// Stage the keyword list, one `meta:keyword` element per keyword; an
    /// empty list removes every keyword.
    ///
    /// # Errors
    ///
    /// Returns [`Error::Core`] for text XML cannot carry or a keyword over
    /// 1 MiB.
    pub fn set_keywords(&mut self, keywords: &[String]) -> Result<&mut Self, Error> {
        self.assignments.set_keywords(keywords)?;
        self.count += 1;
        Ok(self)
    }

    /// Publish the staged properties as one new package.
    ///
    /// # Errors
    ///
    /// Returns [`Error::Core`] when the package has no `meta.xml`, the part
    /// is malformed, or a changed edit meets a signed or encrypted package;
    /// the leaf's ingress errors when it refuses the published candidate.
    pub fn commit(self) -> Result<Commit<P>, Error> {
        let package = self.source.metadata_source()?;
        let bytes = match self.assignments.plan_publication(&package)? {
            Publication::Splices(publication) => {
                refuse_envelope(&package)?;
                Some(rebuild_package_with_xml_splices(
                    &package,
                    vec![publication],
                    MAX_PACKAGE_BYTES,
                )?)
            },
            Publication::Unchanged => None,
            Publication::MissingPart => self.assignments.created_package(&package)?,
        };
        let Some(bytes) = bytes else {
            return Ok(Commit {
                package: self.source.clone(),
                patch: Patch {
                    source: self.source.clone(),
                    target: self.source,
                    changed: false,
                },
                diagnostics: Diagnostics {
                    changed: false,
                    fields: self.count,
                },
            });
        };
        let target = self.source.adopt_metadata(bytes)?;
        Ok(Commit {
            package: target.clone(),
            patch: Patch {
                source: self.source,
                target,
                changed: true,
            },
            diagnostics: Diagnostics {
                changed: true,
                fields: self.count,
            },
        })
    }
}

/// The result of one metadata commit.
pub struct Commit<P> {
    package: P,
    patch: Patch<P>,
    diagnostics: Diagnostics,
}

impl<P> Commit<P> {
    /// The committed package.
    #[must_use]
    pub const fn package(&self) -> &P {
        &self.package
    }

    /// The exact source-checked patch.
    #[must_use]
    pub const fn patch(&self) -> &Patch<P> {
        &self.patch
    }

    /// Content-free facts about the commit.
    #[must_use]
    pub const fn diagnostics(&self) -> Diagnostics {
        self.diagnostics
    }

    /// The committed package.
    #[must_use]
    pub fn into_package(self) -> P {
        self.package
    }

    /// The committed package and its patch.
    #[must_use]
    pub fn into_parts(self) -> (P, Patch<P>) {
        (self.package, self.patch)
    }
}

/// Content-free facts about one metadata commit.
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub struct Diagnostics {
    changed: bool,
    fields: usize,
}

impl Diagnostics {
    /// Whether the package bytes changed.
    #[must_use]
    pub const fn changed(&self) -> bool {
        self.changed
    }

    /// The number of field assignments staged.
    #[must_use]
    pub const fn field_count(&self) -> usize {
        self.fields
    }
}

/// An exact-byte metadata patch, applicable only to its source snapshot.
#[derive(Clone)]
pub struct Patch<P> {
    source: P,
    target: P,
    changed: bool,
}

impl<P: Owner> Patch<P> {
    /// Whether applying the patch preserves the source bytes.
    #[must_use]
    pub const fn is_noop(&self) -> bool {
        !self.changed
    }

    /// The patch that restores the source bytes.
    #[must_use]
    pub fn inverse(&self) -> Self {
        Self {
            source: self.target.clone(),
            target: self.source.clone(),
            changed: self.changed,
        }
    }

    /// Replay the patch onto `package`; leaves expose it as
    /// `Package::apply_metadata`. The commit publishes the patch's target
    /// package with the patch and its change flag.
    ///
    /// # Errors
    ///
    /// Returns [`Error::PatchConflict`] when `package` is not the patch's
    /// exact source snapshot.
    pub fn apply(&self, package: &P) -> Result<Commit<P>, Error> {
        if package.metadata_bytes() != self.source.metadata_bytes() {
            return Err(Error::patch_conflict(
                "metadata.apply",
                "the package is not the patch source",
            ));
        }
        Ok(Commit {
            package: self.target.clone(),
            patch: self.clone(),
            diagnostics: Diagnostics {
                changed: self.changed,
                fields: 0,
            },
        })
    }
}

impl<P> std::fmt::Debug for Patch<P> {
    fn fmt(&self, formatter: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        formatter
            .debug_struct("Patch")
            .field("noop", &!self.changed)
            .finish_non_exhaustive()
    }
}

impl<P> std::fmt::Debug for Edit<P> {
    fn fmt(&self, formatter: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        formatter
            .debug_struct("Edit")
            .field("fields", &self.count)
            .finish_non_exhaustive()
    }
}

impl<P> std::fmt::Debug for Commit<P> {
    fn fmt(&self, formatter: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        formatter
            .debug_struct("Commit")
            .field("diagnostics", &self.diagnostics)
            .finish_non_exhaustive()
    }
}

fn invalid(message: &str) -> Error {
    Error::Core(CoreError::InvalidFormat(message.to_owned()))
}

/// Refuse text no XML 1.0 document can carry.
fn check_text(text: &str) -> Result<(), Error> {
    if text.len() > MAX_VALUE_BYTES {
        return Err(invalid("an ODF metadata value exceeds 1 MiB"));
    }
    if text
        .chars()
        .any(|c| matches!(c as u32, 0..=0x8 | 0xb | 0xc | 0xe..=0x1f | 0xfffe | 0xffff))
    {
        return Err(invalid(
            "an ODF metadata value contains a character XML cannot carry",
        ));
    }
    Ok(())
}

/// A changed metadata write refuses signed and encrypted packages: their
/// envelope changes only through the explicit security verbs.
fn refuse_envelope(package: &OwnedPackage) -> Result<(), Error> {
    let archive = package.package()?;
    if archive
        .manifest()
        .entries
        .values()
        .any(|entry| entry.encryption.is_some())
        || package.has_zip_encrypted_entries()
    {
        return Err(Error::refused(
            litchi_core::edit::RefusalKind::UnsupportedOperation,
            "metadata.commit",
            "a changed metadata edit of an encrypted package is not supported",
        ));
    }
    if archive.has_file("META-INF/documentsignatures.xml")
        || archive.has_file("META-INF/macrosignatures.xml")
    {
        return Err(Error::Signed);
    }
    Ok(())
}

/// One located `office:meta` child.
struct Child {
    namespace: Vec<u8>,
    local: Vec<u8>,
    range: Range<usize>,
}

/// The splices (source range, replacement markup) the assignments compile
/// to over `source`; unchanged ranges are dropped.
fn plan(
    source: &[u8],
    fields: &BTreeMap<Field, Option<Vec<String>>>,
) -> Result<Vec<(Range<usize>, String)>, Error> {
    let full = std::str::from_utf8(source).map_err(|_| invalid("ODF meta.xml is not UTF-8"))?;
    // The reader skips a byte-order mark: plan in the text after it, then
    // shift the ranges back onto the part's bytes.
    let text = full.strip_prefix('\u{feff}').unwrap_or(full);
    let base = full.len() - text.len();
    let mut reader = NsReader::from_str(text);
    reader.config_mut().trim_text(false);
    let mut depth = 0usize;
    let mut meta: Option<(Range<usize>, bool)> = None;
    let mut meta_close = None;
    let mut children: Vec<Child> = Vec::new();
    let mut open_child: Option<(Vec<u8>, Vec<u8>, usize)> = None;
    let mut prefixes: [Option<String>; 2] = [None, None];
    loop {
        let start = usize::try_from(reader.buffer_position()).unwrap_or(usize::MAX);
        let (resolved, event) = reader
            .read_resolved_event()
            .map_err(|_| invalid("ODF meta.xml is malformed"))?;
        let namespace = match resolved {
            ResolveResult::Bound(Namespace(uri)) => uri.to_vec(),
            _ => Vec::new(),
        };
        let end = usize::try_from(reader.buffer_position()).unwrap_or(usize::MAX);
        match event {
            Event::Start(ref tag) | Event::Empty(ref tag) => {
                let empty = matches!(event, Event::Empty(_));
                let local = tag.local_name().as_ref().to_vec();
                if depth == 1 && meta.is_none() && namespace == OFFICE && local == b"meta" {
                    meta = Some((start..end, empty));
                    for (prefix, uri) in reader.resolver().bindings() {
                        let PrefixDeclaration::Named(prefix) = prefix else {
                            continue;
                        };
                        let slot = if uri.as_ref() == DC {
                            0
                        } else if uri.as_ref() == META {
                            1
                        } else {
                            continue;
                        };
                        prefixes[slot] = std::str::from_utf8(prefix).ok().map(str::to_owned);
                    }
                } else if depth == 2 && meta.is_some() && meta_close.is_none() {
                    if empty {
                        children.push(Child {
                            namespace,
                            local,
                            range: start..end,
                        });
                    } else {
                        open_child = Some((namespace, local, start));
                    }
                }
                if !empty {
                    depth += 1;
                }
            },
            Event::End(_) => {
                depth = depth.saturating_sub(1);
                if depth == 2 {
                    if let Some((namespace, local, begin)) = open_child.take() {
                        children.push(Child {
                            namespace,
                            local,
                            range: begin..end,
                        });
                    }
                }
                if depth == 1 && meta.is_some() && meta_close.is_none() {
                    meta_close = Some(start);
                }
            },
            Event::Eof => break,
            _ => {},
        }
    }
    let Some((meta_range, meta_empty)) = meta else {
        return Err(invalid("ODF meta.xml has no office:meta element"));
    };
    let mut splices: Vec<(Range<usize>, String)> = Vec::new();
    let mut appended = String::new();
    for (field, value) in fields {
        let (uri, local) = field.qualified();
        let existing: Vec<&Range<usize>> = children
            .iter()
            .filter(|child| child.namespace == uri && child.local == local.as_bytes())
            .map(|child| &child.range)
            .collect();
        let (prefix, declare) = match prefixes[usize::from(uri == META)].as_deref() {
            Some(prefix) => (prefix, false),
            None => (if uri == META { "meta" } else { "dc" }, true),
        };
        let mut rendered = String::new();
        if let Some(values) = value {
            for value in values {
                render(&mut rendered, prefix, local, declare.then_some(uri), value);
            }
        }
        match existing.split_first() {
            Some((first, rest)) => {
                splices.push(((*first).clone(), rendered));
                for range in rest {
                    splices.push(((*range).clone(), String::new()));
                }
            },
            None => appended.push_str(&rendered),
        }
    }
    if !appended.is_empty() {
        if meta_empty {
            let tag = &text[meta_range.clone()];
            let open = tag
                .strip_suffix("/>")
                .ok_or_else(|| invalid("ODF office:meta is not an empty element"))?
                .trim_end();
            let name_end = open
                .find(|c: char| c.is_ascii_whitespace())
                .unwrap_or(open.len());
            let name = &open[1..name_end];
            splices.push((meta_range, format!("{open}>{appended}</{name}>")));
        } else {
            let close = meta_close.ok_or_else(|| invalid("ODF office:meta is not closed"))?;
            splices.push((close..close, appended));
        }
    }
    splices.retain(|(range, bytes)| text.get(range.clone()) != Some(bytes.as_str()));
    splices.sort_by_key(|(range, _)| (range.start, range.end));
    for (range, _) in &mut splices {
        *range = range.start + base..range.end + base;
    }
    Ok(splices)
}

/// Write one property element.
fn render(out: &mut String, prefix: &str, local: &str, declare: Option<&[u8]>, value: &str) {
    out.push('<');
    out.push_str(prefix);
    out.push(':');
    out.push_str(local);
    if let Some(uri) = declare {
        out.push_str(" xmlns:");
        out.push_str(prefix);
        out.push_str("=\"");
        out.push_str(std::str::from_utf8(uri).unwrap_or_default());
        out.push('"');
    }
    if value.is_empty() {
        out.push_str("/>");
        return;
    }
    out.push('>');
    for character in value.chars() {
        match character {
            '&' => out.push_str("&amp;"),
            '<' => out.push_str("&lt;"),
            '>' => out.push_str("&gt;"),
            other => out.push(other),
        }
    }
    out.push_str("</");
    out.push_str(prefix);
    out.push(':');
    out.push_str(local);
    out.push('>');
}

#[cfg(test)]
mod tests {
    use super::*;

    fn fields(pairs: &[(Field, Option<&str>)]) -> BTreeMap<Field, Option<Vec<String>>> {
        let mut assignments = Assignments::default();
        for (field, value) in pairs {
            assignments.set(*field, *value).unwrap();
        }
        assignments.fields
    }

    const META_XML: &str = concat!(
        "<?xml version=\"1.0\" encoding=\"UTF-8\"?>\n",
        "<office:document-meta xmlns:office=\"urn:oasis:names:tc:opendocument:xmlns:office:1.0\" ",
        "xmlns:dc=\"http://purl.org/dc/elements/1.1/\" ",
        "xmlns:meta=\"urn:oasis:names:tc:opendocument:xmlns:meta:1.0\" office:version=\"1.3\">",
        "<office:meta><meta:generator>LO</meta:generator><dc:title>Old</dc:title>",
        "<meta:keyword>a</meta:keyword><meta:keyword>b</meta:keyword></office:meta>",
        "</office:document-meta>"
    );

    fn apply(source: &str, splices: &[(Range<usize>, String)]) -> String {
        let mut out = String::new();
        let mut cursor = 0;
        for (range, bytes) in splices {
            out.push_str(&source[cursor..range.start]);
            out.push_str(bytes);
            cursor = range.end;
        }
        out.push_str(&source[cursor..]);
        out
    }

    #[test]
    fn replaces_appends_and_removes_children() {
        let splices = plan(
            META_XML.as_bytes(),
            &fields(&[
                (Field::Title, Some("New & <x>")),
                (Field::Keywords, Some("c; d,")),
                (Field::Subject, Some("S")),
            ]),
        )
        .unwrap();
        let out = apply(META_XML, &splices);
        assert!(out.contains("<dc:title>New &amp; &lt;x&gt;</dc:title>"));
        assert!(out.contains("<meta:keyword>c</meta:keyword><meta:keyword>d</meta:keyword>"));
        assert!(!out.contains(">b<"));
        assert!(out.contains("<dc:subject>S</dc:subject></office:meta>"));
        assert!(out.starts_with("<?xml version=\"1.0\" encoding=\"UTF-8\"?>\n"));
        let removed = apply(
            META_XML,
            &plan(META_XML.as_bytes(), &fields(&[(Field::Title, None)])).unwrap(),
        );
        assert!(!removed.contains("dc:title"));
    }

    #[test]
    fn plans_on_part_bytes_after_a_byte_order_mark() {
        let source = format!("\u{feff}{META_XML}");
        let splices = plan(source.as_bytes(), &fields(&[(Field::Subject, Some("S"))])).unwrap();
        let out = apply(&source, &splices);
        assert!(
            out.contains("<dc:subject>S</dc:subject></office:meta>"),
            "{out}"
        );
    }

    #[test]
    fn unchanged_values_plan_nothing() {
        let splices = plan(META_XML.as_bytes(), &fields(&[(Field::Title, Some("Old"))])).unwrap();
        assert!(splices.is_empty());
    }

    #[test]
    fn expands_an_empty_meta_and_declares_missing_prefixes() {
        let source = concat!(
            "<office:document-meta xmlns:office=\"urn:oasis:names:tc:opendocument:xmlns:office:1.0\">",
            "<office:meta/></office:document-meta>"
        );
        let out = apply(
            source,
            &plan(source.as_bytes(), &fields(&[(Field::Author, Some("Ann"))])).unwrap(),
        );
        assert!(out.contains(
            "<office:meta><meta:initial-creator xmlns:meta=\"urn:oasis:names:tc:opendocument:xmlns:meta:1.0\">Ann</meta:initial-creator></office:meta>"
        ));
    }

    #[test]
    fn refuses_text_xml_cannot_carry() {
        let mut assignments = Assignments::default();
        assert!(assignments.set(Field::Title, Some("a\u{1}b")).is_err());
        assert!(assignments.is_empty());
    }
}
