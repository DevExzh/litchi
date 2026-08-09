//! Source-bound host for `PowerPoint` 2020 Designer shape properties.
//!
//! The `p202:designPr` payload belongs below exactly one selected shape's
//! direct `p:nvPr` extension list.  This module deliberately scans the whole
//! slide, rather than a detached shape fragment, so inherited namespace
//! prefixes are proven before the payload codec accepts them.

use std::ops::Range;
use std::sync::Arc;

use litchi_opc::constants::content_type as ct;
use litchi_opc::{OpcPackage, PackURI};
use quick_xml::events::{BytesStart, Event};
use quick_xml::name::{Namespace, ResolveResult};
use quick_xml::reader::NsReader;

use super::{DrawingProperties, Limits, P202_NAMESPACE, PROPERTIES_EXTENSION_URI};
use crate::{Error, Result};

const PML_TRANSITIONAL: &[u8] = b"http://schemas.openxmlformats.org/presentationml/2006/main";
const PML_STRICT: &[u8] = b"http://purl.oclc.org/ooxml/presentationml/main";

/// A selected shape's source-bound optional `p202:designPr` value.
///
/// A snapshot owns a bounded copy of its containing slide.  It is therefore
/// intentionally not cloneable: a commit can only be published against the
/// exact source from which it was read.
#[derive(Debug)]
pub struct PropertiesSnapshot {
    owner: PackURI,
    source: Arc<Vec<u8>>,
    shape: Range<usize>,
    occurrences: Vec<PropertySource>,
    limits: Limits,
}

#[derive(Debug)]
struct PropertySource {
    value: Option<DrawingProperties>,
    inner_extensions: Option<Vec<u8>>,
}

impl PropertiesSnapshot {
    /// Borrow the selected shape's optional typed properties.
    #[inline]
    #[must_use]
    pub fn properties(&self) -> Option<&DrawingProperties> {
        match self.occurrences.as_slice() {
            [occurrence] => occurrence.value.as_ref(),
            _ => None,
        }
    }

    /// Return whether any matching extension has a typed `designPr` payload.
    #[inline]
    #[must_use]
    pub fn is_present(&self) -> bool {
        self.occurrences
            .iter()
            .any(|occurrence| occurrence.value.is_some())
    }

    /// Return the number of matching outer Designer property extensions.
    ///
    /// More than one occurrence is preserved for inspection, but singular
    /// mutation is refused because the intended owner would be ambiguous.
    #[inline]
    #[must_use]
    pub fn occurrence_count(&self) -> usize {
        self.occurrences.len()
    }

    /// Iterate over matching outer extensions in document order.
    ///
    /// An item is `None` when the matching extension has no correctly
    /// namespace-qualified `p202:designPr` payload.
    #[must_use]
    pub fn occurrences(&self) -> impl ExactSizeIterator<Item = Option<&DrawingProperties>> + '_ {
        self.occurrences
            .iter()
            .map(|occurrence| occurrence.value.as_ref())
    }

    /// Start a move-only edit tied to this exact slide source.
    #[inline]
    #[must_use]
    pub fn edit(self) -> PropertiesEdit {
        PropertiesEdit {
            snapshot: self,
            state: EditState::Unchanged,
            changed: false,
        }
    }
}

#[derive(Debug)]
enum EditState {
    Unchanged,
    Set(DrawingProperties),
    Remove,
}

/// A move-only edit of one shape's optional Designer properties.
#[derive(Debug)]
pub struct PropertiesEdit {
    snapshot: PropertiesSnapshot,
    state: EditState,
    changed: bool,
}

impl PropertiesEdit {
    /// Borrow the currently projected optional properties.
    #[inline]
    #[must_use]
    pub fn properties(&self) -> Option<&DrawingProperties> {
        match &self.state {
            EditState::Unchanged => self.snapshot.properties(),
            EditState::Set(value) => Some(value),
            EditState::Remove => None,
        }
    }

    /// Replace the optional properties after validating them under this
    /// source's finite limits.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation fails.
    pub fn set(&mut self, value: DrawingProperties) -> Result<()> {
        validate_properties(&value, self.snapshot.limits)?;
        self.changed =
            self.snapshot.occurrence_count() != 1 || self.snapshot.properties() != Some(&value);
        self.state = EditState::Set(value);
        Ok(())
    }

    /// Remove the `designPr` payload, leaving all unrelated extension markup.
    pub fn remove(&mut self) {
        self.changed = self.snapshot.is_present();
        self.state = EditState::Remove;
    }

    /// Return whether this edit will change the selected source.
    #[inline]
    #[must_use]
    pub fn is_changed(&self) -> bool {
        self.changed
    }

    /// Consume this edit into a move-only source-bound commit.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation fails.
    pub fn commit(self) -> Result<PropertiesCommit> {
        if self.changed && self.snapshot.occurrence_count() > 1 {
            return Err(Error::UnsafeEdit {
                operation: "commit_shape_designer_properties",
                reason: "the selected shape has multiple matching Designer properties extensions",
            });
        }
        if let EditState::Set(value) = &self.state {
            validate_properties(value, self.snapshot.limits)?;
        }
        let value = match self.state {
            EditState::Unchanged | EditState::Remove => None,
            EditState::Set(value) => Some(value),
        };
        Ok(PropertiesCommit {
            snapshot: self.snapshot,
            value,
            changed: self.changed,
        })
    }
}

/// A move-only staged `designPr` publication.
#[derive(Debug)]
pub struct PropertiesCommit {
    snapshot: PropertiesSnapshot,
    value: Option<DrawingProperties>,
    changed: bool,
}

impl PropertiesCommit {
    /// Return whether publishing this commit is a byte-preserving no-op.
    #[inline]
    #[must_use]
    pub fn is_noop(&self) -> bool {
        !self.changed
    }
}

/// Load source-bound properties for one selected shape with default limits.
pub(crate) fn load_properties<'k>(
    package: &OpcPackage,
    owner: &PackURI,
    key: impl Into<crate::shape::Key<'k>>,
) -> Result<PropertiesSnapshot> {
    load_properties_with_limits(package, owner, key, Limits::default())
}

/// Load source-bound properties under explicit finite limits.
pub(crate) fn load_properties_with_limits<'k>(
    package: &OpcPackage,
    owner: &PackURI,
    key: impl Into<crate::shape::Key<'k>>,
    limits: Limits,
) -> Result<PropertiesSnapshot> {
    let part = package.get_part(owner)?;
    crate::parts::validate_content_type(part, ct::PML_SLIDE)?;
    let source = copy_owner(part.blob(), limits)?;
    let shape = crate::tag::shape::selected_raw_span(&source, key.into())?;
    let located = locate(&source, shape.clone(), limits)?;
    let occurrences = parse_properties(&source, &located, limits)?;
    Ok(PropertiesSnapshot {
        owner: part.partname().clone(),
        source,
        shape,
        occurrences,
        limits,
    })
}

/// Apply a move-only commit atomically to its originating slide.
pub(crate) fn apply_properties(
    package: &mut OpcPackage,
    commit: PropertiesCommit,
) -> Result<PropertiesSnapshot> {
    let PropertiesCommit {
        snapshot,
        value,
        changed,
    } = commit;
    if !changed {
        validate_current_source(package, &snapshot)?;
        return Ok(snapshot);
    }
    if let Some(value) = value.as_ref() {
        validate_properties(value, snapshot.limits)?;
    }
    let original = snapshot.source.as_slice();
    let staged = stage(
        original,
        &snapshot.shape,
        value.as_ref(),
        &snapshot,
        snapshot.limits,
    )?;
    let shape_end = snapshot
        .shape
        .end
        .checked_add(staged.len())
        .and_then(|value| value.checked_sub(original.len()))
        .ok_or_else(|| {
            limit(
                "Designer properties shape bytes",
                snapshot.limits.xml_bytes(),
            )
        })?;
    let staged_shape = snapshot.shape.start..shape_end;
    let located = locate(&staged, staged_shape.clone(), snapshot.limits)?;
    let occurrences = parse_properties(&staged, &located, snapshot.limits)?;
    let parsed = match occurrences.as_slice() {
        [occurrence] => occurrence.value.as_ref(),
        _ => None,
    };
    if parsed != value.as_ref() {
        return Err(Error::Invalid(
            "staged Designer properties did not round-trip their typed value".into(),
        ));
    }

    let part = package.get_part_mut(&snapshot.owner)?;
    crate::parts::validate_content_type(part, ct::PML_SLIDE)?;
    if part.blob() != original {
        return Err(Error::UnsafeEdit {
            operation: "apply_shape_designer_properties",
            reason: "the selected slide changed during Designer properties staging",
        });
    }
    if staged == original {
        return Ok(snapshot);
    }
    let staged = Arc::new(staged);
    part.set_blob_shared(Arc::clone(&staged));
    package.unsign();
    Ok(PropertiesSnapshot {
        owner: snapshot.owner,
        source: staged,
        shape: staged_shape,
        occurrences,
        limits: snapshot.limits,
    })
}

fn validate_current_source(package: &OpcPackage, snapshot: &PropertiesSnapshot) -> Result<()> {
    let part = package.get_part(&snapshot.owner)?;
    crate::parts::validate_content_type(part, ct::PML_SLIDE)?;
    if part.blob() != snapshot.source.as_slice() {
        return Err(Error::UnsafeEdit {
            operation: "apply_shape_designer_properties",
            reason: "the selected slide changed after Designer properties were loaded",
        });
    }
    Ok(())
}

/// Directly replace or create selected shape Designer properties.
#[allow(
    dead_code,
    reason = "kept as the crate-private direct publication wrapper"
)]
pub(crate) fn put_properties<'k>(
    package: &mut OpcPackage,
    owner: &PackURI,
    key: impl Into<crate::shape::Key<'k>>,
    value: DrawingProperties,
) -> Result<PropertiesSnapshot> {
    let mut edit = load_properties(package, owner, key)?.edit();
    edit.set(value)?;
    apply_properties(package, edit.commit()?)
}

/// Directly remove selected shape Designer properties.
#[allow(
    dead_code,
    reason = "kept as the crate-private direct publication wrapper"
)]
pub(crate) fn remove_properties<'k>(
    package: &mut OpcPackage,
    owner: &PackURI,
    key: impl Into<crate::shape::Key<'k>>,
) -> Result<PropertiesSnapshot> {
    let mut edit = load_properties(package, owner, key)?.edit();
    edit.remove();
    apply_properties(package, edit.commit()?)
}

#[derive(Debug, Clone)]
struct Element {
    span: Range<usize>,
    close_start: usize,
    empty: bool,
    prefix: Option<Vec<u8>>,
    pml: Option<PmlNamespace>,
}

#[derive(Debug, Default)]
struct Located {
    nv_pr: Option<Element>,
    ext_lst: Option<Element>,
    ext_other: bool,
    property_exts: Vec<PropertyOccurrence>,
}

#[derive(Debug, Default)]
struct PropertyOccurrence {
    outer: Option<Element>,
    other: bool,
    properties: Option<Element>,
}

#[derive(Clone, Copy, Debug, PartialEq, Eq)]
enum PmlNamespace {
    Transitional,
    Strict,
}

impl PmlNamespace {
    fn as_bytes(self) -> &'static [u8] {
        match self {
            Self::Transitional => PML_TRANSITIONAL,
            Self::Strict => PML_STRICT,
        }
    }
}

#[derive(Clone, Copy, Debug, PartialEq, Eq)]
enum Role {
    Other,
    Target,
    NonVisual,
    NvPr,
    ExtList,
    PropertyExt(usize),
    Properties(usize),
}

#[derive(Debug)]
struct Frame {
    role: Role,
    start: usize,
    empty: bool,
    prefix: Option<Vec<u8>>,
    pml: Option<PmlNamespace>,
}

fn locate(xml: &[u8], shape: Range<usize>, limits: Limits) -> Result<Located> {
    if shape.start >= shape.end || shape.end > xml.len() {
        return Err(Error::Invalid(
            "Designer properties shape span is invalid".into(),
        ));
    }
    if xml.len() > limits.xml_bytes() {
        return Err(limit("Designer slide XML bytes", limits.xml_bytes()));
    }
    let mut reader = NsReader::from_reader(xml);
    reader.config_mut().trim_text(false);
    let mut stack = Vec::<Frame>::new();
    stack
        .try_reserve_exact(32)
        .map_err(|source| Error::Allocation {
            resource: "Designer properties XML frames",
            source,
        })?;
    let mut located = Located::default();
    let mut events = 0usize;
    loop {
        let start = position(&reader)?;
        let decoder = reader.decoder();
        let (namespace, event) = reader.read_resolved_event().map_err(xml_error)?;
        let pml = pml_namespace(&namespace);
        let p202 = is_p202(&namespace);
        let event = event.into_owned();
        let empty = matches!(&event, Event::Empty(_));
        let end = position(&reader)?;
        if !matches!(&event, Event::Eof) {
            events = events
                .checked_add(1)
                .ok_or_else(|| limit("Designer slide XML events", limits.xml_nodes()))?;
            if events > limits.xml_nodes() {
                return Err(limit("Designer slide XML events", limits.xml_nodes()));
            }
        }
        match event {
            Event::Start(element) | Event::Empty(element) => {
                if stack.len() >= limits.xml_depth() {
                    return Err(limit("Designer slide XML depth", limits.xml_depth()));
                }
                let role = start_role(
                    &stack,
                    start,
                    &shape,
                    pml,
                    p202,
                    &element,
                    decoder,
                    &mut located,
                    limits,
                )?;
                let prefix = element
                    .name()
                    .prefix()
                    .map(quick_xml::name::Prefix::into_inner)
                    .map(|value| copy(value, "Designer namespace prefix"))
                    .transpose()?;
                let frame = Frame {
                    role,
                    start,
                    empty,
                    prefix,
                    pml,
                };
                if empty {
                    finish(xml, frame, end, &mut located)?;
                } else {
                    stack.push(frame);
                }
            },
            Event::End(_) => {
                let frame = stack.pop().ok_or_else(|| {
                    Error::Invalid("Designer properties XML stack underflow".into())
                })?;
                finish(xml, frame, end, &mut located)?;
            },
            Event::Text(text) => {
                if !text.decode().map_err(xml_error)?.trim().is_empty() {
                    mark_direct_other(&stack, &mut located);
                }
            },
            Event::Eof => break,
            Event::Comment(_) | Event::PI(_) => mark_direct_other(&stack, &mut located),
            Event::Decl(_) => {},
            Event::CData(_) | Event::DocType(_) | Event::GeneralRef(_) => {
                return Err(Error::Invalid(
                    "Designer properties slide contains forbidden markup".into(),
                ));
            },
        }
    }
    if !stack.is_empty() {
        return Err(Error::Invalid(
            "Designer properties slide XML is unterminated".into(),
        ));
    }
    if located.nv_pr.is_none() {
        return Err(Error::Invalid(
            "selected shape has no direct p:nvPr host".into(),
        ));
    }
    Ok(located)
}

fn start_role(
    stack: &[Frame],
    start: usize,
    shape: &Range<usize>,
    pml: Option<PmlNamespace>,
    p202: bool,
    element: &BytesStart<'_>,
    decoder: quick_xml::encoding::Decoder,
    located: &mut Located,
    limits: Limits,
) -> Result<Role> {
    let local = element.local_name();
    let parent = stack.last().map_or(Role::Other, |frame| frame.role);
    if start == shape.start {
        if pml.is_none()
            || !matches!(
                local.as_ref(),
                b"sp" | b"pic" | b"cxnSp" | b"graphicFrame" | b"grpSp"
            )
        {
            return Err(Error::Invalid(
                "selected Designer properties owner is not a supported shape".into(),
            ));
        }
        return Ok(Role::Target);
    }
    let role = match (parent, pml.is_some(), p202, local.as_ref()) {
        (
            Role::Target,
            true,
            _,
            b"nvSpPr" | b"nvPicPr" | b"nvCxnSpPr" | b"nvGraphicFramePr" | b"nvGrpSpPr",
        ) => Role::NonVisual,
        (Role::NonVisual, true, _, b"nvPr") => Role::NvPr,
        (Role::NvPr, true, _, b"extLst") => {
            if has_extra_attributes(element, None)? {
                located.ext_other = true;
            }
            Role::ExtList
        },
        (Role::ExtList, true, _, b"ext") => {
            let uri = extension_uri(element, decoder, limits)?;
            if uri.as_deref() == Some(PROPERTIES_EXTENSION_URI) {
                located
                    .property_exts
                    .try_reserve(1)
                    .map_err(|source| Error::Allocation {
                        resource: "Designer properties extension inventory",
                        source,
                    })?;
                let index = located.property_exts.len();
                located.property_exts.push(PropertyOccurrence {
                    other: has_extra_attributes(element, Some(b"uri"))?,
                    ..PropertyOccurrence::default()
                });
                Role::PropertyExt(index)
            } else {
                located.ext_other = true;
                Role::Other
            }
        },
        (Role::PropertyExt(index), false, true, b"designPr") => {
            if located.property_exts[index].properties.is_some() {
                return Err(Error::Invalid(
                    "Designer properties extension has duplicate designPr elements".into(),
                ));
            }
            Role::Properties(index)
        },
        (Role::PropertyExt(index), _, _, _) => {
            located.property_exts[index].other = true;
            Role::Other
        },
        (Role::ExtList, _, _, _) => {
            located.ext_other = true;
            Role::Other
        },
        _ => Role::Other,
    };
    Ok(role)
}

fn finish(xml: &[u8], frame: Frame, end: usize, located: &mut Located) -> Result<()> {
    if frame.start >= end || end > xml.len() {
        return Err(Error::Invalid(
            "Designer properties element range is invalid".into(),
        ));
    }
    let element = Element {
        span: frame.start..end,
        close_start: if frame.empty {
            end
        } else {
            close_start(xml, end)
        },
        empty: frame.empty,
        prefix: frame.prefix,
        pml: frame.pml,
    };
    match frame.role {
        Role::NvPr => {
            if located.nv_pr.replace(element).is_some() {
                return Err(Error::Invalid(
                    "selected shape has duplicate direct p:nvPr hosts".into(),
                ));
            }
        },
        Role::ExtList => {
            if located.ext_lst.replace(element).is_some() {
                return Err(Error::Invalid(
                    "selected shape has duplicate direct p:extLst hosts".into(),
                ));
            }
        },
        Role::PropertyExt(index) => located.property_exts[index].outer = Some(element),
        Role::Properties(index) => located.property_exts[index].properties = Some(element),
        _ => {},
    }
    Ok(())
}

fn parse_properties(xml: &[u8], located: &Located, limits: Limits) -> Result<Vec<PropertySource>> {
    let mut result = Vec::new();
    result
        .try_reserve_exact(located.property_exts.len())
        .map_err(|source| Error::Allocation {
            resource: "Designer properties typed occurrence inventory",
            source,
        })?;
    for occurrence in &located.property_exts {
        let Some(element) = occurrence.properties.as_ref() else {
            result.push(PropertySource {
                value: None,
                inner_extensions: None,
            });
            continue;
        };
        let bytes = xml
            .get(element.span.clone())
            .ok_or_else(|| Error::Invalid("Designer properties range is invalid".into()))?;
        let source =
            super::p202::read_properties_with_prefix(bytes, limits, element.prefix.as_deref())?;
        result.push(PropertySource {
            value: Some(source.value),
            inner_extensions: source.inner_extensions,
        });
    }
    Ok(result)
}

fn stage(
    source: &[u8],
    shape: &Range<usize>,
    value: Option<&DrawingProperties>,
    snapshot: &PropertiesSnapshot,
    limits: Limits,
) -> Result<Vec<u8>> {
    let located = locate(source, shape.clone(), limits)?;
    if located.property_exts.len() > 1 {
        return Err(Error::UnsafeEdit {
            operation: "stage_shape_designer_properties",
            reason: "the selected shape has multiple matching Designer properties extensions",
        });
    }
    let occurrence = located.property_exts.first();
    let current = occurrence.and_then(|occurrence| occurrence.properties.as_ref());
    let inner_extensions = snapshot
        .occurrences
        .first()
        .and_then(|occurrence| occurrence.inner_extensions.as_deref());
    match (current, value) {
        (Some(current), Some(value)) => {
            let replacement = super::p202::write_properties(value, inner_extensions, limits)?;
            replace(source, current.span.clone(), &replacement, limits)
        },
        (None, Some(value)) => insert_properties(source, &located, value, inner_extensions, limits),
        (None, None) => fallible_copy_vec(source, "Designer no-op output"),
        (Some(current), None) => {
            let occurrence = occurrence.ok_or(Error::UnsafeEdit {
                operation: "stage_shape_designer_properties",
                reason: "current Designer properties have no outer occurrence",
            })?;
            remove_properties_xml(source, &located, occurrence, current, limits)
        },
    }
}

fn insert_properties(
    source: &[u8],
    located: &Located,
    value: &DrawingProperties,
    inner: Option<&[u8]>,
    limits: Limits,
) -> Result<Vec<u8>> {
    let payload = super::p202::write_properties(value, inner, limits)?;
    if let Some(ext) = located
        .property_exts
        .first()
        .and_then(|occurrence| occurrence.outer.as_ref())
    {
        return insert_inside(source, ext, &payload, limits);
    }
    if let Some(ext_lst) = located.ext_lst.as_ref() {
        let extension = write_extension(&payload, ext_lst, limits)?;
        return insert_inside(source, ext_lst, &extension, limits);
    }
    let nv_pr = located
        .nv_pr
        .as_ref()
        .ok_or_else(|| Error::Invalid("selected shape has no direct p:nvPr host".into()))?;
    let extension = write_extension(&payload, nv_pr, limits)?;
    let list = write_extension_list(&extension, nv_pr, limits)?;
    insert_inside(source, nv_pr, &list, limits)
}

fn remove_properties_xml(
    source: &[u8],
    located: &Located,
    occurrence: &PropertyOccurrence,
    properties: &Element,
    limits: Limits,
) -> Result<Vec<u8>> {
    if occurrence.other {
        return replace(source, properties.span.clone(), &[], limits);
    }
    let extension = occurrence
        .outer
        .as_ref()
        .ok_or_else(|| Error::Invalid("Designer properties extension disappeared".into()))?;
    if located.ext_other || located.property_exts.len() > 1 {
        return replace(source, extension.span.clone(), &[], limits);
    }
    let ext_lst = located
        .ext_lst
        .as_ref()
        .ok_or_else(|| Error::Invalid("Designer properties extension list disappeared".into()))?;
    replace(source, ext_lst.span.clone(), &[], limits)
}

fn insert_inside(
    source: &[u8],
    element: &Element,
    child: &[u8],
    limits: Limits,
) -> Result<Vec<u8>> {
    if element.empty {
        let raw = source.get(element.span.clone()).ok_or_else(|| {
            Error::Invalid("Designer properties empty host range is invalid".into())
        })?;
        let slash = raw
            .iter()
            .rposition(|byte| *byte == b'/')
            .ok_or_else(|| Error::Invalid("Designer properties empty host is malformed".into()))?;
        let qname = raw
            .get(
                1..raw[1..]
                    .iter()
                    .position(|byte| byte.is_ascii_whitespace() || matches!(*byte, b'/' | b'>'))
                    .map(|offset| offset + 1)
                    .ok_or_else(|| {
                        Error::Invalid("Designer properties host name is unterminated".into())
                    })?,
            )
            .ok_or_else(|| Error::Invalid("Designer properties host name is invalid".into()))?;
        let mut replacement = Vec::new();
        let length = raw
            .len()
            .checked_sub(1)
            .and_then(|value| value.checked_add(child.len()))
            .and_then(|value| value.checked_add(qname.len()))
            .and_then(|value| value.checked_add(3))
            .ok_or_else(|| limit("Designer properties output bytes", limits.xml_bytes()))?;
        if length > limits.xml_bytes() {
            return Err(limit(
                "Designer properties output bytes",
                limits.xml_bytes(),
            ));
        }
        replacement
            .try_reserve_exact(length)
            .map_err(|source| Error::Allocation {
                resource: "Designer properties empty-host expansion",
                source,
            })?;
        replacement.extend_from_slice(&raw[..slash]);
        replacement.extend_from_slice(&raw[slash + 1..]);
        replacement.extend_from_slice(child);
        replacement.extend_from_slice(b"</");
        replacement.extend_from_slice(qname);
        replacement.push(b'>');
        return replace(source, element.span.clone(), &replacement, limits);
    }
    replace(
        source,
        element.close_start..element.close_start,
        child,
        limits,
    )
}

fn write_extension(payload: &[u8], host: &Element, limits: Limits) -> Result<Vec<u8>> {
    let namespace = host
        .pml
        .ok_or_else(|| Error::Invalid("Designer properties host has no PML namespace".into()))?;
    let qname_len = host.prefix.as_ref().map_or(3, |prefix| prefix.len() + 4);
    let declaration_prefix_len = host.prefix.as_ref().map_or(0, |prefix| prefix.len() + 1);
    let mut output = Vec::new();
    let size = qname_len
        .checked_add(qname_len)
        .and_then(|value| value.checked_add(PROPERTIES_EXTENSION_URI.len()))
        .and_then(|value| value.checked_add(namespace.as_bytes().len()))
        .and_then(|value| value.checked_add(payload.len()))
        .and_then(|value| value.checked_add(declaration_prefix_len))
        .and_then(|value| value.checked_add(21))
        .ok_or_else(|| limit("Designer properties output bytes", limits.xml_bytes()))?;
    if size > limits.xml_bytes() {
        return Err(limit(
            "Designer properties output bytes",
            limits.xml_bytes(),
        ));
    }
    output
        .try_reserve_exact(size)
        .map_err(|source| Error::Allocation {
            resource: "Designer properties extension",
            source,
        })?;
    output.push(b'<');
    write_qname(&mut output, host.prefix.as_deref(), b"ext");
    output.extend_from_slice(b" xmlns");
    if let Some(prefix) = host.prefix.as_deref() {
        output.push(b':');
        output.extend_from_slice(prefix);
    }
    output.extend_from_slice(b"=\"");
    output.extend_from_slice(namespace.as_bytes());
    output.extend_from_slice(b"\" uri=\"");
    output.extend_from_slice(PROPERTIES_EXTENSION_URI.as_bytes());
    output.extend_from_slice(b"\">");
    output.extend_from_slice(payload);
    output.extend_from_slice(b"</");
    write_qname(&mut output, host.prefix.as_deref(), b"ext");
    output.push(b'>');
    Ok(output)
}

fn write_extension_list(extension: &[u8], host: &Element, limits: Limits) -> Result<Vec<u8>> {
    let namespace = host
        .pml
        .ok_or_else(|| Error::Invalid("Designer properties host has no PML namespace".into()))?;
    let qname_len = host.prefix.as_ref().map_or(6, |prefix| prefix.len() + 7);
    let declaration_prefix_len = host.prefix.as_ref().map_or(0, |prefix| prefix.len() + 1);
    let size = qname_len
        .checked_add(qname_len)
        .and_then(|value| value.checked_add(namespace.as_bytes().len()))
        .and_then(|value| value.checked_add(extension.len()))
        .and_then(|value| value.checked_add(declaration_prefix_len))
        .and_then(|value| value.checked_add(14))
        .ok_or_else(|| limit("Designer properties output bytes", limits.xml_bytes()))?;
    if size > limits.xml_bytes() {
        return Err(limit(
            "Designer properties output bytes",
            limits.xml_bytes(),
        ));
    }
    let mut output = Vec::new();
    output
        .try_reserve_exact(size)
        .map_err(|source| Error::Allocation {
            resource: "Designer properties extension list",
            source,
        })?;
    output.push(b'<');
    write_qname(&mut output, host.prefix.as_deref(), b"extLst");
    output.extend_from_slice(b" xmlns");
    if let Some(prefix) = host.prefix.as_deref() {
        output.push(b':');
        output.extend_from_slice(prefix);
    }
    output.extend_from_slice(b"=\"");
    output.extend_from_slice(namespace.as_bytes());
    output.extend_from_slice(b"\">");
    output.extend_from_slice(extension);
    output.extend_from_slice(b"</");
    write_qname(&mut output, host.prefix.as_deref(), b"extLst");
    output.push(b'>');
    Ok(output)
}

fn write_qname(output: &mut Vec<u8>, prefix: Option<&[u8]>, local: &[u8]) {
    if let Some(prefix) = prefix {
        output.extend_from_slice(prefix);
        output.push(b':');
    }
    output.extend_from_slice(local);
}

fn replace(
    source: &[u8],
    range: Range<usize>,
    replacement: &[u8],
    limits: Limits,
) -> Result<Vec<u8>> {
    let before = source
        .get(..range.start)
        .ok_or_else(|| Error::Invalid("Designer properties replacement range is invalid".into()))?;
    let after = source
        .get(range.end..)
        .ok_or_else(|| Error::Invalid("Designer properties replacement range is invalid".into()))?;
    let len = before
        .len()
        .checked_add(replacement.len())
        .and_then(|value| value.checked_add(after.len()))
        .ok_or_else(|| limit("Designer properties output bytes", limits.xml_bytes()))?;
    if len > limits.xml_bytes() {
        return Err(limit(
            "Designer properties output bytes",
            limits.xml_bytes(),
        ));
    }
    let mut output = Vec::new();
    output
        .try_reserve_exact(len)
        .map_err(|source| Error::Allocation {
            resource: "Designer properties output",
            source,
        })?;
    output.extend_from_slice(before);
    output.extend_from_slice(replacement);
    output.extend_from_slice(after);
    Ok(output)
}

fn extension_uri(
    element: &BytesStart<'_>,
    decoder: quick_xml::encoding::Decoder,
    limits: Limits,
) -> Result<Option<String>> {
    let mut result = None;
    for attribute in element.attributes().with_checks(true) {
        let attribute = attribute.map_err(xml_error)?;
        if attribute.key.as_ref() == b"uri" {
            if result.is_some() {
                return Err(Error::Invalid(
                    "Designer properties extension has duplicate uri attributes".into(),
                ));
            }
            let value = attribute
                .decoded_and_normalized_value(quick_xml::XmlVersion::Explicit1_0, decoder)
                .map_err(xml_error)?;
            if value.len() > limits.attribute_bytes() {
                return Err(limit(
                    "Designer properties extension URI bytes",
                    limits.attribute_bytes(),
                ));
            }
            result = Some(value.into_owned());
        }
    }
    Ok(result)
}

fn has_extra_attributes(element: &BytesStart<'_>, allowed: Option<&[u8]>) -> Result<bool> {
    for attribute in element.attributes().with_checks(true) {
        let attribute = attribute.map_err(xml_error)?;
        let key = attribute.key.as_ref();
        if is_namespace_declaration(key) || allowed.is_some_and(|allowed| key == allowed) {
            continue;
        }
        return Ok(true);
    }
    Ok(false)
}

fn is_namespace_declaration(key: &[u8]) -> bool {
    key == b"xmlns" || key.starts_with(b"xmlns:")
}

fn mark_direct_other(stack: &[Frame], located: &mut Located) {
    match stack.last().map(|frame| frame.role) {
        Some(Role::ExtList) => located.ext_other = true,
        Some(Role::PropertyExt(index)) => located.property_exts[index].other = true,
        _ => {},
    }
}

fn validate_properties(value: &DrawingProperties, limits: Limits) -> Result<()> {
    if let Some(tags) = value.tags() {
        tags.validate(limits)?;
    }
    Ok(())
}

fn copy_owner(value: &[u8], limits: Limits) -> Result<Arc<Vec<u8>>> {
    if value.len() > limits.xml_bytes() {
        return Err(limit("Designer slide XML bytes", limits.xml_bytes()));
    }
    Ok(Arc::new(fallible_copy_vec(value, "Designer slide source")?))
}

fn fallible_copy_vec(value: &[u8], resource: &'static str) -> Result<Vec<u8>> {
    let mut copy = Vec::new();
    copy.try_reserve_exact(value.len())
        .map_err(|source| Error::Allocation { resource, source })?;
    copy.extend_from_slice(value);
    Ok(copy)
}

fn copy(value: &[u8], resource: &'static str) -> Result<Vec<u8>> {
    fallible_copy_vec(value, resource)
}
fn position(reader: &NsReader<&[u8]>) -> Result<usize> {
    usize::try_from(reader.buffer_position())
        .map_err(|_err| Error::Invalid("Designer properties XML offset does not fit usize".into()))
}
fn close_start(xml: &[u8], end: usize) -> usize {
    xml[..end]
        .iter()
        .rposition(|byte| *byte == b'<')
        .unwrap_or(end)
}
fn pml_namespace(namespace: &ResolveResult<'_>) -> Option<PmlNamespace> {
    match namespace {
        ResolveResult::Bound(Namespace(value)) if *value == PML_TRANSITIONAL => {
            Some(PmlNamespace::Transitional)
        },
        ResolveResult::Bound(Namespace(value)) if *value == PML_STRICT => {
            Some(PmlNamespace::Strict)
        },
        _ => None,
    }
}
fn is_p202(namespace: &ResolveResult<'_>) -> bool {
    matches!(namespace, ResolveResult::Bound(Namespace(value)) if *value == P202_NAMESPACE.as_bytes())
}
fn xml_error(error: impl std::fmt::Display) -> Error {
    Error::Xml(error.to_string())
}
fn limit(resource: &'static str, limit: usize) -> Error {
    Error::Limit { resource, limit }
}

#[cfg(test)]
#[allow(
    clippy::unwrap_used,
    clippy::expect_used,
    reason = "test assertions panic on failure by design"
)]
mod tests {
    use super::*;
    use litchi_opc::XmlPart;

    const PML: &str = "http://schemas.openxmlformats.org/presentationml/2006/main";
    fn package(xml: String) -> (OpcPackage, PackURI) {
        let name = PackURI::new("/ppt/slides/slide1.xml").unwrap();
        let mut package = OpcPackage::new();
        package.add_part(Box::new(XmlPart::new(
            name.clone(),
            ct::PML_SLIDE.into(),
            xml.into_bytes(),
        )));
        (package, name)
    }
    fn slide(body: &str) -> String {
        slide_with_namespace(PML, body)
    }
    fn slide_with_namespace(namespace: &str, body: &str) -> String {
        format!(
            r#"<p:sld xmlns:p="{namespace}" xmlns:p202="{P202_NAMESPACE}"><p:cSld><p:spTree><p:nvGrpSpPr/><p:grpSpPr/>{body}</p:spTree></p:cSld></p:sld>"#
        )
    }
    fn shape(body: &str) -> String {
        format!(
            r#"<p:sp><p:nvSpPr><p:cNvPr id="1" name="Box"/><p:cNvSpPr/><p:nvPr>{body}</p:nvPr></p:nvSpPr><p:spPr/></p:sp>"#
        )
    }

    #[test]
    fn edits_removes_and_preserves_siblings() {
        let body = format!(
            r#"<p:extLst><p:ext uri="urn:opaque"><x:keep xmlns:x="urn:x"/></p:ext><p:ext uri="{PROPERTIES_EXTENSION_URI}"><p202:designPr edtDesignElem="false"><p202:designTagLst><p202:designTag name="a" val="b"/></p202:designTagLst><p202:extLst><x:raw xmlns:x="urn:x"/></p202:extLst></p202:designPr></p:ext></p:extLst>"#
        );
        let (mut package, owner) = package(slide(&shape(&body)));
        let snapshot = load_properties(&package, &owner, "Box").unwrap();
        assert!(!snapshot.properties().unwrap().effective_editable());
        let mut edit = snapshot.edit();
        edit.set(DrawingProperties::new().with_editable(Some(true)))
            .unwrap();
        let committed = apply_properties(&mut package, edit.commit().unwrap()).unwrap();
        assert!(committed.properties().unwrap().effective_editable());
        let removed = remove_properties(&mut package, &owner, "Box").unwrap();
        assert!(!removed.is_present());
        let bytes = package.get_part(&owner).unwrap().blob();
        assert!(
            bytes
                .windows(b"uri=\"urn:opaque\"".len())
                .any(|v| v == b"uri=\"urn:opaque\"")
        );
    }

    #[test]
    fn wrong_namespace_is_opaque_and_duplicate_guids_are_inventoried() {
        let wrong = format!(
            r#"<p:extLst><p:ext uri="{PROPERTIES_EXTENSION_URI}"><bad:designPr xmlns:bad="urn:bad"/></p:ext></p:extLst>"#
        );
        let (wrong_package, owner) = package(slide(&shape(&wrong)));
        assert!(load_properties(&wrong_package, &owner, "Box").is_ok());
        let duplicate = format!(
            r#"<p:extLst><p:ext uri="{PROPERTIES_EXTENSION_URI}"/><p:ext uri="{PROPERTIES_EXTENSION_URI}"/></p:extLst>"#
        );
        let (duplicate_package, owner) = package(slide(&shape(&duplicate)));
        let snapshot = load_properties(&duplicate_package, &owner, "Box").unwrap();
        assert_eq!(snapshot.occurrence_count(), 2);
        assert_eq!(snapshot.occurrences().count(), 2);
        let mut edit = snapshot.edit();
        edit.set(DrawingProperties::new()).unwrap();
        assert!(edit.commit().is_err());
    }

    #[test]
    fn no_op_stale_and_explicit_false_are_distinct() {
        let body = format!(
            r#"<p:extLst><p:ext uri="{PROPERTIES_EXTENSION_URI}"><p202:designPr edtDesignElem="false"/></p:ext></p:extLst>"#
        );
        let (mut package, owner) = package(slide(&shape(&body)));
        let snapshot = load_properties(&package, &owner, "Box").unwrap();
        let edit = snapshot.edit();
        let commit = edit.commit().unwrap();
        assert!(commit.is_noop());
        let returned = apply_properties(&mut package, commit).unwrap();
        assert_eq!(returned.properties().unwrap().editable(), Some(false));
        let snapshot = load_properties(&package, &owner, "Box").unwrap();
        let mut edit = snapshot.edit();
        edit.remove();
        let commit = edit.commit().unwrap();
        package
            .get_part_mut(&owner)
            .unwrap()
            .set_blob(slide(&shape("<p:extLst/> ")).into_bytes());
        assert!(apply_properties(&mut package, commit).is_err());
    }

    #[test]
    fn no_op_rejects_stale_owner_and_remove_projects_absence() {
        let body = format!(
            r#"<p:extLst><p:ext uri="{PROPERTIES_EXTENSION_URI}"><p202:designPr/></p:ext></p:extLst>"#
        );
        let (mut package, owner) = package(slide(&shape(&body)));
        let mut removal = load_properties(&package, &owner, "Box").unwrap().edit();
        removal.remove();
        assert!(removal.properties().is_none());

        let commit = load_properties(&package, &owner, "Box")
            .unwrap()
            .edit()
            .commit()
            .unwrap();
        package
            .get_part_mut(&owner)
            .unwrap()
            .set_blob(slide(&shape("<p:nvPr/> ")).into_bytes());
        assert!(apply_properties(&mut package, commit).is_err());
    }

    #[test]
    fn strict_insertion_uses_selected_host_qname_and_namespace() {
        let strict = std::str::from_utf8(PML_STRICT).unwrap();
        let (mut package, owner) = package(slide_with_namespace(strict, &shape("")));
        let mut edit = load_properties(&package, &owner, "Box").unwrap().edit();
        edit.set(DrawingProperties::new().with_editable(Some(true)))
            .unwrap();
        apply_properties(&mut package, edit.commit().unwrap()).unwrap();
        let bytes = package.get_part(&owner).unwrap().blob();
        let text = std::str::from_utf8(bytes).unwrap();
        assert!(text.contains(&format!(r#"<p:extLst xmlns:p="{strict}">"#)));
        assert!(text.contains(&format!(
            r#"<p:ext xmlns:p="{strict}" uri="{PROPERTIES_EXTENSION_URI}">"#
        )));
        assert!(!text.contains(PML));
    }

    #[test]
    fn duplicate_outer_extensions_are_readable_but_not_mutable() {
        let body = format!(
            r#"<p:extLst><p:ext uri="{PROPERTIES_EXTENSION_URI}"><p202:designPr edtDesignElem="false"/></p:ext><p:ext uri="{PROPERTIES_EXTENSION_URI}"><p202:designPr edtDesignElem="true"/></p:ext></p:extLst>"#
        );
        let (package, owner) = package(slide(&shape(&body)));
        let snapshot = load_properties(&package, &owner, "Box").unwrap();
        assert_eq!(snapshot.occurrence_count(), 2);
        let values = snapshot
            .occurrences()
            .map(|value| value.unwrap().effective_editable())
            .collect::<Vec<_>>();
        assert_eq!(values, [false, true]);
        let mut edit = snapshot.edit();
        edit.set(DrawingProperties::new()).unwrap();
        assert!(edit.commit().is_err());
    }

    #[test]
    fn removal_preserves_annotated_containers_and_comments() {
        let body = format!(
            r#"<p:extLst custom="keep"><!--list--><p:ext uri="{PROPERTIES_EXTENSION_URI}" custom="keep"><!--outer--><p202:designPr/></p:ext></p:extLst>"#
        );
        let (mut package, owner) = package(slide(&shape(&body)));
        remove_properties(&mut package, &owner, "Box").unwrap();
        let text = std::str::from_utf8(package.get_part(&owner).unwrap().blob()).unwrap();
        assert!(text.contains(r#"<p:extLst custom="keep">"#));
        assert!(text.contains("<!--list-->"));
        assert!(
            text.contains(r#"<p:ext uri="{E7BDC344-281C-4309-B0C6-D0EE65EED2A8}" custom="keep">"#)
        );
        assert!(text.contains("<!--outer-->"));
        assert!(!text.contains("p202:designPr"));
    }
}
