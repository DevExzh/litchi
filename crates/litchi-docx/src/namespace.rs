#![expect(
    clippy::arbitrary_source_item_ordering,
    reason = "items remain grouped by OOXML schema family and package lifecycle"
)]
#![expect(
    clippy::needless_pass_by_value,
    reason = "the public API shape is retained for compatibility"
)]
#![expect(
    clippy::option_option,
    reason = "nested options distinguish omitted, present-empty, and present-valued XML"
)]
#![expect(
    clippy::ref_option,
    reason = "the public API shape is retained for compatibility"
)]
#![expect(
    clippy::shadow_reuse,
    reason = "parser bindings are intentionally refined after validation"
)]
#![expect(
    clippy::shadow_unrelated,
    reason = "local parser names mirror the OOXML role currently being decoded"
)]
use crate::error::{Error, Result};
use quick_xml::XmlVersion;
use quick_xml::encoding::Decoder;
use quick_xml::events::{BytesStart, Event};
use quick_xml::name::{Namespace, NamespaceResolver, PrefixDeclaration, ResolveResult};
use quick_xml::reader::NsReader;
use std::sync::{Arc, LazyLock};

pub(crate) const WORDPROCESSINGML_NAMESPACE: &[u8] =
    b"http://schemas.openxmlformats.org/wordprocessingml/2006/main";
pub(crate) const STRICT_WORDPROCESSINGML_NAMESPACE: &[u8] =
    b"http://purl.oclc.org/ooxml/wordprocessingml/main";

/// A bounded snapshot of the namespace bindings that are in scope at a
/// retained Word element.
///
/// Element views are zero-copy ranges into a part, so their opening tag does
/// not necessarily carry declarations written on an ancestor.  The package
/// scanner captures this snapshot while it already has the ancestor stack
/// open and shares it with the view.
pub(crate) type NamespaceBindings = Arc<[(Option<Vec<u8>>, Vec<u8>)]>;

static EMPTY_NAMESPACE_BINDINGS: LazyLock<NamespaceBindings> = LazyLock::new(|| Arc::from([]));

const MAX_NAMESPACE_BINDINGS: usize = 256;
const MAX_NAMESPACE_BYTES: usize = 64 * 1024;

#[cfg(test)]
#[path = "namespace_capture_tests.rs"]
mod capture_tests;

/// Reuse unchanged namespace scopes and bound aggregate new captures per scan.
/// The charge includes owned prefix/URI bytes and binding entries, not allocator RSS.
pub(crate) struct NamespaceCapture {
    last: Option<NamespaceBindings>,
    used: usize,
    maximum: usize,
}

impl Default for NamespaceCapture {
    fn default() -> Self {
        Self {
            last: None,
            used: 0,
            maximum: 64 * 1024 * 1024,
        }
    }
}

impl NamespaceCapture {
    pub(crate) fn capture(&mut self, resolver: &NamespaceResolver) -> Result<NamespaceBindings> {
        if let Some(last) = &self.last {
            let mut saved = last.iter();
            let equal = resolver.bindings().all(|(prefix, Namespace(namespace))| {
                saved.next().is_some_and(|(old_prefix, old_namespace)| {
                    let prefix = match prefix {
                        PrefixDeclaration::Default => None,
                        PrefixDeclaration::Named(prefix) => Some(prefix),
                    };
                    prefix == old_prefix.as_deref() && namespace == old_namespace.as_slice()
                })
            }) && saved.next().is_none();
            if equal {
                return Ok(Arc::clone(last));
            }
        }
        let (count, bytes) = namespace_requirements(resolver)?;
        let required = count
            .checked_mul(size_of::<(Option<Vec<u8>>, Vec<u8>)>())
            .and_then(|entries| entries.checked_add(bytes))
            .and_then(|additional| self.used.checked_add(additional))
            .ok_or_else(|| Error::InvalidFormat("Word namespace capture budget overflow".into()))?;
        if required > self.maximum {
            return Err(Error::InvalidFormat(format!(
                "Word namespace captures require {required} bytes, limit {}",
                self.maximum
            )));
        }
        let captured = snapshot_namespace_bindings(resolver, count)?;
        self.used = required;
        self.last = Some(Arc::clone(&captured));
        Ok(captured)
    }
}

impl NamespaceCapture {
    /// Capture the bindings in scope in the shared binding tracker that the
    /// production layout scanner drives (change 0754), under the same reuse,
    /// count, byte and aggregate bounds as [`Self::capture`].
    ///
    /// The tracker visits each prefix once, innermost first, and omits the
    /// reserved prefixes and undeclared bindings, so the snapshot is exactly
    /// the set a retained span needs to resolve every name the way it
    /// resolves in its part.
    pub(crate) fn capture_tracker(
        &mut self,
        tracker: &litchi_ooxml_common::private::BindingTracker,
    ) -> Result<NamespaceBindings> {
        if let Some(last) = &self.last {
            let mut saved = last.iter();
            let mut equal = true;
            tracker.for_each_in_scope(|prefix, namespace| {
                if !equal {
                    return;
                }
                let prefix = (!prefix.is_empty()).then_some(prefix);
                equal = saved.next().is_some_and(|(old_prefix, old_namespace)| {
                    prefix == old_prefix.as_deref() && namespace == old_namespace.as_slice()
                });
            });
            if equal && saved.next().is_none() {
                return Ok(Arc::clone(last));
            }
        }
        let mut count = 0usize;
        let mut bytes = 0usize;
        let mut bound_error = None;
        tracker.for_each_in_scope(|prefix, namespace| {
            if bound_error.is_some() {
                return;
            }
            count += 1;
            if count > MAX_NAMESPACE_BINDINGS {
                bound_error = Some(Error::InvalidFormat(format!(
                    "Word XML namespace binding count exceeds {MAX_NAMESPACE_BINDINGS}"
                )));
                return;
            }
            match bytes
                .checked_add(prefix.len())
                .and_then(|value| value.checked_add(namespace.len()))
            {
                Some(total) if total <= MAX_NAMESPACE_BYTES => bytes = total,
                Some(_) => {
                    bound_error = Some(Error::InvalidFormat(format!(
                        "Word XML namespace bytes exceed {MAX_NAMESPACE_BYTES}"
                    )));
                },
                None => {
                    bound_error = Some(Error::InvalidFormat(
                        "Word XML namespace bytes overflow".into(),
                    ));
                },
            }
        });
        if let Some(error) = bound_error {
            return Err(error);
        }
        let required = count
            .checked_mul(size_of::<(Option<Vec<u8>>, Vec<u8>)>())
            .and_then(|entries| entries.checked_add(bytes))
            .and_then(|additional| self.used.checked_add(additional))
            .ok_or_else(|| Error::InvalidFormat("Word namespace capture budget overflow".into()))?;
        if required > self.maximum {
            return Err(Error::InvalidFormat(format!(
                "Word namespace captures require {required} bytes, limit {}",
                self.maximum
            )));
        }
        let captured = if count == 0 {
            Arc::clone(&EMPTY_NAMESPACE_BINDINGS)
        } else {
            let mut bindings = Vec::new();
            bindings
                .try_reserve_exact(count)
                .map_err(|source| Error::Allocation {
                    resource: "Word namespace bindings",
                    source,
                })?;
            let mut copy_error = None;
            tracker.for_each_in_scope(|prefix, namespace| {
                if copy_error.is_some() {
                    return;
                }
                let copied = (|| -> Result<(Option<Vec<u8>>, Vec<u8>)> {
                    let prefix = if prefix.is_empty() {
                        None
                    } else {
                        Some(copy_namespace_bytes(prefix)?)
                    };
                    Ok((prefix, copy_namespace_bytes(namespace)?))
                })();
                match copied {
                    Ok(binding) => bindings.push(binding),
                    Err(error) => copy_error = Some(error),
                }
            });
            if let Some(error) = copy_error {
                return Err(error);
            }
            Arc::from(bindings.into_boxed_slice())
        };
        self.used = required;
        self.last = Some(Arc::clone(&captured));
        Ok(captured)
    }
}

fn copy_namespace_bytes(bytes: &[u8]) -> Result<Vec<u8>> {
    let mut value = Vec::new();
    value
        .try_reserve_exact(bytes.len())
        .map_err(|source| Error::Allocation {
            resource: "Word namespace bytes",
            source,
        })?;
    value.extend_from_slice(bytes);
    Ok(value)
}

fn namespace_requirements(resolver: &NamespaceResolver) -> Result<(usize, usize)> {
    let mut count = 0usize;
    let mut bytes = 0usize;
    for (prefix, Namespace(namespace)) in resolver.bindings() {
        count += 1;
        if count > MAX_NAMESPACE_BINDINGS {
            return Err(Error::InvalidFormat(format!(
                "Word XML namespace binding count exceeds {MAX_NAMESPACE_BINDINGS}"
            )));
        }
        let prefix_len = match prefix {
            PrefixDeclaration::Default => 0,
            PrefixDeclaration::Named(prefix) => prefix.len(),
        };
        bytes = bytes
            .checked_add(prefix_len)
            .and_then(|value| value.checked_add(namespace.len()))
            .ok_or_else(|| Error::InvalidFormat("Word XML namespace bytes overflow".into()))?;
        if bytes > MAX_NAMESPACE_BYTES {
            return Err(Error::InvalidFormat(format!(
                "Word XML namespace bytes exceed {MAX_NAMESPACE_BYTES}"
            )));
        }
    }
    Ok((count, bytes))
}

/// Snapshot the bindings a resolver holds, under the same binding-count and
/// byte bounds as a captured scan snapshot.
pub(crate) fn resolver_bindings(resolver: &NamespaceResolver) -> Result<NamespaceBindings> {
    let (count, _bytes) = namespace_requirements(resolver)?;
    snapshot_namespace_bindings(resolver, count)
}

fn snapshot_namespace_bindings(
    resolver: &NamespaceResolver,
    count: usize,
) -> Result<NamespaceBindings> {
    fn copy(bytes: &[u8]) -> Result<Vec<u8>> {
        let mut value = Vec::new();
        value
            .try_reserve_exact(bytes.len())
            .map_err(|source| Error::Allocation {
                resource: "Word namespace bytes",
                source,
            })?;
        value.extend_from_slice(bytes);
        Ok(value)
    }
    if count == 0 {
        return Ok(Arc::clone(&EMPTY_NAMESPACE_BINDINGS));
    }
    let mut bindings = Vec::new();
    bindings
        .try_reserve_exact(count)
        .map_err(|source| Error::Allocation {
            resource: "Word namespace bindings",
            source,
        })?;
    for (prefix, Namespace(namespace)) in resolver.bindings() {
        let prefix = match prefix {
            PrefixDeclaration::Default => None,
            PrefixDeclaration::Named(prefix) => Some(copy(prefix)?),
        };
        bindings.push((prefix, copy(namespace)?));
    }
    Ok(Arc::from(bindings.into_boxed_slice()))
}

pub(crate) fn is_wordprocessing_namespace(namespace: &ResolveResult<'_>) -> bool {
    matches!(
        namespace,
        ResolveResult::Bound(Namespace(value))
            if *value == WORDPROCESSINGML_NAMESPACE
                || *value == STRICT_WORDPROCESSINGML_NAMESPACE
    )
}

/// One retained element span of `source`, with every namespace declaration it
/// inherits from `source` re-declared on its own root element.
///
/// Until change 0653 the shared markup-compatibility writer repeated every
/// in-scope declaration on every element it emitted, so any span cut out of a
/// processed part happened to resolve on its own. The writer now declares each
/// namespace once, where XML requires it, so a consumer that hands a span out
/// or parses it standalone asks for the inherited declarations here instead.
pub(crate) fn self_contained_element_xml(
    source: &[u8],
    start: u32,
    length: u32,
) -> Result<Vec<u8>> {
    let start = usize::try_from(start)
        .map_err(|_source_error| Error::InvalidFormat("Word XML offset exceeds usize".into()))?;
    let length = usize::try_from(length).map_err(|_source_error| {
        Error::InvalidFormat("Word XML range length exceeds usize".into())
    })?;
    Ok(litchi_ooxml_common::mce::self_contained_fragment(
        source,
        start,
        length,
        &litchi_ooxml_common::mce::Limits::default(),
    )?
    .into_owned())
}

pub(crate) fn word_attribute_value(
    element: &BytesStart<'_>,
    name: &[u8],
    decoder: Decoder,
    resolver: &NamespaceResolver,
) -> Result<Option<String>> {
    let mut value = None;
    for attribute in element.attributes() {
        let attribute = attribute.map_err(|error| Error::Xml(error.to_string()))?;
        if attribute.key.local_name().as_ref() != name {
            continue;
        }
        let (namespace, _) = resolver.resolve_attribute(attribute.key);
        let is_word_attribute = is_wordprocessing_namespace(&namespace)
            || matches!(namespace, ResolveResult::Unbound)
            || matches!(namespace, ResolveResult::Unknown(prefix) if prefix.as_slice() == b"w");
        if !is_word_attribute {
            continue;
        }
        if value.is_some() {
            return Err(Error::InvalidFormat(format!(
                "duplicate Word attribute '{}'",
                String::from_utf8_lossy(name)
            )));
        }
        value = Some(
            attribute
                .decoded_and_normalized_value(XmlVersion::Explicit1_0, decoder)
                .map_err(|error| Error::Xml(error.to_string()))?
                .into_owned(),
        );
    }
    Ok(value)
}

fn is_fragment_word_namespace(
    namespace: &ResolveResult<'_>,
    fragment_prefix: &Option<Option<Vec<u8>>>,
) -> bool {
    if is_wordprocessing_namespace(namespace) {
        return true;
    }
    match namespace {
        ResolveResult::Unknown(prefix) => {
            fragment_prefix
                .as_ref()
                .and_then(|prefix| prefix.as_deref())
                == Some(prefix.as_slice())
        },
        ResolveResult::Unbound => fragment_prefix == &Some(None),
        ResolveResult::Bound(_) => false,
    }
}

/// Maximum capture nesting depth accepted by the shared `WordprocessingML`
/// element scanner, matching the hardened settings and mail-merge parsers.
const MAX_SCAN_DEPTH: usize = 128;
/// Maximum number of elements scanned in one document part.
const MAX_SCAN_NODES: usize = 1_000_000;

pub(crate) fn scan_word_element_ranges(
    xml_bytes: &[u8],
    targets: &[&[u8]],
    mut emit: impl FnMut(usize, u32, u32) -> Result<()>,
) -> Result<()> {
    scan_word_element_ranges_impl(
        xml_bytes,
        &[],
        targets,
        false,
        |target, start, length, _| emit(target, start, length),
    )
}

/// Scan selected Word elements while retaining the namespace context that was
/// active at each selected element's opening tag.
///
/// The context is captured during this pass, when the resolver already has all
/// ancestor declarations in scope. Callers that retain detached zero-copy
/// ranges can store the returned [`NamespaceBindings`] beside each range and
/// avoid reparsing the complete source part on every semantic query.
pub(crate) fn scan_word_element_ranges_with_context(
    xml_bytes: &[u8],
    inherited_namespaces: &[(Option<Vec<u8>>, Vec<u8>)],
    targets: &[&[u8]],
    emit: impl FnMut(usize, u32, u32, NamespaceBindings) -> Result<()>,
) -> Result<()> {
    scan_word_element_ranges_impl(xml_bytes, inherited_namespaces, targets, true, emit)
}

fn scan_word_element_ranges_impl(
    xml_bytes: &[u8],
    inherited_namespaces: &[(Option<Vec<u8>>, Vec<u8>)],
    targets: &[&[u8]],
    capture_namespaces: bool,
    mut emit: impl FnMut(usize, u32, u32, NamespaceBindings) -> Result<()>,
) -> Result<()> {
    enum ScanEvent {
        Start(usize, NamespaceBindings),
        NestedStart,
        Empty(usize, NamespaceBindings),
        End,
        Eof,
        Other,
    }

    let mut reader = NsReader::from_reader(xml_bytes);
    for (prefix, namespace) in inherited_namespaces {
        let prefix = prefix
            .as_deref()
            .map_or(PrefixDeclaration::Default, |prefix| {
                PrefixDeclaration::Named(prefix)
            });
        reader
            .resolver_mut()
            .add(prefix, Namespace(namespace))
            .map_err(|error| Error::Xml(error.to_string()))?;
    }
    let mut namespace_capture = NamespaceCapture::default();
    let mut first_element = true;
    let mut fragment_prefix: Option<Option<Vec<u8>>> = None;
    let mut capture: Option<(usize, usize, usize, NamespaceBindings)> = None;
    let mut nodes = 0usize;
    let mut total_depth = 0usize;

    loop {
        let event_start = usize::try_from(reader.buffer_position()).map_err(|_source_error| {
            Error::InvalidFormat("Word XML offset does not fit usize".to_string())
        })?;
        let event = {
            let (namespace, event) = reader
                .read_resolved_event()
                .map_err(|error| Error::Xml(error.to_string()))?;

            if matches!(event, Event::Start(_) | Event::Empty(_)) {
                nodes = nodes.checked_add(1).ok_or_else(|| {
                    Error::InvalidFormat("Word XML element counter overflow".to_string())
                })?;
                if nodes > MAX_SCAN_NODES {
                    return Err(Error::InvalidFormat(format!(
                        "Word XML exceeds {MAX_SCAN_NODES} elements"
                    )));
                }
            }
            // Total nesting is tracked separately from capture depth so
            // deeply nested non-target content is rejected before
            // quick-xml's own namespace resolver overflows (u16).
            if matches!(event, Event::Start(_)) {
                total_depth = total_depth.checked_add(1).ok_or_else(|| {
                    Error::InvalidFormat("Word XML nesting is too deep".to_string())
                })?;
                if total_depth > MAX_SCAN_DEPTH {
                    return Err(Error::InvalidFormat(format!(
                        "Word XML nesting exceeds the {MAX_SCAN_DEPTH} depth limit"
                    )));
                }
            }
            if matches!(event, Event::End(_)) {
                total_depth = total_depth
                    .checked_sub(1)
                    .ok_or_else(|| Error::InvalidFormat("invalid Word XML nesting".to_string()))?;
            }

            if first_element && matches!(event, Event::Start(_) | Event::Empty(_)) {
                first_element = false;
                if inherited_namespaces.is_empty()
                    && !matches!(namespace, ResolveResult::Bound(_))
                    && let Event::Start(element) | Event::Empty(element) = &event
                {
                    fragment_prefix = Some(
                        element
                            .name()
                            .prefix()
                            .map(|prefix| prefix.into_inner().to_vec()),
                    );
                }
            }

            match event {
                Event::Start(_) if capture.is_some() => ScanEvent::NestedStart,
                Event::Start(element)
                    if is_fragment_word_namespace(&namespace, &fragment_prefix) =>
                {
                    if let Some(target) = targets
                        .iter()
                        .position(|target| element.local_name().as_ref() == *target)
                    {
                        ScanEvent::Start(
                            target,
                            if capture_namespaces {
                                namespace_capture.capture(reader.resolver())?
                            } else {
                                Arc::clone(&EMPTY_NAMESPACE_BINDINGS)
                            },
                        )
                    } else {
                        ScanEvent::Other
                    }
                },
                Event::Empty(element)
                    if capture.is_none()
                        && is_fragment_word_namespace(&namespace, &fragment_prefix) =>
                {
                    if let Some(target) = targets
                        .iter()
                        .position(|target| element.local_name().as_ref() == *target)
                    {
                        ScanEvent::Empty(
                            target,
                            if capture_namespaces {
                                namespace_capture.capture(reader.resolver())?
                            } else {
                                Arc::clone(&EMPTY_NAMESPACE_BINDINGS)
                            },
                        )
                    } else {
                        ScanEvent::Other
                    }
                },
                Event::End(_) if capture.is_some() => ScanEvent::End,
                Event::Eof => ScanEvent::Eof,
                Event::Start(_)
                | Event::End(_)
                | Event::Empty(_)
                | Event::Text(_)
                | Event::CData(_)
                | Event::Comment(_)
                | Event::Decl(_)
                | Event::PI(_)
                | Event::DocType(_)
                | Event::GeneralRef(_) => ScanEvent::Other,
            }
        };
        let event_end = usize::try_from(reader.buffer_position()).map_err(|_source_error| {
            Error::InvalidFormat("Word XML offset does not fit usize".to_string())
        })?;

        match event {
            ScanEvent::Start(target, namespaces) => {
                capture = Some((target, event_start, 1, namespaces));
            },
            ScanEvent::NestedStart => {
                let Some((_, _, depth, _)) = capture.as_mut() else {
                    return Err(Error::InvalidFormat(
                        "missing captured Word element".to_string(),
                    ));
                };
                *depth = depth.checked_add(1).ok_or_else(|| {
                    Error::InvalidFormat("Word element nesting is too deep".to_string())
                })?;
                if *depth > MAX_SCAN_DEPTH {
                    return Err(Error::InvalidFormat(format!(
                        "Word element nesting exceeds the {MAX_SCAN_DEPTH} depth limit"
                    )));
                }
            },
            ScanEvent::Empty(target, namespaces) => {
                let length =
                    u32::try_from(event_end.checked_sub(event_start).ok_or_else(|| {
                        Error::InvalidFormat("Word element range underflow".into())
                    })?)
                    .map_err(|_source_error| {
                        Error::InvalidFormat("Word element range exceeds u32".into())
                    })?;
                emit(
                    target,
                    u32::try_from(event_start).map_err(|_source_error| {
                        Error::InvalidFormat("Word element offset exceeds u32".into())
                    })?,
                    length,
                    namespaces,
                )?;
            },
            ScanEvent::End => {
                let Some((_, _, depth, _)) = capture.as_mut() else {
                    return Err(Error::InvalidFormat(
                        "missing captured Word element".to_string(),
                    ));
                };
                *depth = depth.checked_sub(1).ok_or_else(|| {
                    Error::InvalidFormat("invalid Word element nesting".to_string())
                })?;
                if *depth == 0 {
                    let Some((target, start, _, namespaces)) = capture.take() else {
                        return Err(Error::InvalidFormat(
                            "missing captured Word element range".to_string(),
                        ));
                    };
                    let start = u32::try_from(start).map_err(|_source_error| {
                        Error::InvalidFormat("Word element offset exceeds u32".into())
                    })?;
                    let length =
                        u32::try_from(event_end.checked_sub(start as usize).ok_or_else(|| {
                            Error::InvalidFormat("Word element range underflow".into())
                        })?)
                        .map_err(|_source_error| {
                            Error::InvalidFormat("Word element range exceeds u32".into())
                        })?;
                    emit(target, start, length, namespaces)?;
                }
            },
            ScanEvent::Eof if capture.is_some() => {
                return Err(Error::InvalidFormat(
                    "unterminated Word element".to_string(),
                ));
            },
            ScanEvent::Eof => break,
            ScanEvent::Other => {},
        }
    }

    Ok(())
}

pub(crate) fn direct_word_property_value(
    xml_bytes: &[u8],
    root_name: &[u8],
    properties_name: &[u8],
    property_name: &[u8],
) -> Result<Option<String>> {
    let mut reader = NsReader::from_reader(xml_bytes);
    let mut fragment_prefix: Option<Option<Vec<u8>>> = None;
    let mut depth = 0usize;
    let mut properties_depth = None;
    let mut saw_properties = false;
    let mut value = None;
    let mut saw_root = false;

    loop {
        let decoder = reader.decoder();
        let event = reader
            .read_event()
            .map_err(|error| Error::Xml(error.to_string()))?
            .into_owned();
        let resolver = reader.resolver().clone();
        let (namespace, event) = resolver.resolve_event(event);

        if fragment_prefix.is_none()
            && depth == 0
            && let Event::Start(element) | Event::Empty(element) = &event
            && !matches!(namespace, ResolveResult::Bound(_))
        {
            fragment_prefix = Some(
                element
                    .name()
                    .prefix()
                    .map(|prefix| prefix.into_inner().to_vec()),
            );
        }

        match event {
            Event::Start(element) => {
                depth = depth.checked_add(1).ok_or_else(|| {
                    Error::InvalidFormat("Word property XML nesting is too deep".into())
                })?;
                let is_word = is_fragment_word_namespace(&namespace, &fragment_prefix);
                if depth == 1 {
                    if saw_root || !is_word || element.local_name().as_ref() != root_name {
                        return Err(Error::InvalidFormat(
                            "Word property XML has an invalid root".into(),
                        ));
                    }
                    saw_root = true;
                } else if depth == 2 && is_word && element.local_name().as_ref() == properties_name
                {
                    if saw_properties {
                        return Err(Error::InvalidFormat(
                            "duplicate Word property container".into(),
                        ));
                    }
                    saw_properties = true;
                    properties_depth = Some(depth);
                } else if depth == 3
                    && properties_depth == Some(2)
                    && is_word
                    && element.local_name().as_ref() == property_name
                {
                    set_direct_property_value(
                        &mut value,
                        &element,
                        decoder,
                        &resolver,
                        &fragment_prefix,
                        property_name,
                    )?;
                }
            },
            Event::Empty(element) => {
                let child_depth = depth.checked_add(1).ok_or_else(|| {
                    Error::InvalidFormat("Word property XML nesting is too deep".into())
                })?;
                let is_word = is_fragment_word_namespace(&namespace, &fragment_prefix);
                if child_depth == 1 {
                    if saw_root || !is_word || element.local_name().as_ref() != root_name {
                        return Err(Error::InvalidFormat(
                            "Word property XML has an invalid root".into(),
                        ));
                    }
                    saw_root = true;
                } else if child_depth == 2
                    && is_word
                    && element.local_name().as_ref() == properties_name
                {
                    if saw_properties {
                        return Err(Error::InvalidFormat(
                            "duplicate Word property container".into(),
                        ));
                    }
                    saw_properties = true;
                } else if child_depth == 3
                    && properties_depth == Some(2)
                    && is_word
                    && element.local_name().as_ref() == property_name
                {
                    set_direct_property_value(
                        &mut value,
                        &element,
                        decoder,
                        &resolver,
                        &fragment_prefix,
                        property_name,
                    )?;
                }
            },
            Event::End(_) => {
                if properties_depth == Some(depth) {
                    properties_depth = None;
                }
                depth = depth.checked_sub(1).ok_or_else(|| {
                    Error::InvalidFormat("invalid Word property XML nesting".into())
                })?;
            },
            Event::Eof if depth != 0 => {
                return Err(Error::InvalidFormat(
                    "unterminated Word property XML".into(),
                ));
            },
            Event::Eof => break,
            Event::Text(_)
            | Event::CData(_)
            | Event::Comment(_)
            | Event::Decl(_)
            | Event::PI(_)
            | Event::DocType(_)
            | Event::GeneralRef(_) => {},
        }
    }

    if !saw_root {
        return Err(Error::InvalidFormat("Word property XML has no root".into()));
    }
    Ok(value)
}

fn set_direct_property_value(
    slot: &mut Option<String>,
    element: &BytesStart<'_>,
    decoder: Decoder,
    resolver: &NamespaceResolver,
    fragment_prefix: &Option<Option<Vec<u8>>>,
    property_name: &[u8],
) -> Result<()> {
    if slot.is_some() {
        return Err(Error::InvalidFormat(format!(
            "duplicate Word property '{}'",
            String::from_utf8_lossy(property_name)
        )));
    }
    let mut value = None;
    for attribute in element.attributes() {
        let attribute = attribute.map_err(|error| Error::Xml(error.to_string()))?;
        if attribute.key.local_name().as_ref() != b"val" {
            continue;
        }
        let (namespace, _) = resolver.resolve_attribute(attribute.key);
        if !is_fragment_word_namespace(&namespace, fragment_prefix) {
            continue;
        }
        if value.is_some() {
            return Err(Error::InvalidFormat(
                "duplicate Word property value attribute".into(),
            ));
        }
        value = Some(
            attribute
                .decoded_and_normalized_value(XmlVersion::Explicit1_0, decoder)
                .map_err(|error| Error::Xml(error.to_string()))?
                .into_owned(),
        );
    }
    *slot = Some(value.ok_or_else(|| {
        Error::InvalidFormat(format!(
            "Word property '{}' requires a value",
            String::from_utf8_lossy(property_name)
        ))
    })?);
    Ok(())
}

pub(crate) fn normalize_xml_integer(value: String, description: &str) -> Result<String> {
    let value = value.trim();
    let digits = value
        .strip_prefix('+')
        .or_else(|| value.strip_prefix('-'))
        .unwrap_or(value);
    if digits.is_empty() || !digits.bytes().all(|byte| byte.is_ascii_digit()) {
        return Err(Error::InvalidFormat(format!(
            "invalid {description} value '{value}'"
        )));
    }
    Ok(value.to_owned())
}

#[cfg(test)]
mod tests {
    use quick_xml::events::Event;
    use quick_xml::name::ResolveResult;
    use quick_xml::reader::NsReader;

    /// Whether the fragment's root element resolves into a namespace.
    fn root_is_bound(fragment: &[u8]) -> bool {
        let mut reader = NsReader::from_reader(fragment);
        let mut buffer = Vec::new();
        loop {
            match reader.read_event_into(&mut buffer) {
                Ok(Event::Start(ref element) | Event::Empty(ref element)) => {
                    return matches!(
                        reader.resolver().resolve_element(element.name()).0,
                        ResolveResult::Bound(_)
                    );
                },
                Ok(Event::Eof) | Err(_) => return false,
                Ok(_) => {},
            }
            buffer.clear();
        }
    }

    /// Change 0653: a `w:p` span cut out of a real, marker-bearing
    /// `word/document.xml` no longer carries `xmlns:w`, because the shared
    /// markup-compatibility writer stopped repeating every in-scope
    /// declaration on every element. `Paragraph::self_contained_xml` restores
    /// it at the slice boundary, which is what `Paragraph::extensions` and
    /// `Row::extension_ids` now parse.
    #[test]
    fn a_retained_span_resolves_only_after_the_inherited_declarations_return() {
        let package = crate::Package::open("../../test-data/ooxml/docx/table-alignment.docx")
            .expect("a Word-authored fixture opens");
        let document = package.document().expect("the document part is readable");
        let paragraph = document
            .paragraph(0)
            .expect("the first paragraph is readable")
            .expect("the fixture has a paragraph");

        assert!(
            !root_is_bound(paragraph.xml_bytes()),
            "the retained span still carries every in-scope declaration"
        );
        let repaired = paragraph
            .self_contained_xml()
            .expect("the span can be made self-contained");
        assert!(
            root_is_bound(&repaired),
            "the repaired span still does not resolve: {}",
            String::from_utf8_lossy(&repaired[..repaired.len().min(200)])
        );
    }
}
