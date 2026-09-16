//! Making one element span of a processed part parse standalone.
//!
//! Until change 0653 the markup-compatibility writer re-declared every in-scope
//! namespace binding on every start tag it emitted, so any element span of the
//! processed buffer happened to be namespace self-contained and a consumer
//! could slice an inner element out of it and parse that range on its own. The
//! writer now declares each namespace once, where XML requires it, which makes
//! the processed part 16x smaller on real producer output and leaves an inner
//! span carrying only the declarations its own subtree makes.
//!
//! This module restores the property **at the slice boundary** instead: one
//! bounded walk of the bytes before the span collects the declarations the span
//! inherits, and [`InScopeNamespaces::make_self_contained`] re-declares exactly
//! those on the span's own root element. That is the shape the crate's
//! already-hardened slicing sites use (`litchi-xlsb`'s drawing-anchor transfer,
//! `litchi-docx`'s section inventory and settings extensions), generalized so
//! every consumer can share one implementation and one set of bounds.

use std::borrow::Cow;

use memchr::memmem;
use quick_xml::events::{BytesStart, Event};
use quick_xml::reader::Reader;

use super::model::{Error, Limits};
use crate::private::BindingTracker;

type R<T> = Result<T, Error>;

fn bad(message: impl Into<String>) -> Error {
    Error::NonConformant(message.into())
}

fn limit(resource: &str) -> Error {
    Error::LimitExceeded(resource.into())
}

fn xerr(error: impl std::fmt::Display) -> Error {
    Error::Xml(error.to_string())
}

fn reserve_exact<T>(values: &mut Vec<T>, additional: usize, resource: &'static str) -> R<()> {
    values
        .try_reserve_exact(additional)
        .map_err(|source| Error::Allocation { resource, source })
}

/// The namespace declarations in scope at one point of an XML document.
///
/// Ordered innermost binding first, each prefix appearing at most once, with
/// the reserved `xml` and `xmlns` prefixes and any prefix left undeclared
/// (`xmlns=""`) excluded. An empty prefix is the default namespace.
#[derive(Clone, Debug, Default, PartialEq, Eq)]
pub struct InScopeNamespaces {
    bindings: Vec<(Box<[u8]>, Box<[u8]>)>,
}

impl InScopeNamespaces {
    /// Collect the declarations in scope at `offset` of `document`.
    ///
    /// `offset` names the first byte of an element's start tag; the walk stops
    /// before the event that begins there, so the result is exactly what that
    /// element inherits from its ancestors, without its own declarations.
    ///
    /// The walk is bounded by `limits`: `max_depth` caps element nesting and
    /// `max_namespace_bindings` caps how many distinct prefixes are returned.
    ///
    /// # Errors
    ///
    /// Returns [`Error::Xml`] when `document` is not well formed before
    /// `offset`, [`Error::NonConformant`] when a namespace declaration is
    /// invalid or `offset` is outside `document`, [`Error::LimitExceeded`] when
    /// a bound in `limits` is reached, and [`Error::Allocation`] when a bounded
    /// buffer cannot be reserved.
    pub fn at_offset(document: &[u8], offset: usize, limits: &Limits) -> R<Self> {
        let head = document
            .get(..offset)
            .ok_or_else(|| bad("fragment offset is outside the document"))?;
        // A document with no declaration before the span binds nothing, so the
        // span already carries every binding it can use.
        if memmem::find(head, b"xmlns").is_none() {
            return Ok(Self::default());
        }
        // Real producer parts declare every prefix on the part's root element.
        // When no `xmlns` occurs between the root start tag and the span, no
        // element between them declares anything, so the root's own
        // declarations are the complete in-scope set and the structural walk
        // below is not needed.
        if let Some(root) = root_start_tag_end(document)?
            && root <= offset
            && let Some(between) = document.get(root..offset)
            && memmem::find(between, b"xmlns").is_none()
        {
            return Self::from_tracker(&root_tracker(document, limits)?, limits);
        }
        Self::from_tracker(&walked_tracker(document, offset, limits)?, limits)
    }

    /// Build from declarations a caller already resolved, innermost first and
    /// each prefix at most once.
    ///
    /// This is the shape [`in_scope_declarations`](crate::private::in_scope_declarations)
    /// returns, so a consumer that already walked the document with its own
    /// namespace-aware reader re-uses that walk instead of paying for a second
    /// one here.
    #[must_use]
    pub fn from_declarations(bindings: Vec<(Box<[u8]>, Box<[u8]>)>) -> Self {
        Self { bindings }
    }

    fn from_tracker(tracker: &BindingTracker, limits: &Limits) -> R<Self> {
        let mut bindings: Vec<(Box<[u8]>, Box<[u8]>)> = Vec::new();
        let mut error = None;
        tracker.for_each_in_scope(|prefix, namespace| {
            if error.is_some() {
                return;
            }
            if bindings.len() >= limits.max_namespace_bindings {
                error = Some(limit("namespace bindings"));
                return;
            }
            if let Err(failure) = reserve_exact(&mut bindings, 1, "MCE in-scope namespaces") {
                error = Some(failure);
                return;
            }
            bindings.push((prefix.into(), namespace.into()));
        });
        match error {
            Some(failure) => Err(failure),
            None => Ok(Self { bindings }),
        }
    }

    /// Whether no declaration is in scope.
    #[must_use]
    pub fn is_empty(&self) -> bool {
        self.bindings.is_empty()
    }

    /// How many distinct prefixes are in scope.
    #[must_use]
    pub fn len(&self) -> usize {
        self.bindings.len()
    }

    /// Visit each binding as `(prefix, namespace)`, innermost first.
    ///
    /// An empty prefix is the default namespace.
    pub fn iter(&self) -> impl Iterator<Item = (&[u8], &[u8])> {
        self.bindings
            .iter()
            .map(|(prefix, namespace)| (prefix.as_ref(), namespace.as_ref()))
    }

    /// The namespace bound to `prefix`, if any. An empty prefix is the default.
    #[must_use]
    pub fn namespace(&self, prefix: &[u8]) -> Option<&[u8]> {
        self.bindings
            .iter()
            .find(|(candidate, _)| candidate.as_ref() == prefix)
            .map(|(_, namespace)| namespace.as_ref())
    }

    /// Copy `fragment` with every inherited binding it does not declare itself
    /// re-declared on its root element.
    ///
    /// The fragment is returned borrowed when nothing has to be added, which is
    /// the case whenever no declaration is in scope or the root already
    /// re-binds every one of them.
    ///
    /// # Errors
    ///
    /// Returns [`Error::Xml`] when `fragment` is not well formed,
    /// [`Error::NonConformant`] when it does not begin with an element,
    /// [`Error::LimitExceeded`] when the result would exceed
    /// `limits.max_output_bytes`, and [`Error::Allocation`] when the result
    /// cannot be reserved.
    pub fn make_self_contained<'a>(&self, fragment: &'a [u8], limits: &Limits) -> R<Cow<'a, [u8]>> {
        if self.bindings.is_empty() {
            return Ok(Cow::Borrowed(fragment));
        }
        let (insertion, declared) = root_insertion_point(fragment)?;
        let missing: Vec<&(Box<[u8]>, Box<[u8]>)> = self
            .bindings
            .iter()
            .filter(|(prefix, _)| !declared.iter().any(|candidate| candidate == prefix))
            .collect();
        if missing.is_empty() {
            return Ok(Cow::Borrowed(fragment));
        }
        let mut extra = 0usize;
        for (prefix, namespace) in &missing {
            // ` xmlns` + optional `:` + prefix + `="` + namespace + `"`.
            let declaration = b" xmlns=\"\"".len()
                + usize::from(!prefix.is_empty())
                + prefix.len()
                + escaped_len(namespace);
            extra = extra
                .checked_add(declaration)
                .ok_or_else(|| limit("output bytes"))?;
        }
        let total = fragment
            .len()
            .checked_add(extra)
            .ok_or_else(|| limit("output bytes"))?;
        if total > limits.max_output_bytes {
            return Err(limit("output bytes"));
        }
        let mut out = Vec::new();
        reserve_exact(&mut out, total, "MCE self-contained fragment")?;
        out.extend_from_slice(&fragment[..insertion]);
        for (prefix, namespace) in missing {
            out.extend_from_slice(b" xmlns");
            if !prefix.is_empty() {
                out.push(b':');
                out.extend_from_slice(prefix);
            }
            out.extend_from_slice(b"=\"");
            escape_into(&mut out, namespace);
            out.push(b'"');
        }
        out.extend_from_slice(&fragment[insertion..]);
        Ok(Cow::Owned(out))
    }
}

/// `document[start .. start + len]`, re-declared so that it parses standalone.
///
/// `start` must be the first byte of an element's start tag inside `document`
/// and `len` its complete span. Every namespace prefix the element inherits and
/// does not re-bind itself is added to its own start tag, so every qualified
/// name in the span resolves to the namespace it resolves to in `document`.
///
/// # Errors
///
/// Returns [`Error::NonConformant`] when the range is outside `document` or
/// does not begin with an element, [`Error::Xml`] when `document` is not well
/// formed before the range, [`Error::LimitExceeded`] when a bound in `limits`
/// is reached, and [`Error::Allocation`] when a bounded buffer cannot be
/// reserved.
pub fn self_contained_fragment<'a>(
    document: &'a [u8],
    start: usize,
    len: usize,
    limits: &Limits,
) -> R<Cow<'a, [u8]>> {
    let end = start
        .checked_add(len)
        .ok_or_else(|| bad("fragment range overflows the document"))?;
    let fragment = document
        .get(start..end)
        .ok_or_else(|| bad("fragment range is outside the document"))?;
    if fragment.first() != Some(&b'<') {
        return Err(bad("fragment does not begin with an element"));
    }
    InScopeNamespaces::at_offset(document, start, limits)?.make_self_contained(fragment, limits)
}

/// The end offset of the document's first start tag, or `None` when it has no
/// element at all.
fn root_start_tag_end(document: &[u8]) -> R<Option<usize>> {
    let mut reader = Reader::from_reader(document);
    reader.config_mut().trim_text(false);
    let mut buffer = Vec::new();
    loop {
        match reader.read_event_into(&mut buffer).map_err(xerr)? {
            Event::Start(_) | Event::Empty(_) => {
                return usize::try_from(reader.buffer_position())
                    .map(Some)
                    .map_err(|_error| bad("document offset does not fit usize"));
            },
            Event::Eof => return Ok(None),
            _ => {},
        }
        buffer.clear();
    }
}

/// A tracker holding only the document root's own declarations.
fn root_tracker(document: &[u8], limits: &Limits) -> R<BindingTracker> {
    let mut reader = Reader::from_reader(document);
    reader.config_mut().trim_text(false);
    let mut tracker = BindingTracker::new();
    let mut buffer = Vec::new();
    loop {
        match reader.read_event_into(&mut buffer).map_err(xerr)? {
            Event::Start(element) | Event::Empty(element) => {
                push(&mut tracker, &element, limits)?;
                return Ok(tracker);
            },
            Event::Eof => return Ok(tracker),
            _ => {},
        }
        buffer.clear();
    }
}

/// A tracker holding the declarations in scope at `offset`.
fn walked_tracker(document: &[u8], offset: usize, limits: &Limits) -> R<BindingTracker> {
    let mut reader = Reader::from_reader(document);
    reader.config_mut().trim_text(false);
    let mut tracker = BindingTracker::new();
    let mut buffer = Vec::new();
    let mut depth = 0usize;
    loop {
        let position = usize::try_from(reader.buffer_position())
            .map_err(|_error| bad("document offset does not fit usize"))?;
        if position >= offset {
            return Ok(tracker);
        }
        match reader.read_event_into(&mut buffer).map_err(xerr)? {
            Event::Start(element) => {
                depth = depth.checked_add(1).ok_or_else(|| limit("depth"))?;
                if depth > limits.max_depth {
                    return Err(limit("depth"));
                }
                push(&mut tracker, &element, limits)?;
            },
            Event::End(_) => {
                depth = depth.checked_sub(1).ok_or_else(|| bad("unexpected end"))?;
                tracker.pop();
            },
            Event::Eof => return Ok(tracker),
            _ => {},
        }
        buffer.clear();
    }
}

fn push(tracker: &mut BindingTracker, element: &BytesStart<'_>, limits: &Limits) -> R<()> {
    tracker
        .push(element)
        .map_err(|error| bad(format!("invalid namespace: {error}")))?;
    if tracker.declaration_count() > limits.max_namespace_bindings {
        return Err(limit("namespace bindings"));
    }
    Ok(())
}

/// Where a declaration may be inserted into a fragment's root start tag, and
/// the prefixes that tag already declares.
fn root_insertion_point(fragment: &[u8]) -> R<(usize, Vec<Box<[u8]>>)> {
    let mut reader = Reader::from_reader(fragment);
    reader.config_mut().trim_text(false);
    let mut buffer = Vec::new();
    loop {
        let event = reader.read_event_into(&mut buffer).map_err(xerr)?;
        let empty = matches!(event, Event::Empty(_));
        match event {
            Event::Start(element) | Event::Empty(element) => {
                let end = usize::try_from(reader.buffer_position())
                    .map_err(|_error| bad("fragment offset does not fit usize"))?;
                // The tag ends with `>`, and `/>` when it is an empty element.
                let closing = if empty { 2 } else { 1 };
                let insertion = end
                    .checked_sub(closing)
                    .ok_or_else(|| bad("fragment root start tag is truncated"))?;
                let mut declared = Vec::new();
                for attribute in element.attributes().with_checks(false) {
                    let Ok(attribute) = attribute else {
                        break;
                    };
                    let key = attribute.key.into_inner();
                    if key == b"xmlns" {
                        reserve_exact(&mut declared, 1, "MCE fragment declarations")?;
                        declared.push(Box::default());
                    } else if let Some(prefix) = key.strip_prefix(b"xmlns:".as_slice()) {
                        reserve_exact(&mut declared, 1, "MCE fragment declarations")?;
                        declared.push(prefix.into());
                    }
                }
                return Ok((insertion, declared));
            },
            Event::Eof => return Err(bad("fragment does not begin with an element")),
            Event::Decl(_) | Event::Comment(_) | Event::Text(_) => {},
            _ => return Err(bad("fragment does not begin with an element")),
        }
        buffer.clear();
    }
}

/// How many bytes a namespace URI occupies once quoted as an attribute value.
///
/// The tracker keeps the declaration's bytes exactly as the source wrote them
/// between its quotes, so they are already in attribute-value form: a `&` is
/// already the start of a reference and must not be escaped a second time, and
/// a whitespace character normalizes the same way here as it did there. Only a
/// literal `"`, legal inside a single-quoted source declaration, has to change
/// so that the rewritten declaration can use double quotes.
fn escaped_len(value: &[u8]) -> usize {
    value.iter().fold(0usize, |total, byte| {
        total.saturating_add(if *byte == b'"' { b"&quot;".len() } else { 1 })
    })
}

fn escape_into(out: &mut Vec<u8>, value: &[u8]) {
    for byte in value {
        if *byte == b'"' {
            out.extend_from_slice(b"&quot;");
        } else {
            out.push(*byte);
        }
    }
}
