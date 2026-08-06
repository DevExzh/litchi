//! Source-checked typed mail-merge edits.

use std::sync::Arc;

use quick_xml::events::Event;
use quick_xml::name::{Namespace, ResolveResult};
use quick_xml::reader::NsReader;

use super::model::{DataSourceObject, FieldMap, STRICT_W, Settings, W};
use super::{Conformance, Recipients};
use crate::{Error, Result};

/// Stable fingerprint of the exact settings and recipient-data source.
pub type Revision = u64;

/// An immutable, source-preserving typed mail-merge owner.
#[derive(Debug, Clone)]
pub struct Snapshot {
    xml: Arc<Vec<u8>>,
    settings: Settings,
    recipients_xml: Option<Arc<Vec<u8>>>,
    recipients: Option<Recipients>,
    conformance: Conformance,
    revision: Revision,
}

impl Snapshot {
    /// Parse a complete settings XML source containing `w:mailMerge`.
    pub fn from_xml(xml: impl Into<Vec<u8>>) -> Result<Self> {
        Self::from_parts(xml.into(), None)
    }

    /// Parse settings and an optional inert recipient-data XML source.
    pub fn from_parts(xml: impl Into<Vec<u8>>, recipients_xml: Option<Vec<u8>>) -> Result<Self> {
        let xml = xml.into();
        let settings = super::parse_settings_mail_merge(&xml)?
            .ok_or_else(|| invalid("settings XML does not contain mailMerge"))?;
        let conformance = detect_conformance(&xml)?;
        let recipients = recipients_xml
            .as_deref()
            .map(Recipients::parse_xml)
            .transpose()?;
        let xml = Arc::new(xml);
        let recipients_xml = recipients_xml.map(Arc::new);
        let revision = fingerprint(&xml, recipients_xml.as_deref().map(Vec::as_slice));
        Ok(Self {
            xml,
            settings,
            recipients_xml,
            recipients,
            conformance,
            revision,
        })
    }

    /// Borrow the exact settings source bytes.
    #[inline]
    #[must_use]
    pub fn xml_bytes(&self) -> &[u8] {
        self.xml.as_slice()
    }

    /// Borrow the optional exact recipient-data source bytes.
    #[inline]
    #[must_use]
    pub fn recipients_xml(&self) -> Option<&[u8]> {
        self.recipients_xml.as_deref().map(Vec::as_slice)
    }

    /// Borrow the typed settings projection.
    #[inline]
    #[must_use]
    pub const fn settings(&self) -> &Settings {
        &self.settings
    }

    /// Borrow the typed recipient projection.
    #[inline]
    #[must_use]
    pub const fn recipients(&self) -> Option<&Recipients> {
        self.recipients.as_ref()
    }

    /// Return the source conformance family.
    #[inline]
    #[must_use]
    pub const fn conformance(&self) -> Conformance {
        self.conformance
    }

    /// Return the source fingerprint used by stale checks.
    #[inline]
    #[must_use]
    pub const fn revision(&self) -> Revision {
        self.revision
    }

    /// Start an isolated typed edit.
    #[must_use]
    pub fn edit(&self) -> Transaction {
        Transaction {
            original: self.clone(),
            settings: self.settings.clone(),
            recipients: self.recipients.clone(),
        }
    }

    fn same_source(&self, other: &Self) -> bool {
        self.xml.as_slice() == other.xml.as_slice()
            && self.recipients_xml.as_deref() == other.recipients_xml.as_deref()
            && self.settings == other.settings
            && self.recipients == other.recipients
            && self.conformance == other.conformance
    }
}

/// A typed mail-merge edit that has not yet been published.
#[derive(Debug, Clone)]
pub struct Transaction {
    original: Snapshot,
    settings: Settings,
    recipients: Option<Recipients>,
}

impl Transaction {
    /// Borrow the staged settings projection.
    #[inline]
    #[must_use]
    pub const fn settings(&self) -> &Settings {
        &self.settings
    }

    /// Borrow the staged recipient projection.
    #[inline]
    #[must_use]
    pub const fn recipients(&self) -> Option<&Recipients> {
        self.recipients.as_ref()
    }

    /// Apply an atomic mutation to the typed settings model.
    pub fn edit_settings(
        &mut self,
        edit: impl FnOnce(&mut Settings) -> Result<()>,
    ) -> Result<&mut Self> {
        let mut candidate = self.settings.clone();
        edit(&mut candidate)?;
        candidate.to_xml(self.original.conformance)?;
        self.settings = candidate;
        Ok(self)
    }

    /// Replace the typed settings model after bounded validation.
    pub fn replace_settings(&mut self, value: Settings) -> Result<&mut Self> {
        value.to_xml(self.original.conformance)?;
        self.settings = value;
        Ok(self)
    }

    /// Replace the inert recipient inclusion metadata.
    pub fn set_recipients(&mut self, value: Option<Recipients>) -> Result<&mut Self> {
        if let Some(value) = &value {
            value.to_xml(self.original.conformance)?;
        }
        self.recipients = value;
        Ok(self)
    }

    /// Set one recipient inclusion flag without opening any data source.
    pub fn set_recipient_active(&mut self, index: usize, active: bool) -> Result<&mut Self> {
        let mut recipients = self
            .recipients
            .clone()
            .ok_or_else(|| invalid("mail-merge recipient data is absent"))?;
        recipients.set_recipient_active(index, active)?;
        self.recipients = Some(recipients);
        Ok(self)
    }

    /// Replace one ODSO field-map entry, preserving the surrounding model.
    pub fn set_field_map(&mut self, index: usize, value: Option<FieldMap>) -> Result<&mut Self> {
        let mut settings = self.settings.clone();
        let odso = settings.odso.get_or_insert_with(DataSourceObject::default);
        match value {
            Some(value) if index == odso.field_maps.len() => odso.field_maps.push(value),
            Some(value) if index < odso.field_maps.len() => odso.field_maps[index] = value,
            Some(_) => return Err(invalid("mail-merge field-map index is out of range")),
            None if index < odso.field_maps.len() => {
                odso.field_maps.remove(index);
            },
            None => return Err(invalid("mail-merge field-map index is out of range")),
        }
        settings.to_xml(self.original.conformance)?;
        self.settings = settings;
        Ok(self)
    }

    /// Validate and consume this edit into a reversible commit.
    pub fn commit(self) -> Result<Commit> {
        self.settings.to_xml(self.original.conformance)?;
        let semantic_changed =
            self.settings != self.original.settings || self.recipients != self.original.recipients;
        if !semantic_changed {
            return Ok(Commit {
                snapshot: self.original.clone(),
                patch: Patch {
                    before: self.original.clone(),
                    after: self.original,
                },
                changed: false,
            });
        }

        let settings_xml = if self.settings == self.original.settings {
            self.original.xml.as_ref().clone()
        } else {
            rewrite_settings(
                self.original.xml_bytes(),
                &self.original.settings,
                &self.settings,
                self.original.conformance,
            )?
        };
        let recipients_xml = match (&self.original.recipients, &self.recipients) {
            (Some(before), Some(after)) if before != after => Some(rewrite_recipients(
                self.original
                    .recipients_xml
                    .as_ref()
                    .ok_or_else(|| invalid("recipient metadata has no source XML"))?
                    .as_slice(),
                before,
                after,
                self.original.conformance,
            )?),
            (None, Some(after)) => Some(after.to_xml(self.original.conformance)?.into_bytes()),
            (_, None) => None,
            (Some(_), Some(_)) => self
                .original
                .recipients_xml
                .as_ref()
                .map(|xml| xml.as_ref().clone()),
        };
        let snapshot = Snapshot::from_parts(settings_xml, recipients_xml)?;
        let patch = Patch {
            before: self.original,
            after: snapshot.clone(),
        };
        Ok(Commit {
            snapshot,
            patch,
            changed: true,
        })
    }
}

/// A successful typed mail-merge publication.
#[derive(Debug, Clone)]
pub struct Commit {
    snapshot: Snapshot,
    patch: Patch,
    changed: bool,
}

impl Commit {
    #[inline]
    #[must_use]
    pub const fn changed(&self) -> bool {
        self.changed
    }

    #[inline]
    #[must_use]
    pub const fn snapshot(&self) -> &Snapshot {
        &self.snapshot
    }

    #[inline]
    #[must_use]
    pub const fn patch(&self) -> &Patch {
        &self.patch
    }

    #[must_use]
    pub fn into_parts(self) -> (Snapshot, Patch) {
        (self.snapshot, self.patch)
    }

    #[must_use]
    pub fn into_patch(self) -> Patch {
        self.patch
    }
}

/// A reversible source-checked typed patch.
#[derive(Debug, Clone)]
pub struct Patch {
    before: Snapshot,
    after: Snapshot,
}

impl Patch {
    #[inline]
    #[must_use]
    pub const fn before(&self) -> &Snapshot {
        &self.before
    }

    #[inline]
    #[must_use]
    pub const fn after(&self) -> &Snapshot {
        &self.after
    }

    #[inline]
    #[must_use]
    pub const fn expected_revision(&self) -> Revision {
        self.before.revision
    }

    #[inline]
    #[must_use]
    pub fn is_empty(&self) -> bool {
        self.before.same_source(&self.after)
    }

    #[must_use]
    pub fn inverse(&self) -> Self {
        Self {
            before: self.after.clone(),
            after: self.before.clone(),
        }
    }

    /// Apply only to the exact captured source snapshot.
    pub fn apply(&self, source: &Snapshot) -> Result<Snapshot> {
        if !source.same_source(&self.before) {
            return Err(invalid("mail-merge patch source is stale"));
        }
        Ok(self.after.clone())
    }
}

fn rewrite_settings(
    source: &[u8],
    before: &Settings,
    after: &Settings,
    conformance: Conformance,
) -> Result<Vec<u8>> {
    if super::parse_settings_mail_merge(source)?.as_ref() != Some(before) {
        return Err(invalid("mail-merge settings source is stale"));
    }
    let layout = scan(source)?;
    let mail = layout
        .mail_merge
        .ok_or_else(|| invalid("mailMerge element is missing from its settings source"))?;
    let replacement = opaque_preserving_fragment(
        source,
        &layout,
        mail,
        after.to_xml(conformance)?.into_bytes(),
    );
    let mut output = Vec::with_capacity(source.len() - (mail.end - mail.start) + replacement.len());
    output.extend_from_slice(&source[..mail.start]);
    output.extend_from_slice(&replacement);
    output.extend_from_slice(&source[mail.end..]);
    Ok(output)
}

fn rewrite_recipients(
    source: &[u8],
    before: &Recipients,
    after: &Recipients,
    conformance: Conformance,
) -> Result<Vec<u8>> {
    if Recipients::parse_xml(source)? != *before {
        return Err(invalid("mail-merge recipient source is stale"));
    }
    let layout = scan(source)?;
    let root = layout
        .root
        .ok_or_else(|| invalid("recipient root is missing"))?;
    let canonical = after.to_xml(conformance)?.into_bytes();
    let replacement = opaque_preserving_fragment(source, &layout, root, canonical);
    let mut output = Vec::with_capacity(source.len() - (root.end - root.start) + replacement.len());
    output.extend_from_slice(&source[..root.start]);
    output.extend_from_slice(&replacement);
    output.extend_from_slice(&source[root.end..]);
    Ok(output)
}

fn opaque_preserving_fragment(
    source: &[u8],
    layout: &Layout,
    root: Span,
    mut canonical: Vec<u8>,
) -> Vec<u8> {
    let mut opaque = Vec::new();
    for span in &layout.spans {
        if span.parent == Some(root.index) && !span.word {
            opaque.extend_from_slice(&source[span.start..span.end]);
        }
    }
    if opaque.is_empty() {
        return canonical;
    }
    if let Some(close) = canonical.windows(2).position(|window| window == b"</") {
        canonical.splice(close..close, opaque);
    }
    canonical
}

#[derive(Debug, Clone, Copy)]
struct Span {
    index: usize,
    start: usize,
    end: usize,
    parent: Option<usize>,
    word: bool,
}

#[derive(Debug, Default)]
struct Layout {
    spans: Vec<Span>,
    root: Option<Span>,
    mail_merge: Option<Span>,
}

fn scan(xml: &[u8]) -> Result<Layout> {
    let mut reader = NsReader::from_reader(xml);
    reader.config_mut().check_end_names = true;
    reader.config_mut().trim_text(false);
    let mut layout = Layout::default();
    let mut stack = Vec::new();
    loop {
        let start = usize::try_from(reader.buffer_position())
            .map_err(|_| invalid("mail-merge XML offset is too large"))?;
        let event = reader
            .read_event()
            .map_err(|error| Error::Xml(error.to_string()))?
            .into_owned();
        let end = usize::try_from(reader.buffer_position())
            .map_err(|_| invalid("mail-merge XML offset is too large"))?;
        let resolver = reader.resolver().clone();
        let (namespace, event) = resolver.resolve_event(event);
        match event {
            Event::Start(_element) => {
                let index = layout.spans.len();
                let span = Span {
                    index,
                    start,
                    end: 0,
                    parent: stack.last().copied(),
                    word: is_word(&namespace),
                };
                layout.spans.push(span);
                if stack.is_empty() {
                    layout.root = Some(span);
                }
                stack.push(index);
            },
            Event::Empty(element) => {
                let index = layout.spans.len();
                let span = Span {
                    index,
                    start,
                    end,
                    parent: stack.last().copied(),
                    word: is_word(&namespace),
                };
                layout.spans.push(span);
                if stack.is_empty() {
                    layout.root = Some(span);
                }
                let _ = element;
            },
            Event::End(_) => {
                let index = stack
                    .pop()
                    .ok_or_else(|| invalid("mail-merge XML has an unexpected end"))?;
                layout.spans[index].end = end;
            },
            Event::Eof => break,
            _ => {},
        }
    }
    if !stack.is_empty() {
        return Err(invalid("mail-merge XML has an unterminated element"));
    }
    let root = layout
        .root
        .ok_or_else(|| invalid("mail-merge XML has no root"))?;
    let root_index = root.index;
    let root = layout.spans[root_index];
    layout.root = Some(root);
    layout.mail_merge = layout.spans.iter().copied().find(|span| {
        span.parent == Some(root_index) && span.word && local_name(xml, *span) == b"mailMerge"
    });
    if root.word && local_name(xml, root) == b"recipients" {
        layout.mail_merge = None;
    }
    Ok(layout)
}

fn local_name(xml: &[u8], span: Span) -> &[u8] {
    let bytes = &xml[span.start..span.end.min(xml.len())];
    let begin = bytes
        .iter()
        .position(|byte| *byte == b'<')
        .map_or(0, |i| i + 1);
    let end = bytes[begin..]
        .iter()
        .position(|byte| byte.is_ascii_whitespace() || *byte == b'/' || *byte == b'>')
        .map_or(bytes.len(), |i| begin + i);
    let name = &bytes[begin..end];
    name.rsplit(|byte| *byte == b':').next().unwrap_or(name)
}

fn is_word(namespace: &ResolveResult<'_>) -> bool {
    matches!(
        namespace,
        ResolveResult::Bound(Namespace(value))
            if *value == W.as_bytes() || *value == STRICT_W.as_bytes()
    )
}

fn detect_conformance(xml: &[u8]) -> Result<Conformance> {
    if xml
        .windows(STRICT_W.len())
        .any(|window| window == STRICT_W.as_bytes())
    {
        Ok(Conformance::Strict)
    } else if xml.windows(W.len()).any(|window| window == W.as_bytes()) {
        Ok(Conformance::Transitional)
    } else {
        Err(invalid("mail-merge XML has no recognized Word namespace"))
    }
}

fn fingerprint(xml: &[u8], recipients_xml: Option<&[u8]>) -> Revision {
    let mut hash = 0xcbf29ce484222325u64;
    for bytes in [Some(xml), recipients_xml].into_iter().flatten() {
        for byte in bytes {
            hash ^= u64::from(*byte);
            hash = hash.wrapping_mul(0x100000001b3);
        }
    }
    hash
}

fn invalid(message: impl Into<String>) -> Error {
    Error::Invalid(message.into())
}
