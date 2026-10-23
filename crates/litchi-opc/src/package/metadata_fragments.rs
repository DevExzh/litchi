//! Shared bounded lexical transport for OPC metadata entries.

use super::relationships::{clone_string_bounded, relationship_id};
use crate::xml_attributes::BytesStartExt as _;
use crate::{OpcError, OwnedXmlPart, PackURI, Result, SourceLimitKind, SourceLimits};
use litchi_core::xml::ReaderOrigin;
use quick_xml::{
    events::Event,
    name::{Namespace, ResolveResult},
    reader::NsReader,
};
use std::{
    collections::{HashMap, HashSet},
    ops::Range,
    sync::Arc,
};

#[derive(Debug)]
struct Entry {
    id: String,
    span: Range<usize>,
    events: usize,
}

/// Bounded lexical entries decoded from one metadata member.
///
/// Decoding grants no package replacement authority. Fragment edits must still
/// match every touched destination entry and its inherited namespace context.
pub(super) struct Fragments {
    kind: Kind,
    bytes: Arc<Vec<u8>>,
    entries: Vec<Entry>,
    by_id: HashMap<String, usize>,
    namespaces: Vec<(String, String)>,
    root: Range<usize>,
    root_name: Vec<u8>,
    close: usize,
    empty: bool,
    events: usize,
}

impl std::fmt::Debug for Fragments {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        f.debug_struct("Fragments")
            .field("bytes", &self.bytes.len())
            .field("entries", &self.entries.len())
            .finish_non_exhaustive()
    }
}

/// Borrowed exact entry, including its following-entry and trailing-byte anchor.
#[derive(Clone, Copy, Debug)]
pub(super) struct Fragment<'a> {
    catalog: &'a Fragments,
    index: usize,
}

impl<'a> Fragment<'a> {
    /// Private native key within this manifest.
    pub(super) fn id(&self) -> &'a str {
        &self.catalog.entries[self.index].id
    }
    pub(super) fn byte_range(&self) -> Range<usize> {
        self.catalog.entries[self.index].span.clone()
    }
    /// Exact element bytes, including attributes, paired tags and child markup.
    pub(super) fn bytes(&self) -> &'a [u8] {
        &self.catalog.bytes[self.catalog.entries[self.index].span.clone()]
    }
    pub(super) fn kind(&self) -> &'a Kind {
        &self.catalog.kind
    }
    /// Next entry in source order, or no following entry.
    pub(super) fn following_id(&self) -> Option<&'a str> {
        self.catalog
            .entries
            .get(self.index + 1)
            .map(|entry| entry.id.as_str())
    }
    /// Exact bytes between this element and its following anchor.
    pub(super) fn trailing_bytes(&self) -> &'a [u8] {
        let end = self
            .catalog
            .entries
            .get(self.index + 1)
            .map_or(self.catalog.close, |entry| entry.span.start);
        &self.catalog.bytes[self.catalog.entries[self.index].span.end..end]
    }
    // Prove the terminal insertion slot from retained left-hand gaps. This also
    // restores a deleted suffix before unrelated entries appended afterward.
    fn terminal_position(
        &self,
        destination: &Fragments,
        additions: &HashMap<&str, usize>,
    ) -> Option<usize> {
        if let Some(current) = destination.get(self.id()) {
            return Some(destination.entries[current.index].span.end);
        }
        let mut first = self.index;
        let mut position = destination
            .root
            .end
            .saturating_sub(usize::from(destination.empty) * 2);
        while first > 0 {
            let previous = &self.catalog.entries[first - 1];
            if let Some(current) = destination.get(&previous.id) {
                position = destination.entries[current.index].span.end;
                break;
            }
            if !additions.contains_key(previous.id.as_str()) {
                return None;
            }
            first -= 1;
        }
        for index in first..=self.index {
            let start = if index == 0 {
                self.catalog.root.end
            } else {
                self.catalog.entries[index - 1].span.end
            };
            let gap = &self.catalog.bytes[start..self.catalog.entries[index].span.start];
            let end = position.checked_add(gap.len())?;
            if destination.bytes.get(position..end) != Some(gap) {
                return None;
            }
            position = end;
        }
        Some(position)
    }
}

impl Fragments {
    /// Decode untrusted member XML under explicit byte, event and graph bounds.
    /// Duplicate native keys and nonstandard lexical roots are not editable.
    pub(super) fn from_bytes(
        kind: Kind,
        bytes: Arc<Vec<u8>>,
        limits: SourceLimits,
    ) -> Result<Self> {
        let uri = kind.uri(limits)?;
        OwnedXmlPart::check_capture_size(&uri, bytes.len(), limits)?;
        let parsed = kind.parse(&bytes, limits)?;
        OwnedXmlPart::preflight_capture_with_limits(&uri, kind.content_type(), &bytes, limits)?;
        let mut entries = Vec::new();
        entries
            .try_reserve_exact(parsed.count())
            .map_err(allocation)?;
        let mut by_id = HashMap::new();
        by_id.try_reserve(parsed.count()).map_err(allocation)?;
        let mut namespaces = Vec::new();
        let mut root = None;
        let mut root_name = Vec::new();
        let mut close = None;
        let mut empty = false;
        let mut active: Option<(String, usize, usize)> = None;
        let mut depth = 0usize;
        let mut events = 0usize;
        let mut reader = NsReader::from_reader(bytes.as_slice());
        reader.config_mut().trim_text(false);
        let origin = ReaderOrigin::of(&bytes);
        loop {
            events = events
                .checked_add(1)
                .ok_or_else(|| kind.invalid("metadata fragment events overflow"))?;
            limits.check(
                SourceLimitKind::Events,
                events as u64,
                limits.max_events() as u64,
            )?;
            let start = origin
                .offset(reader.buffer_position())
                .ok_or_else(|| kind.invalid("metadata XML offset overflow"))?;
            let (resolved, event) = reader.read_resolved_event()?;
            let related = matches!(resolved, ResolveResult::Bound(Namespace(value)) if value == kind.namespace());
            let decoder = reader.decoder();
            let end = origin
                .offset(reader.buffer_position())
                .ok_or_else(|| kind.invalid("metadata XML offset overflow"))?;
            match event {
                Event::Start(element) | Event::Empty(element) => {
                    let is_empty = bytes.get(end.saturating_sub(2)..end) == Some(b"/>");
                    if depth == 0 {
                        if root.is_some()
                            || !related
                            || element.local_name().as_ref() != kind.root_name()
                        {
                            return Err(kind.invalid(
                                "metadata fragments require the expected namespaced root",
                            ));
                        }
                        root = Some(start..end);
                        root_name
                            .try_reserve_exact(element.name().as_ref().len())
                            .map_err(allocation)?;
                        root_name.extend_from_slice(element.name().as_ref());
                        namespaces.try_reserve(1).map_err(allocation)?;
                        namespaces.push((String::new(), String::new()));
                        for attribute in element.checked_attributes() {
                            let attribute =
                                attribute.map_err(|error| kind.invalid(error.to_string()))?;
                            let key = attribute.key.as_ref();
                            let prefix = if key == b"xmlns" {
                                Some(b"".as_slice())
                            } else {
                                key.strip_prefix(b"xmlns:").or_else(|| {
                                    matches!(key, b"xml:base" | b"xml:lang" | b"xml:space")
                                        .then_some(key)
                                })
                            };
                            if let Some(prefix) = prefix {
                                let prefix = std::str::from_utf8(prefix)
                                    .map_err(|error| kind.invalid(error.to_string()))?;
                                let value = attribute
                                    .decoded_and_normalized_value(
                                        quick_xml::XmlVersion::Implicit1_0,
                                        decoder,
                                    )
                                    .map_err(|error| kind.invalid(error.to_string()))?;
                                let value = clone_string_bounded(&value, "OPC metadata namespace")?;
                                if prefix.is_empty() {
                                    namespaces[0].1 = value;
                                } else {
                                    namespaces.try_reserve(1).map_err(allocation)?;
                                    namespaces.push((
                                        clone_string_bounded(prefix, "OPC metadata prefix")?,
                                        value,
                                    ));
                                }
                            }
                        }
                        namespaces.sort_unstable();
                        empty = is_empty;
                        if empty {
                            close = Some(end - 2);
                        }
                    } else if depth == 1 && related && kind.is_entry(element.local_name().as_ref())
                    {
                        let id = kind.entry_id(&element, decoder)?;
                        if by_id.contains_key(&id) || !parsed.contains(&id) {
                            return Err(
                                kind.invalid("metadata fragment identity is duplicate or repaired")
                            );
                        }
                        by_id.try_reserve(1).map_err(allocation)?;
                        by_id.insert(
                            clone_string_bounded(&id, "OPC metadata fragment ID")?,
                            entries.len(),
                        );
                        if is_empty {
                            entries.try_reserve(1).map_err(allocation)?;
                            entries.push(Entry {
                                id,
                                span: start..end,
                                events: 1,
                            });
                        } else {
                            active = Some((id, start, events));
                        }
                    }
                    if !is_empty {
                        depth = depth
                            .checked_add(1)
                            .ok_or_else(|| kind.invalid("metadata fragment depth overflow"))?;
                        limits.check(
                            SourceLimitKind::Depth,
                            depth as u64,
                            limits.max_depth() as u64,
                        )?;
                    }
                },
                Event::End(_) => {
                    if depth == 2 {
                        if let Some((id, start, first)) = active.take() {
                            entries.try_reserve(1).map_err(allocation)?;
                            entries.push(Entry {
                                id,
                                span: start..end,
                                events: events - first + 1,
                            });
                        }
                    }
                    if depth == 1 {
                        close = Some(start);
                    }
                    depth = depth
                        .checked_sub(1)
                        .ok_or_else(|| kind.invalid("metadata fragment depth underflow"))?;
                },
                Event::Eof => break,
                _ => {},
            }
        }
        if entries.len() != parsed.count() || depth != 0 {
            return Err(kind.invalid("metadata fragment catalog differs from the parsed graph"));
        }
        Ok(Self {
            bytes,
            entries,
            by_id,
            namespaces,
            root: root.ok_or_else(|| kind.invalid("missing metadata root"))?,
            root_name,
            close: close.ok_or_else(|| kind.invalid("missing metadata root close"))?,
            empty,
            events,
            kind,
        })
    }

    /// Exact source-order entry views, without copying their XML.
    pub(super) fn iter(&self) -> impl ExactSizeIterator<Item = Fragment<'_>> {
        (0..self.entries.len()).map(|index| Fragment {
            catalog: self,
            index,
        })
    }
    /// Resolve one normalized native metadata key.
    pub(super) fn get(&self, id: &str) -> Option<Fragment<'_>> {
        self.by_id.get(id).map(|&index| Fragment {
            catalog: self,
            index,
        })
    }
}

/// One guarded entry delta. Replacement IDs and owners must agree.
#[derive(Clone, Copy, Debug)]
pub(super) struct FragmentEdit<'a> {
    pub before: Option<Fragment<'a>>,
    pub after: Option<Fragment<'a>>,
}

#[derive(Debug)]
struct Splice<'a> {
    span: Range<usize>,
    output: &'a [u8],
    rank: usize,
    id: &'a str,
}

/// Source-bound batched lexical edit with exact byte/count admission.
///
/// The event bound is conservative: removing an element may merge adjacent
/// character-data events. No output buffer is allocated until materialization.
pub(super) struct FragmentPlan<'a> {
    source: &'a OwnedXmlPart,
    catalog: Fragments,
    splices: Vec<Splice<'a>>,
    final_len: usize,
    entry_count: usize,
    event_bound: usize,
    noop: bool,
    limits: SourceLimits,
}

impl std::fmt::Debug for FragmentPlan<'_> {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        f.debug_struct("FragmentPlan")
            .field("final_len", &self.final_len)
            .field("entry_count", &self.entry_count)
            .field("event_upper_bound", &self.event_bound)
            .field("noop", &self.noop)
            .finish_non_exhaustive()
    }
}

impl<'a> FragmentPlan<'a> {
    pub(super) fn new(
        source: &'a OwnedXmlPart,
        kind: Kind,
        edits: &[FragmentEdit<'a>],
        limits: SourceLimits,
    ) -> Result<Self> {
        // A durable plan may tighten the source token's admission, but a
        // caller-provided profile must never widen the policy that admitted
        // the token in the first place.
        let limits = source.limits().min(limits);
        kind.admit_edit_count(edits.len(), limits)?;
        let catalog = Fragments::from_bytes(kind.clone(), Arc::clone(&source.bytes), limits)?;
        let mut touched = HashSet::new();
        touched.try_reserve(edits.len()).map_err(allocation)?;
        let mut additions = HashMap::new();
        additions.try_reserve(edits.len()).map_err(allocation)?;
        let mut splices = Vec::new();
        splices
            .try_reserve_exact(
                edits
                    .len()
                    .checked_mul(2)
                    .ok_or_else(|| kind.invalid("metadata splice count overflow"))?,
            )
            .map_err(allocation)?;
        let mut final_len = source.bytes().len();
        let mut entry_count = catalog.entries.len();
        let mut event_bound = catalog.events;
        for (index, edit) in edits.iter().enumerate() {
            let selected = edit
                .before
                .or(edit.after)
                .ok_or_else(|| kind.invalid("empty metadata fragment edit"))?;
            if !touched.insert(selected.id()) {
                return Err(kind.invalid("duplicate metadata fragment edit"));
            }
            for fragment in [edit.before, edit.after].into_iter().flatten() {
                if fragment.kind() != &kind || fragment.id() != selected.id() {
                    return Err(kind.invalid("metadata fragment owner or ID differs"));
                }
                if fragment.catalog.namespaces != catalog.namespaces {
                    return Err(kind.invalid("metadata fragment namespace context is stale"));
                }
            }
            let current = catalog.get(selected.id());
            match (edit.before, current) {
                (Some(before), Some(current)) if before.bytes() == current.bytes() => {
                    let entry = &catalog.entries[current.index];
                    final_len = final_len
                        .checked_sub(entry.span.len())
                        .ok_or_else(|| kind.invalid("metadata fragment byte underflow"))?;
                    event_bound = event_bound
                        .checked_sub(entry.events)
                        .ok_or_else(|| kind.invalid("metadata fragment event underflow"))?;
                    entry_count -= 1;
                    splices.push(Splice {
                        span: entry.span.clone(),
                        output: edit.after.map_or(&[], |after| after.bytes()),
                        rank: 0,
                        id: selected.id(),
                    });
                },
                (None, None) => {},
                _ => return Err(kind.invalid("metadata fragment before-state is stale")),
            }
            if let Some(after) = edit.after {
                if edit.before.is_none() {
                    additions.insert(after.id(), index);
                }
                final_len = final_len
                    .checked_add(after.bytes().len())
                    .ok_or_else(|| kind.invalid("metadata fragment output overflows"))?;
                event_bound = event_bound
                    .checked_add(after.catalog.entries[after.index].events)
                    .ok_or_else(|| kind.invalid("metadata fragment event overflow"))?;
                entry_count = entry_count
                    .checked_add(1)
                    .ok_or_else(|| kind.invalid("metadata fragment count overflow"))?;
            }
        }
        let mut resolved = Vec::new();
        resolved
            .try_reserve_exact(edits.len())
            .map_err(allocation)?;
        resolved.resize(edits.len(), None);
        let mut visiting = Vec::new();
        visiting
            .try_reserve_exact(edits.len())
            .map_err(allocation)?;
        visiting.resize(edits.len(), false);
        let mut stack = Vec::new();
        stack.try_reserve_exact(edits.len()).map_err(allocation)?;
        for index in 0..edits.len() {
            if edits[index].before.is_some()
                || edits[index].after.is_none()
                || resolved[index].is_some()
            {
                continue;
            }
            let mut next = index;
            loop {
                if resolved[next].is_some() {
                    break;
                }
                if visiting[next] {
                    return Err(kind.invalid("cyclic metadata fragment anchors"));
                }
                visiting[next] = true;
                stack.push(next);
                let after = edits[next]
                    .after
                    .ok_or_else(|| kind.invalid("missing metadata insertion"))?;
                if let Some(other) = after.following_id().and_then(|id| additions.get(id)) {
                    next = *other;
                } else {
                    break;
                }
            }
            while let Some(index) = stack.pop() {
                let after = edits[index]
                    .after
                    .ok_or_else(|| kind.invalid("missing metadata insertion"))?;
                let (anchor, rank) = match after.following_id() {
                    Some(id) if additions.contains_key(id) => resolved[*additions
                        .get(id)
                        .ok_or_else(|| kind.invalid("missing insertion anchor"))?]
                    .ok_or_else(|| kind.invalid("unresolved insertion anchor"))?,
                    Some(id) => (
                        catalog.entries[*catalog
                            .by_id
                            .get(id)
                            .ok_or_else(|| kind.invalid("metadata insertion anchor is missing"))?]
                        .span
                        .start,
                        0usize,
                    ),
                    None => (
                        after
                            .terminal_position(&catalog, &additions)
                            .and_then(|position| position.checked_add(after.trailing_bytes().len()))
                            .ok_or_else(|| {
                                kind.invalid("metadata insertion left-byte anchor is stale")
                            })?,
                        0usize,
                    ),
                };
                let tail = after.trailing_bytes();
                let position = anchor
                    .checked_sub(tail.len())
                    .ok_or_else(|| kind.invalid("metadata insertion anchor underflow"))?;
                if source.bytes().get(position..anchor) != Some(tail)
                    || position
                        < catalog
                            .root
                            .end
                            .saturating_sub(usize::from(catalog.empty) * 2)
                {
                    return Err(kind.invalid("metadata insertion trailing-byte anchor is stale"));
                }
                let previous = catalog
                    .entries
                    .partition_point(|entry| entry.span.start < position);
                if previous > 0 && catalog.entries[previous - 1].span.end > position {
                    return Err(kind.invalid("metadata insertion anchor overlaps an entry"));
                }
                resolved[index] = Some((
                    position,
                    rank.checked_add(1)
                        .ok_or_else(|| kind.invalid("metadata anchor rank overflow"))?,
                ));
                visiting[index] = false;
            }
        }
        for (index, edit) in edits.iter().enumerate() {
            if edit.before.is_some() {
                continue;
            }
            if let Some(after) = edit.after {
                let (position, rank) =
                    resolved[index].ok_or_else(|| kind.invalid("unresolved metadata insertion"))?;
                splices.push(Splice {
                    span: position..position,
                    output: after.bytes(),
                    rank,
                    id: after.id(),
                });
            }
        }
        splices.sort_unstable_by(|left, right| {
            left.span
                .start
                .cmp(&right.span.start)
                .then_with(|| left.span.is_empty().cmp(&right.span.is_empty()).reverse())
                .then_with(|| right.rank.cmp(&left.rank))
                .then_with(|| left.id.cmp(right.id))
        });
        let mut cursor = 0;
        for splice in &splices {
            if splice.span.start < cursor {
                return Err(kind.invalid("metadata fragment edits overlap"));
            }
            cursor = splice.span.end;
        }
        if catalog.empty && !additions.is_empty() {
            final_len = final_len
                .checked_add(catalog.root_name.len())
                .and_then(|value| value.checked_add(2))
                .ok_or_else(|| kind.invalid("metadata root expansion overflow"))?;
            event_bound = event_bound
                .checked_add(1)
                .ok_or_else(|| kind.invalid("metadata event overflow"))?;
        }
        let mut plan = FragmentPlan {
            source,
            catalog,
            splices,
            final_len,
            entry_count,
            event_bound,
            noop: false,
            limits,
        };
        plan.admit(limits)?;
        let mut offset = 0usize;
        let mut equal = true;
        plan.chunks(|chunk| {
            equal &= source
                .bytes()
                .get(offset..offset.saturating_add(chunk.len()))
                == Some(chunk);
            offset = offset
                .checked_add(chunk.len())
                .ok_or_else(|| kind.invalid("metadata output offset overflow"))?;
            Ok(())
        })?;
        plan.noop = equal && offset == source.bytes().len();
        Ok(plan)
    }
}

impl FragmentPlan<'_> {
    pub(super) const fn final_len(&self) -> usize {
        self.final_len
    }
    pub(super) const fn entry_count(&self) -> usize {
        self.entry_count
    }
    pub(super) const fn event_upper_bound(&self) -> usize {
        self.event_bound
    }
    pub(super) const fn is_noop(&self) -> bool {
        self.noop
    }
    fn admit(&self, limits: SourceLimits) -> Result<()> {
        OwnedXmlPart::check_capture_size(&self.source.name, self.final_len, limits)?;
        self.catalog
            .kind
            .admit(self.final_len, self.entry_count, self.event_bound, limits)
    }

    fn chunks(&self, mut emit: impl FnMut(&[u8]) -> Result<()>) -> Result<()> {
        let source = self.source.bytes();
        if self.catalog.empty && !self.splices.is_empty() {
            emit(&source[..self.catalog.close])?;
            emit(b">")?;
            for splice in &self.splices {
                emit(splice.output)?;
            }
            emit(b"</")?;
            emit(&self.catalog.root_name)?;
            emit(b">")?;
            return emit(&source[self.catalog.root.end..]);
        }
        let mut cursor = 0;
        for splice in &self.splices {
            emit(&source[cursor..splice.span.start])?;
            emit(splice.output)?;
            cursor = splice.span.end;
        }
        emit(&source[cursor..])
    }

    /// Revalidate the final graph and return an owned XML token plus binding.
    pub(super) fn materialize(self, limits: SourceLimits) -> Result<(OwnedXmlPart, Parsed)> {
        let limits = self.limits.min(limits);
        self.admit(limits)?;
        let kind = &self.catalog.kind;
        if self.noop {
            OwnedXmlPart::preflight_capture_with_limits(
                &self.source.name,
                &self.source.content_type,
                self.source.bytes(),
                limits,
            )?;
            let parsed = kind.parse(self.source.bytes(), limits)?;
            return Ok((self.source.with_limits(limits), parsed));
        }
        let mut output = Vec::new();
        output
            .try_reserve_exact(self.final_len)
            .map_err(allocation)?;
        self.chunks(|chunk| {
            output.extend_from_slice(chunk);
            Ok(())
        })?;
        if output.len() != self.final_len {
            return Err(kind.invalid("metadata fragment output size differs from plan"));
        }
        let parsed = kind.parse(&output, limits)?;
        if parsed.count() != self.entry_count {
            return Err(kind.invalid("metadata fragment graph differs from plan"));
        }
        let xml = OwnedXmlPart::capture_with_limits(
            self.source.name.clone(),
            self.source.content_type.clone(),
            Arc::new(output),
            limits,
        )?;
        Ok((xml, parsed))
    }
}

fn allocation(source: std::collections::TryReserveError) -> OpcError {
    OpcError::Allocation {
        resource: "OPC metadata fragments",
        source,
    }
}

#[derive(Clone, Debug, PartialEq, Eq)]
pub(super) enum Kind {
    Relationships(PackURI),
    ContentTypes,
}

pub(super) enum Parsed {
    Relationships(crate::Relationships),
    ContentTypes(crate::content_type::ContentTypeMap),
}

impl Parsed {
    fn count(&self) -> usize {
        match self {
            Self::Relationships(value) => value.len(),
            Self::ContentTypes(value) => value.mapping_count(),
        }
    }
    fn contains(&self, key: &str) -> bool {
        match self {
            Self::Relationships(value) => value.get(key).is_some(),
            Self::ContentTypes(value) => key
                .split_at_checked(1)
                .is_some_and(|(kind, key)| value.contains_fragment_key(kind == "D", key)),
        }
    }
}

impl Kind {
    pub(super) fn admit_edit_count(&self, count: usize, limits: SourceLimits) -> Result<()> {
        let maximum = self
            .max_entries(limits)
            .checked_mul(2)
            .ok_or_else(|| self.invalid("metadata edit count overflow"))?;
        limits.check(self.count_limit(), count as u64, maximum as u64)
    }

    fn invalid(&self, message: impl Into<String>) -> OpcError {
        match self {
            Self::Relationships(_) => OpcError::InvalidRelationship(message.into()),
            Self::ContentTypes => OpcError::InvalidContentTypesManifest(message.into()),
        }
    }
    fn uri(&self, limits: SourceLimits) -> Result<PackURI> {
        match self {
            Self::Relationships(owner) => {
                OwnedXmlPart::check_derived_capture_member_name(owner, limits)?;
                owner.rels_uri().map_err(OpcError::InvalidPackUri)
            },
            Self::ContentTypes => {
                let name = PackURI::new(crate::packuri::CONTENT_TYPES_URI)
                    .map_err(OpcError::InvalidPackUri)?;
                OwnedXmlPart::check_capture_member_name(&name, limits)?;
                Ok(name)
            },
        }
    }
    fn namespace(&self) -> &'static [u8] {
        match self {
            Self::Relationships(_) => {
                b"http://schemas.openxmlformats.org/package/2006/relationships"
            },
            Self::ContentTypes => b"http://schemas.openxmlformats.org/package/2006/content-types",
        }
    }
    fn root_name(&self) -> &'static [u8] {
        match self {
            Self::Relationships(_) => b"Relationships",
            Self::ContentTypes => b"Types",
        }
    }
    fn is_entry(&self, name: &[u8]) -> bool {
        match self {
            Self::Relationships(_) => name == b"Relationship",
            Self::ContentTypes => matches!(name, b"Default" | b"Override"),
        }
    }
    fn content_type(&self) -> &'static str {
        match self {
            Self::Relationships(_) => crate::constants::content_type::OPC_RELATIONSHIPS,
            Self::ContentTypes => crate::constants::content_type::XML,
        }
    }
    fn entry_id(
        &self,
        element: &quick_xml::events::BytesStart<'_>,
        decoder: quick_xml::Decoder,
    ) -> Result<String> {
        if matches!(self, Self::Relationships(_)) {
            return relationship_id(element, decoder);
        }
        let default = element.local_name().as_ref() == b"Default";
        let attribute_name: &[u8] = if default { b"Extension" } else { b"PartName" };
        for attribute in element.checked_attributes() {
            let attribute = attribute?;
            if attribute.key.as_ref() == attribute_name {
                let value = attribute
                    .decoded_and_normalized_value(quick_xml::XmlVersion::Implicit1_0, decoder)?;
                return content_key(default, &value);
            }
        }
        Err(self.invalid("content-type fragment key is missing"))
    }
    fn parse(&self, bytes: &[u8], limits: SourceLimits) -> Result<Parsed> {
        match self {
            Self::Relationships(owner) => {
                crate::pkgreader::PackageReader::parse_owned_relationships(bytes, owner, limits)
                    .map(Parsed::Relationships)
            },
            Self::ContentTypes => crate::content_type::ContentTypeMap::from_xml(bytes, limits)
                .map(Parsed::ContentTypes),
        }
    }
    fn count_limit(&self) -> SourceLimitKind {
        match self {
            Self::Relationships(_) => SourceLimitKind::RelationshipsPerPart,
            Self::ContentTypes => SourceLimitKind::ContentTypeMappings,
        }
    }
    fn max_entries(&self, limits: SourceLimits) -> usize {
        match self {
            Self::Relationships(_) => limits.max_relationships_per_part(),
            Self::ContentTypes => limits.max_content_type_mappings(),
        }
    }
    fn admit(
        &self,
        bytes: usize,
        entries: usize,
        events: usize,
        limits: SourceLimits,
    ) -> Result<()> {
        limits.check(
            self.count_limit(),
            entries as u64,
            self.max_entries(limits) as u64,
        )?;
        limits.check(
            SourceLimitKind::Events,
            events as u64,
            limits.max_events() as u64,
        )?;
        match self {
            Self::Relationships(_) => {
                limits.check(
                    SourceLimitKind::RelationshipXmlBytes,
                    bytes as u64,
                    limits.max_relationship_xml_bytes() as u64,
                )?;
                limits.check(
                    SourceLimitKind::TotalRelationshipXmlBytes,
                    bytes as u64,
                    limits.max_total_relationship_xml_bytes() as u64,
                )?;
                limits.check(
                    SourceLimitKind::TotalRelationships,
                    entries as u64,
                    limits.max_total_relationships() as u64,
                )?;
                limits.check(
                    SourceLimitKind::TotalRelationshipXmlEvents,
                    events as u64,
                    limits.max_total_relationship_xml_events() as u64,
                )?;
            },
            Self::ContentTypes => limits.check(
                SourceLimitKind::ContentTypesBytes,
                bytes as u64,
                limits.max_content_types_bytes() as u64,
            )?,
        }
        Ok(())
    }
}

/// Prefixes keep extension and part-name key spaces separate; spelling follows
/// ContentTypeMap exactly (XML attribute normalization, then ASCII folding).
pub(super) fn content_key(default: bool, value: &str) -> Result<String> {
    let mut key = String::new();
    let length = value.len().checked_add(1).ok_or_else(|| {
        OpcError::InvalidContentTypesManifest("content-type key length overflow".into())
    })?;
    key.try_reserve_exact(length).map_err(allocation)?;
    key.push(if default { 'D' } else { 'O' });
    key.push_str(value);
    key[1..].make_ascii_lowercase();
    Ok(key)
}
