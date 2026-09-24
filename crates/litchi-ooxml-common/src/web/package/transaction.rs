use super::super::codec::{invalid, limit};
use super::super::model::Limits;
use super::super::{
    Arc, BTreeSet, Error, HashMap, HashSet, OpcPackage, PackURI, Part, Result,
    TASK_PANES_RELATIONSHIP, VecDeque, WEB_EXTENSION_NAMESPACE,
};
use super::{PlannedPart, PlannedRelationship, fold_part_name, folded_name_conflicts};
use litchi_opc::{
    BlobPart, OpcError, OwnedContentTypes, OwnedRelationships, OwnedXmlPart, TargetMode,
};
use quick_xml::events::Event;
use quick_xml::name::{Namespace, ResolveResult};
use quick_xml::reader::NsReader;
use std::ops::Range;
/// Opaque, exact, reversible task-pane graph transaction.
///
/// A patch records the precise source and destination state of every affected
/// relationship and part. Applying a stale patch fails before mutation. Empty
/// patches preserve package signatures; a changed signed package requires an
/// explicit signature disposition before this low-level patch is applied.
#[must_use = "a planned Web Extensions patch has no effect until it is applied"]
#[derive(Clone, Default)]
pub struct Patch {
    pub(in crate::web) root: Option<RootChange>,
    pub(in crate::web) parts: Box<[PartChange]>,
    pub(in crate::web) protection: Option<ProtectionGuard>,
    pub(in crate::web::package) lexical: Option<LexicalChange>,
    pub(in crate::web) durable_intent: Option<Arc<super::durable::WebIntent>>,
    pub(in crate::web) durable_reversed: bool,
}

pub(in crate::web) struct PatchPlan {
    pub(in crate::web) before: PlannedGraph,
    pub(in crate::web) after: PlannedGraph,
    pub(in crate::web) parts: Vec<PlannedPart>,
    pub(in crate::web) deletions: Vec<PackURI>,
    pub(in crate::web) limits: Limits,
}

pub(in crate::web) struct PlannedGraph {
    pub(in crate::web) root: Option<RelationshipState>,
    pub(in crate::web) owned_parts: Vec<PackURI>,
}

impl std::fmt::Debug for Patch {
    fn fmt(&self, formatter: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        formatter
            .debug_struct("Patch")
            .field("empty", &self.is_empty())
            .field(
                "affected_parts",
                &self
                    .parts
                    .iter()
                    .filter(|part| part.before != part.after)
                    .count(),
            )
            .finish_non_exhaustive()
    }
}

impl Patch {
    pub(in crate::web) fn planned(package: &OpcPackage, plan: PatchPlan) -> Result<Self> {
        let PatchPlan {
            before,
            after,
            parts: planned,
            deletions,
            limits,
        } = plan;
        let affected_parts = planned
            .len()
            .checked_add(deletions.len())
            .ok_or(Error::Limit {
                resource: "Web Extensions patch parts",
                max: usize::MAX,
                actual: usize::MAX,
            })?;
        let mut parts = Vec::with_capacity(affected_parts);
        let mut planned_names = HashSet::with_capacity(planned.len());
        for part in planned {
            let name = part.name.clone();
            planned_names.insert(fold_part_name(&name));
            let before = match package.get_part(&name) {
                Ok(part) => Some(PartState::capture(part)),
                Err(OpcError::PartNotFound(_)) => None,
                Err(error) => return Err(error.into()),
            };
            parts.push(PartChange {
                name,
                before,
                after: Some(PartState::from_planned(part)),
            });
        }
        for name in deletions {
            if planned_names.contains(&fold_part_name(&name)) {
                return invalid("planned Web Extensions part is also marked for deletion".into());
            }
            let before = match package.get_part(&name) {
                Ok(part) => PartState::capture(part),
                Err(OpcError::PartNotFound(_)) => {
                    return Err(Error::Missing(
                        "Web Extensions patch source part disappeared".into(),
                    ));
                },
                Err(error) => return Err(error.into()),
            };
            parts.push(PartChange {
                name,
                before: Some(before),
                after: None,
            });
        }
        parts.sort_by(|left, right| left.name.as_str().cmp(right.name.as_str()));
        let protection = ProtectionGuard {
            source: GraphScope {
                root_relationship_id: before.root.as_ref().map(|root| root.id.clone()),
                owned_parts: before.owned_parts.into_boxed_slice(),
            },
            destination: GraphScope {
                root_relationship_id: after.root.as_ref().map(|root| root.id.clone()),
                owned_parts: after.owned_parts.into_boxed_slice(),
            },
            limits,
        };
        let mut patch = Self {
            root: Some(RootChange {
                before: before.root,
                after: after.root,
            }),
            parts: parts.into_boxed_slice(),
            protection: Some(protection),
            lexical: None,
            durable_intent: None,
            durable_reversed: false,
        };
        // Keep the source proof even when the semantic request is an exact
        // no-op.  Planning callers use that proof to distinguish a no-op
        // against the captured package from an unbound `Patch::default()`.
        let source_parts = patch
            .protection
            .as_ref()
            .map(|guard| guard.source.owned_parts.as_ref())
            .unwrap_or(&[]);
        let destination_parts = patch
            .protection
            .as_ref()
            .map(|guard| guard.destination.owned_parts.as_ref())
            .unwrap_or(&[]);
        patch.lexical = Some(LexicalChange::capture(
            package,
            &patch.parts,
            patch.root.as_ref(),
            source_parts,
            destination_parts,
            limits,
        )?);
        Ok(patch)
    }

    /// Return whether this patch makes no changes.
    #[must_use]
    pub fn is_empty(&self) -> bool {
        !self.has_changes()
    }

    /// Build the exact inverse without copying part payloads.
    #[must_use = "the inverse has no effect until it is applied"]
    pub fn inverse(&self) -> Self {
        Self {
            root: self.root.as_ref().map(|root| RootChange {
                before: root.after.clone(),
                after: root.before.clone(),
            }),
            parts: self
                .parts
                .iter()
                .map(|part| PartChange {
                    name: part.name.clone(),
                    before: part.after.clone(),
                    after: part.before.clone(),
                })
                .collect::<Vec<_>>()
                .into_boxed_slice(),
            protection: self.protection.as_ref().map(ProtectionGuard::inverse),
            lexical: self.lexical.as_ref().map(LexicalChange::inverse),
            durable_intent: self.durable_intent.clone(),
            durable_reversed: !self.durable_reversed,
        }
    }

    /// Apply this patch after checking its exact source graph.
    ///
    /// Returns `true` when the package changed. Source checks, destination-name
    /// checks, and relationship staging all finish before the first mutation.
    /// # Errors
    ///
    /// Returns an error when input violates OOXML constraints, exceeds a configured
    /// bound, or an underlying XML or package operation fails.
    pub fn apply(&self, package: &mut OpcPackage) -> Result<bool> {
        if self.is_empty() {
            if self.is_source_bound() {
                self.validate_source(package)?;
                self.validate_protection(package)?;
            }
            return Ok(false);
        }
        self.validate_source(package)?;
        self.validate_protection(package)?;
        self.validate_destination_names(package)?;
        if package.is_signed() || package.requires_signature_edit_policy() {
            return Err(Error::Opc(OpcError::SignedSourceRequiresExplicitPolicy));
        }
        // The exact token APIs can still fail on a bounded reservation.
        // Publish through an Arc-sharing package clone so every failure,
        // including a late lexical replacement, is atomic.
        let mut candidate = package.clone();
        self.apply_to_candidate(&mut candidate)?;
        *package = candidate;
        Ok(true)
    }

    pub(in crate::web) fn has_changes(&self) -> bool {
        self.root
            .as_ref()
            .is_some_and(|root| root.before != root.after)
            || self.parts.iter().any(|part| part.before != part.after)
    }

    fn is_source_bound(&self) -> bool {
        self.root.is_some()
            || !self.parts.is_empty()
            || self.protection.is_some()
            || self.lexical.is_some()
    }

    pub(in crate::web) fn validate_source(&self, package: &OpcPackage) -> Result<()> {
        if let Some(root) = &self.root {
            let mut task_relationships = package
                .rels()
                .iter()
                .filter(|relationship| relationship.reltype() == TASK_PANES_RELATIONSHIP);
            let first = task_relationships.next();
            let unique = task_relationships.next().is_none();
            let root_matches = match &root.before {
                Some(before) => first.is_some_and(|actual| before.matches(actual)) && unique,
                None => first.is_none(),
            };
            if !root_matches {
                return Err(Error::Relationship(
                    "Web Extensions patch source relationship changed".into(),
                ));
            }
            if let Some(after) = &root.after {
                let reuses_source_id = root
                    .before
                    .as_ref()
                    .is_some_and(|before| before.id == after.id);
                if !reuses_source_id && package.rels().get(&after.id).is_some() {
                    return Err(Error::Relationship(
                        "Web Extensions patch destination relationship is occupied".into(),
                    ));
                }
            }
        }
        for change in &self.parts {
            let actual = match package.get_part(&change.name) {
                Ok(part) => Some(part),
                Err(OpcError::PartNotFound(_)) => None,
                Err(error) => return Err(error.into()),
            };
            let matches = match (&change.before, actual) {
                (Some(before), Some(actual)) => before.matches(actual),
                (None, None) => true,
                _ => false,
            };
            if !matches {
                return Err(Error::Relationship(
                    "Web Extensions patch source part changed".into(),
                ));
            }
        }
        if let Some(lexical) = &self.lexical {
            lexical.validate_before(package)?;
        }
        Ok(())
    }

    fn apply_to_candidate(&self, package: &mut OpcPackage) -> Result<()> {
        let staged = self
            .parts
            .iter()
            .filter(|part| part.before != part.after)
            .filter_map(|part| {
                part.after
                    .as_ref()
                    .map(|after| after.blob_part(part.name.clone()))
            })
            .collect::<Result<Vec<_>>>()?;

        for part in staged {
            package.add_part(Box::new(part));
        }
        if let Some(root) = self.root.as_ref().filter(|root| root.before != root.after) {
            if let Some(before) = &root.before {
                package.rels_mut().remove(&before.id);
            }
            if let Some(after) = &root.after {
                package.rels_mut().add_relationship(
                    after.relationship_type.clone(),
                    after.target.clone(),
                    after.id.clone(),
                    after.external,
                );
            }
        }
        for part in self
            .parts
            .iter()
            .filter(|part| part.before != part.after && part.after.is_none())
        {
            package.remove_part(&part.name);
        }

        if let Some(lexical) = &self.lexical {
            for xml in &lexical.parts {
                let Some(after) = &xml.after else {
                    continue;
                };
                let part = package.get_part(&xml.name)?;
                if part.blob() != after.bytes() {
                    return Err(Error::Relationship(
                        "Web Extensions patch XML destination changed while staging".into(),
                    ));
                }
                package.try_replace_owned_xml_part(after.bytes(), after.clone())?;
            }
            for relationships in &lexical.relationships {
                let Some(after) = &relationships.after else {
                    continue;
                };
                let current = package.source_relationships(after.owner())?;
                package.try_replace_relationships(&current, after)?;
            }
            let current = package.source_content_types()?;
            package.try_replace_content_types(current.bytes(), &lexical.content_types.after)?;

            lexical.validate_after(package)?;
        }
        Ok(())
    }

    pub(in crate::web) fn validate_protection(&self, package: &OpcPackage) -> Result<()> {
        let Some(guard) = &self.protection else {
            return Ok(());
        };
        let changed_existing: HashSet<_> = self
            .parts
            .iter()
            .filter(|part| part.before.is_some() && part.before != part.after)
            .map(|part| fold_part_name(&part.name))
            .collect();
        if !changed_existing.is_empty() {
            let protected = protected_parts(package, &guard.source, &guard.limits)?;
            if changed_existing.iter().any(|name| protected.contains(name)) {
                return Err(Error::Relationship(
                    "Web Extensions patch would change a newly shared source part".into(),
                ));
            }
        }

        let changed_destinations: HashSet<_> = self
            .parts
            .iter()
            .filter(|part| part.after.is_some() && part.before != part.after)
            .map(|part| fold_part_name(&part.name))
            .collect();
        if !changed_destinations.is_empty() {
            let protected = protected_parts(package, &guard.destination, &guard.limits)?;
            if changed_destinations
                .iter()
                .any(|name| protected.contains(name))
            {
                return Err(Error::Relationship(
                    "Web Extensions patch would create a newly shared destination part".into(),
                ));
            }
        }
        Ok(())
    }

    pub(in crate::web) fn validate_destination_names(&self, package: &OpcPackage) -> Result<()> {
        let changed = self
            .parts
            .iter()
            .filter(|change| change.before != change.after && change.after.is_some());
        let mut destinations = BTreeSet::new();
        let mut replacements = HashMap::new();
        for change in changed {
            let folded = fold_part_name(&change.name);
            if folded_name_conflicts(&destinations, &folded) {
                return invalid("Web Extensions patch destination parts conflict".into());
            }
            destinations.insert(folded.clone());
            if change.before.is_some() {
                replacements.insert(folded, &change.name);
            }
        }
        if destinations.is_empty() {
            return Ok(());
        }
        for existing in package.iter_parts() {
            let folded = fold_part_name(existing.partname());
            if !folded_name_conflicts(&destinations, &folded) {
                continue;
            }
            let is_replaced_source = replacements
                .get(&folded)
                .is_some_and(|name| *name == existing.partname());
            if !is_replaced_source {
                return invalid("Web Extensions patch destination part is occupied".into());
            }
        }
        Ok(())
    }
}

#[derive(Clone)]
pub(in crate::web::package) struct LexicalChange {
    pub(in crate::web::package) content_types: ContentTypesChange,
    pub(in crate::web::package) relationships: Box<[RelationshipTokenChange]>,
    pub(in crate::web::package) parts: Box<[XmlTokenChange]>,
}

#[derive(Clone)]
pub(in crate::web::package) struct ContentTypesChange {
    pub(in crate::web::package) before: OwnedContentTypes,
    pub(in crate::web::package) after: OwnedContentTypes,
}

#[derive(Clone)]
pub(in crate::web::package) struct RelationshipTokenChange {
    pub(in crate::web::package) owner: PackURI,
    pub(in crate::web::package) before: Option<OwnedRelationships>,
    pub(in crate::web::package) after: Option<OwnedRelationships>,
}

#[derive(Clone)]
pub(in crate::web::package) struct XmlTokenChange {
    pub(in crate::web::package) name: PackURI,
    pub(in crate::web::package) before: Option<OwnedXmlPart>,
    pub(in crate::web::package) after: Option<OwnedXmlPart>,
    pub(in crate::web::package) before_exists: bool,
    pub(in crate::web::package) after_exists: bool,
}

#[derive(Clone, Copy, PartialEq, Eq)]
enum ProvenSourceElement {
    Contains,
    Background,
}

#[derive(Clone, Copy, PartialEq, Eq)]
enum ProvenSourceEvent {
    Eof,
    Start {
        empty: bool,
        element: Option<ProvenSourceElement>,
    },
    Other,
}

struct ProvenSourceAttribute<'a> {
    name: &'a [u8],
    value: &'a [u8],
    value_range: Range<usize>,
    remove_range: Range<usize>,
}

struct ProvenSourceTag<'a> {
    name: &'a [u8],
    attributes: Vec<ProvenSourceAttribute<'a>>,
}

struct ProvenSourceTagEdit {
    tag_range: Range<usize>,
    insertion: Option<(String, Vec<u8>)>,
    replacements: Vec<(Range<usize>, Vec<u8>)>,
}

/// Reconstruct a source token only for the narrow custom-function rewrite
/// emitted by the semantic writer.  The writer changes values or inserts and
/// removes the unqualified `val`, `state`, and `runtimeId` attributes on the
/// two recognized web-extension elements; all other source events must remain
/// byte-identical.  This deliberately refuses text, child, namespace, and
/// arbitrary authored XML changes.
fn reconstruct_proven_attribute_source(
    before: &OwnedXmlPart,
    after: &[u8],
    limits: Limits,
) -> Result<Option<OwnedXmlPart>> {
    if before.bytes() == after {
        return Ok(Some(before.clone()));
    }
    if after.len() > limits.xml_bytes {
        return limit("Web Extensions XML bytes", limits.xml_bytes, after.len());
    }

    let before_bytes = before.bytes();
    let mut before_reader = NsReader::from_reader(before_bytes);
    let mut after_reader = NsReader::from_reader(after);
    before_reader.config_mut().trim_text(false);
    after_reader.config_mut().trim_text(false);
    before_reader.config_mut().check_end_names = true;
    after_reader.config_mut().check_end_names = true;

    let mut edits = Vec::new();
    loop {
        let (before_start, before_end, before_event) =
            read_proven_source_event(&mut before_reader, before_bytes)?;
        let (after_start, after_end, after_event) =
            read_proven_source_event(&mut after_reader, after)?;
        if before_event == ProvenSourceEvent::Eof && after_event == ProvenSourceEvent::Eof {
            break;
        }
        if before_event == ProvenSourceEvent::Eof || after_event == ProvenSourceEvent::Eof {
            return Ok(None);
        }
        let before_raw = &before_bytes[before_start..before_end];
        let after_raw = &after[after_start..after_end];
        match (before_event, after_event) {
            (
                ProvenSourceEvent::Start {
                    empty: before_empty,
                    element: Some(before_element),
                },
                ProvenSourceEvent::Start {
                    empty: after_empty,
                    element: Some(after_element),
                },
            ) if before_empty == after_empty && before_element == after_element => {
                let Some(edit) = source_tag_edit(
                    before_element,
                    before_raw,
                    after_raw,
                    before_start..before_end,
                    limits,
                )?
                else {
                    return Ok(None);
                };
                if edits.len() >= limits.nodes {
                    return limit(
                        "Web Extensions source attribute edits",
                        limits.nodes,
                        edits.len().saturating_add(1),
                    );
                }
                edits.try_reserve(1).map_err(|source| Error::Allocation {
                    resource: "Web Extensions source attribute edits",
                    source,
                })?;
                edits.push(edit);
            },
            _ if before_raw == after_raw => {},
            _ => return Ok(None),
        }
    }
    if edits.is_empty() {
        return Ok(None);
    }

    // Apply later tags first so ranges in earlier tags remain source-relative.
    edits.sort_unstable_by_key(|edit| std::cmp::Reverse(edit.tag_range.start));
    let mut result = before.clone();
    for edit in edits {
        if let Some((name, value)) = edit.insertion {
            result = result.insert_unqualified_attribute(edit.tag_range.clone(), &name, &value)?;
        }
        if !edit.replacements.is_empty() {
            result = result.replace_attributes(&edit.replacements)?;
        }
        if result.bytes().len() > limits.xml_bytes {
            return limit(
                "Web Extensions XML bytes",
                limits.xml_bytes,
                result.bytes().len(),
            );
        }
    }
    if result.bytes() != after {
        return Ok(None);
    }
    Ok(Some(result))
}

fn read_proven_source_event(
    reader: &mut NsReader<&[u8]>,
    bytes: &[u8],
) -> Result<(usize, usize, ProvenSourceEvent)> {
    let start =
        usize::try_from(reader.buffer_position()).map_err(|error| Error::Xml(error.to_string()))?;
    let event = reader
        .read_event()
        .map_err(|error| Error::Xml(error.to_string()))?;
    let resolver = reader.resolver().clone();
    let (namespace, event) = resolver.resolve_event(event);
    let end =
        usize::try_from(reader.buffer_position()).map_err(|error| Error::Xml(error.to_string()))?;
    if end < start || end > bytes.len() {
        return Err(Error::Xml("source XML event range is invalid".into()));
    }
    let event = match event {
        Event::Start(element) => ProvenSourceEvent::Start {
            empty: false,
            element: proven_source_element(&namespace, element.local_name().as_ref()),
        },
        Event::Empty(element) => ProvenSourceEvent::Start {
            empty: true,
            element: proven_source_element(&namespace, element.local_name().as_ref()),
        },
        Event::Eof => ProvenSourceEvent::Eof,
        _ => ProvenSourceEvent::Other,
    };
    Ok((start, end, event))
}

fn proven_source_element(
    namespace: &ResolveResult<'_>,
    local_name: &[u8],
) -> Option<ProvenSourceElement> {
    if !matches!(
        namespace,
        ResolveResult::Bound(Namespace(value)) if *value == WEB_EXTENSION_NAMESPACE.as_bytes()
    ) {
        return None;
    }
    match local_name {
        b"containsCustomFunctions" => Some(ProvenSourceElement::Contains),
        b"backgroundAppData" => Some(ProvenSourceElement::Background),
        _ => None,
    }
}

fn source_tag_edit(
    element: ProvenSourceElement,
    before_raw: &[u8],
    after_raw: &[u8],
    tag_range: Range<usize>,
    limits: Limits,
) -> Result<Option<ProvenSourceTagEdit>> {
    let before = parse_proven_source_tag(before_raw, limits)?;
    let after = parse_proven_source_tag(after_raw, limits)?;
    if before.name != after.name {
        return Ok(None);
    }
    let mut replacements = Vec::new();
    let mut insertion = None;
    for attribute in &before.attributes {
        let matching = after
            .attributes
            .iter()
            .find(|candidate| candidate.name == attribute.name);
        match matching {
            Some(matching) if attribute.value == matching.value => {},
            Some(matching) => {
                if !proven_attribute(element, attribute.name) {
                    return Ok(None);
                }
                replacements
                    .try_reserve(1)
                    .map_err(|source| Error::Allocation {
                        resource: "Web Extensions source attribute replacements",
                        source,
                    })?;
                replacements.push((
                    tag_range.start + attribute.value_range.start
                        ..tag_range.start + attribute.value_range.end,
                    copy_proven_bytes(
                        matching.value,
                        "Web Extensions source attribute replacement",
                    )?,
                ));
            },
            None => {
                if !proven_attribute(element, attribute.name) {
                    return Ok(None);
                }
                replacements
                    .try_reserve(1)
                    .map_err(|source| Error::Allocation {
                        resource: "Web Extensions source attribute removals",
                        source,
                    })?;
                replacements.push((
                    tag_range.start + attribute.remove_range.start
                        ..tag_range.start + attribute.remove_range.end,
                    Vec::new(),
                ));
            },
        }
    }
    for attribute in &after.attributes {
        if before
            .attributes
            .iter()
            .any(|candidate| candidate.name == attribute.name)
        {
            continue;
        }
        if !proven_attribute(element, attribute.name) || insertion.is_some() {
            return Ok(None);
        }
        let name_bytes =
            std::str::from_utf8(attribute.name).map_err(|error| Error::Xml(error.to_string()))?;
        let mut name = String::new();
        name.try_reserve(name_bytes.len())
            .map_err(|source| Error::Allocation {
                resource: "Web Extensions source attribute name",
                source,
            })?;
        name.push_str(name_bytes);
        insertion = Some((
            name,
            copy_proven_bytes(attribute.value, "Web Extensions source attribute insertion")?,
        ));
    }
    if replacements.len() > limits.items {
        return limit(
            "Web Extensions source attribute edits",
            limits.items,
            replacements.len(),
        );
    }
    if replacements.is_empty() && insertion.is_none() && before_raw != after_raw {
        return Ok(None);
    }
    Ok(Some(ProvenSourceTagEdit {
        tag_range,
        insertion,
        replacements,
    }))
}

fn proven_attribute(element: ProvenSourceElement, name: &[u8]) -> bool {
    match element {
        ProvenSourceElement::Contains => name == b"val",
        ProvenSourceElement::Background => name == b"state" || name == b"runtimeId",
    }
}

fn parse_proven_source_tag(raw: &[u8], limits: Limits) -> Result<ProvenSourceTag<'_>> {
    if raw.len() < 3 || raw[0] != b'<' {
        return Err(Error::Xml("source XML start tag is invalid".into()));
    }
    let mut cursor = 1usize;
    let name_start = cursor;
    while cursor < raw.len()
        && !is_xml_space(raw[cursor])
        && raw[cursor] != b'>'
        && raw[cursor] != b'/'
    {
        cursor += 1;
    }
    if cursor == name_start {
        return Err(Error::Xml("source XML start tag has no name".into()));
    }
    let name = &raw[name_start..cursor];
    let mut attributes = Vec::new();
    loop {
        let before_space = cursor;
        while cursor < raw.len() && is_xml_space(raw[cursor]) {
            cursor += 1;
        }
        if cursor >= raw.len() {
            return Err(Error::Xml("source XML start tag is unterminated".into()));
        }
        if raw[cursor] == b'>' || (raw[cursor] == b'/' && raw.get(cursor + 1) == Some(&b'>')) {
            break;
        }
        if cursor == before_space {
            return Err(Error::Xml("source XML attributes are not separated".into()));
        }
        let attribute_start = cursor;
        while cursor < raw.len()
            && !is_xml_space(raw[cursor])
            && raw[cursor] != b'='
            && raw[cursor] != b'>'
            && raw[cursor] != b'/'
        {
            cursor += 1;
        }
        if cursor == attribute_start {
            return Err(Error::Xml("source XML attribute has no name".into()));
        }
        if attributes.len() >= limits.items {
            return limit(
                "Web Extensions source tag attributes",
                limits.items,
                attributes.len().saturating_add(1),
            );
        }
        let name = &raw[attribute_start..cursor];
        while cursor < raw.len() && is_xml_space(raw[cursor]) {
            cursor += 1;
        }
        if raw.get(cursor) != Some(&b'=') {
            return Err(Error::Xml("source XML attribute has no equals sign".into()));
        }
        cursor += 1;
        while cursor < raw.len() && is_xml_space(raw[cursor]) {
            cursor += 1;
        }
        let quote = *raw
            .get(cursor)
            .ok_or_else(|| Error::Xml("source XML attribute has no quote".into()))?;
        if quote != b'\'' && quote != b'"' {
            return Err(Error::Xml("source XML attribute quote is invalid".into()));
        }
        cursor += 1;
        let value_start = cursor;
        while cursor < raw.len() && raw[cursor] != quote {
            cursor += 1;
        }
        let value_end = cursor;
        if cursor >= raw.len() {
            return Err(Error::Xml(
                "source XML attribute has no closing quote".into(),
            ));
        }
        cursor += 1;
        if attributes
            .iter()
            .any(|attribute: &ProvenSourceAttribute<'_>| attribute.name == name)
        {
            return Err(Error::Xml("source XML attribute is duplicated".into()));
        }
        attributes
            .try_reserve(1)
            .map_err(|source| Error::Allocation {
                resource: "Web Extensions source tag attributes",
                source,
            })?;
        attributes.push(ProvenSourceAttribute {
            name,
            value: &raw[value_start..value_end],
            value_range: value_start..value_end,
            remove_range: attribute_start.saturating_sub(1)..cursor,
        });
    }
    Ok(ProvenSourceTag { name, attributes })
}

fn copy_proven_bytes(bytes: &[u8], resource: &'static str) -> Result<Vec<u8>> {
    let mut copy = Vec::new();
    copy.try_reserve_exact(bytes.len())
        .map_err(|source| Error::Allocation { resource, source })?;
    copy.extend_from_slice(bytes);
    Ok(copy)
}

fn is_xml_space(byte: u8) -> bool {
    matches!(byte, b' ' | b'\t' | b'\n' | b'\r')
}

impl LexicalChange {
    fn capture(
        package: &OpcPackage,
        parts: &[PartChange],
        root: Option<&RootChange>,
        source_parts: &[PackURI],
        destination_parts: &[PackURI],
        limits: Limits,
    ) -> Result<Self> {
        let before_content_types = package.source_content_types()?;
        check_lexical_bytes(
            before_content_types.bytes().len(),
            &limits,
            "content-types XML",
        )?;

        let mut candidate = package.clone();
        for change in parts.iter().filter(|change| change.before != change.after) {
            if let Some(after) = &change.after {
                candidate.add_part(Box::new(after.blob_part(change.name.clone())?));
            } else {
                candidate.remove_part(&change.name);
            }
        }
        if let Some(root) = root.filter(|root| root.before != root.after) {
            if let Some(before) = &root.before {
                candidate.rels_mut().remove(&before.id);
            }
            if let Some(after) = &root.after {
                candidate.rels_mut().add_relationship(
                    after.relationship_type.clone(),
                    after.target.clone(),
                    after.id.clone(),
                    after.external,
                );
            }
        }

        let content_types_after = transition_content_types(&before_content_types, parts, &limits)?;
        check_lexical_bytes(
            content_types_after.bytes().len(),
            &limits,
            "content-types XML",
        )?;

        let scope_capacity = source_parts
            .len()
            .checked_add(destination_parts.len())
            .and_then(|count| count.checked_add(parts.len()))
            .ok_or(Error::Limit {
                resource: "Web Extensions patch lexical scope",
                max: limits.package_parts,
                actual: usize::MAX,
            })?;
        let mut scope_names = Vec::new();
        scope_names
            .try_reserve_exact(scope_capacity)
            .map_err(|source| Error::Allocation {
                resource: "Web Extensions patch lexical scope",
                source,
            })?;
        let mut seen_scope_names = HashSet::new();
        seen_scope_names
            .try_reserve(scope_capacity)
            .map_err(|source| Error::Allocation {
                resource: "Web Extensions patch lexical scope",
                source,
            })?;
        for name in source_parts
            .iter()
            .chain(destination_parts)
            .chain(parts.iter().map(|change| &change.name))
        {
            if seen_scope_names.insert(fold_part_name(name)) {
                scope_names.push(name.clone());
            }
        }

        let mut relationship_changes = Vec::new();
        relationship_changes
            .try_reserve_exact(
                scope_names
                    .len()
                    .checked_add(if root.is_some() { 1 } else { 0 })
                    .ok_or(Error::Limit {
                        resource: "Web Extensions patch relationship tokens",
                        max: limits.package_relationships,
                        actual: usize::MAX,
                    })?,
            )
            .map_err(|source| Error::Allocation {
                resource: "Web Extensions patch relationship tokens",
                source,
            })?;
        let mut xml_changes = Vec::new();
        xml_changes
            .try_reserve_exact(scope_names.len())
            .map_err(|source| Error::Allocation {
                resource: "Web Extensions patch XML tokens",
                source,
            })?;
        let mut total_xml_bytes = before_content_types.bytes().len();
        total_xml_bytes = total_xml_bytes
            .checked_add(content_types_after.bytes().len())
            .ok_or(Error::Limit {
                resource: "Web Extensions patch lexical XML",
                max: limits.total_xml_bytes,
                actual: usize::MAX,
            })?;

        if root.is_some() {
            let owner = PackURI::new("/").map_err(|error| Error::Uri(error.to_string()))?;
            let before = package.source_relationships(&owner)?;
            let after = match root {
                Some(root) if root.before != root.after => transition_relationships(
                    &before,
                    root.before.as_ref().map(std::slice::from_ref),
                    root.after.as_ref().map(std::slice::from_ref),
                    &limits,
                )?,
                _ => candidate.source_relationships(&owner)?,
            };
            charge_lexical_bytes(&mut total_xml_bytes, before.bytes().len(), &limits)?;
            charge_lexical_bytes(&mut total_xml_bytes, after.bytes().len(), &limits)?;
            relationship_changes.push(RelationshipTokenChange {
                owner,
                before: Some(before),
                after: Some(after),
            });
        }

        for name in &scope_names {
            let change = parts
                .iter()
                .find(|change| fold_part_name(&change.name) == fold_part_name(name));
            let changed = change.is_some_and(|change| change.before != change.after);
            let before_relationships = optional_relationships(package, name)?;
            let candidate_relationships = optional_relationships(&candidate, name)?;
            let after_relationships =
                match (changed, &before_relationships, &candidate_relationships) {
                    (true, Some(before), Some(_after)) => {
                        let Some(change) = change else {
                            return Err(Error::Missing(
                                "Web Extensions patch relationship owner disappeared".into(),
                            ));
                        };
                        Some(transition_relationships(
                            before,
                            change
                                .before
                                .as_ref()
                                .map(|part| part.relationships.as_ref()),
                            change
                                .after
                                .as_ref()
                                .map(|part| part.relationships.as_ref()),
                            &limits,
                        )?)
                    },
                    (_, _, after) => after.clone(),
                };
            if let Some(before) = &before_relationships {
                charge_lexical_bytes(&mut total_xml_bytes, before.bytes().len(), &limits)?;
            }
            if let Some(after) = &after_relationships {
                charge_lexical_bytes(&mut total_xml_bytes, after.bytes().len(), &limits)?;
            }
            if before_relationships.is_some() || after_relationships.is_some() {
                let owner = before_relationships
                    .as_ref()
                    .map(|token| token.owner().clone())
                    .or_else(|| {
                        after_relationships
                            .as_ref()
                            .map(|token| token.owner().clone())
                    })
                    .ok_or_else(|| {
                        Error::Missing("Web Extensions patch relationship owner disappeared".into())
                    })?;
                relationship_changes.push(RelationshipTokenChange {
                    owner,
                    before: before_relationships,
                    after: after_relationships,
                });
            }

            let before_exists = optional_part_exists(package, name)?;
            let after_exists = optional_part_exists(&candidate, name)?;
            let before_xml = optional_xml_part(package, name)?;
            let after_xml = match optional_xml_part(&candidate, name) {
                Ok(xml) => xml,
                Err(Error::Opc(OpcError::XmlError(_))) => {
                    if let Some(before) = before_xml.as_ref() {
                        let after_bytes = candidate.get_part(name)?.blob();
                        let Some(after) =
                            reconstruct_proven_attribute_source(before, after_bytes, limits)?
                        else {
                            return Err(Error::Opc(OpcError::XmlError(
                                "Web Extensions noncompact XML edit has no authorized source proof"
                                    .into(),
                            )));
                        };
                        Some(after)
                    } else {
                        // A newly authored XML part remains subject to the
                        // ordinary OPC authored-XML publication audit.
                        None
                    }
                },
                Err(error) => return Err(error),
            };
            if let Some(xml) = &before_xml {
                charge_lexical_bytes(&mut total_xml_bytes, xml.bytes().len(), &limits)?;
            }
            if let Some(xml) = &after_xml {
                charge_lexical_bytes(&mut total_xml_bytes, xml.bytes().len(), &limits)?;
            }
            if before_xml.is_some() || after_xml.is_some() {
                xml_changes.push(XmlTokenChange {
                    name: name.clone(),
                    before: before_xml,
                    after: after_xml,
                    before_exists,
                    after_exists,
                });
            }
        }

        Ok(Self {
            content_types: ContentTypesChange {
                before: before_content_types,
                after: content_types_after,
            },
            relationships: relationship_changes.into_boxed_slice(),
            parts: xml_changes.into_boxed_slice(),
        })
    }

    fn inverse(&self) -> Self {
        Self {
            content_types: ContentTypesChange {
                before: self.content_types.after.clone(),
                after: self.content_types.before.clone(),
            },
            relationships: self
                .relationships
                .iter()
                .map(|change| RelationshipTokenChange {
                    owner: change.owner.clone(),
                    before: change.after.clone(),
                    after: change.before.clone(),
                })
                .collect::<Vec<_>>()
                .into_boxed_slice(),
            parts: self
                .parts
                .iter()
                .map(|change| XmlTokenChange {
                    name: change.name.clone(),
                    before: change.after.clone(),
                    after: change.before.clone(),
                    before_exists: change.after_exists,
                    after_exists: change.before_exists,
                })
                .collect::<Vec<_>>()
                .into_boxed_slice(),
        }
    }

    fn validate_before(&self, package: &OpcPackage) -> Result<()> {
        self.validate_tokens(package, &self.content_types.before, false)
    }

    fn validate_after(&self, package: &OpcPackage) -> Result<()> {
        self.validate_tokens(package, &self.content_types.after, true)
    }

    fn validate_tokens(
        &self,
        package: &OpcPackage,
        content_types: &OwnedContentTypes,
        after: bool,
    ) -> Result<()> {
        if package.source_content_types()?.bytes() != content_types.bytes() {
            return Err(Error::Relationship(
                if after {
                    "Web Extensions patch destination content-types source changed"
                } else {
                    "Web Extensions patch source content-types changed"
                }
                .into(),
            ));
        }
        for change in &self.relationships {
            let expected = if after { &change.after } else { &change.before };
            let actual = optional_relationships(package, &change.owner)?;
            if actual.as_ref() != expected.as_ref() {
                return Err(Error::Relationship(
                    if after {
                        "Web Extensions patch destination relationship source changed"
                    } else {
                        "Web Extensions patch source relationship metadata changed"
                    }
                    .into(),
                ));
            }
        }
        for change in &self.parts {
            let expected = if after { &change.after } else { &change.before };
            let expected_exists = if after {
                change.after_exists
            } else {
                change.before_exists
            };
            let actual_part = package.get_part(&change.name);
            if !expected_exists {
                if actual_part.is_ok() {
                    return Err(Error::Relationship(
                        if after {
                            "Web Extensions patch destination XML part unexpectedly exists"
                        } else {
                            "Web Extensions patch source XML part unexpectedly exists"
                        }
                        .into(),
                    ));
                }
                continue;
            }
            let Some(expected) = expected else {
                // The part exists but its destination is authored rather than
                // source-provenanced. The semantic PartState check remains the
                // authoritative byte guard for this direction.
                continue;
            };
            let actual = optional_xml_part(package, &change.name)?;
            if actual.as_ref() != Some(expected) {
                return Err(Error::Relationship(
                    if after {
                        "Web Extensions patch destination XML source changed"
                    } else {
                        "Web Extensions patch source XML metadata changed"
                    }
                    .into(),
                ));
            }
        }
        Ok(())
    }
}

fn transition_content_types(
    before: &OwnedContentTypes,
    parts: &[PartChange],
    limits: &Limits,
) -> Result<OwnedContentTypes> {
    let mut removed = Vec::new();
    removed
        .try_reserve_exact(parts.len())
        .map_err(|source| Error::Allocation {
            resource: "Web Extensions patch content-type removals",
            source,
        })?;
    let mut additions = Vec::new();
    additions
        .try_reserve_exact(parts.len())
        .map_err(|source| Error::Allocation {
            resource: "Web Extensions patch content-type additions",
            source,
        })?;
    for change in parts.iter().filter(|change| change.before != change.after) {
        match (&change.before, &change.after) {
            (Some(before), Some(after)) if before.content_type == after.content_type => {},
            (Some(_), Some(after)) => {
                removed.push(change.name.clone());
                additions.push((&change.name, after.content_type.as_str()));
            },
            (Some(_), None) => removed.push(change.name.clone()),
            (None, Some(after)) => additions.push((&change.name, after.content_type.as_str())),
            (None, None) => {},
        }
    }
    let mut result = before.without_parts(&removed, limits.xml_bytes)?;
    if !additions.is_empty() {
        result = result.with_part_overrides(&additions, limits.xml_bytes)?;
    }
    Ok(result)
}

fn transition_relationships(
    before: &OwnedRelationships,
    before_state: Option<&[RelationshipState]>,
    after_state: Option<&[RelationshipState]>,
    limits: &Limits,
) -> Result<OwnedRelationships> {
    let before_state = before_state.unwrap_or(&[]);
    let after_state = after_state.unwrap_or(&[]);
    let mut result = before.clone();
    let mut additions = Vec::new();
    additions
        .try_reserve_exact(after_state.len())
        .map_err(|source| Error::Allocation {
            resource: "Web Extensions patch relationship additions",
            source,
        })?;
    for old in before_state {
        let unchanged = after_state
            .iter()
            .find(|new| new.id == old.id)
            .is_some_and(|new| new == old);
        if !unchanged {
            result = result.without_relationship(&old.id, limits.xml_bytes)?;
        }
    }
    for new in after_state {
        let unchanged = before_state
            .iter()
            .find(|old| old.id == new.id)
            .is_some_and(|old| old == new);
        if !unchanged {
            additions.push((
                new.relationship_type.as_str(),
                new.target.as_str(),
                new.id.as_str(),
                if new.external {
                    TargetMode::External
                } else {
                    TargetMode::Internal
                },
            ));
        }
    }
    if !additions.is_empty() {
        result = result.with_relationships(&additions, limits.xml_bytes)?;
    }
    Ok(result)
}

fn optional_relationships(
    package: &OpcPackage,
    owner: &PackURI,
) -> Result<Option<OwnedRelationships>> {
    if owner.as_str() == "/" {
        return package
            .source_relationships(owner)
            .map(Some)
            .map_err(Into::into);
    }
    match package.get_part(owner) {
        Ok(_) => package
            .source_relationships(owner)
            .map(Some)
            .map_err(Into::into),
        Err(OpcError::PartNotFound(_)) => Ok(None),
        Err(error) => Err(error.into()),
    }
}

fn optional_part_exists(package: &OpcPackage, name: &PackURI) -> Result<bool> {
    match package.get_part(name) {
        Ok(_) => Ok(true),
        Err(OpcError::PartNotFound(_)) => Ok(false),
        Err(error) => Err(error.into()),
    }
}

fn optional_xml_part(package: &OpcPackage, name: &PackURI) -> Result<Option<OwnedXmlPart>> {
    match package.get_part(name) {
        Ok(part) if is_xml_part(part.content_type(), part.partname()) => {
            package.source_xml_part(name).map(Some).map_err(Into::into)
        },
        Ok(_) => Ok(None),
        Err(OpcError::PartNotFound(_)) => Ok(None),
        Err(error) => Err(error.into()),
    }
}

fn is_xml_part(content_type: &str, name: &PackURI) -> bool {
    content_type.eq_ignore_ascii_case("application/xml")
        || content_type.ends_with("+xml")
        || name.as_str().rsplit('/').next().is_some_and(|part| {
            part.len() >= 4 && part[part.len() - 4..].eq_ignore_ascii_case(".xml")
        })
}

fn check_lexical_bytes(bytes: usize, limits: &Limits, resource: &'static str) -> Result<()> {
    if bytes > limits.xml_bytes {
        return limit(resource, limits.xml_bytes, bytes);
    }
    Ok(())
}

fn charge_lexical_bytes(total: &mut usize, bytes: usize, limits: &Limits) -> Result<()> {
    *total = total.checked_add(bytes).ok_or(Error::Limit {
        resource: "Web Extensions patch lexical XML",
        max: limits.total_xml_bytes,
        actual: usize::MAX,
    })?;
    if *total > limits.total_xml_bytes {
        return limit(
            "Web Extensions patch lexical XML",
            limits.total_xml_bytes,
            *total,
        );
    }
    Ok(())
}

#[cfg(test)]
#[path = "transaction_tests.rs"]
mod transaction_tests;

#[derive(Clone, PartialEq, Eq)]
pub(in crate::web) struct RootChange {
    pub(in crate::web) before: Option<RelationshipState>,
    pub(in crate::web) after: Option<RelationshipState>,
}

#[derive(Clone)]
pub(in crate::web) struct ProtectionGuard {
    pub(in crate::web) source: GraphScope,
    pub(in crate::web) destination: GraphScope,
    pub(in crate::web) limits: Limits,
}

impl ProtectionGuard {
    pub(in crate::web) fn inverse(&self) -> Self {
        Self {
            source: self.destination.clone(),
            destination: self.source.clone(),
            limits: self.limits,
        }
    }
}

#[derive(Clone)]
pub(in crate::web) struct GraphScope {
    pub(in crate::web) root_relationship_id: Option<String>,
    pub(in crate::web) owned_parts: Box<[PackURI]>,
}

pub(in crate::web) fn protected_parts(
    package: &OpcPackage,
    scope: &GraphScope,
    limits: &Limits,
) -> Result<HashSet<String>> {
    let part_count = package.part_count();
    if part_count > limits.package_parts {
        return limit("package parts", limits.package_parts, part_count);
    }
    if scope.owned_parts.is_empty() {
        return Ok(HashSet::new());
    }

    let mut names = Vec::with_capacity(scope.owned_parts.len());
    let mut by_name = HashMap::with_capacity(scope.owned_parts.len());
    for name in &scope.owned_parts {
        let folded = fold_part_name(name);
        if by_name.insert(folded.clone(), names.len()).is_some() {
            return invalid("duplicate part in Web Extensions patch protection scope".into());
        }
        names.push(folded);
    }

    let mut outbound = vec![Vec::new(); names.len()];
    let mut protected = vec![false; names.len()];
    let mut queue = VecDeque::new();
    let mut relationships = 0usize;

    for relationship in package.rels().iter() {
        charge_patch_relationship(&mut relationships, limits)?;
        if relationship.is_external() {
            continue;
        }
        let Ok(target) = relationship.target_partname() else {
            continue;
        };
        let Some(&target) = by_name.get(&fold_part_name(&target)) else {
            continue;
        };
        if scope.root_relationship_id.as_deref() != Some(relationship.r_id()) && !protected[target]
        {
            protected[target] = true;
            queue.push_back(target);
        }
    }

    for part in package.iter_parts() {
        let source = by_name.get(&fold_part_name(part.partname())).copied();
        for relationship in part.rels().iter() {
            charge_patch_relationship(&mut relationships, limits)?;
            if relationship.is_external() {
                continue;
            }
            let Ok(target_name) = relationship.target_partname() else {
                continue;
            };
            let Some(&target) = by_name.get(&fold_part_name(&target_name)) else {
                continue;
            };
            if let Some(source) = source {
                outbound[source].push(target);
            } else if !protected[target] {
                protected[target] = true;
                queue.push_back(target);
            }
        }
    }

    while let Some(source) = queue.pop_front() {
        for &target in &outbound[source] {
            if !protected[target] {
                protected[target] = true;
                queue.push_back(target);
            }
        }
    }

    Ok(names
        .into_iter()
        .enumerate()
        .filter_map(|(index, name)| protected[index].then_some(name))
        .collect())
}

pub(in crate::web) fn charge_patch_relationship(
    relationships: &mut usize,
    limits: &Limits,
) -> Result<()> {
    *relationships = relationships.checked_add(1).ok_or(Error::Limit {
        resource: "package relationships",
        max: limits.package_relationships,
        actual: usize::MAX,
    })?;
    if *relationships > limits.package_relationships {
        return limit(
            "package relationships",
            limits.package_relationships,
            *relationships,
        );
    }
    Ok(())
}

#[derive(Clone, PartialEq, Eq)]
pub(in crate::web) struct PartChange {
    pub(in crate::web) name: PackURI,
    pub(in crate::web) before: Option<PartState>,
    pub(in crate::web) after: Option<PartState>,
}

#[derive(Clone, PartialEq, Eq)]
pub(in crate::web) struct PartState {
    pub(in crate::web) content_type: String,
    pub(in crate::web) data: Arc<Vec<u8>>,
    pub(in crate::web) relationships: Box<[RelationshipState]>,
}

impl PartState {
    pub(in crate::web) fn capture(part: &dyn Part) -> Self {
        let mut relationships = part
            .rels()
            .iter()
            .map(RelationshipState::capture)
            .collect::<Vec<_>>();
        relationships.sort_by(|left, right| left.id.cmp(&right.id));
        Self {
            content_type: part.content_type().to_owned(),
            data: part.blob_arc(),
            relationships: relationships.into_boxed_slice(),
        }
    }

    pub(in crate::web) fn from_planned(part: PlannedPart) -> Self {
        let mut relationships = part
            .relationships
            .into_iter()
            .map(RelationshipState::from_planned)
            .collect::<Vec<_>>();
        relationships.sort_by(|left, right| left.id.cmp(&right.id));
        Self {
            content_type: part.content_type,
            data: part.data,
            relationships: relationships.into_boxed_slice(),
        }
    }

    pub(in crate::web) fn matches(&self, part: &dyn Part) -> bool {
        self.content_type == part.content_type()
            && self.data.as_slice() == part.blob()
            && self.relationships.len() == part.rels().len()
            && self.relationships.iter().all(|expected| {
                part.rels()
                    .get(&expected.id)
                    .is_some_and(|actual| expected.matches(actual))
            })
    }

    pub(in crate::web) fn blob_part(&self, name: PackURI) -> Result<BlobPart> {
        let mut part =
            BlobPart::new_shared(name, self.content_type.clone(), Arc::clone(&self.data));
        for relationship in &self.relationships {
            part.rels_mut().try_add_relationship(
                relationship.relationship_type.clone(),
                relationship.target.clone(),
                relationship.id.clone(),
                if relationship.external {
                    TargetMode::External
                } else {
                    TargetMode::Internal
                },
            )?;
        }
        Ok(part)
    }
}

#[derive(Clone, PartialEq, Eq)]
pub(in crate::web) struct RelationshipState {
    pub(in crate::web) id: String,
    pub(in crate::web) relationship_type: String,
    pub(in crate::web) target: String,
    pub(in crate::web) external: bool,
}

impl RelationshipState {
    pub(in crate::web) fn capture(relationship: &litchi_opc::Relationship) -> Self {
        Self {
            id: relationship.r_id().to_owned(),
            relationship_type: relationship.reltype().to_owned(),
            target: relationship.target_ref().to_owned(),
            external: relationship.is_external(),
        }
    }

    pub(in crate::web) fn from_planned(relationship: PlannedRelationship) -> Self {
        Self {
            id: relationship.id,
            relationship_type: relationship.relationship_type,
            target: relationship.target,
            external: relationship.external,
        }
    }

    pub(in crate::web) fn matches(&self, relationship: &litchi_opc::Relationship) -> bool {
        self.id == relationship.r_id()
            && self.relationship_type == relationship.reltype()
            && self.target == relationship.target_ref()
            && self.external == relationship.is_external()
    }
}
