//! Metadata images are provenance; entry keys are the merge conflict domain.
use super::{
    Input, Limits, add_len, invalid, parse_part_name, put_bytes32, put_u32, require_capture_bytes,
    stale,
};
use crate::{Error, Result};
use litchi_opc::{
    ContentTypeFragmentKey, ContentTypeFragments, OpcPackage, PackURI, RelationshipFragments,
    SourceLimits,
};
use std::{ops::Range, sync::Arc};

#[derive(Clone, PartialEq, Eq)]
pub(super) enum Owner {
    ContentTypes,
    Relationships(PackURI),
}
impl Ord for Owner {
    fn cmp(&self, other: &Self) -> std::cmp::Ordering {
        match (self, other) {
            (Self::ContentTypes, Self::ContentTypes) => std::cmp::Ordering::Equal,
            (Self::ContentTypes, _) => std::cmp::Ordering::Less,
            (_, Self::ContentTypes) => std::cmp::Ordering::Greater,
            (Self::Relationships(a), Self::Relationships(b)) => a.as_str().cmp(b.as_str()),
        }
    }
}
impl PartialOrd for Owner {
    fn partial_cmp(&self, other: &Self) -> Option<std::cmp::Ordering> {
        Some(self.cmp(other))
    }
}
#[derive(Clone, PartialEq, Eq)]
struct Image {
    present: bool,
    bytes: Arc<Vec<u8>>,
}
#[derive(Clone, PartialEq, Eq, PartialOrd, Ord, Hash)]
struct Key {
    kind: u8,
    text: String,
}
#[derive(Clone, PartialEq, Eq)]
struct Change {
    key: Key,
    before: Option<Range<usize>>,
    after: Option<Range<usize>>,
}
#[derive(Clone, PartialEq, Eq)]
struct Member {
    owner: Owner,
    before: Option<Image>,
    after: Option<Image>,
    changes: Vec<Change>,
    opaque: bool,
}
#[derive(Clone, Default, PartialEq, Eq)]
pub(super) struct Metadata {
    members: Vec<Member>,
}
impl std::fmt::Debug for Metadata {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        f.debug_struct("Metadata")
            .field("members", &self.members.len())
            .finish_non_exhaustive()
    }
}
struct Row {
    key: Key,
    span: Range<usize>,
}
enum Catalog {
    Types(ContentTypeFragments),
    Relationships(RelationshipFragments),
}
struct Table {
    catalog: Catalog,
    rows: Vec<Row>,
}

fn allocation(source: std::collections::TryReserveError) -> Error {
    Error::Allocation {
        resource: "opened-presentation metadata",
        source,
    }
}
fn text(value: &str) -> Result<String> {
    let mut output = String::new();
    output.try_reserve_exact(value.len()).map_err(allocation)?;
    output.push_str(value);
    Ok(output)
}
fn source_limits(limits: Limits, maximum: usize) -> Result<SourceLimits> {
    let maximum = maximum.min(u32::MAX as usize);
    if maximum == 0 {
        return Err(Error::InvalidLimit {
            kind: crate::opened::LimitKind::Codec(crate::opened::CodecLimitKind::XmlBytes),
            value: 0,
            maximum: 0,
        });
    }
    Ok(SourceLimits::builder()
        .max_part_bytes(maximum as u64)?
        .max_entry_bytes(maximum as u64)?
        .max_content_types_bytes(maximum)?
        .max_relationship_xml_bytes(maximum)?
        .max_total_relationship_xml_bytes(maximum)?
        .max_content_type_mappings(limits.max_parts().saturating_add(2))?
        .max_relationships_per_part(limits.max_parts())?
        .max_total_relationships(limits.max_parts())?
        .build()?)
}
impl Table {
    fn new(owner: &Owner, image: &Image, limits: SourceLimits) -> Result<Self> {
        let catalog = match owner {
            Owner::ContentTypes => {
                if !image.present {
                    return Err(invalid("content-types metadata cannot be absent"));
                }
                Catalog::Types(ContentTypeFragments::from_bytes(
                    Arc::clone(&image.bytes),
                    limits,
                )?)
            },
            Owner::Relationships(owner) => Catalog::Relationships(
                RelationshipFragments::from_bytes(owner.clone(), Arc::clone(&image.bytes), limits)?,
            ),
        };
        let count = match &catalog {
            Catalog::Types(value) => value.iter().len(),
            Catalog::Relationships(value) => value.iter().len(),
        };
        if !image.present && count != 0 {
            return Err(invalid("absent metadata member contains relationships"));
        }
        if !image.present {
            let Owner::Relationships(owner) = owner else {
                return Err(invalid("content-types metadata cannot be absent"));
            };
            if owner.as_str() == "/" {
                return Err(invalid("root relationship member cannot be absent"));
            }
            let graph = litchi_opc::pkgreader::PackageReader::parse_owned_relationships(
                &image.bytes,
                owner,
                limits,
            )?;
            if graph
                .plan_canonical(limits)?
                .materialize(limits)?
                .as_slice()
                != image.bytes.as_slice()
            {
                return Err(invalid(
                    "absent relationship member requires canonical empty XML",
                ));
            }
        }
        let mut rows = Vec::new();
        rows.try_reserve_exact(count).map_err(allocation)?;
        match &catalog {
            Catalog::Types(value) => {
                for entry in value.iter() {
                    let (kind, value) = match entry.key() {
                        ContentTypeFragmentKey::Default(value) => (0, value),
                        ContentTypeFragmentKey::Override(value) => (1, value),
                    };
                    rows.push(Row {
                        key: Key {
                            kind,
                            text: text(value)?,
                        },
                        span: entry.byte_range(),
                    });
                }
            },
            Catalog::Relationships(value) => {
                for entry in value.iter() {
                    rows.push(Row {
                        key: Key {
                            kind: 2,
                            text: text(entry.id())?,
                        },
                        span: entry.byte_range(),
                    });
                }
            },
        }
        rows.sort_unstable_by(|a, b| a.key.cmp(&b.key));
        Ok(Self { catalog, rows })
    }
    fn get(&self, key: &Key) -> Option<&Range<usize>> {
        self.rows
            .binary_search_by(|row| row.key.cmp(key))
            .ok()
            .map(|index| &self.rows[index].span)
    }
}
fn entry<'a>(image: &'a Option<Image>, span: &Option<Range<usize>>) -> Option<&'a [u8]> {
    image
        .as_ref()
        .zip(span.as_ref())
        .map(|(image, span)| &image.bytes[span.clone()])
}
fn residue<'a>(
    image: &'a Image,
    changes: &[Change],
    before: bool,
) -> Result<impl Iterator<Item = u8> + 'a> {
    let mut ranges = Vec::new();
    ranges
        .try_reserve_exact(changes.len().saturating_add(1))
        .map_err(allocation)?;
    ranges.extend(changes.iter().filter_map(|change| {
        if before {
            change.before.clone()
        } else {
            change.after.clone()
        }
    }));
    ranges.sort_unstable_by_key(|range| range.start);
    ranges.push(image.bytes.len()..image.bytes.len());
    Ok(ranges
        .into_iter()
        .scan(0, move |cursor, span| {
            let bytes = &image.bytes[*cursor..span.start];
            *cursor = span.end;
            Some(bytes.iter().copied())
        })
        .flatten())
}
impl Member {
    fn new(
        owner: Owner,
        before: Option<Image>,
        after: Option<Image>,
        limits: SourceLimits,
    ) -> Result<Self> {
        if (matches!(&owner, Owner::ContentTypes)
            || matches!(&owner, Owner::Relationships(name) if name.as_str() == "/"))
            && (before.is_none() || after.is_none())
        {
            return Err(invalid("package metadata owner cannot be absent"));
        }
        if before == after && owner != Owner::ContentTypes {
            return Err(invalid("metadata delta is a no-op"));
        }
        let left = before
            .as_ref()
            .map(|image| Table::new(&owner, image, limits))
            .transpose()?;
        let right = after
            .as_ref()
            .map(|image| Table::new(&owner, image, limits))
            .transpose()?;
        let mut changes = Vec::new();
        let maximum = left
            .as_ref()
            .map_or(0, |table| table.rows.len())
            .checked_add(right.as_ref().map_or(0, |table| table.rows.len()))
            .ok_or_else(|| invalid("metadata key count overflow"))?;
        changes.try_reserve_exact(maximum).map_err(allocation)?;
        if let Some(left) = &left {
            for row in &left.rows {
                let next = right
                    .as_ref()
                    .and_then(|right| right.get(&row.key))
                    .cloned();
                let previous = Some(row.span.clone());
                if entry(&before, &previous) != entry(&after, &next) {
                    changes.push(Change {
                        key: row.key.clone(),
                        before: previous,
                        after: next,
                    });
                }
            }
        }
        if let Some(right) = &right {
            for row in &right.rows {
                if left.as_ref().and_then(|left| left.get(&row.key)).is_none() {
                    changes.push(Change {
                        key: row.key.clone(),
                        before: None,
                        after: Some(row.span.clone()),
                    });
                }
            }
        }
        changes.sort_unstable_by(|a, b| a.key.cmp(&b.key));
        let opaque = before == after
            || match (&before, &after) {
                (Some(before), Some(after)) => {
                    !residue(before, &changes, true)?.eq(residue(after, &changes, false)?)
                },
                _ => true,
            };
        Ok(Self {
            owner,
            before,
            after,
            changes,
            opaque,
        })
    }
    fn inverse(&self) -> Self {
        Self {
            owner: self.owner.clone(),
            before: self.after.clone(),
            after: self.before.clone(),
            changes: self
                .changes
                .iter()
                .map(|change| Change {
                    key: change.key.clone(),
                    before: change.after.clone(),
                    after: change.before.clone(),
                })
                .collect(),
            opaque: self.opaque,
        }
    }
    fn conflicts(&self, other: &Self) -> bool {
        if self.owner != other.owner || self == other {
            return false;
        }
        self.before != other.before || self.entry_conflicts(other)
    }
    // A decoded inverse keeps each original group's after image as provenance.
    // Such groups share an exact after image instead of an exact before image.
    fn entry_conflicts(&self, other: &Self) -> bool {
        if self.owner != other.owner || self == other {
            return false;
        }
        if (self.before != other.before && self.after != other.after) || self.opaque || other.opaque
        {
            return true;
        }
        // Removing the physical member cannot be combined with another edit
        // that needs that member, even if the relationship IDs are disjoint.
        if [self, other].iter().any(|member| {
            member.before.as_ref().is_some_and(|image| image.present)
                && member.after.as_ref().is_some_and(|image| !image.present)
        }) {
            return true;
        }
        self.changes.iter().any(|left| {
            other
                .changes
                .binary_search_by(|right| right.key.cmp(&left.key))
                .ok()
                .is_some_and(|index| {
                    let right = &other.changes[index];
                    entry(&self.before, &left.before) != entry(&other.before, &right.before)
                        || entry(&self.after, &left.after) != entry(&other.after, &right.after)
                })
        })
    }
}
fn is_root(owner: &Owner) -> bool {
    matches!(owner, Owner::Relationships(name) if name.as_str() == "/")
}
fn image(package: &OpcPackage, owner: &Owner, limits: SourceLimits) -> Result<Option<Image>> {
    Ok(match owner {
        Owner::ContentTypes => Some(Image {
            present: true,
            bytes: package.source_content_types(limits)?.bytes_arc(),
        }),
        Owner::Relationships(owner) => {
            if owner.as_str() != "/" && package.part_metadata(owner).is_none() {
                None
            } else {
                let token = package.source_relationships(owner, limits)?;
                Some(Image {
                    present: token.member_present(),
                    bytes: token.bytes_arc(),
                })
            }
        },
    })
}
impl Metadata {
    pub(super) fn group_count(&self) -> usize {
        self.members.len()
    }
    pub(super) fn joined_count(&self, other: &Self) -> usize {
        self.members.len()
            + other
                .members
                .iter()
                .filter(|member| !self.members.contains(member))
                .count()
    }
    pub(super) fn conflicting_part_owners(&self, other: &Self) -> Result<Vec<PackURI>> {
        let mut owners = Vec::new();
        for member in &self.members {
            let Owner::Relationships(owner) = &member.owner else {
                continue;
            };
            if owner.as_str() != "/"
                && !owners.contains(owner)
                && other.members.iter().any(|other| member.conflicts(other))
            {
                owners.try_reserve(1).map_err(allocation)?;
                owners.push(owner.clone());
            }
        }
        Ok(owners)
    }
    pub(super) fn root_conflicts(&self, other: &Self) -> bool {
        self.members
            .iter()
            .filter(|member| is_root(&member.owner))
            .any(|left| other.members.iter().any(|right| left.conflicts(right)))
    }
    pub(super) fn partition_root(&self, limits: Limits) -> Result<(Self, Self)> {
        self.partition_relationships(&PackURI::new("/").map_err(Error::Invalid)?, limits)
    }
    pub(super) fn partition_relationships(
        &self,
        name: &PackURI,
        limits: Limits,
    ) -> Result<(Self, Self)> {
        self.check(limits)?;
        let selected =
            |member: &Member| matches!(&member.owner, Owner::Relationships(owner) if owner == name);
        let count = self
            .members
            .iter()
            .filter(|member| selected(member))
            .count();
        let mut root = Vec::new();
        let mut other = Vec::new();
        root.try_reserve_exact(count).map_err(allocation)?;
        other
            .try_reserve_exact(self.members.len() - count)
            .map_err(allocation)?;
        for member in &self.members {
            if selected(member) {
                root.push(member.clone());
            } else {
                other.push(member.clone());
            }
        }
        Ok((Self { members: other }, Self { members: root }))
    }
    pub(super) fn is_empty(&self) -> bool {
        self.members.is_empty()
    }
    pub(super) fn require_coverage(
        &self,
        deltas: &[super::Delta],
        root: Option<&(Vec<super::RelationshipState>, Vec<super::RelationshipState>)>,
        limits: Limits,
    ) -> Result<()> {
        if root.is_some() && !self.has_relationships(&PackURI::new("/").map_err(Error::Invalid)?) {
            return Err(invalid("root relationship delta lacks metadata provenance"));
        }
        for delta in deltas {
            if delta.before.is_none() != delta.after.is_none()
                && !self
                    .members
                    .iter()
                    .any(|member| member.owner == Owner::ContentTypes)
            {
                return Err(invalid(
                    "resource topology delta lacks content-type provenance",
                ));
            }
            match (&delta.before, &delta.after) {
                (Some(before), Some(after)) => {
                    if before.relationships != after.relationships
                        && !self.has_relationships(&delta.name)
                    {
                        return Err(invalid("relationship delta lacks metadata provenance"));
                    }
                    if before.content_type != after.content_type
                        && !self
                            .members
                            .iter()
                            .any(|member| member.owner == Owner::ContentTypes)
                    {
                        return Err(invalid("content-type change lacks metadata provenance"));
                    }
                },
                _ if !self.has_relationships(&delta.name) => {
                    return Err(invalid("resource topology delta lacks metadata provenance"));
                },
                _ => {},
            }
        }
        let policy = source_limits(limits, limits.max_patch_bytes())?;
        for member in &self.members {
            if member.before == member.after
                && !deltas.iter().any(|delta| {
                    delta.before.as_ref().map(|s| &s.content_type)
                        != delta.after.as_ref().map(|s| &s.content_type)
                })
            {
                return Err(invalid(
                    "unchanged metadata proof lacks a content-type dependency",
                ));
            }
            let Owner::Relationships(owner) = &member.owner else {
                continue;
            };
            let decode = |image: &Option<Image>| {
                image
                    .as_ref()
                    .map(|image| {
                        litchi_opc::pkgreader::PackageReader::parse_owned_relationships(
                            &image.bytes,
                            owner,
                            policy,
                        )
                        .map_err(Error::from)
                    })
                    .transpose()
            };
            let before = decode(&member.before)?;
            let after = decode(&member.after)?;
            let declared = if owner.as_str() == "/" {
                root.map(|(a, b)| (Some(a.as_slice()), Some(b.as_slice())))
            } else {
                deltas
                    .iter()
                    .find(|delta| &delta.name == owner)
                    .map(|delta| {
                        (
                            delta
                                .before
                                .as_ref()
                                .map(|state| state.relationships.as_slice()),
                            delta
                                .after
                                .as_ref()
                                .map(|state| state.relationships.as_slice()),
                        )
                    })
            };
            if let Some((expected_before, expected_after)) = declared {
                for (actual, expected) in [(&before, expected_before), (&after, expected_after)] {
                    match (actual, expected) {
                        (None, None) => {},
                        (Some(actual), Some(expected)) => {
                            // Joined groups share a base but retain separate
                            // after images. Bind each group's selected IDs to
                            // the merged graph, leaving other groups independent.
                            let matches = if member.opaque {
                                super::relationships_match(actual, expected)
                            } else {
                                member.changes.iter().all(|change| {
                                    let wanted =
                                        expected.iter().find(|value| value.id == change.key.text);
                                    match (actual.get(&change.key.text), wanted) {
                                        (None, None) => true,
                                        (Some(a), Some(b)) => {
                                            a.reltype() == b.kind
                                                && a.target_ref() == b.target
                                                && a.is_external() == b.external
                                        },
                                        _ => false,
                                    }
                                })
                            };
                            if !matches {
                                return Err(invalid(
                                    "metadata disagrees with declared relationships",
                                ));
                            }
                        },
                        _ => {
                            return Err(invalid("metadata disagrees with declared owner topology"));
                        },
                    }
                }
            } else if !before
                .as_ref()
                .zip(after.as_ref())
                .is_some_and(|(a, b)| super::relationships_equal(a, b))
            {
                return Err(invalid(
                    "semantic metadata change lacks declared relationships",
                ));
            }
        }
        Ok(())
    }
    pub(super) fn additional_resources(&self, deltas: &[super::Delta], root: bool) -> usize {
        self.members
            .iter()
            .enumerate()
            .filter(|(index, member)| {
                if self.members[..*index]
                    .iter()
                    .any(|previous| previous.owner == member.owner)
                {
                    return false;
                }
                match &member.owner {
                    Owner::ContentTypes => !deltas.iter().any(|delta| {
                        delta.before.as_ref().map(|state| &state.content_type)
                            != delta.after.as_ref().map(|state| &state.content_type)
                    }),
                    Owner::Relationships(owner) if owner.as_str() == "/" => !root,
                    Owner::Relationships(owner) => !deltas.iter().any(|delta| &delta.name == owner),
                }
            })
            .count()
    }
    pub(super) fn has_relationships(&self, name: &PackURI) -> bool {
        self.members
            .iter()
            .any(|member| matches!(&member.owner, Owner::Relationships(owner) if owner == name))
    }
    pub(super) fn inverse(&self) -> Self {
        Self {
            members: self.members.iter().map(Member::inverse).collect(),
        }
    }
    pub(super) fn conflicts(&self, other: &Self) -> bool {
        self.members
            .iter()
            .any(|left| other.members.iter().any(|right| left.conflicts(right)))
    }
    pub(super) fn join(&self, other: &Self, limits: Limits) -> Result<Self> {
        if self.conflicts(other) {
            return Err(invalid("patches conflict on metadata keys or owner markup"));
        }
        self.joined_len(other, limits)?;
        let count = self.members.len()
            + other
                .members
                .iter()
                .filter(|member| !self.members.contains(member))
                .count();
        let mut members = Vec::new();
        members.try_reserve_exact(count).map_err(allocation)?;
        members.extend(self.members.iter().cloned());
        for member in &other.members {
            if !members.contains(member) {
                members.push(member.clone());
            }
        }
        members.sort_by(|a, b| {
            a.owner.cmp(&b.owner).then_with(|| {
                a.changes
                    .iter()
                    .map(|c| &c.key)
                    .cmp(b.changes.iter().map(|c| &c.key))
                    .then_with(|| {
                        a.after
                            .as_ref()
                            .map(|i| (i.present, i.bytes.as_slice()))
                            .cmp(&b.after.as_ref().map(|i| (i.present, i.bytes.as_slice())))
                    })
            })
        });
        let result = Self { members };
        result.check(limits)?;
        Ok(result)
    }
    pub(super) fn joined_len(&self, other: &Self, limits: Limits) -> Result<usize> {
        let mut count = self.members.len();
        let mut length = self.encoded_len()?;
        for member in &other.members {
            if !self.members.contains(member) {
                count = count
                    .checked_add(1)
                    .ok_or_else(|| invalid("metadata count overflow"))?;
                let mut member_len = 3;
                if let Owner::Relationships(owner) = &member.owner {
                    add_len(&mut member_len, 4 + owner.as_str().len())?;
                }
                for image in [&member.before, &member.after].into_iter().flatten() {
                    add_len(&mut member_len, 5)?;
                    add_len(&mut member_len, image.bytes.len())?;
                }
                add_len(&mut length, member_len)?;
            }
        }
        if count > limits.max_parts().saturating_mul(2).saturating_add(2) {
            return Err(Error::Limit {
                kind: crate::opened::LimitKind::Parts,
                observed: count,
                maximum: limits.max_parts(),
            });
        }
        require_capture_bytes(length, limits)?;
        self.check(limits)?;
        other.check(limits)?;
        Ok(length)
    }
    pub(super) fn check(&self, limits: Limits) -> Result<()> {
        if self.members.len() > limits.max_parts().saturating_mul(2).saturating_add(2) {
            return Err(Error::Limit {
                kind: crate::opened::LimitKind::Parts,
                observed: self.members.len(),
                maximum: limits.max_parts(),
            });
        }
        require_capture_bytes(self.encoded_len()?, limits)?;
        let policy = source_limits(limits, limits.max_patch_bytes())?;
        for member in &self.members {
            for image in [&member.before, &member.after].into_iter().flatten() {
                Table::new(&member.owner, image, policy)?;
            }
        }
        Ok(())
    }
    pub(super) fn capture(
        source: &OpcPackage,
        target: &OpcPackage,
        deltas: &[(
            &PackURI,
            Option<super::BorrowedState<'_>>,
            Option<super::BorrowedState<'_>>,
        )],
        limits: Limits,
        base_len: usize,
    ) -> Result<Self> {
        let mut result = Self::default();
        let types_changed = deltas
            .iter()
            .any(|(_, a, b)| a.map(|v| v.content_type) != b.map(|v| v.content_type))
            || !source.same_retained_content_types(target)?;
        let mut owners = Vec::new();
        owners
            .try_reserve_exact(deltas.len().saturating_add(2))
            .map_err(allocation)?;
        let mut push_owner = |owner: Owner| -> Result<()> {
            if owners.len() >= limits.max_parts().saturating_mul(2).saturating_add(2) {
                return Err(Error::Limit {
                    kind: crate::opened::LimitKind::Parts,
                    observed: owners.len().checked_add(1).unwrap_or(usize::MAX),
                    maximum: limits.max_parts(),
                });
            }
            owners.try_reserve(1).map_err(allocation)?;
            owners.push(owner);
            Ok(())
        };
        if types_changed {
            push_owner(Owner::ContentTypes)?;
        }
        let root = PackURI::new("/").map_err(Error::Invalid)?;
        if !super::relationships_equal(source.rels(), target.rels())
            || !source.same_retained_relationships(target, &root)?
        {
            push_owner(Owner::Relationships(root))?;
        }
        for part in source.iter_parts() {
            let other = target.part_metadata(part.partname());
            if other.is_none_or(|other| !super::relationships_equal(part.rels(), other.rels()))
                || !source.same_retained_relationships(target, part.partname())?
            {
                push_owner(Owner::Relationships(part.partname().clone()))?;
            }
        }
        for part in target.iter_parts() {
            if source.part_metadata(part.partname()).is_none() {
                push_owner(Owner::Relationships(part.partname().clone()))?;
            }
        }
        owners.sort_unstable();
        let mut total = base_len
            .checked_add(4)
            .ok_or_else(|| invalid("metadata wire size overflow"))?;
        require_capture_bytes(total, limits)?;
        for owner in owners {
            let mut overhead = 3usize;
            if let Owner::Relationships(name) = &owner {
                add_len(&mut overhead, 4 + name.as_str().len())?;
            }
            add_len(&mut total, overhead)?;
            require_capture_bytes(total, limits)?;
            let mut capture = |package: &OpcPackage| -> Result<Option<Image>> {
                if let Owner::Relationships(owner) = &owner
                    && owner.as_str() != "/"
                    && package.part_metadata(owner).is_none()
                {
                    return Ok(None);
                }
                let policy = source_limits(limits, limits.max_patch_bytes().saturating_sub(total))?;
                let image = image(package, &owner, policy)?;
                if let Some(image) = &image {
                    add_len(&mut total, 5)?;
                    add_len(&mut total, image.bytes.len())?;
                    require_capture_bytes(total, limits)?;
                }
                Ok(image)
            };
            let before = capture(source)?;
            let after = capture(target)?;
            if before != after || (owner == Owner::ContentTypes && types_changed) {
                let member = Member::new(
                    owner,
                    before,
                    after,
                    source_limits(limits, limits.max_patch_bytes())?,
                )?;
                result.members.try_reserve(1).map_err(allocation)?;
                result.members.push(member);
            }
        }
        result.check(limits)?;
        Ok(result)
    }
    pub(super) fn encoded_len(&self) -> Result<usize> {
        u32::try_from(self.members.len())
            .map_err(|_| invalid("metadata count exceeds wire width"))?;
        let mut size = 4;
        for member in &self.members {
            add_len(&mut size, 3)?;
            if let Owner::Relationships(owner) = &member.owner {
                u32::try_from(owner.as_str().len())
                    .map_err(|_| invalid("metadata owner exceeds wire width"))?;
                add_len(&mut size, 4 + owner.as_str().len())?;
            }
            for image in [&member.before, &member.after].into_iter().flatten() {
                u32::try_from(image.bytes.len())
                    .map_err(|_| invalid("metadata image exceeds wire width"))?;
                add_len(&mut size, 5)?;
                add_len(&mut size, image.bytes.len())?;
            }
        }
        Ok(size)
    }
    pub(super) fn encode(&self, output: &mut Vec<u8>) -> Result<()> {
        put_u32(output, self.members.len())?;
        for member in &self.members {
            match &member.owner {
                Owner::ContentTypes => output.push(0),
                Owner::Relationships(owner) => {
                    output.push(1);
                    put_bytes32(output, owner.as_str().as_bytes())?;
                },
            }
            for image in [&member.before, &member.after] {
                output.push(u8::from(image.is_some()));
                if let Some(image) = image {
                    output.push(u8::from(image.present));
                    put_bytes32(output, &image.bytes)?;
                }
            }
        }
        Ok(())
    }
    pub(super) fn preflight(input: &mut Input<'_>, limits: Limits) -> Result<()> {
        let count = input.usize32()?;
        if count > limits.max_parts().saturating_mul(2).saturating_add(2) {
            return Err(Error::Limit {
                kind: crate::opened::LimitKind::Parts,
                observed: count,
                maximum: limits.max_parts(),
            });
        }
        for _ in 0..count {
            match input.byte()? {
                0 => {},
                1 => {
                    input.bytes32()?;
                },
                _ => return Err(invalid("invalid metadata owner tag")),
            }
            for _ in 0..2 {
                match input.byte()? {
                    0 => {},
                    1 => {
                        if input.byte()? > 1 {
                            return Err(invalid("invalid metadata presence"));
                        }
                        input.bytes32()?;
                    },
                    _ => return Err(invalid("invalid metadata image tag")),
                }
            }
        }
        Ok(())
    }
    pub(super) fn parse(input: &mut Input<'_>, limits: Limits) -> Result<Self> {
        let count = input.usize32()?;
        if count > limits.max_parts().saturating_mul(2).saturating_add(2) {
            return Err(Error::Limit {
                kind: crate::opened::LimitKind::Parts,
                observed: count,
                maximum: limits.max_parts(),
            });
        }
        let mut members = Vec::new();
        members.try_reserve_exact(count).map_err(allocation)?;
        for _ in 0..count {
            let owner = match input.byte()? {
                0 => Owner::ContentTypes,
                1 => Owner::Relationships(parse_part_name(input.bytes32()?)?),
                _ => return Err(invalid("invalid metadata owner tag")),
            };
            let mut read = || -> Result<Option<Image>> {
                match input.byte()? {
                    0 => Ok(None),
                    1 => {
                        let present = match input.byte()? {
                            0 => false,
                            1 => true,
                            _ => return Err(invalid("invalid metadata presence")),
                        };
                        let bytes = input.bytes32()?;
                        let mut owned = Vec::new();
                        owned.try_reserve_exact(bytes.len()).map_err(allocation)?;
                        owned.extend_from_slice(bytes);
                        Ok(Some(Image {
                            present,
                            bytes: Arc::new(owned),
                        }))
                    },
                    _ => Err(invalid("invalid metadata image tag")),
                }
            };
            let before = read()?;
            let after = read()?;
            let member = Member::new(
                owner,
                before,
                after,
                source_limits(limits, limits.max_patch_bytes())?,
            )?;
            if members
                .iter()
                .any(|other: &Member| other == &member || other.entry_conflicts(&member))
            {
                return Err(invalid("duplicate or conflicting metadata deltas"));
            }
            members.push(member);
        }
        let value = Self { members };
        value.check(limits)?;
        Ok(value)
    }
    pub(super) fn validate(&self, package: &OpcPackage, after: bool, limits: Limits) -> Result<()> {
        let policy = source_limits(limits, limits.max_patch_bytes())?;
        for member in &self.members {
            let expected = if after { &member.after } else { &member.before };
            let current = image(package, &member.owner, policy)?;
            if &current == expected {
                continue;
            }
            if member.opaque || current.is_none() || expected.is_none() {
                return Err(stale());
            }
            let (Some(current), Some(expected)) = (current, expected) else {
                return Err(stale());
            };
            if member.before.as_ref().map(|i| i.present) != member.after.as_ref().map(|i| i.present)
                && current.present != expected.present
            {
                return Err(stale());
            }
            let table = Table::new(&member.owner, &current, policy)?;
            for change in &member.changes {
                let span = if after { &change.after } else { &change.before };
                if table
                    .get(&change.key)
                    .map(|span| &current.bytes[span.clone()])
                    != span.as_ref().map(|span| &expected.bytes[span.clone()])
                {
                    return Err(stale());
                }
            }
        }
        Ok(())
    }
}

pub(super) struct Prepared {
    members: Vec<(Owner, Option<Image>)>,
}
impl Metadata {
    /// Plan against the original destination, before any payload mutation.
    pub(super) fn prepare(&self, package: &OpcPackage, limits: Limits) -> Result<Prepared> {
        self.validate(package, false, limits)?;
        let mut members = Vec::new();
        members
            .try_reserve_exact(self.members.len())
            .map_err(allocation)?;
        let mut charged = 0usize;
        for member in &self.members {
            if members.iter().any(|(owner, _)| owner == &member.owner) {
                continue;
            }
            let policy = source_limits(limits, limits.max_patch_bytes().saturating_sub(charged))?;
            let mut groups = Vec::new();
            groups
                .try_reserve_exact(self.members.len())
                .map_err(allocation)?;
            groups.extend(
                self.members
                    .iter()
                    .filter(|other| other.owner == member.owner),
            );
            let current = image(package, &member.owner, policy)?;
            let desired = if groups.len() == 1 && current == member.before {
                member.after.clone()
            } else {
                if groups.iter().any(|group| group.opaque) {
                    return Err(stale());
                }
                let current = current.ok_or_else(stale)?;
                let mut tables = Vec::new();
                tables.try_reserve_exact(groups.len()).map_err(allocation)?;
                let mut present = current.present;
                let mut count = 0usize;
                for group in &groups {
                    let before = group.before.as_ref().ok_or_else(stale)?;
                    let after = group.after.as_ref().ok_or_else(stale)?;
                    if before.present != after.present {
                        present = after.present;
                    }
                    count = count
                        .checked_add(group.changes.len())
                        .ok_or_else(|| invalid("metadata edit count overflow"))?;
                    tables.push((
                        Table::new(&group.owner, before, policy)?,
                        Table::new(&group.owner, after, policy)?,
                    ));
                }
                let mut seen = std::collections::HashSet::new();
                seen.try_reserve(count).map_err(allocation)?;
                let bytes = match &member.owner {
                    Owner::ContentTypes => {
                        let source = package.source_content_types(policy)?;
                        let mut edits = Vec::new();
                        edits.try_reserve_exact(count).map_err(allocation)?;
                        for (group, (before, after)) in groups.iter().zip(&tables) {
                            let (Catalog::Types(before), Catalog::Types(after)) =
                                (&before.catalog, &after.catalog)
                            else {
                                return Err(invalid("metadata catalog kind differs"));
                            };
                            for change in &group.changes {
                                if !seen.insert(&change.key) {
                                    continue;
                                }
                                let key = if change.key.kind == 0 {
                                    ContentTypeFragmentKey::Default(&change.key.text)
                                } else {
                                    ContentTypeFragmentKey::Override(&change.key.text)
                                };
                                edits.push(litchi_opc::ContentTypeFragmentEdit {
                                    before: before.get(key)?,
                                    after: after.get(key)?,
                                });
                            }
                        }
                        source
                            .plan_fragment_edits(&edits, policy)?
                            .materialize(policy)?
                            .bytes_arc()
                    },
                    Owner::Relationships(owner) => {
                        let source = package.source_relationships(owner, policy)?;
                        let mut edits = Vec::new();
                        edits.try_reserve_exact(count).map_err(allocation)?;
                        for (group, (before, after)) in groups.iter().zip(&tables) {
                            let (Catalog::Relationships(before), Catalog::Relationships(after)) =
                                (&before.catalog, &after.catalog)
                            else {
                                return Err(invalid("metadata catalog kind differs"));
                            };
                            for change in &group.changes {
                                if !seen.insert(&change.key) {
                                    continue;
                                }
                                edits.push(litchi_opc::RelationshipFragmentEdit {
                                    before: before.get(&change.key.text),
                                    after: after.get(&change.key.text),
                                });
                            }
                        }
                        source
                            .plan_fragment_edits(&edits, policy)?
                            .materialize(policy)?
                            .bytes_arc()
                    },
                };
                Some(Image { present, bytes })
            };
            if let Some(image) = &desired {
                add_len(&mut charged, image.bytes.len())?;
                require_capture_bytes(charged, limits)?;
            }
            members.push((member.owner.clone(), desired));
        }
        Ok(Prepared { members })
    }
}
impl Prepared {
    /// Install into an unpublished candidate after resource payload edits.
    pub(super) fn install(self, candidate: &mut OpcPackage, limits: Limits) -> Result<()> {
        let policy = source_limits(limits, limits.max_patch_bytes())?;
        for (owner, desired) in self.members {
            match (owner, desired) {
                (Owner::ContentTypes, Some(image)) => {
                    let current = candidate.source_content_types(policy)?;
                    candidate.try_replace_content_types_bytes(
                        current.bytes(),
                        image.bytes,
                        policy,
                    )?;
                },
                (Owner::Relationships(owner), Some(image)) => {
                    let current = candidate.source_relationships(&owner, policy)?;
                    candidate.try_replace_relationships_bytes(
                        &current,
                        image.bytes,
                        image.present,
                        policy,
                    )?;
                },
                (Owner::Relationships(owner), None)
                    if owner.as_str() != "/" && candidate.part_metadata(&owner).is_none() => {},
                _ => return Err(invalid("metadata target owner presence differs")),
            }
        }
        Ok(())
    }
}

#[cfg(test)]
mod tests {
    #![allow(clippy::unwrap_used, reason = "test assertions")]
    use super::*;

    fn image(body: &str, present: bool) -> Image {
        Image {
            present,
            bytes: Arc::new(format!("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?><Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">{body}</Relationships>").into_bytes()),
        }
    }
    fn edge(id: &str) -> String {
        format!(
            "<Relationship Id=\"{id}\" Type=\"urn:test\" Target=\"https://example.com/{id}\" TargetMode=\"External\"/>"
        )
    }
    fn owner() -> Owner {
        Owner::Relationships(PackURI::new("/custom/test.xml").unwrap())
    }
    fn member(before: Image, after: Image) -> Member {
        Member::new(owner(), Some(before), Some(after), SourceLimits::default()).unwrap()
    }

    #[test]
    fn disjoint_metadata_joins_require_the_same_full_before_image() {
        let a = member(
            image("<!--a-->", true),
            image(&format!("{}<!--a-->", edge("one")), true),
        );
        let b = member(
            image("<!--b-->", true),
            image(&format!("{}<!--b-->", edge("two")), true),
        );
        assert!(!a.opaque && !b.opaque);
        assert!(a.conflicts(&b));
    }

    #[test]
    fn removing_member_presence_conflicts_with_another_entry_edit() {
        let name = PackURI::new("/custom/test.xml").unwrap();
        let graph = litchi_opc::Relationships::new(name.base_uri().to_owned());
        let canonical = Image {
            present: false,
            bytes: Arc::new(
                graph
                    .plan_canonical(SourceLimits::default())
                    .unwrap()
                    .materialize(SourceLimits::default())
                    .unwrap(),
            ),
        };
        let present = Image {
            present: true,
            ..canonical.clone()
        };
        let removal = member(present.clone(), canonical);
        let token = RelationshipFragments::from_bytes(
            name,
            Arc::clone(&present.bytes),
            SourceLimits::default(),
        )
        .unwrap();
        assert_eq!(token.iter().len(), 0);
        let addition = member(present, image(&edge("two"), true));
        assert!(removal.conflicts(&addition));
    }

    #[test]
    fn absent_images_require_canonical_empty_relationships() {
        assert!(
            Table::new(
                &owner(),
                &image("<!--noncanonical-->", false),
                SourceLimits::default()
            )
            .is_err()
        );
        assert!(
            Table::new(
                &Owner::Relationships(PackURI::new("/").unwrap()),
                &image("", false),
                SourceLimits::default()
            )
            .is_err()
        );
    }

    #[test]
    fn undeclared_relationship_changes_cannot_hide_in_metadata_only_wire() {
        let metadata = Metadata {
            members: vec![member(image("", true), image(&edge("added"), true))],
        };
        assert!(
            metadata
                .require_coverage(&[], None, Limits::default())
                .is_err()
        );
        let spelling = Metadata {
            members: vec![member(image("", true), image("<!--spelling-->", true))],
        };
        spelling
            .require_coverage(&[], None, Limits::default())
            .unwrap();
    }
    #[test]
    fn topology_requires_content_type_provenance_even_when_default_already_covers_it() {
        use super::super::Patch;
        let source = OpcPackage::from_bytes(
            &crate::Package::new().unwrap().to_bytes().unwrap(),
            SourceLimits::default(),
        )
        .unwrap();
        let mut target = source.clone();
        target
            .try_add_part(Box::new(litchi_opc::BlobPart::new(
                PackURI::new("/custom/proof.xml").unwrap(),
                "application/xml".into(),
                b"<proof/>".to_vec(),
            )))
            .unwrap();
        let patch = Patch::capture(
            &source,
            &target,
            PackURI::new("/ppt/presentation.xml").unwrap(),
            Limits::default(),
        )
        .unwrap();
        let proof = patch
            .metadata
            .members
            .iter()
            .find(|member| member.owner == Owner::ContentTypes)
            .unwrap();
        assert!(
            proof.before == proof.after,
            "fixture must exercise an unchanged default binding"
        );
        assert!(proof.opaque);
        let decoded = Patch::from_bytes(&patch.to_bytes().unwrap(), Limits::default()).unwrap();
        let mut replayed = source.clone();
        super::super::apply(&mut replayed, &decoded, true).unwrap();
        let mut overlay = source.clone();
        let token = overlay.source_content_types(overlay.read_limits()).unwrap();
        let mut xml = token.bytes().to_vec();
        xml.extend_from_slice(b"<!--unrelated trailing comment-->");
        let limits = overlay.read_limits();
        overlay
            .try_replace_content_types_bytes(token.bytes(), Arc::new(xml), limits)
            .unwrap();
        let original = litchi_opc::PackageWriter::to_bytes(&overlay).unwrap();
        assert!(super::super::apply(&mut overlay, &decoded, true).is_err());
        assert_eq!(
            original,
            litchi_opc::PackageWriter::to_bytes(&overlay).unwrap()
        );
        let mut forged = patch;
        forged
            .metadata
            .members
            .retain(|member| member.owner != Owner::ContentTypes);
        assert!(Patch::from_bytes(&forged.to_bytes().unwrap(), Limits::default()).is_err());
        let mut inverse = decoded.inverse();
        inverse
            .metadata
            .members
            .retain(|member| member.owner != Owner::ContentTypes);
        assert!(Patch::from_bytes(&inverse.to_bytes().unwrap(), Limits::default()).is_err());
    }
}
