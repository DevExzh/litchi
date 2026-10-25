//! Retained relationship closures used by source-backed semantic replay.
//!
//! The index is built once while an OPC package enters the facade.  Published
//! package versions path-copy the changed keys, so a replay validates only its
//! affected closure instead of walking every unrelated part again.

use std::cmp::Ordering;
use std::collections::{HashMap, HashSet};
use std::sync::Arc;

use litchi_opc::constants::content_type as ct;
use litchi_opc::{OpcPackage, PackURI, Part, Relationships};

use crate::{Error, Result};

/// A small immutable AVL map keeps replay index updates proportional to the
/// changed keys.  The index used to retain a linked overlay chain and flatten
/// it periodically; that made an otherwise local edit copy every indexed
/// part at the checkpoint.  Path-copying this map retains only O(log n) nodes
/// per changed key and has no publication-history chain to compact.
#[derive(Clone, Debug)]
struct PersistentMap<V> {
    root: Option<Arc<PersistentNode<V>>>,
}

#[derive(Clone, Debug)]
struct PersistentNode<V> {
    key: PackURI,
    value: V,
    height: usize,
    left: Option<Arc<PersistentNode<V>>>,
    right: Option<Arc<PersistentNode<V>>>,
}

impl<V> Default for PersistentMap<V> {
    fn default() -> Self {
        Self { root: None }
    }
}

impl<V> PersistentMap<V> {
    fn from_hash_map(entries: HashMap<PackURI, V>) -> Self {
        let mut entries: Vec<Option<(PackURI, V)>> = entries.into_iter().map(Some).collect();
        entries.sort_unstable_by(|left, right| {
            let (Some((left, _)), Some((right, _))) = (left.as_ref(), right.as_ref()) else {
                return Ordering::Equal;
            };
            left.as_str().cmp(right.as_str())
        });
        Self {
            root: build_balanced(&mut entries),
        }
    }
}

impl<V: Clone> PersistentMap<V> {
    fn get(&self, key: &PackURI) -> Option<&V> {
        let mut current = self.root.as_ref();
        while let Some(node) = current {
            match key.as_str().cmp(node.key.as_str()) {
                Ordering::Less => current = node.left.as_ref(),
                Ordering::Equal => return Some(&node.value),
                Ordering::Greater => current = node.right.as_ref(),
            }
        }
        None
    }

    fn insert(&self, key: PackURI, value: V) -> Self {
        Self {
            root: Some(insert_persistent(self.root.as_ref(), key, value)),
        }
    }

    fn remove(&self, key: &PackURI) -> Self {
        if self.get(key).is_none() {
            return self.clone();
        }
        Self {
            root: remove_persistent(self.root.as_ref(), key),
        }
    }

    fn first_key(&self) -> Option<&PackURI> {
        let mut current = self.root.as_ref();
        let mut first = None;
        while let Some(node) = current {
            first = Some(&node.key);
            current = node.left.as_ref();
        }
        first
    }
}

fn build_balanced<V>(entries: &mut [Option<(PackURI, V)>]) -> Option<Arc<PersistentNode<V>>> {
    if entries.is_empty() {
        return None;
    }
    let middle = entries.len() / 2;
    let left = build_balanced(&mut entries[..middle]);
    let right = build_balanced(&mut entries[middle + 1..]);
    let (key, value) = entries[middle].take()?;
    Some(make_node(key, value, left, right))
}

fn node_height<V>(node: Option<&Arc<PersistentNode<V>>>) -> usize {
    node.map_or(0, |node| node.height)
}

fn make_node<V>(
    key: PackURI,
    value: V,
    left: Option<Arc<PersistentNode<V>>>,
    right: Option<Arc<PersistentNode<V>>>,
) -> Arc<PersistentNode<V>> {
    let height = 1 + node_height(left.as_ref()).max(node_height(right.as_ref()));
    Arc::new(PersistentNode {
        key,
        value,
        height,
        left,
        right,
    })
}

fn balance_factor<V>(node: &Arc<PersistentNode<V>>) -> isize {
    node_height(node.left.as_ref()) as isize - node_height(node.right.as_ref()) as isize
}

fn rotate_right<V: Clone>(node: Arc<PersistentNode<V>>) -> Arc<PersistentNode<V>> {
    let Some(pivot) = node.left.as_ref() else {
        return node;
    };
    let new_right = make_node(
        node.key.clone(),
        node.value.clone(),
        pivot.right.clone(),
        node.right.clone(),
    );
    make_node(
        pivot.key.clone(),
        pivot.value.clone(),
        pivot.left.clone(),
        Some(new_right),
    )
}

fn rotate_left<V: Clone>(node: Arc<PersistentNode<V>>) -> Arc<PersistentNode<V>> {
    let Some(pivot) = node.right.as_ref() else {
        return node;
    };
    let new_left = make_node(
        node.key.clone(),
        node.value.clone(),
        node.left.clone(),
        pivot.left.clone(),
    );
    make_node(
        pivot.key.clone(),
        pivot.value.clone(),
        Some(new_left),
        pivot.right.clone(),
    )
}

fn rebalance<V: Clone>(mut node: Arc<PersistentNode<V>>) -> Arc<PersistentNode<V>> {
    let factor = balance_factor(&node);
    if factor > 1 {
        if node
            .left
            .as_ref()
            .is_some_and(|left| balance_factor(left) < 0)
        {
            let Some(left_node) = node.left.as_ref().cloned() else {
                return node;
            };
            let left = rotate_left(left_node);
            node = make_node(
                node.key.clone(),
                node.value.clone(),
                Some(left),
                node.right.clone(),
            );
        }
        return rotate_right(node);
    }
    if factor < -1 {
        if node
            .right
            .as_ref()
            .is_some_and(|right| balance_factor(right) > 0)
        {
            let Some(right_node) = node.right.as_ref().cloned() else {
                return node;
            };
            let right = rotate_right(right_node);
            node = make_node(
                node.key.clone(),
                node.value.clone(),
                node.left.clone(),
                Some(right),
            );
        }
        return rotate_left(node);
    }
    node
}

fn insert_persistent<V: Clone>(
    node: Option<&Arc<PersistentNode<V>>>,
    key: PackURI,
    value: V,
) -> Arc<PersistentNode<V>> {
    let Some(node) = node else {
        return make_node(key, value, None, None);
    };
    match key.as_str().cmp(node.key.as_str()) {
        Ordering::Less => rebalance(make_node(
            node.key.clone(),
            node.value.clone(),
            Some(insert_persistent(node.left.as_ref(), key, value)),
            node.right.clone(),
        )),
        Ordering::Equal => make_node(
            node.key.clone(),
            value,
            node.left.clone(),
            node.right.clone(),
        ),
        Ordering::Greater => rebalance(make_node(
            node.key.clone(),
            node.value.clone(),
            node.left.clone(),
            Some(insert_persistent(node.right.as_ref(), key, value)),
        )),
    }
}

fn take_min<V: Clone>(
    node: &Arc<PersistentNode<V>>,
) -> (PackURI, V, Option<Arc<PersistentNode<V>>>) {
    let Some(left) = node.left.as_ref() else {
        return (node.key.clone(), node.value.clone(), node.right.clone());
    };
    let (key, value, new_left) = take_min(left);
    let rebuilt = rebalance(make_node(
        node.key.clone(),
        node.value.clone(),
        new_left,
        node.right.clone(),
    ));
    (key, value, Some(rebuilt))
}

fn remove_persistent<V: Clone>(
    node: Option<&Arc<PersistentNode<V>>>,
    key: &PackURI,
) -> Option<Arc<PersistentNode<V>>> {
    let node = node?;
    match key.as_str().cmp(node.key.as_str()) {
        Ordering::Less => Some(rebalance(make_node(
            node.key.clone(),
            node.value.clone(),
            remove_persistent(node.left.as_ref(), key),
            node.right.clone(),
        ))),
        Ordering::Greater => Some(rebalance(make_node(
            node.key.clone(),
            node.value.clone(),
            node.left.clone(),
            remove_persistent(node.right.as_ref(), key),
        ))),
        Ordering::Equal => match (node.left.as_ref(), node.right.as_ref()) {
            (None, None) => None,
            (Some(left), None) => Some(left.clone()),
            (None, Some(right)) => Some(right.clone()),
            (Some(left), Some(right)) => {
                let (successor_key, successor_value, successor_right) = take_min(right);
                Some(rebalance(make_node(
                    successor_key,
                    successor_value,
                    Some(left.clone()),
                    successor_right,
                )))
            },
        },
    }
}

#[derive(Debug, Clone, PartialEq, Eq)]
enum NotesEntry {
    Slide { bytes: u64 },
    Master { bytes: u64, themes: Arc<[PackURI]> },
}

/// The source-backed closures needed by the notes and media replay guards.
///
/// Every map is a persistent path-copying map. A package edit therefore copies
/// only the changed key paths; lookups stay logarithmic in the retained graph,
/// and a long publication history cannot retain an unbounded overlay chain or
/// trigger a periodic full-index flatten. The cached notes total is updated as
/// a delta and never recomputed from the package.
#[derive(Debug, Clone)]
pub(crate) struct ReplayIndex {
    notes_entries: PersistentMap<NotesEntry>,
    notes_theme_users: PersistentMap<Arc<[PackURI]>>,
    missing_themes: PersistentMap<()>,
    missing_theme_hint: Option<PackURI>,
    media_slides: PersistentMap<Arc<[PackURI]>>,
    notes_total_bytes: u64,
}

impl ReplayIndex {
    /// Build the initial index during package ingress.  This is the only
    /// package-wide walk; later package versions use [`Self::update`].
    pub(crate) fn from_package(package: &OpcPackage) -> Result<Self> {
        // Keep ingress reservations proportional to the expected relevant
        // edges. The maps grow fallibly as the graph reveals more owners.
        let part_count = package.part_count().min(256);
        let mut notes_entries: HashMap<PackURI, NotesEntry> = HashMap::new();
        notes_entries
            .try_reserve(part_count)
            .map_err(|source| Error::Allocation {
                resource: "PPTX replay notes index",
                source,
            })?;
        let mut notes_theme_users: HashMap<PackURI, Vec<PackURI>> = HashMap::new();
        notes_theme_users
            .try_reserve(part_count)
            .map_err(|source| Error::Allocation {
                resource: "PPTX replay notes theme index",
                source,
            })?;
        let mut media_slides: HashMap<PackURI, Vec<PackURI>> = HashMap::new();
        media_slides
            .try_reserve(part_count)
            .map_err(|source| Error::Allocation {
                resource: "PPTX replay media relationship index",
                source,
            })?;
        let mut notes_total_bytes = 0_u64;

        for part in package.iter_parts() {
            let name = part.partname();
            let bytes = metadata_bytes(part)?;
            match part.content_type() {
                ct::PML_NOTES_SLIDE => {
                    notes_total_bytes = notes_total_bytes.checked_add(bytes).ok_or_else(|| {
                        Error::Invalid("PPTX replay notes index byte count overflow".into())
                    })?;
                    insert_notes_entry(
                        &mut notes_entries,
                        name.clone(),
                        NotesEntry::Slide { bytes },
                    )?;
                },
                ct::PML_NOTES_MASTER => {
                    let themes = notes_theme_targets(part.rels())?;
                    notes_total_bytes = notes_total_bytes.checked_add(bytes).ok_or_else(|| {
                        Error::Invalid("PPTX replay notes index byte count overflow".into())
                    })?;
                    for theme in &themes {
                        notes_total_bytes = notes_total_bytes
                            .checked_add(theme_bytes_if_present(package, theme)?)
                            .ok_or_else(|| {
                                Error::Invalid("PPTX replay notes index byte count overflow".into())
                            })?;
                    }
                    insert_notes_entry(
                        &mut notes_entries,
                        name.clone(),
                        NotesEntry::Master {
                            bytes,
                            themes: Arc::from(themes.clone()),
                        },
                    )?;
                    for theme in themes {
                        push_theme_user(&mut notes_theme_users, theme, name.clone())?;
                    }
                },
                _ => {},
            }
            if part.content_type() == ct::PML_SLIDE {
                for target in media_targets(part.rels())? {
                    push_media_slide(&mut media_slides, target, name.clone())?;
                }
            }
        }

        let missing_themes =
            persistent_from_hash_map(missing_theme_map(package, &notes_theme_users)?);
        let missing_theme_hint = missing_themes.first_key().cloned();
        let notes_entries = persistent_from_hash_map(notes_entries);
        let notes_theme_users = notes_theme_users
            .into_iter()
            .filter(|(_, users)| !users.is_empty())
            .fold(PersistentMap::default(), |map, (name, users)| {
                map.insert(name, Arc::from(users))
            });
        let media_slides = media_slides
            .into_iter()
            .fold(PersistentMap::default(), |map, (name, slides)| {
                map.insert(name, Arc::from(slides))
            });
        Ok(Self {
            notes_total_bytes,
            notes_entries,
            notes_theme_users,
            missing_themes,
            missing_theme_hint,
            media_slides,
        })
    }

    /// The retained notes aggregate, updated without reading unrelated parts.
    pub(crate) const fn notes_total_bytes(&self) -> u64 {
        self.notes_total_bytes
    }

    /// Slides that currently reference a media resource.
    pub(crate) fn media_slides(&self, resource: &PackURI) -> Option<Arc<[PackURI]>> {
        self.media_slides.get(resource).cloned()
    }

    fn notes_entry(&self, name: &PackURI) -> Option<NotesEntry> {
        self.notes_entries.get(name).cloned()
    }

    fn theme_users(&self, name: &PackURI) -> Arc<[PackURI]> {
        self.notes_theme_users
            .get(name)
            .cloned()
            .unwrap_or_else(|| Arc::from([]))
    }

    pub(crate) fn missing_theme_name(&self) -> Result<Option<PackURI>> {
        Ok(self.missing_theme_hint.clone())
    }

    /// Apply changed part names to the retained index.
    ///
    /// `changed` is the exact `PartPatch` write set.  Notes totals are
    /// adjusted for changed note slides/masters/themes; media reverse edges
    /// are adjusted only for changed slide relationship graphs.
    pub(crate) fn update(
        &self,
        source: &OpcPackage,
        target: &OpcPackage,
        changed: &[PackURI],
    ) -> Result<Self> {
        if changed.is_empty() {
            return Ok(self.clone());
        }
        let mut notes_entries = self.notes_entries.clone();
        let mut notes_theme_users: HashMap<PackURI, Vec<PackURI>> = HashMap::new();
        let mut changed_masters = HashSet::new();
        let mut changed_themes = HashSet::new();
        let mut affected_themes = HashSet::new();
        changed_masters
            .try_reserve(changed.len())
            .map_err(|source| Error::Allocation {
                resource: "PPTX replay changed notes masters",
                source,
            })?;
        changed_themes
            .try_reserve(changed.len())
            .map_err(|source| Error::Allocation {
                resource: "PPTX replay changed notes themes",
                source,
            })?;
        affected_themes
            .try_reserve(changed.len())
            .map_err(|source| Error::Allocation {
                resource: "PPTX replay affected notes themes",
                source,
            })?;
        for name in changed {
            let old = optional_part(source, name)?;
            let new = optional_part(target, name)?;
            if old
                .as_ref()
                .is_some_and(|part| part.content_type() == ct::PML_NOTES_MASTER)
                || new
                    .as_ref()
                    .is_some_and(|part| part.content_type() == ct::PML_NOTES_MASTER)
            {
                changed_masters.insert(name.clone());
            }
            if old
                .as_ref()
                .is_some_and(|part| part.content_type() == ct::OFC_THEME)
                || new
                    .as_ref()
                    .is_some_and(|part| part.content_type() == ct::OFC_THEME)
            {
                changed_themes.insert(name.clone());
            }
        }

        let mut notes_total_bytes = self.notes_total_bytes;
        for name in changed {
            let old = optional_part(source, name)?;
            let new = optional_part(target, name)?;
            let old_entry = self.notes_entry(name);
            let new_entry = if let Some(part) = new {
                match part.content_type() {
                    ct::PML_NOTES_SLIDE => Some(NotesEntry::Slide {
                        bytes: part_bytes(part)?,
                    }),
                    ct::PML_NOTES_MASTER => Some(NotesEntry::Master {
                        bytes: part_bytes(part)?,
                        themes: Arc::from(notes_theme_targets(part.rels())?),
                    }),
                    _ => None,
                }
            } else {
                None
            };
            if old_entry.is_some() || new_entry.is_some() {
                let old_value = notes_contribution_if_present(old_entry.as_ref(), source)?;
                let new_value = notes_contribution(new_entry.as_ref(), target)?;
                notes_total_bytes = adjust_total(notes_total_bytes, old_value, new_value)?;
                notes_entries = match new_entry.as_ref() {
                    Some(entry) => notes_entries.insert(name.clone(), entry.clone()),
                    None => notes_entries.remove(name),
                };
            } else if old
                .as_ref()
                .is_some_and(|part| part.content_type() == ct::PML_NOTES_SLIDE)
                || new
                    .as_ref()
                    .is_some_and(|part| part.content_type() == ct::PML_NOTES_SLIDE)
            {
                notes_entries = match new_entry.as_ref() {
                    Some(entry) => notes_entries.insert(name.clone(), entry.clone()),
                    None => notes_entries.remove(name),
                };
            }
            if changed_masters.contains(name) {
                let old_themes = master_themes(old_entry.as_ref());
                let new_themes = master_themes(new_entry.as_ref());
                if old_themes != new_themes {
                    for theme in old_themes {
                        affected_themes.insert(theme.clone());
                        remove_theme_user(&mut notes_theme_users, self, theme.clone(), name)?;
                    }
                    for theme in new_themes {
                        affected_themes.insert(theme.clone());
                        add_theme_user(&mut notes_theme_users, self, theme.clone(), name)?;
                    }
                }
            }
        }

        affected_themes.extend(changed_themes.iter().cloned());

        let mut missing_themes = self.missing_themes.clone();
        for theme in &affected_themes {
            let users = notes_theme_users
                .get(theme)
                .cloned()
                .unwrap_or_else(|| self.theme_users(theme).to_vec());
            let new_missing = !users.is_empty() && optional_part(target, theme)?.is_none();
            if new_missing {
                missing_themes = missing_themes.insert(theme.clone(), ());
            } else {
                missing_themes = missing_themes.remove(theme);
            }
        }

        let missing_theme_hint = missing_themes.first_key().cloned();

        // A changed theme contributes to each unchanged notes master that
        // references it. Masters changed in this same patch were already
        // charged their complete before/after contribution above.
        for theme in &changed_themes {
            let users = notes_theme_users
                .get(theme)
                .cloned()
                .unwrap_or_else(|| self.theme_users(theme).to_vec());
            if users.is_empty() {
                continue;
            }
            let old_bytes = theme_bytes_if_present(source, theme)?;
            let new_bytes = theme_bytes(target, theme)?;
            if old_bytes == new_bytes {
                continue;
            }
            for user in users {
                if changed_masters.contains(&user) {
                    continue;
                }
                notes_total_bytes = adjust_total(notes_total_bytes, old_bytes, new_bytes)?;
            }
        }

        let mut media_updates = HashMap::new();
        media_updates
            .try_reserve(changed.len())
            .map_err(|source| Error::Allocation {
                resource: "PPTX replay changed media relationships",
                source,
            })?;
        for name in changed {
            let old = optional_part(source, name)?;
            let new = optional_part(target, name)?;
            let old_slide = old
                .as_ref()
                .is_some_and(|part| part.content_type() == ct::PML_SLIDE);
            let new_slide = new
                .as_ref()
                .is_some_and(|part| part.content_type() == ct::PML_SLIDE);
            if !old_slide && !new_slide {
                continue;
            }
            let old_targets = old
                .filter(|part| part.content_type() == ct::PML_SLIDE)
                .map(|part| media_targets(part.rels()))
                .transpose()?
                .unwrap_or_default();
            let new_targets = new
                .filter(|part| part.content_type() == ct::PML_SLIDE)
                .map(|part| media_targets(part.rels()))
                .transpose()?
                .unwrap_or_default();
            let mut resources = old_targets.clone();
            resources.extend(new_targets.iter().cloned());
            resources.sort_unstable_by(|left, right| left.as_str().cmp(right.as_str()));
            resources.dedup();
            for resource in resources {
                media_updates
                    .try_reserve(1)
                    .map_err(|source| Error::Allocation {
                        resource: "PPTX replay changed media relationships",
                        source,
                    })?;
                let users = media_updates.entry(resource.clone()).or_insert_with(|| {
                    self.media_slides(&resource)
                        .map_or_else(Vec::new, |value| value.to_vec())
                });
                users.retain(|slide| slide != name);
                if new_targets.iter().any(|target| target == &resource) {
                    if !users.iter().any(|slide| slide == name) {
                        users.try_reserve(1).map_err(|source| Error::Allocation {
                            resource: "PPTX replay media relationship users",
                            source,
                        })?;
                        users.push(name.clone());
                    }
                }
            }
        }

        let mut media_slides = self.media_slides.clone();
        for (name, slides) in media_updates {
            media_slides = if slides.is_empty() {
                media_slides.remove(&name)
            } else {
                media_slides.insert(name, Arc::from(slides))
            };
        }

        Ok(Self {
            notes_total_bytes,
            notes_entries,
            notes_theme_users: notes_theme_users.into_iter().fold(
                self.notes_theme_users.clone(),
                |map, (name, users)| {
                    if users.is_empty() {
                        map.remove(&name)
                    } else {
                        map.insert(name, Arc::from(users))
                    }
                },
            ),
            missing_themes,
            missing_theme_hint,
            media_slides,
        })
    }
}

fn persistent_from_hash_map<V>(entries: HashMap<PackURI, V>) -> PersistentMap<V> {
    PersistentMap::from_hash_map(entries)
}

fn optional_part<'a>(package: &'a OpcPackage, name: &PackURI) -> Result<Option<&'a dyn Part>> {
    match package.get_part(name) {
        Ok(part) => Ok(Some(part)),
        Err(litchi_opc::OpcError::PartNotFound(_)) => Ok(None),
        Err(error) => Err(error.into()),
    }
}

fn part_bytes(part: &dyn Part) -> Result<u64> {
    u64::try_from(part.blob().len())
        .map_err(|_| Error::Invalid("PPTX replay index byte count overflow".into()))
}

fn metadata_bytes(part: litchi_opc::PartMetadata<'_>) -> Result<u64> {
    u64::try_from(part.payload_len()?)
        .map_err(|_| Error::Invalid("PPTX replay index byte count overflow".into()))
}

fn master_themes(entry: Option<&NotesEntry>) -> &[PackURI] {
    match entry {
        Some(NotesEntry::Master { themes, .. }) => themes.as_ref(),
        _ => &[],
    }
}

fn notes_contribution(entry: Option<&NotesEntry>, package: &OpcPackage) -> Result<u64> {
    let Some(entry) = entry else {
        return Ok(0);
    };
    match entry {
        NotesEntry::Slide { bytes } => Ok(*bytes),
        NotesEntry::Master { bytes, themes } => {
            let mut total = *bytes;
            for theme in themes.iter() {
                total = total
                    .checked_add(theme_bytes(package, theme)?)
                    .ok_or_else(|| Error::Invalid("PPTX replay notes total overflow".into()))?;
            }
            Ok(total)
        },
    }
}

fn notes_contribution_if_present(entry: Option<&NotesEntry>, package: &OpcPackage) -> Result<u64> {
    let Some(entry) = entry else {
        return Ok(0);
    };
    match entry {
        NotesEntry::Slide { bytes } => Ok(*bytes),
        NotesEntry::Master { bytes, themes } => {
            let mut total = *bytes;
            for theme in themes.iter() {
                total = total
                    .checked_add(theme_bytes_if_present(package, theme)?)
                    .ok_or_else(|| Error::Invalid("PPTX replay notes total overflow".into()))?;
            }
            Ok(total)
        },
    }
}

fn theme_bytes(package: &OpcPackage, name: &PackURI) -> Result<u64> {
    let part = package
        .part_metadata(name)
        .ok_or_else(|| litchi_opc::OpcError::PartNotFound(name.to_string()))?;
    if part.content_type() != ct::OFC_THEME {
        return Ok(0);
    }
    metadata_bytes(part)
}

fn theme_bytes_if_present(package: &OpcPackage, name: &PackURI) -> Result<u64> {
    let Some(part) = package.part_metadata(name) else {
        return Ok(0);
    };
    if part.content_type() != ct::OFC_THEME {
        return Ok(0);
    }
    metadata_bytes(part)
}

fn notes_theme_targets(rels: &Relationships) -> Result<Vec<PackURI>> {
    let mut result = Vec::new();
    for relationship in rels
        .iter()
        .filter(|relationship| relationship.reltype().ends_with("/theme"))
        .filter(|relationship| !relationship.is_external())
    {
        result.try_reserve(1).map_err(|source| Error::Allocation {
            resource: "PPTX replay notes theme targets",
            source,
        })?;
        result.push(relationship.target_partname().map_err(Error::Opc)?);
    }
    Ok(result)
}

fn missing_theme_map(
    package: &OpcPackage,
    users: &HashMap<PackURI, Vec<PackURI>>,
) -> Result<HashMap<PackURI, ()>> {
    let mut missing = HashMap::new();
    missing
        .try_reserve(users.len())
        .map_err(|source| Error::Allocation {
            resource: "PPTX replay missing notes themes",
            source,
        })?;
    for (theme, owners) in users {
        if !owners.is_empty() && package.part_metadata(theme).is_none() {
            missing.try_reserve(1).map_err(|source| Error::Allocation {
                resource: "PPTX replay missing notes themes",
                source,
            })?;
            missing.insert(theme.clone(), ());
        }
    }
    Ok(missing)
}

fn media_targets(rels: &Relationships) -> Result<Vec<PackURI>> {
    let mut result = Vec::new();
    for relationship in rels.iter() {
        if relationship.is_external() {
            continue;
        }
        let Some(target) = relationship.target_partname().ok() else {
            continue;
        };
        if !target.as_str().starts_with("/ppt/media/") {
            continue;
        }
        result.try_reserve(1).map_err(|source| Error::Allocation {
            resource: "PPTX replay media relationship targets",
            source,
        })?;
        if !result.iter().any(|current| current == &target) {
            result.push(target);
        }
    }
    Ok(result)
}

fn push_media_slide(
    users: &mut HashMap<PackURI, Vec<PackURI>>,
    resource: PackURI,
    slide: PackURI,
) -> Result<()> {
    users.try_reserve(1).map_err(|source| Error::Allocation {
        resource: "PPTX replay media relationship users",
        source,
    })?;
    let values = users.entry(resource).or_default();
    if values.iter().any(|current| current == &slide) {
        return Ok(());
    }
    values.try_reserve(1).map_err(|source| Error::Allocation {
        resource: "PPTX replay media relationship users",
        source,
    })?;
    values.push(slide);
    Ok(())
}

fn insert_notes_entry(
    entries: &mut HashMap<PackURI, NotesEntry>,
    name: PackURI,
    value: NotesEntry,
) -> Result<()> {
    entries.try_reserve(1).map_err(|source| Error::Allocation {
        resource: "PPTX replay notes entry",
        source,
    })?;
    entries.insert(name, value);
    Ok(())
}

fn push_theme_user(
    users: &mut HashMap<PackURI, Vec<PackURI>>,
    theme: PackURI,
    master: PackURI,
) -> Result<()> {
    users.try_reserve(1).map_err(|source| Error::Allocation {
        resource: "PPTX replay notes theme users",
        source,
    })?;
    let values = users.entry(theme).or_default();
    values.try_reserve(1).map_err(|source| Error::Allocation {
        resource: "PPTX replay notes theme users",
        source,
    })?;
    values.push(master);
    Ok(())
}

fn adjust_total(total: u64, old: u64, new: u64) -> Result<u64> {
    total
        .checked_sub(old)
        .and_then(|value| value.checked_add(new))
        .ok_or_else(|| Error::Invalid("PPTX replay notes total overflow".into()))
}

fn add_theme_user(
    updates: &mut HashMap<PackURI, Vec<PackURI>>,
    index: &ReplayIndex,
    theme: PackURI,
    master: &PackURI,
) -> Result<()> {
    updates.try_reserve(1).map_err(|source| Error::Allocation {
        resource: "PPTX replay notes theme users",
        source,
    })?;
    let users = updates
        .entry(theme.clone())
        .or_insert_with(|| index.theme_users(&theme).to_vec());
    users.try_reserve(1).map_err(|source| Error::Allocation {
        resource: "PPTX replay notes theme users",
        source,
    })?;
    users.push(master.clone());
    Ok(())
}

fn remove_theme_user(
    updates: &mut HashMap<PackURI, Vec<PackURI>>,
    index: &ReplayIndex,
    theme: PackURI,
    master: &PackURI,
) -> Result<()> {
    updates.try_reserve(1).map_err(|source| Error::Allocation {
        resource: "PPTX replay notes theme users",
        source,
    })?;
    let users = updates
        .entry(theme.clone())
        .or_insert_with(|| index.theme_users(&theme).to_vec());
    if let Some(position) = users.iter().position(|user| user == master) {
        users.remove(position);
    }
    Ok(())
}

#[cfg(test)]
#[allow(
    clippy::expect_used,
    clippy::unwrap_used,
    reason = "persistent-map assertions use panic-on-failure test helpers"
)]
mod tests {
    use super::*;
    use std::collections::HashMap as StdHashMap;

    fn name(value: &str) -> PackURI {
        PackURI::new(format!("/{value}")).expect("valid test URI")
    }

    fn check_tree<V>(node: Option<&Arc<PersistentNode<V>>>) -> usize {
        let Some(node) = node else {
            return 0;
        };
        let left = check_tree(node.left.as_ref());
        let right = check_tree(node.right.as_ref());
        assert!(left.abs_diff(right) <= 1, "AVL balance invariant failed");
        assert_eq!(node.height, 1 + left.max(right));
        1 + left.max(right)
    }

    fn collect_keys<V>(node: Option<&Arc<PersistentNode<V>>>, keys: &mut Vec<String>) {
        let Some(node) = node else {
            return;
        };
        collect_keys(node.left.as_ref(), keys);
        keys.push(node.key.as_str().to_owned());
        collect_keys(node.right.as_ref(), keys);
    }

    fn assert_valid<V>(map: &PersistentMap<V>) {
        check_tree(map.root.as_ref());
        let mut keys = Vec::new();
        collect_keys(map.root.as_ref(), &mut keys);
        assert!(keys.windows(2).all(|window| window[0] < window[1]));
    }

    fn exercise_insertions(names: &[&str]) {
        let mut map = PersistentMap::default();
        let mut expected = StdHashMap::new();
        for (index, value) in names.iter().enumerate() {
            let key = name(value);
            map = map.insert(key.clone(), index);
            expected.insert(key.clone(), index);
            assert_eq!(map.get(&key), expected.get(&key));
            assert_valid(&map);
        }
    }

    #[test]
    fn persistent_map_keeps_avl_balance_for_ordered_and_zig_zag_inserts() {
        exercise_insertions(&["a", "b", "c", "d", "e", "f", "g", "h"]);
        exercise_insertions(&["h", "g", "f", "e", "d", "c", "b", "a"]);
        exercise_insertions(&["m", "a", "z", "b", "y", "c", "x", "d", "w"]);
    }

    #[test]
    fn persistent_map_replacement_deletion_and_old_root_persistence() {
        let mut base = PersistentMap::default();
        for (index, value) in ["d", "b", "f", "a", "c", "e", "g", "h"]
            .into_iter()
            .enumerate()
        {
            base = base.insert(name(value), index);
        }
        assert_valid(&base);

        let replaced_key = name("d");
        let replaced = base.insert(replaced_key.clone(), 99);
        let mut expected = StdHashMap::new();
        for (index, value) in ["d", "b", "f", "a", "c", "e", "g", "h"]
            .into_iter()
            .enumerate()
        {
            expected.insert(name(value), index);
        }
        expected.insert(replaced_key.clone(), 99);
        assert_eq!(base.get(&replaced_key), Some(&0));
        assert_eq!(replaced.get(&replaced_key), expected.get(&replaced_key));
        assert_valid(&replaced);

        let all_keys = ["d", "b", "f", "a", "c", "e", "g", "h"];
        for value in ["a", "g", "b", "d"] {
            let key = name(value);
            let before = replaced.clone();
            let after = replaced.remove(&key);
            let mut after_expected = expected.clone();
            after_expected.remove(&key);
            assert!(before.get(&key).is_some());
            assert!(after.get(&key).is_none());
            for candidate in all_keys {
                let candidate = name(candidate);
                assert_eq!(after.get(&candidate), after_expected.get(&candidate));
            }
            assert_eq!(before.get(&replaced_key), expected.get(&replaced_key));
            assert_valid(&before);
            assert_valid(&after);
        }

        let absent = name("absent");
        let unchanged = replaced.remove(&absent);
        assert!(
            matches!((&replaced.root, &unchanged.root), (Some(left), Some(right)) if Arc::ptr_eq(left, right))
        );
        assert_valid(&unchanged);

        let deleted = replaced.remove(&replaced_key);
        assert!(deleted.get(&replaced_key).is_none());
        assert_eq!(replaced.get(&replaced_key), Some(&99));
        assert_valid(&deleted);
    }

    #[test]
    fn persistent_map_distinguishes_absent_from_present_option_value() {
        let present = name("present-none");
        let absent = name("absent");
        let map = PersistentMap::default().insert(present.clone(), None::<usize>);
        assert!(map.get(&present).is_some_and(|value| value.is_none()));
        assert!(map.get(&absent).is_none());
        assert_valid(&map);

        let removed = map.remove(&present);
        assert!(removed.get(&present).is_none());
        assert!(map.get(&present).is_some());
        assert_valid(&removed);
    }
}
