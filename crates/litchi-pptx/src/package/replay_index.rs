//! Retained relationship closures used by source-backed semantic replay.
//!
//! The index is built once while an OPC package enters the facade.  Published
//! package versions retain the previous index and add a small overlay for the
//! changed parts, so a replay validates only its affected closure instead of
//! walking every unrelated part again.

use std::collections::{HashMap, HashSet};
use std::sync::Arc;

use litchi_opc::constants::content_type as ct;
use litchi_opc::{OpcPackage, PackURI, Part};

use crate::{Error, Result};

#[derive(Debug, Clone, PartialEq, Eq)]
enum NotesEntry {
    Slide { bytes: u64 },
    Master { bytes: u64, theme: Option<PackURI> },
}

/// The source-backed closures needed by the notes and media replay guards.
///
/// `parent` is intentionally persistent.  A package edit therefore copies
/// only entries for changed parts; lookups walk the bounded publication chain
/// until they find the newest entry for a name.  The cached notes total is
/// updated as a delta and never recomputed from the package.
#[derive(Debug, Clone)]
pub(crate) struct ReplayIndex {
    parent: Option<Arc<Self>>,
    notes_total_bytes: u64,
    notes_entries: Arc<HashMap<PackURI, Option<NotesEntry>>>,
    notes_theme_users: Arc<HashMap<PackURI, Option<Arc<[PackURI]>>>>,
    media_slides: Arc<HashMap<PackURI, Arc<[PackURI]>>>,
}

impl ReplayIndex {
    /// Build the initial index during package ingress.  This is the only
    /// package-wide walk; later package versions use [`Self::update`].
    pub(crate) fn from_package(package: &OpcPackage) -> Result<Self> {
        // Keep ingress reservations proportional to the expected relevant
        // edges. The maps grow fallibly as the graph reveals more owners.
        let part_count = package.part_count().min(256);
        let mut notes_entries = HashMap::new();
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

        for part in package.try_iter_parts() {
            let part = part?;
            let name = part.partname().clone();
            let bytes = u64::try_from(part.blob().len()).map_err(|_| {
                Error::Invalid("PPTX replay notes index byte count overflow".into())
            })?;
            match part.content_type() {
                ct::PML_NOTES_SLIDE => {
                    notes_total_bytes = notes_total_bytes.checked_add(bytes).ok_or_else(|| {
                        Error::Invalid("PPTX replay notes index byte count overflow".into())
                    })?;
                    insert_notes_entry(
                        &mut notes_entries,
                        name,
                        Some(NotesEntry::Slide { bytes }),
                    )?;
                },
                ct::PML_NOTES_MASTER => {
                    let theme = notes_theme_target(part)?;
                    let theme_bytes = theme_bytes(package, theme.as_ref())?;
                    notes_total_bytes = notes_total_bytes
                        .checked_add(bytes)
                        .and_then(|total| total.checked_add(theme_bytes))
                        .ok_or_else(|| {
                            Error::Invalid("PPTX replay notes index byte count overflow".into())
                        })?;
                    insert_notes_entry(
                        &mut notes_entries,
                        name.clone(),
                        Some(NotesEntry::Master {
                            bytes,
                            theme: theme.clone(),
                        }),
                    )?;
                    if let Some(theme) = theme {
                        push_theme_user(&mut notes_theme_users, theme, name)?;
                    }
                },
                _ => {},
            }
            if part.content_type() == ct::PML_SLIDE {
                for target in media_targets(part)? {
                    push_media_slide(&mut media_slides, target, part.partname().clone())?;
                }
            }
        }

        Ok(Self {
            parent: None,
            notes_total_bytes,
            notes_entries: Arc::new(notes_entries),
            notes_theme_users: Arc::new(
                notes_theme_users
                    .into_iter()
                    .map(|(name, users)| (name, Some(Arc::from(users))))
                    .collect(),
            ),
            media_slides: Arc::new(
                media_slides
                    .into_iter()
                    .map(|(name, slides)| (name, Arc::from(slides)))
                    .collect(),
            ),
        })
    }

    /// The retained notes aggregate, updated without reading unrelated parts.
    pub(crate) const fn notes_total_bytes(&self) -> u64 {
        self.notes_total_bytes
    }

    /// Slides that currently reference a media resource.
    pub(crate) fn media_slides(&self, resource: &PackURI) -> Option<Arc<[PackURI]>> {
        let mut current = Some(self);
        while let Some(index) = current {
            if let Some(value) = index.media_slides.get(resource) {
                return Some(Arc::clone(value));
            }
            current = index.parent.as_deref();
        }
        None
    }

    fn notes_entry(&self, name: &PackURI) -> Option<NotesEntry> {
        let mut current = Some(self);
        while let Some(index) = current {
            if let Some(value) = index.notes_entries.get(name) {
                return value.clone();
            }
            current = index.parent.as_deref();
        }
        None
    }

    fn theme_users(&self, name: &PackURI) -> Arc<[PackURI]> {
        let mut current = Some(self);
        while let Some(index) = current {
            if let Some(value) = index.notes_theme_users.get(name) {
                return value.clone().unwrap_or_else(|| Arc::from([]));
            }
            current = index.parent.as_deref();
        }
        Arc::from([])
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
        let mut notes_entries = HashMap::new();
        notes_entries
            .try_reserve(changed.len())
            .map_err(|source| Error::Allocation {
                resource: "PPTX replay changed notes entries",
                source,
            })?;
        let mut notes_theme_users: HashMap<PackURI, Option<Vec<PackURI>>> = HashMap::new();
        let mut changed_masters = HashSet::new();
        let mut changed_themes = HashSet::new();
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
            let new_entry = if let Some(part) = new.as_deref() {
                match part.content_type() {
                    ct::PML_NOTES_SLIDE => Some(NotesEntry::Slide {
                        bytes: part_bytes(part)?,
                    }),
                    ct::PML_NOTES_MASTER => Some(NotesEntry::Master {
                        bytes: part_bytes(part)?,
                        theme: notes_theme_target(part)?,
                    }),
                    _ => None,
                }
            } else {
                None
            };
            if old_entry.is_some() || new_entry.is_some() {
                let old_value = notes_contribution(old_entry.as_ref(), source)?;
                let new_value = notes_contribution(new_entry.as_ref(), target)?;
                notes_total_bytes = adjust_total(notes_total_bytes, old_value, new_value)?;
                insert_notes_entry(&mut notes_entries, name.clone(), new_entry.clone())?;
            } else if old
                .as_ref()
                .is_some_and(|part| part.content_type() == ct::PML_NOTES_SLIDE)
                || new
                    .as_ref()
                    .is_some_and(|part| part.content_type() == ct::PML_NOTES_SLIDE)
            {
                insert_notes_entry(&mut notes_entries, name.clone(), new_entry.clone())?;
            }
            if changed_masters.contains(name) {
                let old_theme = match old_entry.as_ref() {
                    Some(NotesEntry::Master { theme, .. }) => theme.clone(),
                    _ => None,
                };
                let new_theme = match new_entry.as_ref() {
                    Some(NotesEntry::Master { theme, .. }) => theme.clone(),
                    _ => None,
                };
                if old_theme != new_theme {
                    if let Some(theme) = old_theme {
                        remove_theme_user(&mut notes_theme_users, self, theme, name)?;
                    }
                    if let Some(theme) = new_theme {
                        add_theme_user(&mut notes_theme_users, self, theme, name)?;
                    }
                }
            }
        }

        // A changed theme contributes to each unchanged notes master that
        // references it. Masters changed in this same patch were already
        // charged their complete before/after contribution above.
        for theme in &changed_themes {
            let old_bytes = theme_bytes(source, Some(theme))?;
            let new_bytes = theme_bytes(target, Some(theme))?;
            if old_bytes == new_bytes {
                continue;
            }
            let users = notes_theme_users
                .get(theme)
                .and_then(|users| users.as_ref())
                .cloned()
                .unwrap_or_else(|| self.theme_users(theme).to_vec());
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
                .as_deref()
                .filter(|part| part.content_type() == ct::PML_SLIDE)
                .map(media_targets)
                .transpose()?
                .unwrap_or_default();
            let new_targets = new
                .as_deref()
                .filter(|part| part.content_type() == ct::PML_SLIDE)
                .map(media_targets)
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

        Ok(Self {
            parent: Some(Arc::new(self.clone())),
            notes_total_bytes,
            notes_entries: Arc::new(notes_entries),
            notes_theme_users: Arc::new(
                notes_theme_users
                    .into_iter()
                    .map(|(name, users)| (name, users.map(Arc::from)))
                    .collect(),
            ),
            media_slides: Arc::new(
                media_updates
                    .into_iter()
                    .map(|(name, slides)| (name, Arc::from(slides)))
                    .collect(),
            ),
        })
    }
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

fn notes_contribution(entry: Option<&NotesEntry>, package: &OpcPackage) -> Result<u64> {
    let Some(entry) = entry else {
        return Ok(0);
    };
    match entry {
        NotesEntry::Slide { bytes } => Ok(*bytes),
        NotesEntry::Master { bytes, theme } => {
            let theme_bytes = theme_bytes(package, theme.as_ref())?;
            bytes
                .checked_add(theme_bytes)
                .ok_or_else(|| Error::Invalid("PPTX replay notes total overflow".into()))
        },
    }
}

fn theme_bytes(package: &OpcPackage, name: Option<&PackURI>) -> Result<u64> {
    let Some(name) = name else {
        return Ok(0);
    };
    let Some(part) = optional_part(package, name)? else {
        return Ok(0);
    };
    if part.content_type() != ct::OFC_THEME {
        return Ok(0);
    }
    part_bytes(part)
}

fn notes_theme_target(part: &dyn Part) -> Result<Option<PackURI>> {
    part.rels()
        .iter()
        .filter(|relationship| relationship.reltype().ends_with("/theme"))
        .filter(|relationship| !relationship.is_external())
        .map(|relationship| relationship.target_partname().map_err(Error::Opc))
        .next()
        .transpose()
}

fn media_targets(part: &dyn Part) -> Result<Vec<PackURI>> {
    let mut result = Vec::new();
    for relationship in part.rels().iter() {
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
    let values = users.entry(resource).or_insert_with(Vec::new);
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
    entries: &mut HashMap<PackURI, Option<NotesEntry>>,
    name: PackURI,
    value: Option<NotesEntry>,
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
    let values = users.entry(theme).or_insert_with(Vec::new);
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
    updates: &mut HashMap<PackURI, Option<Vec<PackURI>>>,
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
        .or_insert_with(|| Some(index.theme_users(&theme).to_vec()))
        .as_mut()
        .ok_or_else(|| Error::Invalid("PPTX replay notes theme users missing".into()))?;
    if !users.iter().any(|user| user == master) {
        users.try_reserve(1).map_err(|source| Error::Allocation {
            resource: "PPTX replay notes theme users",
            source,
        })?;
        users.push(master.clone());
    }
    Ok(())
}

fn remove_theme_user(
    updates: &mut HashMap<PackURI, Option<Vec<PackURI>>>,
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
        .or_insert_with(|| Some(index.theme_users(&theme).to_vec()))
        .as_mut()
        .ok_or_else(|| Error::Invalid("PPTX replay notes theme users missing".into()))?;
    users.retain(|user| user != master);
    Ok(())
}
