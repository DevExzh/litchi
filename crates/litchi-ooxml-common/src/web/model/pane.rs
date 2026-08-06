use super::super::codec::*;
use super::super::validation::*;
use super::super::*;
use super::*;
#[derive(Debug, Clone, PartialEq)]
pub struct Pane {
    pub(in crate::web) dock_state: Dock,
    pub(in crate::web) visible: bool,
    pub(in crate::web) width: f64,
    pub(in crate::web) row: u32,
    pub(in crate::web) locked: bool,
    pub(in crate::web) relationship_id: String,
    pub(in crate::web) add_in: AddIn,
    pub(in crate::web) snapshot_resources: Vec<SnapshotResource>,
    pub(in crate::web) extension_list: Option<ExtList>,
}

impl Pane {
    /// Create a right-docked, visible pane with a schema-valid default width.
    #[must_use]
    pub fn new(add_in: AddIn) -> Self {
        Self {
            dock_state: Dock::Right,
            visible: true,
            width: 320.0,
            row: 0,
            locked: false,
            relationship_id: String::new(),
            add_in,
            snapshot_resources: Vec::new(),
            extension_list: None,
        }
    }

    #[must_use]
    pub fn show(mut self, visible: bool) -> Self {
        self.set_visible(visible);
        self
    }

    pub fn set_visible(&mut self, visible: bool) -> &mut Self {
        self.visible = visible;
        self
    }

    pub fn width(mut self, width: f64) -> Result<Self> {
        self.set_width(width)?;
        Ok(self)
    }

    pub fn set_width(&mut self, width: f64) -> Result<&mut Self> {
        if !width.is_finite() || width <= 0.0 {
            return invalid("task-pane width must be finite and positive".into());
        }
        self.width = width;
        Ok(self)
    }

    pub fn dock(mut self, state: impl AsRef<str>) -> Result<Self> {
        self.set_dock(state)?;
        Ok(self)
    }

    pub fn set_dock(&mut self, state: impl AsRef<str>) -> Result<&mut Self> {
        self.dock_state = Dock::parse(state.as_ref())?;
        Ok(self)
    }

    pub fn set_row(&mut self, row: u32) -> &mut Self {
        self.row = row;
        self
    }

    pub fn set_locked(&mut self, locked: bool) -> &mut Self {
        self.locked = locked;
        self
    }

    /// Attach an embedded image using shared storage and a semantic part name.
    pub fn embed(
        mut self,
        part_name: impl AsRef<str>,
        content_type: impl Into<String>,
        data: impl Into<Arc<Vec<u8>>>,
    ) -> Result<Self> {
        self.set_image(part_name, content_type, data)?;
        Ok(self)
    }

    /// Attach or replace the embedded image in place.
    pub fn set_image(
        &mut self,
        part_name: impl AsRef<str>,
        content_type: impl Into<String>,
        data: impl Into<Arc<Vec<u8>>>,
    ) -> Result<&mut Self> {
        let part_name = PackURI::new(part_name.as_ref().to_owned()).map_err(Error::Uri)?;
        if part_name.as_str() == "/" {
            return invalid("snapshot image cannot target the package root".into());
        }
        let content_type = content_type.into();
        validate_image_content_type(&content_type)?;
        let data = data.into();
        if data.len() > MAX_WEB_EXTENSION_SNAPSHOT_BYTES {
            return limit(
                "web extension snapshot bytes",
                MAX_WEB_EXTENSION_SNAPSHOT_BYTES,
                data.len(),
            );
        }
        let linked_id = self
            .add_in
            .snapshot
            .as_ref()
            .and_then(|snapshot| snapshot.linked_relationship_id.as_deref());
        let existing_id = self
            .add_in
            .snapshot
            .as_ref()
            .and_then(|snapshot| snapshot.embedded_relationship_id.as_deref())
            .filter(|id| Some(*id) != linked_id)
            .map(str::to_owned);
        let relationship_id = match existing_id {
            Some(id) => id,
            None => self.next_snapshot_relationship_id("rIdSnapshot")?,
        };
        self.snapshot_resources
            .retain(|resource| resource.relationship_id != relationship_id);
        self.snapshot_resources.push(SnapshotResource {
            relationship_id: relationship_id.clone(),
            target: SnapshotTarget::Internal {
                part_name,
                content_type,
                data,
            },
        });
        self.add_in
            .snapshot
            .get_or_insert_with(Snapshot::default)
            .embedded_relationship_id = Some(relationship_id);
        Ok(self)
    }

    /// Retain an external image link without resolving or contacting it.
    pub fn linked(mut self, target: impl Into<String>) -> Result<Self> {
        self.set_external_link(target)?;
        Ok(self)
    }

    /// Attach or replace an inert external image link in place.
    pub fn set_external_link(&mut self, target: impl Into<String>) -> Result<&mut Self> {
        let target = target.into();
        validate_external_uri_reference(&target)?;
        let embedded_id = self
            .add_in
            .snapshot
            .as_ref()
            .and_then(|snapshot| snapshot.embedded_relationship_id.as_deref());
        let existing_id = self
            .add_in
            .snapshot
            .as_ref()
            .and_then(|snapshot| snapshot.linked_relationship_id.as_deref())
            .filter(|id| Some(*id) != embedded_id)
            .map(str::to_owned);
        let relationship_id = match existing_id {
            Some(id) => id,
            None => self.next_snapshot_relationship_id("rIdSnapshotLink")?,
        };
        self.snapshot_resources
            .retain(|resource| resource.relationship_id != relationship_id);
        self.snapshot_resources.push(SnapshotResource {
            relationship_id: relationship_id.clone(),
            target: SnapshotTarget::External { target },
        });
        self.add_in
            .snapshot
            .get_or_insert_with(Snapshot::default)
            .linked_relationship_id = Some(relationship_id);
        Ok(self)
    }

    /// Attach an internal linked image without exposing its relationship ID.
    pub fn linked_image(
        mut self,
        part_name: impl AsRef<str>,
        content_type: impl Into<String>,
        data: impl Into<Arc<Vec<u8>>>,
    ) -> Result<Self> {
        self.set_linked_image(part_name, content_type, data)?;
        Ok(self)
    }

    /// Attach or replace an internal linked image in place.
    pub fn set_linked_image(
        &mut self,
        part_name: impl AsRef<str>,
        content_type: impl Into<String>,
        data: impl Into<Arc<Vec<u8>>>,
    ) -> Result<&mut Self> {
        let part_name = PackURI::new(part_name.as_ref().to_owned()).map_err(Error::Uri)?;
        if part_name.as_str() == "/" {
            return invalid("linked snapshot image cannot target the package root".into());
        }
        let content_type = content_type.into();
        validate_image_content_type(&content_type)?;
        let data = data.into();
        if data.len() > MAX_WEB_EXTENSION_SNAPSHOT_BYTES {
            return limit(
                "web extension snapshot bytes",
                MAX_WEB_EXTENSION_SNAPSHOT_BYTES,
                data.len(),
            );
        }
        let embedded_id = self
            .add_in
            .snapshot
            .as_ref()
            .and_then(|snapshot| snapshot.embedded_relationship_id.as_deref());
        let existing_id = self
            .add_in
            .snapshot
            .as_ref()
            .and_then(|snapshot| snapshot.linked_relationship_id.as_deref())
            .filter(|id| Some(*id) != embedded_id)
            .map(str::to_owned);
        let relationship_id = match existing_id {
            Some(id) => id,
            None => self.next_snapshot_relationship_id("rIdSnapshotLink")?,
        };
        self.snapshot_resources
            .retain(|resource| resource.relationship_id != relationship_id);
        self.snapshot_resources.push(SnapshotResource {
            relationship_id: relationship_id.clone(),
            target: SnapshotTarget::Internal {
                part_name,
                content_type,
                data,
            },
        });
        self.add_in
            .snapshot
            .get_or_insert_with(Snapshot::default)
            .linked_relationship_id = Some(relationship_id);
        Ok(self)
    }

    /// Set snapshot compression metadata, creating an empty snapshot if needed.
    #[must_use]
    pub fn compress(mut self, compression: Compression) -> Self {
        self.set_compression(Some(compression));
        self
    }

    pub fn set_compression(&mut self, compression: Option<Compression>) -> &mut Self {
        if let Some(compression) = compression {
            self.snapshot_mut().set_compression(Some(compression));
        } else if let Some(snapshot) = self.add_in.snapshot.as_mut() {
            snapshot.set_compression(None);
        }
        self
    }

    /// Append one validated DrawingML effect.
    pub fn effect(mut self, effect: Effect) -> Result<Self> {
        self.push_effect(effect)?;
        Ok(self)
    }

    pub fn push_effect(&mut self, effect: Effect) -> Result<&mut Self> {
        self.snapshot_mut().push_effect(effect)?;
        Ok(self)
    }

    pub fn replace_effect(&mut self, index: usize, effect: Effect) -> Result<Option<Effect>> {
        let Some(snapshot) = self.add_in.snapshot.as_mut() else {
            return Ok(None);
        };
        snapshot.replace_effect(index, effect)
    }

    pub fn remove_effect(&mut self, index: usize) -> Option<Effect> {
        self.add_in
            .snapshot
            .as_mut()
            .and_then(|snapshot| snapshot.remove_effect(index))
    }

    pub fn snapshot_mut(&mut self) -> &mut Snapshot {
        self.add_in.snapshot.get_or_insert_with(Snapshot::default)
    }

    /// Inspect the embedded image without exposing its relationship ID.
    #[must_use]
    pub fn image(&self) -> Option<Image<'_>> {
        let id = self
            .add_in
            .snapshot
            .as_ref()?
            .embedded_relationship_id
            .as_deref()?;
        self.snapshot_resources.iter().find_map(|resource| {
            if resource.relationship_id != id {
                return None;
            }
            match &resource.target {
                SnapshotTarget::Internal {
                    part_name,
                    content_type,
                    data,
                } => Some(Image {
                    part_name,
                    content_type,
                    data,
                }),
                SnapshotTarget::External { .. } => None,
            }
        })
    }

    /// Inspect an internal or inert external link without exposing its relationship ID.
    #[must_use]
    pub fn link(&self) -> Option<Link<'_>> {
        let id = self
            .add_in
            .snapshot
            .as_ref()?
            .linked_relationship_id
            .as_deref()?;
        self.snapshot_resources.iter().find_map(|resource| {
            if resource.relationship_id != id {
                return None;
            }
            match &resource.target {
                SnapshotTarget::External { target } => Some(Link::External(target)),
                SnapshotTarget::Internal {
                    part_name,
                    content_type,
                    data,
                } => Some(Link::Internal(Image {
                    part_name,
                    content_type,
                    data,
                })),
            }
        })
    }

    /// Remove the embedded image, returning whether one existed.
    pub fn clear_image(&mut self) -> bool {
        let Some(id) = self
            .add_in
            .snapshot
            .as_mut()
            .and_then(|snapshot| snapshot.embedded_relationship_id.take())
        else {
            return false;
        };
        let old_len = self.snapshot_resources.len();
        self.snapshot_resources
            .retain(|resource| resource.relationship_id != id);
        old_len != self.snapshot_resources.len()
    }

    /// Remove an internal or external linked image, returning whether one existed.
    pub fn clear_link(&mut self) -> bool {
        let Some(id) = self
            .add_in
            .snapshot
            .as_mut()
            .and_then(|snapshot| snapshot.linked_relationship_id.take())
        else {
            return false;
        };
        let old_len = self.snapshot_resources.len();
        self.snapshot_resources
            .retain(|resource| resource.relationship_id != id);
        old_len != self.snapshot_resources.len()
    }

    pub fn clear_compression(&mut self) -> bool {
        self.add_in
            .snapshot
            .as_mut()
            .and_then(|snapshot| snapshot.compression_state.take())
            .is_some()
    }

    pub fn clear_effects(&mut self) -> bool {
        let Some(snapshot) = self.add_in.snapshot.as_mut() else {
            return false;
        };
        snapshot.clear_effects()
    }

    /// Remove all snapshot XML metadata and embedded/external resources.
    pub fn clear_snapshot(&mut self) -> bool {
        let had_snapshot = self.add_in.snapshot.take().is_some();
        let had_resources = !self.snapshot_resources.is_empty();
        self.snapshot_resources.clear();
        had_snapshot || had_resources
    }

    #[must_use]
    pub const fn visible(&self) -> bool {
        self.visible
    }

    #[must_use]
    pub const fn add_in(&self) -> &AddIn {
        &self.add_in
    }

    pub const fn add_in_mut(&mut self) -> &mut AddIn {
        &mut self.add_in
    }

    #[must_use]
    pub const fn dock_kind(&self) -> &Dock {
        &self.dock_state
    }

    #[must_use]
    pub fn dock_state(&self) -> &str {
        self.dock_state.as_str()
    }

    #[must_use]
    pub const fn pane_width(&self) -> f64 {
        self.width
    }

    #[must_use]
    pub const fn row(&self) -> u32 {
        self.row
    }

    #[must_use]
    pub const fn locked(&self) -> bool {
        self.locked
    }

    #[must_use]
    pub const fn ext(&self) -> Option<&ExtList> {
        self.extension_list.as_ref()
    }

    pub fn set_ext(&mut self, extension: ExtList) -> Result<&mut Self> {
        validate_extension_list(Some(&extension), &[ExtKind::TaskPane])?;
        self.extension_list = Some(extension);
        Ok(self)
    }

    pub fn clear_ext(&mut self) -> Option<ExtList> {
        self.extension_list.take()
    }

    pub(in crate::web) fn next_snapshot_relationship_id(&self, base: &str) -> Result<String> {
        let attempts = self
            .snapshot_resources
            .len()
            .checked_add(2)
            .ok_or(Error::Limit {
                resource: "snapshot relationship IDs",
                max: usize::MAX,
                actual: usize::MAX,
            })?;
        for index in 0..attempts {
            let candidate = if index == 0 {
                base.to_owned()
            } else {
                format!("{base}{index}")
            };
            let used_by_snapshot = self.add_in.snapshot.as_ref().is_some_and(|snapshot| {
                snapshot.embedded_relationship_id.as_deref() == Some(candidate.as_str())
                    || snapshot.linked_relationship_id.as_deref() == Some(candidate.as_str())
            });
            if !used_by_snapshot
                && self
                    .snapshot_resources
                    .iter()
                    .all(|resource| resource.relationship_id != candidate)
            {
                return Ok(candidate);
            }
        }
        Err(Error::Relationship(
            "unable to allocate a snapshot relationship ID".into(),
        ))
    }
}
