//! Source-checked, source-preserving color-map edits.

use std::sync::Arc;

use super::codec;
use super::model::{Map, Override, Role, Slot, Value};
use crate::{Error, Result};

/// A stable fingerprint of the exact color-map owner bytes.
pub type Revision = u64;

/// An immutable typed color-map view bound to one exact XML source.
#[derive(Debug, Clone)]
pub struct Snapshot {
    source_xml: Arc<Vec<u8>>,
    located: codec::Located,
    revision: Revision,
}

impl Snapshot {
    /// Parse a slide-master color map and retain its exact source bytes.
    ///
    /// # Errors
    ///
    /// Returns an error if the input cannot be read or is malformed.
    pub fn from_master(source: impl AsRef<[u8]>) -> Result<Self> {
        Self::from_master_owned(source.as_ref().to_vec())
    }

    /// Parse an owned slide-master color map without another source copy.
    ///
    /// # Errors
    ///
    /// Returns an error if the input cannot be read or is malformed.
    pub fn from_master_owned(source: Vec<u8>) -> Result<Self> {
        let source_xml = Arc::new(source);
        let located = codec::locate_master(source_xml.as_slice())?;
        Ok(Self::from_located(source_xml, located))
    }

    /// Parse a slide or slide-layout color-map override and retain its exact
    /// source bytes.
    ///
    /// # Errors
    ///
    /// Returns an error if the input cannot be read or is malformed.
    pub fn from_override(
        source: impl AsRef<[u8]>,
        root_name: impl AsRef<[u8]>,
        root_label: impl Into<String>,
    ) -> Result<Self> {
        Self::from_override_owned(source.as_ref().to_vec(), root_name, root_label)
    }

    /// Parse an owned color-map override without another source copy.
    ///
    /// # Errors
    ///
    /// Returns an error if the input cannot be read or is malformed.
    pub fn from_override_owned(
        source: Vec<u8>,
        root_name: impl AsRef<[u8]>,
        root_label: impl Into<String>,
    ) -> Result<Self> {
        let source_xml = Arc::new(source);
        let root_label = root_label.into();
        let located =
            codec::locate_override(source_xml.as_slice(), root_name.as_ref(), &root_label)?;
        Ok(Self::from_located(source_xml, located))
    }

    /// Return the typed value represented by this source.
    #[inline]
    #[must_use]
    pub fn value(&self) -> Value {
        self.located.value
    }

    /// Return the explicit map when this source owns one.
    #[inline]
    #[must_use]
    pub fn map(&self) -> Option<Map> {
        mapped_value(&self.located.value)
    }

    /// Return whether this snapshot is a slide-master source.
    #[inline]
    #[must_use]
    pub fn is_master(&self) -> bool {
        matches!(&self.located.source, codec::Source::Master)
    }

    /// Return whether this snapshot is a slide or slide-layout override.
    #[inline]
    #[must_use]
    pub fn is_override(&self) -> bool {
        matches!(&self.located.source, codec::Source::Override { .. })
    }

    /// Borrow the exact source XML captured by this snapshot.
    #[inline]
    #[must_use]
    pub fn source_xml(&self) -> &[u8] {
        self.source_xml.as_slice()
    }

    /// Return the source fingerprint used for stale-source checks.
    #[inline]
    #[must_use]
    pub const fn revision(&self) -> Revision {
        self.revision
    }

    /// Start an atomic detached edit over this typed color-map value.
    #[inline]
    #[must_use]
    pub fn edit(&self) -> Transaction {
        Transaction {
            original: self.clone(),
            working: self.located.value,
        }
    }

    fn from_located(source_xml: Arc<Vec<u8>>, located: codec::Located) -> Self {
        let revision = fingerprint(source_xml.as_slice());
        Self {
            source_xml,
            located,
            revision,
        }
    }

    fn same_source(&self, other: &Self) -> bool {
        self.source_xml.as_slice() == other.source_xml.as_slice()
            && self.revision == other.revision
            && self.located.source == other.located.source
    }
}

/// A bounded edit staged against one color-map source.
#[derive(Debug, Clone)]
pub struct Transaction {
    original: Snapshot,
    working: Value,
}

impl Transaction {
    /// Borrow the projected typed value.
    #[inline]
    #[must_use]
    pub fn snapshot(&self) -> &Value {
        &self.working
    }

    /// Return the projected typed value by copy.
    #[inline]
    #[must_use]
    pub fn value(&self) -> Value {
        self.working
    }

    /// Return the projected explicit map, when one is available.
    #[inline]
    #[must_use]
    pub fn map(&self) -> Option<Map> {
        mapped_value(&self.working)
    }

    /// Return whether the staged value differs from the source value.
    #[inline]
    #[must_use]
    pub fn is_changed(&self) -> bool {
        self.original.value() != self.working
    }

    /// Set one mapped color role without touching any other source bytes.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation fails.
    pub fn set_color(&mut self, slot: Slot, role: Role) -> Result<bool> {
        let map = self.map_mut()?;
        if map.color(slot) == role {
            return Ok(false);
        }
        map.set_color(slot, role);
        Ok(true)
    }

    /// Replace the explicit map while retaining the source's mapping kind.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation fails.
    pub fn replace_map(&mut self, map: Map) -> Result<bool> {
        let current = self.map_mut()?;
        if *current == map {
            return Ok(false);
        }
        *current = map;
        Ok(true)
    }

    /// Replace the typed value without changing its source shape.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation fails.
    pub fn replace(&mut self, value: Value) -> Result<bool> {
        if !same_shape(&self.working, &value) {
            return Err(invalid(
                "a bounded color-map edit cannot create, remove, or switch its mapping kind",
            ));
        }
        if self.working == value {
            return Ok(false);
        }
        self.working = value;
        Ok(true)
    }

    /// Validate and consume this detached edit into a source-checked commit.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation fails.
    pub fn commit(self) -> Result<Commit> {
        if !self.is_changed() {
            let patch = Patch::new(self.original.clone(), self.original.clone());
            return Ok(Commit {
                snapshot: self.original,
                patch,
                changed: false,
            });
        }

        let updated = codec::rewrite(
            self.original.source_xml.as_slice(),
            &self.original.located,
            self.working,
        )?;
        let updated = Arc::new(updated);
        let located = codec::locate_source(updated.as_slice(), &self.original.located.source)?;
        if located.value != self.working {
            return Err(invalid("color-map serialization changed the typed value"));
        }
        let snapshot = Snapshot::from_located(Arc::clone(&updated), located);
        let patch = Patch::new(self.original, snapshot.clone());
        Ok(Commit {
            snapshot,
            patch,
            changed: true,
        })
    }

    fn map_mut(&mut self) -> Result<&mut Map> {
        match &mut self.working {
            Value::Master(map) | Value::Override(Some(Override::Override(map))) => Ok(map),
            Value::Override(None | Some(Override::Master)) => {
                Err(invalid("the color-map source has no explicit map to edit"))
            },
        }
    }
}

/// A committed color-map snapshot and its reversible source patch.
#[derive(Debug, Clone)]
pub struct Commit {
    snapshot: Snapshot,
    patch: Patch,
    changed: bool,
}

impl Commit {
    /// Return whether publication changes the exact source bytes.
    #[inline]
    #[must_use]
    pub const fn changed(&self) -> bool {
        self.changed
    }

    /// Alias for [`Self::changed`].
    #[inline]
    #[must_use]
    pub const fn is_changed(&self) -> bool {
        self.changed
    }

    /// Borrow the projected post-edit snapshot.
    #[inline]
    #[must_use]
    pub fn snapshot(&self) -> &Snapshot {
        &self.snapshot
    }

    /// Borrow the reversible source-checked patch.
    #[inline]
    #[must_use]
    pub fn patch(&self) -> &Patch {
        &self.patch
    }

    /// Consume the commit into its patch.
    #[inline]
    #[must_use]
    pub fn into_patch(self) -> Patch {
        self.patch
    }
}

/// A reversible, source-checked replacement of one color-map XML source.
#[derive(Debug, Clone)]
pub struct Patch {
    before: Snapshot,
    after: Snapshot,
}

impl Patch {
    fn new(before: Snapshot, after: Snapshot) -> Self {
        Self { before, after }
    }

    /// Borrow the exact source context required before publication.
    #[inline]
    #[must_use]
    pub fn before(&self) -> &Snapshot {
        &self.before
    }

    /// Borrow the exact source context produced by publication.
    #[inline]
    #[must_use]
    pub fn after(&self) -> &Snapshot {
        &self.after
    }

    /// Return whether this patch is an exact no-op.
    #[inline]
    #[must_use]
    pub fn is_empty(&self) -> bool {
        self.before.same_source(&self.after)
    }

    /// Return whether this patch changes the source bytes.
    #[inline]
    #[must_use]
    pub fn is_changed(&self) -> bool {
        !self.is_empty()
    }

    /// Return the exact inverse patch.
    #[must_use]
    pub fn inverse(&self) -> Self {
        Self {
            before: self.after.clone(),
            after: self.before.clone(),
        }
    }

    /// Return the source fingerprint required for publication.
    #[inline]
    #[must_use]
    pub const fn expected_revision(&self) -> Revision {
        self.before.revision
    }

    /// Apply this patch atomically to an exact XML source buffer.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation fails.
    pub fn apply(&self, target: &mut Vec<u8>) -> Result<Snapshot> {
        if target.as_slice() != self.before.source_xml.as_slice() {
            return Err(invalid("color-map source is stale"));
        }
        let current = codec::locate_source(target.as_slice(), &self.before.located.source)
            .map_err(|_err| invalid("color-map source is stale"))?;
        if current.value != self.before.value() {
            return Err(invalid("color-map source is stale"));
        }
        if self.is_empty() {
            return Ok(self.before.clone());
        }

        let located =
            codec::locate_source(self.after.source_xml.as_slice(), &self.after.located.source)?;
        if located.value != self.after.value() {
            return Err(invalid("published color-map source differs from the patch"));
        }
        let result = Snapshot::from_located(Arc::clone(&self.after.source_xml), located);
        if !result.same_source(&self.after) {
            return Err(invalid("published color-map source differs from the patch"));
        }
        let replacement = self.after.source_xml.as_ref().clone();
        *target = replacement;
        Ok(result)
    }

    /// Alias for [`Self::apply`] emphasizing the detached source target.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation fails.
    #[inline]
    pub fn apply_to(&self, target: &mut Vec<u8>) -> Result<Snapshot> {
        self.apply(target)
    }
}

fn mapped_value(value: &Value) -> Option<Map> {
    match value {
        Value::Master(map) => Some(*map),
        Value::Override(Some(Override::Override(map))) => Some(*map),
        Value::Override(None | Some(Override::Master)) => None,
    }
}

fn same_shape(left: &Value, right: &Value) -> bool {
    matches!(
        (left, right),
        (Value::Master(_), Value::Master(_))
            | (Value::Override(None), Value::Override(None))
            | (
                Value::Override(Some(Override::Master)),
                Value::Override(Some(Override::Master))
            )
            | (
                Value::Override(Some(Override::Override(_))),
                Value::Override(Some(Override::Override(_)))
            )
    )
}

fn fingerprint(bytes: &[u8]) -> Revision {
    let mut hash = 0xcbf2_9ce4_8422_2325u64;
    for byte in bytes {
        hash ^= u64::from(*byte);
        hash = hash.wrapping_mul(0x100_0000_01b3);
    }
    hash
}

fn invalid(message: impl Into<String>) -> Error {
    Error::Invalid(message.into())
}

#[cfg(test)]
#[allow(
    clippy::unwrap_used,
    clippy::expect_used,
    reason = "test assertions panic on failure by design"
)]
mod tests {
    use super::*;
    use crate::presentation_properties::metadata::color_map::{Role, Slot, Value};

    const PML: &str = "http://schemas.openxmlformats.org/presentationml/2006/main";
    const DML: &str = "http://schemas.openxmlformats.org/drawingml/2006/main";

    fn master_source() -> Vec<u8> {
        format!(
            r#"<?xml version="1.0"?>
<p:sldMaster xmlns:p="{PML}" xmlns:v="urn:vendor" v:root="keep">
  <p:clrMap v:unknown="keep" bg1='lt1' tx1='dk1' bg2='lt2' tx2='dk2'
      accent1='accent1' accent2='accent2' accent3='accent3' accent4='accent4'
      accent5='accent5' accent6='accent6' hlink='hlink' folHlink='folHlink'><v:extra/></p:clrMap>
  <v:tail>untouched</v:tail>
</p:sldMaster>"#
        )
        .into_bytes()
    }

    fn override_source() -> Vec<u8> {
        format!(
            r#"<p:sldLayout xmlns:p="{PML}" xmlns:a="{DML}" xmlns:v="urn:vendor">
  <p:clrMapOvr v:unknown="keep"><a:overrideClrMapping v:map="keep"
      bg1="lt1" tx1="dk1" bg2="lt2" tx2="dk2" accent1="accent1" accent2="accent2"
      accent3="accent3" accent4="accent4" accent5="accent5" accent6="accent6"
      hlink="hlink" folHlink="folHlink"/></p:clrMapOvr>
  <v:tail>untouched</v:tail>
</p:sldLayout>"#
        )
        .into_bytes()
    }

    #[test]
    fn no_op_commit_and_patch_preserve_exact_source() {
        let source = master_source();
        let snapshot = Snapshot::from_master(&source).unwrap();
        let commit = snapshot.edit().commit().unwrap();

        assert!(!commit.changed());
        assert!(commit.patch().is_empty());
        assert_eq!(commit.snapshot().source_xml(), source.as_slice());

        let mut target = source.clone();
        let applied = commit.patch().apply(&mut target).unwrap();
        assert_eq!(target, source);
        assert_eq!(applied.source_xml(), source.as_slice());
    }

    #[test]
    fn typed_edits_preserve_unknown_xml_and_apply_to_override_sources() {
        let source = master_source();
        let snapshot = Snapshot::from_master(&source).unwrap();
        let mut edit = snapshot.edit();
        assert!(edit.set_color(Slot::Accent1, Role::Accent2).unwrap());
        let commit = edit.commit().unwrap();
        let output = commit.snapshot().source_xml();
        assert!(
            output
                .windows(b"accent1='accent2'".len())
                .any(|window| { window == b"accent1='accent2'" })
        );
        assert!(
            output
                .windows(b"v:extra".len())
                .any(|window| window == b"v:extra")
        );
        assert!(
            output
                .windows(b"<v:tail>untouched</v:tail>".len())
                .any(|window| { window == b"<v:tail>untouched</v:tail>" })
        );

        let override_bytes = override_source();
        let override_snapshot =
            Snapshot::from_override(&override_bytes, b"sldLayout", "slide layout").unwrap();
        let mut override_edit = override_snapshot.edit();
        override_edit
            .set_color(Slot::FollowedHyperlink, Role::Accent6)
            .unwrap();
        let override_commit = override_edit.commit().unwrap();
        assert!(
            override_commit
                .snapshot()
                .source_xml()
                .windows(b"folHlink=\"accent6\"".len())
                .any(|window| window == b"folHlink=\"accent6\"")
        );
        assert!(
            override_commit
                .snapshot()
                .source_xml()
                .windows(b"v:map=\"keep\"".len())
                .any(|window| window == b"v:map=\"keep\"")
        );
    }

    #[test]
    fn self_closing_master_maps_are_editable() {
        let source = String::from_utf8(master_source())
            .unwrap()
            .replace("><v:extra/></p:clrMap>", "/>")
            .into_bytes();
        let snapshot = Snapshot::from_master(&source).unwrap();
        let mut edit = snapshot.edit();
        edit.set_color(Slot::Accent2, Role::Accent3).unwrap();
        let commit = edit.commit().unwrap();
        assert!(
            commit
                .snapshot()
                .source_xml()
                .windows(b"accent2='accent3'".len())
                .any(|window| window == b"accent2='accent3'")
        );
    }

    #[test]
    fn stale_replay_is_atomic_and_inverse_restores_exact_bytes() {
        let source = master_source();
        let snapshot = Snapshot::from_master(&source).unwrap();
        let mut edit = snapshot.edit();
        edit.set_color(Slot::Text1, Role::Light1).unwrap();
        let patch = edit.commit().unwrap().into_patch();

        let mut stale = source.clone();
        stale.extend_from_slice(b" ");
        let stale_before = stale.clone();
        assert!(patch.apply(&mut stale).is_err());
        assert_eq!(stale, stale_before);

        let mut target = source.clone();
        patch.apply(&mut target).unwrap();
        assert_ne!(target, source);
        patch.inverse().apply(&mut target).unwrap();
        assert_eq!(target, source);
    }

    #[test]
    fn failed_edits_leave_the_staged_value_unchanged() {
        let source = format!(
            r#"<p:sld xmlns:p="{PML}" xmlns:a="{DML}"><p:clrMapOvr>
              <a:masterClrMapping/>
            </p:clrMapOvr></p:sld>"#
        );
        let snapshot = Snapshot::from_override(source.as_bytes(), b"sld", "slide").unwrap();
        assert_eq!(snapshot.value(), Value::Override(Some(Override::Master)));

        let mut edit = snapshot.edit();
        let before = *edit.snapshot();
        assert!(edit.set_color(Slot::Accent1, Role::Accent2).is_err());
        assert_eq!(*edit.snapshot(), before);
        assert!(edit.replace(Value::Override(None)).is_err());
        assert_eq!(*edit.snapshot(), before);
    }
}
