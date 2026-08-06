use super::super::codec::*;
use super::super::validation::*;
use super::super::*;
use super::*;
/// DrawingML `CT_Blip` metadata used by a web-extension snapshot.
#[derive(Debug, Clone, Default, PartialEq, Eq)]
pub struct Snapshot {
    pub(in crate::web) embedded_relationship_id: Option<String>,
    pub(in crate::web) linked_relationship_id: Option<String>,
    pub(in crate::web) compression_state: Option<Compression>,
    pub(in crate::web) effects: Vec<Effect>,
    pub(in crate::web) extension_list: Option<ExtList>,
}

impl Snapshot {
    #[must_use]
    pub const fn new() -> Self {
        Self {
            embedded_relationship_id: None,
            linked_relationship_id: None,
            compression_state: None,
            effects: Vec::new(),
            extension_list: None,
        }
    }

    #[must_use]
    pub const fn compression(&self) -> Option<Compression> {
        self.compression_state
    }

    #[must_use]
    pub fn effects(&self) -> &[Effect] {
        &self.effects
    }

    pub fn set_compression(&mut self, compression: Option<Compression>) -> &mut Self {
        self.compression_state = compression;
        self
    }

    pub fn push_effect(&mut self, effect: Effect) -> Result<&mut Self> {
        if self.effects.len() >= MAX_WEB_EXTENSION_ITEMS {
            return limit(
                "snapshot effects",
                MAX_WEB_EXTENSION_ITEMS,
                self.effects.len().saturating_add(1),
            );
        }
        let reparsed = Effect::from_xml(effect.xml.as_bytes())?;
        if reparsed.kind != effect.kind {
            return invalid("snapshot effect kind does not match its XML root".into());
        }
        self.effects.push(effect);
        Ok(self)
    }

    pub fn replace_effect(&mut self, index: usize, effect: Effect) -> Result<Option<Effect>> {
        let reparsed = Effect::from_xml(effect.xml.as_bytes())?;
        if reparsed.kind != effect.kind {
            return invalid("snapshot effect kind does not match its XML root".into());
        }
        let Some(slot) = self.effects.get_mut(index) else {
            return Ok(None);
        };
        Ok(Some(std::mem::replace(slot, effect)))
    }

    pub fn remove_effect(&mut self, index: usize) -> Option<Effect> {
        (index < self.effects.len()).then(|| self.effects.remove(index))
    }

    pub fn clear_effects(&mut self) -> bool {
        let changed = !self.effects.is_empty();
        self.effects.clear();
        changed
    }

    #[must_use]
    pub const fn ext(&self) -> Option<&ExtList> {
        self.extension_list.as_ref()
    }

    pub fn set_ext(&mut self, extension: ExtList) -> Result<&mut Self> {
        validate_extension_list(
            Some(&extension),
            &[ExtKind::DrawingMl, ExtKind::StrictDrawingMl],
        )?;
        self.extension_list = Some(extension);
        Ok(self)
    }

    pub fn clear_ext(&mut self) -> Option<ExtList> {
        self.extension_list.take()
    }
}

/// One inert image relationship owned by a web-extension snapshot.
#[derive(Debug, Clone, PartialEq, Eq)]
pub(in crate::web) struct SnapshotResource {
    pub(in crate::web) relationship_id: String,
    pub(in crate::web) target: SnapshotTarget,
}

/// Internal image bytes or an external linked image target.
///
/// External targets are retained as strings and are never fetched.
#[derive(Debug, Clone, PartialEq, Eq)]
pub(in crate::web) enum SnapshotTarget {
    Internal {
        part_name: PackURI,
        content_type: String,
        data: Arc<Vec<u8>>,
    },
    External {
        target: String,
    },
}

/// A borrowed embedded snapshot. Cloning `shared` clones only the `Arc`.
#[derive(Debug, Clone, Copy)]
pub struct Image<'a> {
    pub(in crate::web) part_name: &'a PackURI,
    pub(in crate::web) content_type: &'a str,
    pub(in crate::web) data: &'a Arc<Vec<u8>>,
}

impl Image<'_> {
    #[must_use]
    pub fn name(&self) -> &PackURI {
        self.part_name
    }

    #[must_use]
    pub fn content_type(&self) -> &str {
        self.content_type
    }

    #[must_use]
    pub fn bytes(&self) -> &[u8] {
        self.data.as_slice()
    }

    /// Share the backing allocation without copying the image payload.
    #[must_use]
    pub fn shared(&self) -> Arc<Vec<u8>> {
        Arc::clone(self.data)
    }
}

/// A borrowed linked-image target. External targets remain inert and are never fetched.
#[derive(Debug, Clone, Copy)]
pub enum Link<'a> {
    Internal(Image<'a>),
    External(&'a str),
}

impl<'a> Link<'a> {
    #[must_use]
    pub const fn internal(self) -> Option<Image<'a>> {
        match self {
            Self::Internal(image) => Some(image),
            Self::External(_) => None,
        }
    }

    #[must_use]
    pub const fn external(self) -> Option<&'a str> {
        match self {
            Self::External(target) => Some(target),
            Self::Internal(_) => None,
        }
    }
}
