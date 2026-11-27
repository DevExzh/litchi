//! Package-independent values for `PresentationML` sections.

use std::collections::{HashMap, HashSet};

/// A logical group of presentation slides.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Section {
    /// Optional display name from `CT_Section`.
    pub name: Option<String>,
    /// Optional section GUID from `CT_Section`.
    pub id: Option<String>,
    /// Presentation slide identifiers in source order.
    pub slide_ids: Vec<u32>,
    /// Optional, inert `p:extLst` permitted by `CT_Section`.
    pub extension_xml: Option<Vec<u8>>,
}

impl Section {
    /// Create a named section with a GUID.
    pub fn new(name: impl Into<String>, id: impl Into<String>) -> Self {
        Self {
            name: Some(name.into()),
            id: Some(id.into()),
            slide_ids: Vec::new(),
            extension_xml: None,
        }
    }

    /// Add a presentation slide identifier.
    pub fn add_slide(&mut self, slide_id: u32) {
        self.slide_ids.push(slide_id);
    }

    /// Add presentation slide identifiers.
    #[must_use]
    pub fn with_slides(mut self, slide_ids: impl IntoIterator<Item = u32>) -> Self {
        self.slide_ids.extend(slide_ids);
        self
    }
}

/// Ordered presentation sections.
#[derive(Debug, Clone, Default, PartialEq, Eq)]
pub struct List {
    pub(crate) sections: Vec<Section>,
}

impl List {
    #[must_use]
    pub fn new() -> Self {
        Self::default()
    }

    pub fn add_section(&mut self, section: Section) {
        self.sections.push(section);
    }

    #[must_use]
    pub fn sections(&self) -> &[Section] {
        &self.sections
    }

    /// Get mutable access to the ordered sections.
    pub fn sections_mut(&mut self) -> &mut [Section] {
        &mut self.sections
    }

    /// Find a section by its stable GUID.
    #[must_use]
    pub fn get_by_id(&self, id: &str) -> Option<&Section> {
        self.sections
            .iter()
            .find(|section| section.id.as_deref() == Some(id))
    }

    /// Find a mutable section by its stable GUID.
    pub fn get_by_id_mut(&mut self, id: &str) -> Option<&mut Section> {
        self.sections
            .iter_mut()
            .find(|section| section.id.as_deref() == Some(id))
    }

    /// Remove a section by its stable GUID.
    pub fn remove_by_id(&mut self, id: &str) -> Option<Section> {
        self.sections
            .iter()
            .position(|section| section.id.as_deref() == Some(id))
            .map(|offset| self.sections.remove(offset))
    }

    /// Reorder sections by a complete stable-GUID permutation.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation fails.
    pub fn reorder(&mut self, ordered_ids: &[String]) -> crate::Result<()> {
        let expected = self
            .sections
            .iter()
            .filter_map(|section| section.id.clone())
            .collect::<HashSet<_>>();
        let actual = ordered_ids.iter().cloned().collect::<HashSet<_>>();
        if expected.len() != self.sections.len()
            || expected != actual
            || ordered_ids.len() != self.sections.len()
        {
            return Err(crate::Error::Invalid(
                "section reorder is not a GUID permutation".into(),
            ));
        }
        self.sections = ordered_ids
            .iter()
            .map(|id| self.get_by_id(id).cloned())
            .collect::<Option<Vec<_>>>()
            .ok_or_else(|| crate::Error::Invalid("section reorder lost a validated GUID".into()))?;
        Ok(())
    }

    /// Keep section membership in presentation order.
    pub(crate) fn sort_slide_membership(&mut self, ordered_slide_ids: &[u32]) {
        let positions = ordered_slide_ids
            .iter()
            .enumerate()
            .map(|(offset, id)| (*id, offset))
            .collect::<HashMap<_, _>>();
        for section in &mut self.sections {
            section
                .slide_ids
                .sort_by_key(|id| positions.get(id).copied().unwrap_or(usize::MAX));
        }
    }

    #[must_use]
    pub fn len(&self) -> usize {
        self.sections.len()
    }

    #[must_use]
    pub fn is_empty(&self) -> bool {
        self.sections.is_empty()
    }
}
