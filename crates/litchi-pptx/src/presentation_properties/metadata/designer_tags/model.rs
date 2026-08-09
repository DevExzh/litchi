//! Source-bound slide-ID Designer-tag state.

use std::sync::Arc;

use super::codec::Layout;
use super::{Edit, Limits, Tags};
use crate::{Error, Result};

/// The exact relationship and part binding of one stable slide ID.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Binding {
    pub(crate) relationship_id: String,
    pub(crate) relationship_type: String,
    pub(crate) relationship_target: String,
    pub(crate) part_name: String,
    pub(crate) part_content_type: String,
}

impl Binding {
    /// Return the presentation relationship ID stored on `p:sldId`.
    #[inline]
    #[must_use]
    pub fn relationship_id(&self) -> &str {
        &self.relationship_id
    }

    /// Return the relationship type exactly as stored in the OPC graph.
    #[inline]
    #[must_use]
    pub fn relationship_type(&self) -> &str {
        &self.relationship_type
    }

    /// Return the relationship target exactly as stored in the OPC graph.
    #[inline]
    #[must_use]
    pub fn relationship_target(&self) -> &str {
        &self.relationship_target
    }

    /// Return the normalized related slide-part name.
    #[inline]
    #[must_use]
    pub fn part_name(&self) -> &str {
        &self.part_name
    }

    /// Return the related slide-part content type.
    #[inline]
    #[must_use]
    pub fn part_content_type(&self) -> &str {
        &self.part_content_type
    }
}

/// A bounded inventory of Designer-tag extensions on one stable slide ID.
///
/// Zero occurrences means absent. One empty [`Tags`] value means explicitly
/// present but empty. More than one occurrence is retained for inspection,
/// while [`Snapshot::edit`] refuses the ambiguous singular mutation.
#[derive(Debug)]
pub struct Snapshot {
    pub(crate) presentation_part_name: String,
    pub(crate) presentation_content_type: String,
    pub(crate) source_xml: Arc<Vec<u8>>,
    pub(crate) slide_id: u32,
    pub(crate) binding: Binding,
    pub(crate) occurrences: Vec<Tags>,
    pub(crate) layout: Layout,
    pub(crate) limits: Limits,
    pub(crate) revision: super::Revision,
}

impl Snapshot {
    /// Load one stable slide-ID owner under safe default bounds.
    ///
    /// # Errors
    ///
    /// Returns an error if the input cannot be read or is malformed.
    #[inline]
    pub fn load(package: &litchi_opc::OpcPackage, slide_id: u32) -> Result<Self> {
        super::load_snapshot(package, slide_id)
    }

    /// Load one stable slide-ID owner under caller-supplied bounds.
    ///
    /// # Errors
    ///
    /// Returns an error if the input cannot be read or is malformed.
    #[inline]
    pub fn load_with_limits(
        package: &litchi_opc::OpcPackage,
        slide_id: u32,
        limits: Limits,
    ) -> Result<Self> {
        super::load_snapshot_with_limits(package, slide_id, limits)
    }

    /// Return the stable `p:sldId/@id` key.
    #[inline]
    #[must_use]
    pub const fn slide_id(&self) -> u32 {
        self.slide_id
    }

    /// Return the validated relationship/part binding.
    #[inline]
    #[must_use]
    pub fn binding(&self) -> &Binding {
        &self.binding
    }

    /// Return the number of matching outer extension entries.
    #[inline]
    #[must_use]
    pub fn occurrence_count(&self) -> usize {
        self.occurrences.len()
    }

    /// Iterate over every matching extension in source order.
    #[inline]
    #[must_use]
    pub fn occurrences(&self) -> impl ExactSizeIterator<Item = &Tags> {
        self.occurrences.iter()
    }

    /// Return the optional singular tag list, refusing duplicate extensions.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation fails.
    pub fn tags(&self) -> Result<Option<&Tags>> {
        match self.occurrences.as_slice() {
            [] => Ok(None),
            [value] => Ok(Some(value)),
            values => Err(ambiguous(values.len())),
        }
    }

    /// Borrow the exact owning presentation XML captured by this snapshot.
    #[inline]
    #[must_use]
    pub fn source_xml(&self) -> &[u8] {
        self.source_xml.as_slice()
    }

    /// Return the selected-host revision used for optimistic stale checks.
    #[inline]
    #[must_use]
    pub const fn revision(&self) -> super::Revision {
        self.revision
    }

    /// Consume this source-bound snapshot into an isolated edit.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation fails.
    pub fn edit(self) -> Result<Edit> {
        if self.occurrences.len() > 1 {
            return Err(ambiguous(self.occurrences.len()));
        }
        Ok(Edit::new(self))
    }

    pub(crate) fn duplicate(&self) -> Self {
        Self {
            presentation_part_name: self.presentation_part_name.clone(),
            presentation_content_type: self.presentation_content_type.clone(),
            source_xml: Arc::clone(&self.source_xml),
            slide_id: self.slide_id,
            binding: self.binding.clone(),
            occurrences: self.occurrences.clone(),
            layout: self.layout.clone(),
            limits: self.limits,
            revision: self.revision,
        }
    }

    pub(crate) fn singular(&self) -> Result<Option<&Tags>> {
        self.tags()
    }

    pub(crate) fn same_selected_source(&self, other: &Self) -> bool {
        self.presentation_part_name == other.presentation_part_name
            && self.presentation_content_type == other.presentation_content_type
            && self.slide_id == other.slide_id
            && self.binding == other.binding
            && self.layout.host_bytes(self.source_xml.as_slice())
                == other.layout.host_bytes(other.source_xml.as_slice())
            && self.revision == other.revision
    }
}

pub(crate) fn ambiguous(count: usize) -> Error {
    Error::Invalid(format!(
        "slide ID has {count} Designer-tag extensions; singular mutation is ambiguous"
    ))
}
