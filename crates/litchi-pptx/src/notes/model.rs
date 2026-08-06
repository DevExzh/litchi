/// PresentationML namespace and relationship conformance used by a notes graph.
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub enum Conformance {
    Transitional,
    Strict,
}

impl Conformance {
    pub(crate) fn p(self) -> &'static str {
        if self == Self::Strict {
            super::PS
        } else {
            super::P
        }
    }
    pub(crate) fn a(self) -> &'static str {
        if self == Self::Strict {
            super::AS
        } else {
            super::A
        }
    }
    pub(crate) fn r(self) -> &'static str {
        if self == Self::Strict {
            super::RS
        } else {
            super::R
        }
    }
    pub(crate) fn notes_slide_rel(self) -> &'static str {
        if self == Self::Strict {
            litchi_opc::constants::relationship_type::STRICT_NOTES_SLIDE
        } else {
            litchi_opc::constants::relationship_type::NOTES_SLIDE
        }
    }
    pub(crate) fn notes_master_rel(self) -> &'static str {
        if self == Self::Strict {
            litchi_opc::constants::relationship_type::STRICT_NOTES_MASTER
        } else {
            litchi_opc::constants::relationship_type::NOTES_MASTER
        }
    }
    pub(crate) fn slide_rel(self) -> &'static str {
        if self == Self::Strict {
            "http://purl.oclc.org/ooxml/officeDocument/relationships/slide"
        } else {
            litchi_opc::constants::relationship_type::SLIDE
        }
    }
    pub(crate) fn theme_rel(self) -> &'static str {
        if self == Self::Strict {
            "http://purl.oclc.org/ooxml/officeDocument/relationships/theme"
        } else {
            litchi_opc::constants::relationship_type::THEME
        }
    }
}

use litchi_opc::TargetMode;

/// An opaque OPC relationship retained by a notes resource.
///
/// The notes semantic layer interprets only the edges required by the
/// PresentationML graph. Every other relationship is retained as inert
/// metadata so a transaction can round-trip vendor extensions and related
/// parts without activating or rendering them.
#[derive(Clone, Debug, PartialEq, Eq)]
pub struct Link {
    pub(crate) id: String,
    pub(crate) relationship_type: String,
    pub(crate) target_ref: String,
    pub(crate) target_mode: TargetMode,
}

impl Link {
    pub(crate) fn new(
        id: impl Into<String>,
        relationship_type: impl Into<String>,
        target_ref: impl Into<String>,
        target_mode: TargetMode,
    ) -> Self {
        Self {
            id: id.into(),
            relationship_type: relationship_type.into(),
            target_ref: target_ref.into(),
            target_mode,
        }
    }

    /// Relationship identifier in the owning `.rels` part.
    pub fn id(&self) -> &str {
        &self.id
    }

    /// Relationship type URI, retained without vocabulary interpretation.
    pub fn relationship_type(&self) -> &str {
        &self.relationship_type
    }

    /// Exact authored target reference.
    pub fn target(&self) -> &str {
        &self.target_ref
    }

    /// Whether the target is external and therefore never resolved.
    pub const fn target_mode(&self) -> TargetMode {
        self.target_mode
    }

    /// Whether the target is external and remains inert.
    pub const fn is_external(&self) -> bool {
        matches!(self.target_mode, TargetMode::External)
    }
}

/// Owned notes-master theme resource.
#[derive(Clone, Debug, PartialEq, Eq)]
pub struct Theme {
    pub(crate) relationship_id: String,
    pub(crate) part_name: String,
    pub(crate) content_type: String,
    pub(crate) data: Vec<u8>,
    pub(crate) relationships: Vec<Link>,
}

impl Theme {
    /// Return the validated package part name for diagnostics.
    pub fn part(&self) -> &str {
        &self.part_name
    }

    /// Return the validated resource content type.
    pub fn content_type(&self) -> &str {
        &self.content_type
    }

    /// Lend the inert theme XML payload.
    pub fn xml(&self) -> &[u8] {
        &self.data
    }

    /// Replace the inert theme XML, returning the previous allocation.
    pub fn replace_xml(&mut self, xml: Vec<u8>) -> Vec<u8> {
        std::mem::replace(&mut self.data, xml)
    }

    /// Lend all outbound relationships, including opaque vendor edges.
    pub fn relationships(&self) -> &[Link] {
        &self.relationships
    }
}

/// Owned notes-master resource and its theme.
#[derive(Clone, Debug, PartialEq, Eq)]
pub struct Master {
    pub(crate) presentation_relationship_id: String,
    pub(crate) part_name: String,
    pub(crate) content_type: String,
    pub(crate) data: Vec<u8>,
    pub(crate) relationships: Vec<Link>,
    pub(crate) theme: Theme,
}

impl Master {
    /// Return the validated package part name for diagnostics.
    pub fn part(&self) -> &str {
        &self.part_name
    }

    /// Return the validated resource content type.
    pub fn content_type(&self) -> &str {
        &self.content_type
    }

    /// Lend the inert notes-master XML payload.
    pub fn xml(&self) -> &[u8] {
        &self.data
    }

    /// Replace the inert notes-master XML, returning the previous allocation.
    pub fn replace_xml(&mut self, xml: Vec<u8>) -> Vec<u8> {
        std::mem::replace(&mut self.data, xml)
    }

    /// Lend the owned notes-master theme resource.
    pub fn theme(&self) -> &Theme {
        &self.theme
    }

    /// Mutably lend the owned notes-master theme resource.
    pub fn theme_mut(&mut self) -> &mut Theme {
        &mut self.theme
    }

    /// Lend all outbound relationships, including the required theme edge.
    pub fn relationships(&self) -> &[Link] {
        &self.relationships
    }
}

/// Owned speaker-notes resource for one slide.
#[derive(Clone, Debug, PartialEq, Eq)]
pub struct Slide {
    pub(crate) slide_part_name: String,
    pub(crate) slide_relationship_id: String,
    pub(crate) part_name: String,
    pub(crate) content_type: String,
    pub(crate) data: Vec<u8>,
    pub(crate) relationships: Vec<Link>,
    pub(crate) backlink_relationship_id: String,
    pub(crate) notes_master_relationship_id: String,
}

impl Slide {
    /// Return the validated owning slide part name for diagnostics.
    pub fn owner(&self) -> &str {
        &self.slide_part_name
    }

    /// Return the validated notes-slide part name for diagnostics.
    pub fn part(&self) -> &str {
        &self.part_name
    }

    /// Return the validated resource content type.
    pub fn content_type(&self) -> &str {
        &self.content_type
    }

    /// Lend the inert notes-slide XML payload.
    pub fn xml(&self) -> &[u8] {
        &self.data
    }

    /// Replace the inert notes-slide XML, returning the previous allocation.
    pub fn replace_xml(&mut self, xml: Vec<u8>) -> Vec<u8> {
        std::mem::replace(&mut self.data, xml)
    }

    /// Lend all outbound relationships, including opaque vendor edges.
    pub fn relationships(&self) -> &[Link] {
        &self.relationships
    }
}

/// Complete owned notes graph for one presentation.
///
/// Topology identities remain private. XML payload replacement is explicit,
/// and [`crate::notes::put`] consumes the graph so successful storage moves
/// those buffers.
#[derive(Clone, Debug, PartialEq, Eq)]
pub struct Graph {
    pub(crate) conformance: Conformance,
    pub(crate) master: Master,
    pub(crate) slides: Vec<Slide>,
}

impl Graph {
    /// Return the graph's Strict or Transitional namespace profile.
    pub fn conformance(&self) -> Conformance {
        self.conformance
    }

    /// Lend the shared notes-master resource.
    pub fn master(&self) -> &Master {
        &self.master
    }

    /// Mutably lend the shared notes-master resource.
    pub fn master_mut(&mut self) -> &mut Master {
        &mut self.master
    }

    /// Lend notes slides in presentation order.
    pub fn slides(&self) -> &[Slide] {
        &self.slides
    }

    /// Mutably lend notes slides in presentation order.
    pub fn slides_mut(&mut self) -> &mut [Slide] {
        &mut self.slides
    }
}
