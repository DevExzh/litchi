use super::super::{Conformance, Source};
use std::ops::Range;

pub(super) struct Attached {
    pub(super) relationship_type: String,
    pub(super) source: Source,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub(super) enum OwnerKind {
    Presentation,
    CommonSlide,
}

#[derive(Debug)]
pub(super) struct OwnerXml {
    pub(super) conformance: Conformance,
    pub(super) insertion: usize,
    pub(super) container: Option<Container>,
    pub(super) anchor: Option<Anchor>,
}

#[derive(Debug)]
pub(super) struct Container {
    pub(super) span: Range<usize>,
    pub(super) close_start: usize,
    pub(super) empty: bool,
    pub(super) qualified_name: Vec<u8>,
    pub(super) child_elements: usize,
    pub(super) other_content: bool,
    pub(super) preserve_when_empty: bool,
}

#[derive(Debug)]
pub(crate) struct Anchor {
    pub(crate) id: String,
    pub(crate) span: Range<usize>,
    pub(crate) id_value: Range<usize>,
}

pub(super) struct OpenContainer {
    pub(super) start: usize,
    pub(super) depth: usize,
    pub(super) qualified_name: Vec<u8>,
    pub(super) child_elements: usize,
    pub(super) other_content: bool,
    pub(super) preserve_when_empty: bool,
    pub(super) tags_seen: bool,
}

pub(super) struct OpenAnchor {
    pub(super) start: usize,
    pub(super) depth: usize,
    pub(super) id: String,
    pub(super) id_value: Range<usize>,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub(super) enum CommonSlidePhase {
    Start,
    Background,
    Shapes,
    CustomerData,
    Controls,
    Extensions,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub(super) enum OwnerMapName {
    Presentation,
    Slide,
    SlideLayout,
    SlideMaster,
    Notes,
    NotesMaster,
    HandoutMaster,
    ShapeTree,
    CustomerDataList,
    Tags,
    Kinsoku,
    DefaultTextStyle,
    ModifyVerifier,
    Extensions,
}

impl OwnerMapName {
    pub(super) fn from_local(local: &[u8]) -> Option<Self> {
        match local {
            b"presentation" => Some(Self::Presentation),
            b"sld" => Some(Self::Slide),
            b"sldLayout" => Some(Self::SlideLayout),
            b"sldMaster" => Some(Self::SlideMaster),
            b"notes" => Some(Self::Notes),
            b"notesMaster" => Some(Self::NotesMaster),
            b"handoutMaster" => Some(Self::HandoutMaster),
            b"spTree" => Some(Self::ShapeTree),
            b"custDataLst" => Some(Self::CustomerDataList),
            b"tags" => Some(Self::Tags),
            b"kinsoku" => Some(Self::Kinsoku),
            b"defaultTextStyle" => Some(Self::DefaultTextStyle),
            b"modifyVerifier" => Some(Self::ModifyVerifier),
            b"extLst" => Some(Self::Extensions),
            _ => None,
        }
    }

    pub(super) fn is_owner_root(self) -> bool {
        matches!(
            self,
            Self::Presentation
                | Self::Slide
                | Self::SlideLayout
                | Self::SlideMaster
                | Self::Notes
                | Self::NotesMaster
                | Self::HandoutMaster
        )
    }

    pub(super) fn is_presentation_later(self) -> bool {
        matches!(
            self,
            Self::Kinsoku | Self::DefaultTextStyle | Self::ModifyVerifier | Self::Extensions
        )
    }
}

#[derive(Debug)]
pub(super) struct OwnerMapElement {
    pub(super) name: OwnerMapName,
    pub(super) conformance: Conformance,
    pub(super) span: Range<usize>,
    pub(super) open_end: usize,
    pub(super) close_start: usize,
    pub(super) empty: bool,
    pub(super) qualified_name: Vec<u8>,
    pub(super) preserve_when_empty: bool,
}

pub(super) struct AnchorIdentity {
    pub(super) value: String,
    pub(super) id_value: Range<usize>,
}
