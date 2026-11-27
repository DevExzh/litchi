use std::ops::Range;

use super::super::{Conformance, Source};

pub(super) struct Attached {
    pub(super) relationship_type: String,
    pub(super) source: Source,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub(super) enum Family {
    Shape,
    Picture,
    Connector,
    GraphicFrame,
    Group,
}

impl Family {
    pub(super) fn from_local(local: &[u8]) -> Option<Self> {
        match local {
            b"sp" => Some(Self::Shape),
            b"pic" => Some(Self::Picture),
            b"cxnSp" => Some(Self::Connector),
            b"graphicFrame" => Some(Self::GraphicFrame),
            b"grpSp" => Some(Self::Group),
            _ => None,
        }
    }

    pub(super) fn non_visual(self) -> &'static [u8] {
        match self {
            Self::Shape => b"nvSpPr",
            Self::Picture => b"nvPicPr",
            Self::Connector => b"nvCxnSpPr",
            Self::GraphicFrame => b"nvGraphicFramePr",
            Self::Group => b"nvGrpSpPr",
        }
    }

    pub(super) fn permits_placeholder(self) -> bool {
        !matches!(self, Self::Connector | Self::Group)
    }
}

#[derive(Debug, Clone)]
pub(super) struct Element {
    pub(super) span: Range<usize>,
    pub(super) open_end: usize,
    pub(super) close_start: usize,
    pub(super) empty: bool,
}

#[derive(Debug, Clone)]
pub(super) struct Container {
    pub(super) element: Element,
    pub(super) child_elements: usize,
    pub(super) preserve_when_empty: bool,
}

#[derive(Debug, Clone)]
pub(super) struct Anchor {
    pub(super) id: String,
    pub(super) span: Range<usize>,
    pub(super) id_value: Range<usize>,
}

#[derive(Debug)]
pub(super) struct Layout {
    pub(super) conformance: Conformance,
    pub(super) nv_pr: Element,
    pub(super) insertion: usize,
    pub(super) container: Option<Container>,
    pub(super) anchor: Option<Anchor>,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub(super) enum NvPhase {
    Start,
    Placeholder,
    Media,
    CustomerData,
    Extensions,
}

#[derive(Debug)]
pub(super) enum NodeKind {
    Root(Family),
    NonVisual,
    NvPr {
        phase: NvPhase,
        ext_start: Option<usize>,
    },
    Container {
        children: usize,
        tags_seen: bool,
        preserve_when_empty: bool,
    },
    Anchor {
        id: String,
        id_value: Range<usize>,
    },
    Opaque,
}

#[derive(Debug)]
pub(super) struct Node {
    pub(super) kind: NodeKind,
    pub(super) start: usize,
    pub(super) open_end: usize,
}

#[derive(Debug, Clone, Copy)]
pub(super) struct RawFrame {
    pub(super) semantic: bool,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub(super) enum RawCandidateKind {
    Supported(Family),
    Unsupported,
}

impl RawCandidateKind {
    pub(super) fn is_group(self) -> bool {
        matches!(self, Self::Supported(Family::Group))
    }

    pub(super) fn family(self) -> Option<Family> {
        match self {
            Self::Supported(family) => Some(family),
            Self::Unsupported => None,
        }
    }
}

#[derive(Debug, Clone, Copy)]
pub(super) struct RawMapFrame {
    pub(super) active_tree: bool,
    pub(super) active_candidate: Option<RawCandidateKind>,
    pub(super) selected_start: Option<usize>,
}
