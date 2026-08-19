//! Layered ODP parser codec facade.
//!
//! XML traversal, semantic model assembly, and structural validation are kept in
//! separate implementation seams while this module owns the parser type.

mod xml;

pub(super) use super::model::{
    Element, ParagraphText, ShapeBuilder, ShapeContainerScope, ShapeElement,
    TransitionStyleDefinition, TransitionStyles,
};
pub(super) use crate::model::animation::ANIMATION_NAMESPACE;
pub(super) use crate::model::legacy_animation::validate_legacy_animation_root;
pub(super) use crate::model::legacy_animation::{Kind as AnimationKind, Node as AnimationNode};
pub(super) use crate::model::{
    Action, Actuate, Attribute, Direction, DrawingAttribute, DrawingAttributeNamespace,
    DrawingHyperlink, DrawingShapeKind, Effect, EffectDirection, EnhancedGeometry,
    EnhancedGeometryChild, EnhancedGeometryChildKind, EventListener, HyperlinkShow, Kind,
    Namespace, Node, Parameter, Reference, ScriptEventListener, Shape, ShapeEventListener, Show,
    Slide, Sound, SoundShow, Speed, Style, Transition, Type,
};
pub(super) use litchi_core::{Error, Result, ShapeType};
pub(super) use quick_xml::XmlVersion;
pub(super) use quick_xml::events::attributes::{Attribute as RawAttribute, Attributes};
pub(super) use quick_xml::events::{BytesRef, BytesStart, Event};
pub(super) use quick_xml::name::{LocalName, Namespace as XmlNamespace, ResolveResult};
pub(super) use quick_xml::reader::NsReader;
pub(super) use std::collections::{HashMap, HashSet};

pub(super) const DRAW_NAMESPACE: &[u8] = b"urn:oasis:names:tc:opendocument:xmlns:drawing:1.0";
pub(super) const DR3D_NAMESPACE: &[u8] = b"urn:oasis:names:tc:opendocument:xmlns:dr3d:1.0";
pub(super) const OFFICE_NAMESPACE: &[u8] = b"urn:oasis:names:tc:opendocument:xmlns:office:1.0";
pub(super) const PRESENTATION_NAMESPACE: &[u8] =
    b"urn:oasis:names:tc:opendocument:xmlns:presentation:1.0";
pub(super) const SCRIPT_NAMESPACE: &[u8] = b"urn:oasis:names:tc:opendocument:xmlns:script:1.0";
pub(super) const STYLE_NAMESPACE: &[u8] = b"urn:oasis:names:tc:opendocument:xmlns:style:1.0";
pub(super) const SMIL_NAMESPACE: &[u8] =
    b"urn:oasis:names:tc:opendocument:xmlns:smil-compatible:1.0";
pub(super) const SVG_NAMESPACE: &[u8] = b"urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0";
pub(super) const TABLE_NAMESPACE: &[u8] = b"urn:oasis:names:tc:opendocument:xmlns:table:1.0";
pub(super) const TEXT_NAMESPACE: &[u8] = b"urn:oasis:names:tc:opendocument:xmlns:text:1.0";
pub(super) const XLINK_NAMESPACE: &[u8] = b"http://www.w3.org/1999/xlink";
pub(super) const XML_NAMESPACE: &[u8] = b"http://www.w3.org/XML/1998/namespace";
pub(super) const ANIMATION_NAMESPACE_BYTES: &[u8] = ANIMATION_NAMESPACE.as_bytes();

/// Copyable snapshot of an element's resolved namespace.
///
/// [`ResolveResult`] borrows the reader's namespace buffer, keeping the reader
/// exclusively borrowed while the value is live. The fused content.xml pass
/// interleaves collector feeds (which need the reader) with namespace checks,
/// so those checks run against this token instead. The known namespace URIs
/// are pairwise distinct, so a bound namespace maps to exactly one class;
/// unbound and unknown prefixes map to [`NsClass::Other`], matching
/// `is_namespace` returning false for every known URI.
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub(super) enum NsClass {
    Animation,
    Drawing,
    Dr3d,
    Office,
    Presentation,
    Script,
    Style,
    Table,
    Text,
    Other,
}

impl NsClass {
    pub(super) fn from_resolve(namespace: &ResolveResult<'_>) -> Self {
        match namespace {
            ResolveResult::Bound(XmlNamespace(uri)) => Self::from_uri(uri),
            ResolveResult::Unbound | ResolveResult::Unknown(_) => Self::Other,
        }
    }

    pub(super) fn from_uri(uri: &[u8]) -> Self {
        if uri == ANIMATION_NAMESPACE_BYTES {
            Self::Animation
        } else if uri == DRAW_NAMESPACE {
            Self::Drawing
        } else if uri == DR3D_NAMESPACE {
            Self::Dr3d
        } else if uri == OFFICE_NAMESPACE {
            Self::Office
        } else if uri == PRESENTATION_NAMESPACE {
            Self::Presentation
        } else if uri == SCRIPT_NAMESPACE {
            Self::Script
        } else if uri == STYLE_NAMESPACE {
            Self::Style
        } else if uri == TABLE_NAMESPACE {
            Self::Table
        } else if uri == TEXT_NAMESPACE {
            Self::Text
        } else {
            Self::Other
        }
    }
}

/// Parser for ODP-specific structures.
///
/// This provides parsing logic specific to presentations,
/// including slide and shape parsing.
pub(crate) struct Parser;
