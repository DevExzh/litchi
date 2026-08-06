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
pub(super) use quick_xml::events::{BytesRef, BytesStart, Event};
pub(super) use quick_xml::name::{Namespace as XmlNamespace, ResolveResult};
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

/// Parser for ODP-specific structures.
///
/// This provides parsing logic specific to presentations,
/// including slide and shape parsing.
pub(crate) struct Parser;
