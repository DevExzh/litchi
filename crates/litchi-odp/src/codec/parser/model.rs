//! Parser-owned state and builders for ODP XML decoding.

use crate::model::animation::Kind;
use crate::model::legacy_animation::Kind as AnimationKind;
use crate::model::{
    DrawingAttribute, DrawingHyperlink, DrawingShapeKind, EnhancedGeometry, Reference, Shape,
    ShapeEventListener, Transition,
};
use litchi_core::ShapeType;
use std::collections::HashMap;

#[derive(Clone, Debug, Default, PartialEq, Eq)]
pub(super) struct TransitionStyleDefinition {
    pub(super) parent: Option<String>,
    pub(super) transition: Transition,
}

#[derive(Debug, Default, PartialEq, Eq)]
pub(super) struct TransitionStyles {
    pub(super) named: HashMap<String, TransitionStyleDefinition>,
    pub(super) default: Transition,
}

#[derive(Clone, Copy)]
pub(super) enum ShapeElement {
    Frame,
    Rect,
    Ellipse,
    Line,
    CustomShape,
    Circle,
    Path,
    Polygon,
    Polyline,
    RegularPolygon,
    PageThumbnail,
    Measure,
    Caption,
    Connector,
    Control,
    Group,
    ThreeDimensionalScene,
    ThreeDimensionalLight,
    ThreeDimensionalCube,
    ThreeDimensionalSphere,
    ThreeDimensionalExtrude,
    ThreeDimensionalRotate,
}

/// Container scope that supplies top-level drawing shapes.
#[derive(Clone, Copy, PartialEq, Eq)]
pub(super) enum ShapeContainerScope {
    /// `draw:page` elements in presentations and drawings.
    DrawPages,
    /// `table:shapes` children of top-level spreadsheet tables.
    SpreadsheetTables,
}

#[derive(Clone, Copy)]
pub(super) enum Element {
    Page,
    Notes,
    SheetShapes,
    SpreadsheetRoot,
    Shape(ShapeElement),
    Image,
    Table,
    Object,
    Plugin,
    PluginParameter,
    DrawingHyperlink,
    EnhancedGeometry,
    EnhancedEquation,
    EnhancedHandle,
    EventListeners,
    EventListener,
    ScriptEventListener,
    Sound,
    TextParagraph,
    TextSpace,
    TextTab,
    TextLineBreak,
    Animation(Kind),
    UnknownAnimation,
    LegacyAnimation(AnimationKind),
    Other,
}

/// Internal structure for building shapes during parsing
#[allow(
    dead_code,
    clippy::struct_excessive_bools,
    reason = "parser-internal builder accumulates independent XML parse flags and fields consumed by later pipeline stages; splitting them would complicate the parser without changing behavior"
)]
pub(super) struct ShapeBuilder {
    pub(super) shape_type: ShapeType,
    pub(super) drawing_kind: Option<DrawingShapeKind>,
    pub(super) drawing_attributes: Vec<DrawingAttribute>,
    pub(super) children: Vec<Shape>,
    pub(super) enhanced_geometry: Option<EnhancedGeometry>,
    pub(super) text: String,
    pub(super) name: Option<String>,
    pub(super) x: Option<String>,
    pub(super) y: Option<String>,
    pub(super) width: Option<String>,
    pub(super) height: Option<String>,
    pub(super) style_name: Option<String>,
    pub(super) layer: Option<String>,
    pub(super) z_index: Option<String>,
    pub(super) transform: Option<String>,
    pub(super) presentation_class: Option<String>,
    pub(super) presentation_placeholder: Option<bool>,
    pub(super) presentation_user_transformed: Option<bool>,
    pub(super) image_href: Option<String>,
    pub(super) media: Option<Reference>,
    pub(super) hyperlink: Option<DrawingHyperlink>,
    pub(super) event_listeners: Vec<ShapeEventListener>,
    pub(super) event_listeners_seen: bool,
    pub(super) is_frame: bool,
    pub(super) is_title: bool,
    pub(super) has_paragraph: bool,
}

#[derive(Default)]
pub(super) struct ParagraphText {
    pub(super) value: String,
    pub(super) trailing_collapsible_space: bool,
}

impl ParagraphText {
    pub(super) fn push_text(&mut self, text: &str) {
        for character in text.chars() {
            if character.is_whitespace() {
                if !self.value.is_empty()
                    && !self
                        .value
                        .chars()
                        .next_back()
                        .is_some_and(char::is_whitespace)
                {
                    self.value.push(' ');
                    self.trailing_collapsible_space = true;
                }
            } else {
                self.value.push(character);
                self.trailing_collapsible_space = false;
            }
        }
    }

    pub(super) fn push_explicit(&mut self, character: char, count: usize) {
        self.value.extend(std::iter::repeat_n(character, count));
        self.trailing_collapsible_space = false;
    }

    pub(super) fn finish(mut self) -> String {
        if self.trailing_collapsible_space {
            self.value.pop();
        }
        self.value
    }
}

#[allow(
    dead_code,
    reason = "parser-internal builder; not every helper is exercised on all code paths"
)]
impl ShapeBuilder {
    pub(super) fn new() -> Self {
        Self {
            shape_type: ShapeType::AutoShape,
            drawing_kind: None,
            drawing_attributes: Vec::new(),
            children: Vec::new(),
            enhanced_geometry: None,
            text: String::new(),
            name: None,
            x: None,
            y: None,
            width: None,
            height: None,
            style_name: None,
            layer: None,
            z_index: None,
            transform: None,
            presentation_class: None,
            presentation_placeholder: None,
            presentation_user_transformed: None,
            image_href: None,
            media: None,
            hyperlink: None,
            event_listeners: Vec::new(),
            event_listeners_seen: false,
            is_frame: false,
            is_title: false,
            has_paragraph: false,
        }
    }

    pub(super) fn build(self) -> Shape {
        Shape {
            shape_type: self.shape_type,
            drawing_kind: self.drawing_kind,
            drawing_attributes: self.drawing_attributes,
            children: self.children,
            enhanced_geometry: self.enhanced_geometry,
            text: self.text,
            name: self.name,
            x: self.x,
            y: self.y,
            width: self.width,
            height: self.height,
            style_name: self.style_name,
            layer: self.layer,
            z_index: self.z_index,
            transform: self.transform,
            presentation_class: self.presentation_class,
            presentation_placeholder: self.presentation_placeholder,
            presentation_user_transformed: self.presentation_user_transformed,
            image_href: self.image_href,
            media: self.media,
            hyperlink: self.hyperlink,
            event_listeners: self.event_listeners,
        }
    }

    pub(super) fn push_paragraph(&mut self, text: &str) {
        if self.has_paragraph {
            self.text.push('\n');
        }
        self.text.push_str(text);
        self.has_paragraph = true;
    }
}
