pub mod action;
pub mod animation;
pub mod declaration;
pub mod legacy_animation;
pub mod media;
pub mod page_layout;
pub mod page_metadata;
pub mod settings;
pub mod slide;
pub mod transition;

pub use action::{
    Action, DrawingHyperlink, Effect, EffectDirection, EventListener, HyperlinkShow,
    ScriptEventListener, ShapeEventListener,
};
pub use animation::{
    AnimationAttribute, AnimationAttributeNamespace, AnimationKind, AnimationNode,
};
pub use declaration::{
    DateTimeDeclaration, DateTimeSource, DeclarationBinding, DeclarationTarget, Declarations,
    TextDeclaration, parse as parse_declarations,
};
pub use legacy_animation::{LegacyAnimationKind, LegacyAnimationNode};
pub use media::{MediaActuate, MediaParameter, MediaReference, MediaShow};
pub use page_layout::{
    Layouts, Measure, PageLayout, Placeholder, PlaceholderClass, Unit, parse as parse_page_layouts,
    remove_xml as remove_page_layout_xml, set_xml as set_page_layout_xml,
};
pub use page_metadata::{PageMetadata, PageMetadataCollection, parse as parse_page_metadata};
pub use settings::{CustomShow, FeatureState, Settings, parse as parse_settings};
pub use slide::{
    DrawingAttribute, DrawingAttributeNamespace, DrawingShapeKind, EnhancedGeometry,
    EnhancedGeometryChild, EnhancedGeometryChildKind, Shape, Slide,
};
pub use transition::{
    SlideTransition, TransitionDirection, TransitionSound, TransitionSoundShow, TransitionSpeed,
    TransitionStyle, TransitionType,
};
