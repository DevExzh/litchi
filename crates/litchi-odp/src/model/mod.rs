pub(crate) mod action;
pub(crate) mod animation;
pub(crate) mod declaration;
pub(crate) mod legacy_animation;
pub(crate) mod media;
mod page_layout;
pub(crate) mod page_metadata;
pub(crate) mod settings;
pub(crate) mod slide;
mod transition;

pub use action::{
    DrawingHyperlink, HyperlinkShow, PresentationAction, PresentationEffect,
    PresentationEffectDirection, PresentationEventListener, ScriptEventListener,
    ShapeEventListener,
};
pub use animation::{
    AnimationAttribute, AnimationAttributeNamespace, AnimationKind, AnimationNode,
};
pub use declaration::{
    PresentationDateTimeDeclaration, PresentationDateTimeSource, PresentationDeclarationBinding,
    PresentationDeclarationTarget, PresentationDeclarations, PresentationTextDeclaration,
    parse_presentation_declarations,
};
pub use legacy_animation::{LegacyAnimationKind, LegacyAnimationNode};
pub use media::{MediaActuate, MediaParameter, MediaReference, MediaShow};
pub use page_layout::{
    PresentationMeasure, PresentationMeasureUnit, PresentationPageLayout, PresentationPageLayouts,
    PresentationPlaceholder, PresentationPlaceholderClass, parse_presentation_page_layouts,
    remove_presentation_page_layout_xml, set_presentation_page_layout_xml,
};
pub use page_metadata::{
    PresentationPageMetadata, PresentationPageMetadataCollection, parse_presentation_page_metadata,
};
pub use settings::{
    CustomPresentationShow, PresentationFeatureState, PresentationSettings,
    parse_presentation_settings,
};
pub use slide::{
    DrawingAttribute, DrawingAttributeNamespace, DrawingShapeKind, EnhancedGeometry,
    EnhancedGeometryChild, EnhancedGeometryChildKind, Shape, Slide,
};
pub use transition::{
    SlideTransition, TransitionDirection, TransitionSound, TransitionSoundShow, TransitionSpeed,
    TransitionStyle, TransitionType,
};
