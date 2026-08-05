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
pub use animation::{Attribute, Kind, Namespace, Node};
pub use declaration::{
    DateTimeDeclaration, DateTimeSource, DeclarationBinding, DeclarationTarget, Declarations,
    TextDeclaration, parse as parse_declarations,
};
pub use legacy_animation::{LegacyAnimationKind, LegacyAnimationNode};
pub use media::{Actuate, Parameter, Reference, Show};
pub use settings::{CustomShow, FeatureState, Settings, parse as parse_settings};
pub use slide::{
    DrawingAttribute, DrawingAttributeNamespace, DrawingShapeKind, EnhancedGeometry,
    EnhancedGeometryChild, EnhancedGeometryChildKind, Shape, Slide,
};
pub use transition::{
    Transition, TransitionDirection, TransitionSound, TransitionSoundShow, TransitionSpeed,
    TransitionStyle, TransitionType,
};
