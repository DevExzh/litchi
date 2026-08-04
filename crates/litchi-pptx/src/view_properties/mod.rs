//! Typed PresentationML view properties.
//!
//! The model, XML codec, and package relationship handling are kept in
//! separate layers while this module preserves the historical public API.

mod codec;
mod model;
mod package;

pub use model::{
    CommonSlideView, CommonView, GridSpacing, Guide, GuideOrientation, NormalView, OutlineSlide,
    OutlineView, Point, Ratio, RestoredPane, SimpleView, SlideLikeView, SorterView, SplitterState,
    ViewKind, ViewProperties,
};
pub use package::load_from_package;
