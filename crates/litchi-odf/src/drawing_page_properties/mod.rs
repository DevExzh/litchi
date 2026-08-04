//! Complete typed ODF `style:drawing-page-properties` support.

mod codec;
mod model;
mod package;

#[cfg(test)]
mod tests;

pub use codec::{parse_drawing_page_style_properties, set_drawing_page_style_properties_xml};
pub use model::{
    BackgroundSize, Color, Duration, Fill, FillRule, ImageRefPoint, LengthOrPercent,
    NonNegativeInteger, Percent, Repeat, Sound, SoundShow, Style, StyleNameRef, StyleProperties,
    Styles, TileDirection, TileRepeatOffset, TransitionDirection, TransitionSpeed, TransitionStyle,
    TransitionType, Visibility,
};

// Historical public names remain aliases; the concise model above is canonical.
pub type DrawingPageBackgroundSize = BackgroundSize;
pub type DrawingPageColor = Color;
pub type DrawingPageDuration = Duration;
pub type DrawingPageFill = Fill;
pub type DrawingPageFillRule = FillRule;
pub type DrawingPageImageRefPoint = ImageRefPoint;
pub type DrawingPageLengthOrPercent = LengthOrPercent;
pub type DrawingPageNonNegativeInteger = NonNegativeInteger;
pub type DrawingPagePercent = Percent;
pub type DrawingPageRepeat = Repeat;
pub type DrawingPageSound = Sound;
pub type DrawingPageSoundShow = SoundShow;
pub type DrawingPageStyle = Style;
pub type DrawingPageStyleNameRef = StyleNameRef;
pub type DrawingPageStyleProperties = StyleProperties;
pub type DrawingPageStyleSet = Styles;
pub type DrawingPageTileDirection = TileDirection;
pub type DrawingPageTileRepeatOffset = TileRepeatOffset;
pub type DrawingPageTransitionDirection = TransitionDirection;
pub type DrawingPageTransitionSpeed = TransitionSpeed;
pub type DrawingPageTransitionStyle = TransitionStyle;
pub type DrawingPageTransitionType = TransitionType;
pub type DrawingPageVisibility = Visibility;
