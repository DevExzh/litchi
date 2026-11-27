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
