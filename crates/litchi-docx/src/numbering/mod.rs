//! Layered WordprocessingML numbering model and XML codec.

mod codec;
mod model;

#[cfg(test)]
mod tests;

pub use codec::parse_numbering;
pub use model::{
    Collection, Definition, Format, Instance, Level, MultiLevel, Override, Paragraph,
    ParseFormatError, ParseMultiLevelError, PictureBullet, Restart, Suffix,
};
