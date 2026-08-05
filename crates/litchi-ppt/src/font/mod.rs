//! Semantic PowerPoint font collection views and bounded record decoding.

mod codec;
mod model;

#[cfg(test)]
mod tests;

pub use model::{
    EmbeddedPowerPointFont, PowerPointFont, PowerPointFontCollection, PowerPointFontCollections,
    PowerPointFontEmbeddingFlags,
};
