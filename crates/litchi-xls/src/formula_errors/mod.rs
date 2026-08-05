//! Layered BIFF8 formula error-checking shared-feature owner (`FeatHdr` + `Feat`).
//!
//! The public formula-error values remain available through the crate facade,
//! while semantic values, BIFF codecs, and regression coverage live in their
//! respective layers.

mod codec;
mod model;

#[cfg(test)]
mod tests;

pub(crate) use codec::FormulaErrorCollector;
pub use model::{
    Checks as FormulaErrorChecks, Feature as FormulaErrorFeature, Header as FormulaErrorHeader,
    Range as FormulaErrorRange,
};
