//! Facade for the namespace-aware ODP XML parser.

mod codec;
mod model;
#[cfg(test)]
mod tests;

pub(crate) use codec::Parser;
