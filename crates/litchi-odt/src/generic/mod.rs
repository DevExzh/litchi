//! Format-neutral access to any packaged `OpenDocument` document family.

mod codec;
mod flat;
mod model;
mod package;

pub(crate) use flat::{ChargedXml, FlatMutationBudget, MemoryLease, allocate_xml};

#[cfg(test)]
mod tests;

pub use model::{Family, FlatDocument, Package};
