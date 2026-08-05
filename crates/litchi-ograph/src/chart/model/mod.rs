//! Host-neutral semantic chart model facade.

mod aggregate;
mod cache;
mod context;
mod groups;
mod inventory;
mod series;

pub use aggregate::Chart;
pub use cache::{Cache, Value, ValueRef, XlValue};
pub use context::{Context, Count, GroupId, Order, Props, Rect};
pub use groups::{Family, Group};
pub use inventory::{Edit, Label, Legend, Raw};
pub use series::{Ai, Binding, CellRef, DataKind, Link, Owner, Role, RowCol, Series, Source};

pub(in crate::chart) use aggregate::{cache_dimensions, dimensions_cover};
pub(in crate::chart) use inventory::Origin;

#[cfg(test)]
mod tests;
