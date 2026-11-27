//! Typed `ChartEx` graph conversion and semantic validation.

mod model;
mod validation;

pub(super) use validation::parse_data_graph;
