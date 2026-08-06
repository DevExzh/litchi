//! Shared `OpenDocument` chart content and borrowed semantic views.
//!
//! The chart reader retains the complete `chart:chart` subtree, including
//! unknown namespaces and extension elements, while the view layer interprets
//! only the small set of standard semantic values needed by all ODF families.
//! Package ownership remains in the concrete family crates; the validated
//! chart-content authoring owner is shared so host families do not depend on
//! the standalone chart family as a peer.

pub mod authoring;
pub mod axis;
pub mod grid;
pub mod legend;
pub mod plot_area;
pub mod reader;
pub mod view;

pub use axis::Dimension;
pub use grid::Class;
pub use legend::Position;
pub use plot_area::Labels;
pub use reader::{Attribute, Element, Kind, read};
pub use view::{Axis, DataPoint, Grid, Legend, PlotArea, Series};
