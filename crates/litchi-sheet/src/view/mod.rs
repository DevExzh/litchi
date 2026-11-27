//! Format-neutral worksheet view state.

mod model;

pub use model::{
    Color, Display, Error, Mode, Pane, Position, Scale, Selection, Split, State, View, Window, Zoom,
};

#[cfg(test)]
mod tests;
