//! Dependency-free paragraph formatting vocabulary shared by iWork formats.

pub mod border;
pub mod direction;
pub mod drop_cap;
pub mod flow;
pub mod format;
pub mod list;
pub mod style;
pub mod tabs;

pub use direction::WritingDirection;
pub use flow::{Flow, Hyphenation};
pub use format::{
    Border, Borders, Format, IndentPoints, Indents, LineSpacing, LineSpacingMultiple,
    LineSpacingPoints, Spacing, SpacingPoints,
};
pub use tabs::{Alignment, DecimalCharacter, DefaultInterval, Leader, Position, Stop, Stops};
