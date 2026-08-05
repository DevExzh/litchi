//! Dependency-free paragraph formatting vocabulary shared by iWork formats.

pub mod direction;
pub mod flow;
pub mod tabs;

pub use direction::WritingDirection;
pub use flow::{Flow, Hyphenation};
pub use tabs::{Alignment, DecimalCharacter, DefaultInterval, Leader, Position, Stop, Stops};
