//! Dependency-free paragraph formatting vocabulary shared by iWork formats.

pub mod direction;
pub mod flow;
pub mod list;
pub mod tabs;

pub use direction::WritingDirection;
pub use flow::{Flow, Hyphenation};
pub use list::{
    ParagraphList, ParagraphListBullet, ParagraphListBulletBaselineOffset,
    ParagraphListBulletGeometry, ParagraphListBulletScale, ParagraphListIndentation,
    ParagraphListLabelColor, ParagraphListLabelIndent, ParagraphListLevel,
    ParagraphListLevelPlacement, ParagraphListNumberFormat, ParagraphListNumberPunctuation,
    ParagraphListNumberScale, ParagraphListNumberSequence, ParagraphListNumberTiering,
    ParagraphListNumbering, ParagraphListPlacement, ParagraphListStart, ParagraphListTextGap,
};
pub use tabs::{Alignment, DecimalCharacter, DefaultInterval, Leader, Position, Stop, Stops};
