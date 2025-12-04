/// Parts for PowerPoint presentation documents.
///
/// This module contains wrapper types for different XML parts in a .pptx package,
/// following the structure of the python-pptx library.
pub mod chart;
pub mod comment;
pub mod presentation;
pub mod slide;
pub mod theme;

pub use chart::{ChartInfo, ChartPart, ChartType};
pub use comment::{
    Comment, CommentAuthor, CommentAuthorsPart, CommentsPart, generate_comment_authors_xml,
    generate_comments_xml,
};
pub use presentation::PresentationPart;
pub use slide::{SlideLayoutPart, SlideMasterPart, SlidePart};
pub use theme::{Theme, ThemeColor, ThemeFont, ThemePart};
