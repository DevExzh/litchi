//! Body text, local formatting, styles, and language behavior.

pub(crate) mod border;
pub(crate) mod character_positioning;
pub(crate) mod document_default_formatting;
pub(crate) mod hyphenation;
pub(crate) mod kinsoku;
pub(crate) mod language;
pub(crate) mod latent_style;
pub(crate) mod paragraph_group;
pub(crate) mod style_list_filter;
pub(crate) mod stylesheet;

pub(crate) use crate::model::types;

pub use crate::api::{
    Break, Format, Inline, Inlines, Paragraph, ParagraphFormat, Paragraphs, Run, Runs, Story,
};
pub use border::{
    Border, BorderStyle, Borders, CharacterBorder, CharacterBorderStyle, CharacterShading, Shading,
    ShadingPattern, TabAlignment, TabLeader, TabStop, TabStops,
};
pub use character_positioning::{
    CharacterBaseline as Baseline, CharacterExpansion as Expansion,
    CharacterPositioning as Positioning,
};
pub use types::{
    Alignment, AnimatedTextEffect, AssociatedCharacterFormatting, CharacterGrid, CharacterType,
    EmphasisMark, FitText, Indentation, ParagraphDropCap, ParagraphDropCapKind,
    ParagraphFontAlignment, ParagraphLineBreaking, ParagraphLogicalIndentation,
    ParagraphSpacingPolicy, ParagraphWrapping, RevisionMetadata, Spacing,
    TextDirection as Direction, UnderlineStyle as Underline,
};
