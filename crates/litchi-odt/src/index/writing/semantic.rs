#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum TextIndexScope {
    Document,
    Chapter,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum TextIndexChapterDisplay {
    Name,
    Number,
    NumberAndName,
    PlainNumber,
    PlainNumberAndName,
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub enum TextIndexTabStop {
    Right {
        leader: Option<char>,
        style_name: Option<String>,
    },
    Left {
        position: String,
        leader: Option<char>,
        style_name: Option<String>,
    },
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub enum TextIndexEntryToken {
    Chapter {
        style_name: Option<String>,
        display: Option<TextIndexChapterDisplay>,
        outline_level: Option<u16>,
    },
    PageNumber {
        style_name: Option<String>,
    },
    Text {
        style_name: Option<String>,
    },
    Span {
        style_name: Option<String>,
        text: String,
    },
    TabStop(TextIndexTabStop),
    LinkStart {
        style_name: Option<String>,
    },
    LinkEnd {
        style_name: Option<String>,
    },
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct TextIndexEntryTemplate {
    pub outline_level: u16,
    pub style_name: String,
    pub tokens: Vec<TextIndexEntryToken>,
}

/// The single, non-outline-level template shared by illustration, table, and object indexes.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct TextIndexSimpleEntryTemplate {
    pub style_name: String,
    pub tokens: Vec<TextIndexEntryToken>,
}

impl TextIndexSimpleEntryTemplate {
    pub fn new(style_name: impl Into<String>) -> Self {
        Self {
            style_name: style_name.into(),
            tokens: Vec::new(),
        }
    }

    pub fn push(&mut self, token: TextIndexEntryToken) -> &mut Self {
        self.tokens.push(token);
        self
    }
}

impl TextIndexEntryTemplate {
    pub fn new(outline_level: u16, style_name: impl Into<String>) -> Self {
        Self {
            outline_level,
            style_name: style_name.into(),
            tokens: Vec::new(),
        }
    }

    pub fn push(&mut self, token: TextIndexEntryToken) -> &mut Self {
        self.tokens.push(token);
        self
    }
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct TextIndexTitleTemplate {
    pub text: String,
    pub style_name: Option<String>,
}

impl TextIndexTitleTemplate {
    pub fn new(text: impl Into<String>) -> Self {
        Self {
            text: text.into(),
            style_name: None,
        }
    }

    pub fn with_style_name(mut self, style_name: impl Into<String>) -> Self {
        self.style_name = Some(style_name.into());
        self
    }
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct TextIndexSourceStyles {
    pub outline_level: u16,
    pub style_names: Vec<String>,
}

impl TextIndexSourceStyles {
    pub fn new(outline_level: u16, style_names: Vec<String>) -> Self {
        Self {
            outline_level,
            style_names,
        }
    }
}

#[derive(Debug, Clone, PartialEq, Eq, Default)]
pub struct TableOfContentsSource {
    pub outline_level: Option<u16>,
    pub use_outline_level: Option<bool>,
    pub use_index_marks: Option<bool>,
    pub use_index_source_styles: Option<bool>,
    pub scope: Option<TextIndexScope>,
    pub relative_tab_stop_position: Option<bool>,
    pub title_template: Option<TextIndexTitleTemplate>,
    pub entry_templates: Vec<TextIndexEntryTemplate>,
    pub source_styles: Vec<TextIndexSourceStyles>,
}

impl TableOfContentsSource {
    pub fn new() -> Self {
        Self::default()
    }

    pub fn push_entry_template(&mut self, template: TextIndexEntryTemplate) -> &mut Self {
        self.entry_templates.push(template);
        self
    }

    pub fn push_source_styles(&mut self, styles: TextIndexSourceStyles) -> &mut Self {
        self.source_styles.push(styles);
        self
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum TextIndexCaptionSequenceFormat {
    Text,
    CategoryAndValue,
    Caption,
}

/// Source policy shared verbatim by ODF illustration and table indexes.
#[derive(Debug, Clone, PartialEq, Eq, Default)]
pub struct IllustrationIndexSource {
    pub scope: Option<TextIndexScope>,
    pub relative_tab_stop_position: Option<bool>,
    pub use_caption: Option<bool>,
    pub caption_sequence_name: Option<String>,
    pub caption_sequence_format: Option<TextIndexCaptionSequenceFormat>,
    pub title_template: Option<TextIndexTitleTemplate>,
    pub entry_template: Option<TextIndexSimpleEntryTemplate>,
}

impl IllustrationIndexSource {
    pub fn new() -> Self {
        Self::default()
    }
}

#[derive(Debug, Clone, PartialEq, Eq, Default)]
pub struct ObjectIndexSource {
    pub scope: Option<TextIndexScope>,
    pub relative_tab_stop_position: Option<bool>,
    pub use_spreadsheet_objects: Option<bool>,
    pub use_math_objects: Option<bool>,
    pub use_draw_objects: Option<bool>,
    pub use_chart_objects: Option<bool>,
    pub use_other_objects: Option<bool>,
    pub title_template: Option<TextIndexTitleTemplate>,
    pub entry_template: Option<TextIndexSimpleEntryTemplate>,
}

impl ObjectIndexSource {
    pub fn new() -> Self {
        Self::default()
    }
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct UserIndexSource {
    pub index_name: String,
    pub scope: Option<TextIndexScope>,
    pub relative_tab_stop_position: Option<bool>,
    pub use_index_marks: Option<bool>,
    pub use_index_source_styles: Option<bool>,
    pub use_graphics: Option<bool>,
    pub use_tables: Option<bool>,
    pub use_floating_frames: Option<bool>,
    pub use_objects: Option<bool>,
    pub copy_outline_levels: Option<bool>,
    pub title_template: Option<TextIndexTitleTemplate>,
    pub entry_templates: Vec<TextIndexEntryTemplate>,
    pub source_styles: Vec<TextIndexSourceStyles>,
}

impl UserIndexSource {
    pub fn new(index_name: impl Into<String>) -> Self {
        Self {
            index_name: index_name.into(),
            scope: None,
            relative_tab_stop_position: None,
            use_index_marks: None,
            use_index_source_styles: None,
            use_graphics: None,
            use_tables: None,
            use_floating_frames: None,
            use_objects: None,
            copy_outline_levels: None,
            title_template: None,
            entry_templates: Vec::new(),
            source_styles: Vec::new(),
        }
    }

    pub fn push_entry_template(&mut self, template: TextIndexEntryTemplate) -> &mut Self {
        self.entry_templates.push(template);
        self
    }

    pub fn push_source_styles(&mut self, styles: TextIndexSourceStyles) -> &mut Self {
        self.source_styles.push(styles);
        self
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum TextAlphabeticalIndexLevel {
    Level1,
    Level2,
    Level3,
    Separator,
}

impl TextAlphabeticalIndexLevel {
    pub(super) fn as_str(self) -> &'static str {
        match self {
            Self::Level1 => "1",
            Self::Level2 => "2",
            Self::Level3 => "3",
            Self::Separator => "separator",
        }
    }
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct TextAlphabeticalIndexEntryTemplate {
    pub level: TextAlphabeticalIndexLevel,
    pub style_name: String,
    pub tokens: Vec<TextIndexEntryToken>,
}

impl TextAlphabeticalIndexEntryTemplate {
    pub fn new(level: TextAlphabeticalIndexLevel, style_name: impl Into<String>) -> Self {
        Self {
            level,
            style_name: style_name.into(),
            tokens: Vec::new(),
        }
    }

    pub fn push(&mut self, token: TextIndexEntryToken) -> &mut Self {
        self.tokens.push(token);
        self
    }
}

#[derive(Debug, Clone, PartialEq, Eq, Default)]
pub struct AlphabeticalIndexSource {
    pub scope: Option<TextIndexScope>,
    pub relative_tab_stop_position: Option<bool>,
    pub ignore_case: Option<bool>,
    pub main_entry_style_name: Option<String>,
    pub alphabetical_separators: Option<bool>,
    pub combine_entries: Option<bool>,
    pub combine_entries_with_dash: Option<bool>,
    pub combine_entries_with_pp: Option<bool>,
    pub use_keys_as_entries: Option<bool>,
    pub capitalize_entries: Option<bool>,
    pub comma_separated: Option<bool>,
    pub language: Option<String>,
    pub country: Option<String>,
    pub script: Option<String>,
    pub rfc_language_tag: Option<String>,
    pub sort_algorithm: Option<String>,
    pub title_template: Option<TextIndexTitleTemplate>,
    pub entry_templates: Vec<TextAlphabeticalIndexEntryTemplate>,
}

impl AlphabeticalIndexSource {
    pub fn new() -> Self {
        Self::default()
    }

    pub fn push_entry_template(
        &mut self,
        template: TextAlphabeticalIndexEntryTemplate,
    ) -> &mut Self {
        self.entry_templates.push(template);
        self
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum TextBibliographyType {
    Article,
    Book,
    Booklet,
    Conference,
    Custom1,
    Custom2,
    Custom3,
    Custom4,
    Custom5,
    Email,
    InBook,
    InCollection,
    InProceedings,
    Journal,
    Manual,
    MastersThesis,
    Misc,
    PhdThesis,
    Proceedings,
    TechReport,
    Unpublished,
    Www,
}

impl TextBibliographyType {
    pub(crate) fn as_str(self) -> &'static str {
        match self {
            Self::Article => "article",
            Self::Book => "book",
            Self::Booklet => "booklet",
            Self::Conference => "conference",
            Self::Custom1 => "custom1",
            Self::Custom2 => "custom2",
            Self::Custom3 => "custom3",
            Self::Custom4 => "custom4",
            Self::Custom5 => "custom5",
            Self::Email => "email",
            Self::InBook => "inbook",
            Self::InCollection => "incollection",
            Self::InProceedings => "inproceedings",
            Self::Journal => "journal",
            Self::Manual => "manual",
            Self::MastersThesis => "mastersthesis",
            Self::Misc => "misc",
            Self::PhdThesis => "phdthesis",
            Self::Proceedings => "proceedings",
            Self::TechReport => "techreport",
            Self::Unpublished => "unpublished",
            Self::Www => "www",
        }
    }
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub enum TextBibliographyEntryToken {
    Field {
        field: crate::bibliography_configuration::Field,
        style_name: Option<String>,
    },
    Span {
        style_name: Option<String>,
        text: String,
    },
    TabStop(TextIndexTabStop),
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct TextBibliographyEntryTemplate {
    pub bibliography_type: TextBibliographyType,
    pub style_name: String,
    pub tokens: Vec<TextBibliographyEntryToken>,
}

impl TextBibliographyEntryTemplate {
    pub fn new(bibliography_type: TextBibliographyType, style_name: impl Into<String>) -> Self {
        Self {
            bibliography_type,
            style_name: style_name.into(),
            tokens: Vec::new(),
        }
    }

    pub fn push(&mut self, token: TextBibliographyEntryToken) -> &mut Self {
        self.tokens.push(token);
        self
    }
}

#[derive(Debug, Clone, PartialEq, Eq, Default)]
pub struct BibliographyIndexSource {
    pub title_template: Option<TextIndexTitleTemplate>,
    pub entry_templates: Vec<TextBibliographyEntryTemplate>,
}

impl BibliographyIndexSource {
    pub fn new() -> Self {
        Self::default()
    }

    pub fn push_entry_template(&mut self, template: TextBibliographyEntryTemplate) -> &mut Self {
        self.entry_templates.push(template);
        self
    }
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct TextIndexBodyTitle {
    pub name: String,
    pub section_style_name: Option<String>,
    pub paragraph_style_name: Option<String>,
    pub text: String,
}

impl TextIndexBodyTitle {
    pub fn new(name: impl Into<String>, text: impl Into<String>) -> Self {
        Self {
            name: name.into(),
            section_style_name: None,
            paragraph_style_name: None,
            text: text.into(),
        }
    }
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct TextIndexBodyParagraph {
    pub style_name: Option<String>,
    pub text: String,
}

impl TextIndexBodyParagraph {
    pub fn new(text: impl Into<String>) -> Self {
        Self {
            style_name: None,
            text: text.into(),
        }
    }
}

#[derive(Debug, Clone, PartialEq, Eq, Default)]
pub struct TextIndexBody {
    pub title: Option<TextIndexBodyTitle>,
    pub paragraphs: Vec<TextIndexBodyParagraph>,
}

impl TextIndexBody {
    pub fn new() -> Self {
        Self::default()
    }

    pub fn push_paragraph(&mut self, paragraph: TextIndexBodyParagraph) -> &mut Self {
        self.paragraphs.push(paragraph);
        self
    }
}
