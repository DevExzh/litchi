//! Typed, schema-directed writing and byte-preserving mutation of text indexes.

use super::{
    TextIndex, TextIndexAttribute, TextIndexContent, TextIndexElement, TextIndexKind,
    parse_text_indexes,
};
use litchi_core::{Error, Result};
use quick_xml::events::Event;
use quick_xml::name::{Namespace, ResolveResult};
use quick_xml::reader::NsReader;
use std::collections::{BTreeMap, BTreeSet};

const TEXT: &str = "urn:oasis:names:tc:opendocument:xmlns:text:1.0";
const STYLE: &str = "urn:oasis:names:tc:opendocument:xmlns:style:1.0";
const FO: &str = "urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0";
const OFFICE: &[u8] = b"urn:oasis:names:tc:opendocument:xmlns:office:1.0";
const MAX_XML_BYTES: usize = 64 * 1024 * 1024;
const MAX_FRAGMENT_BYTES: usize = 16 * 1024 * 1024;
const MAX_TEMPLATES: usize = 4_096;
const MAX_TOKENS: usize = 65_536;
const MAX_BODY_PARAGRAPHS: usize = 65_536;
const MAX_DEPTH: usize = 4_096;

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

pub type TableIndexSource = IllustrationIndexSource;

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
    fn as_str(self) -> &'static str {
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
        field: crate::BibliographyField,
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

impl TextIndex {
    /// Create ODF index markup from caller-provided templates and cached body.
    ///
    /// This does not scan headings, evaluate fields, paginate, or refresh the cache.
    pub fn table_of_contents(
        name: impl Into<String>,
        source: TableOfContentsSource,
        body: TextIndexBody,
    ) -> Result<Self> {
        let name = name.into();
        required(&name, "text index name")?;
        if source.entry_templates.len() > MAX_TEMPLATES
            || source.source_styles.len() > MAX_TEMPLATES
            || body.paragraphs.len() > MAX_BODY_PARAGRAPHS
        {
            return invalid("text index exceeds configured resource limits");
        }

        let mut source_attributes = Vec::new();
        optional_positive(
            &mut source_attributes,
            "outline-level",
            source.outline_level,
        )?;
        optional_bool(
            &mut source_attributes,
            "use-outline-level",
            source.use_outline_level,
        );
        optional_bool(
            &mut source_attributes,
            "use-index-marks",
            source.use_index_marks,
        );
        optional_bool(
            &mut source_attributes,
            "use-index-source-styles",
            source.use_index_source_styles,
        );
        if let Some(scope) = source.scope {
            source_attributes.push(attr(
                TEXT,
                "index-scope",
                match scope {
                    TextIndexScope::Document => "document",
                    TextIndexScope::Chapter => "chapter",
                },
            ));
        }
        optional_bool(
            &mut source_attributes,
            "relative-tab-stop-position",
            source.relative_tab_stop_position,
        );

        let mut source_content = Vec::new();
        if let Some(title) = source.title_template {
            let mut attributes = Vec::new();
            optional_name(&mut attributes, "style-name", title.style_name)?;
            source_content.push(element_content(element(
                TEXT,
                "index-title-template",
                attributes,
                vec![TextIndexContent::Text(title.text)],
            )));
        }

        let mut token_count = 0usize;
        for template in source.entry_templates {
            positive(template.outline_level, "entry template outline level")?;
            required(&template.style_name, "entry template style name")?;
            token_count = token_count
                .checked_add(template.tokens.len())
                .ok_or_else(|| {
                    Error::InvalidFormat("text index token count overflow".to_string())
                })?;
            if token_count > MAX_TOKENS {
                return invalid("text index contains too many entry tokens");
            }
            validate_links(&template.tokens)?;
            let mut tokens = Vec::with_capacity(template.tokens.len());
            for token in template.tokens {
                tokens.push(element_content(token_element(token)?));
            }
            source_content.push(element_content(element(
                TEXT,
                "table-of-content-entry-template",
                vec![
                    attr(TEXT, "outline-level", template.outline_level.to_string()),
                    attr(TEXT, "style-name", template.style_name),
                ],
                tokens,
            )));
        }

        for styles in source.source_styles {
            positive(styles.outline_level, "source styles outline level")?;
            let mut children = Vec::with_capacity(styles.style_names.len());
            for style_name in styles.style_names {
                required(&style_name, "source style name")?;
                children.push(element_content(element(
                    TEXT,
                    "index-source-style",
                    vec![attr(TEXT, "style-name", style_name)],
                    Vec::new(),
                )));
            }
            source_content.push(element_content(element(
                TEXT,
                "index-source-styles",
                vec![attr(
                    TEXT,
                    "outline-level",
                    styles.outline_level.to_string(),
                )],
                children,
            )));
        }

        let mut body_content = Vec::new();
        if let Some(title) = body.title {
            required(&title.name, "index body title name")?;
            let mut title_attrs = vec![attr(TEXT, "name", title.name)];
            optional_name(&mut title_attrs, "style-name", title.section_style_name)?;
            let mut paragraph_attrs = Vec::new();
            optional_name(
                &mut paragraph_attrs,
                "style-name",
                title.paragraph_style_name,
            )?;
            body_content.push(element_content(element(
                TEXT,
                "index-title",
                title_attrs,
                vec![element_content(element(
                    TEXT,
                    "p",
                    paragraph_attrs,
                    vec![TextIndexContent::Text(title.text)],
                ))],
            )));
        }
        for paragraph in body.paragraphs {
            let mut attributes = Vec::new();
            optional_name(&mut attributes, "style-name", paragraph.style_name)?;
            body_content.push(element_content(element(
                TEXT,
                "p",
                attributes,
                vec![TextIndexContent::Text(paragraph.text)],
            )));
        }

        let index = Self {
            kind: TextIndexKind::TableOfContents,
            root: element(
                TEXT,
                "table-of-content",
                vec![attr(TEXT, "name", name)],
                vec![
                    element_content(element(
                        TEXT,
                        "table-of-content-source",
                        source_attributes,
                        source_content,
                    )),
                    element_content(element(TEXT, "index-body", Vec::new(), body_content)),
                ],
            ),
        };
        validate_index(&index)?;
        Ok(index)
    }

    pub fn illustration_index(
        name: impl Into<String>,
        source: IllustrationIndexSource,
        body: TextIndexBody,
    ) -> Result<Self> {
        caption_index(
            TextIndexKind::Illustration,
            "illustration-index",
            "illustration-index-source",
            "illustration-index-entry-template",
            name.into(),
            source,
            body,
        )
    }

    pub fn table_index(
        name: impl Into<String>,
        source: TableIndexSource,
        body: TextIndexBody,
    ) -> Result<Self> {
        caption_index(
            TextIndexKind::Table,
            "table-index",
            "table-index-source",
            "table-index-entry-template",
            name.into(),
            source,
            body,
        )
    }

    pub fn object_index(
        name: impl Into<String>,
        source: ObjectIndexSource,
        body: TextIndexBody,
    ) -> Result<Self> {
        let name = name.into();
        required(&name, "text index name")?;
        let mut attributes =
            common_source_attributes(source.scope, source.relative_tab_stop_position);
        optional_bool(
            &mut attributes,
            "use-spreadsheet-objects",
            source.use_spreadsheet_objects,
        );
        optional_bool(&mut attributes, "use-math-objects", source.use_math_objects);
        optional_bool(&mut attributes, "use-draw-objects", source.use_draw_objects);
        optional_bool(
            &mut attributes,
            "use-chart-objects",
            source.use_chart_objects,
        );
        optional_bool(
            &mut attributes,
            "use-other-objects",
            source.use_other_objects,
        );
        let content = single_template_content(
            source.title_template,
            source.entry_template,
            "object-index-entry-template",
        )?;
        finish_typed_index(
            TextIndexKind::Object,
            "object-index",
            "object-index-source",
            name,
            attributes,
            content,
            body,
        )
    }

    pub fn user_index(
        name: impl Into<String>,
        source: UserIndexSource,
        body: TextIndexBody,
    ) -> Result<Self> {
        let name = name.into();
        required(&name, "text index name")?;
        if source.entry_templates.len() > MAX_TEMPLATES
            || source.source_styles.len() > MAX_TEMPLATES
        {
            return invalid("user index exceeds configured template limits");
        }
        let mut attributes =
            common_source_attributes(source.scope, source.relative_tab_stop_position);
        attributes.push(attr(TEXT, "index-name", source.index_name));
        optional_bool(&mut attributes, "use-index-marks", source.use_index_marks);
        optional_bool(
            &mut attributes,
            "use-index-source-styles",
            source.use_index_source_styles,
        );
        optional_bool(&mut attributes, "use-graphics", source.use_graphics);
        optional_bool(&mut attributes, "use-tables", source.use_tables);
        optional_bool(
            &mut attributes,
            "use-floating-frames",
            source.use_floating_frames,
        );
        optional_bool(&mut attributes, "use-objects", source.use_objects);
        optional_bool(
            &mut attributes,
            "copy-outline-levels",
            source.copy_outline_levels,
        );
        let mut content = Vec::new();
        if let Some(title) = source.title_template {
            content.push(element_content(title_template_element(title)?));
        }
        let mut token_count = 0usize;
        for template in source.entry_templates {
            positive(template.outline_level, "user index entry outline level")?;
            required(&template.style_name, "user index entry style name")?;
            token_count = token_count
                .checked_add(template.tokens.len())
                .ok_or_else(|| {
                    Error::InvalidFormat("user index token count overflow".to_string())
                })?;
            if token_count > MAX_TOKENS {
                return invalid("user index contains too many entry tokens");
            }
            validate_links(&template.tokens)?;
            content.push(element_content(element(
                TEXT,
                "user-index-entry-template",
                vec![
                    attr(TEXT, "outline-level", template.outline_level.to_string()),
                    attr(TEXT, "style-name", template.style_name),
                ],
                template
                    .tokens
                    .into_iter()
                    .map(token_element)
                    .collect::<Result<Vec<_>>>()?
                    .into_iter()
                    .map(element_content)
                    .collect(),
            )));
        }
        for styles in source.source_styles {
            content.push(element_content(source_styles_element(styles)?));
        }
        finish_typed_index(
            TextIndexKind::User,
            "user-index",
            "user-index-source",
            name,
            attributes,
            content,
            body,
        )
    }

    pub fn alphabetical_index(
        name: impl Into<String>,
        source: AlphabeticalIndexSource,
        body: TextIndexBody,
    ) -> Result<Self> {
        let name = name.into();
        required(&name, "text index name")?;
        if source.entry_templates.len() > MAX_TEMPLATES {
            return invalid("alphabetical index has too many templates");
        }
        let mut attributes =
            common_source_attributes(source.scope, source.relative_tab_stop_position);
        optional_bool(&mut attributes, "ignore-case", source.ignore_case);
        optional_name(
            &mut attributes,
            "main-entry-style-name",
            source.main_entry_style_name,
        )?;
        optional_bool(
            &mut attributes,
            "alphabetical-separators",
            source.alphabetical_separators,
        );
        optional_bool(&mut attributes, "combine-entries", source.combine_entries);
        optional_bool(
            &mut attributes,
            "combine-entries-with-dash",
            source.combine_entries_with_dash,
        );
        optional_bool(
            &mut attributes,
            "combine-entries-with-pp",
            source.combine_entries_with_pp,
        );
        optional_bool(
            &mut attributes,
            "use-keys-as-entries",
            source.use_keys_as_entries,
        );
        optional_bool(
            &mut attributes,
            "capitalize-entries",
            source.capitalize_entries,
        );
        optional_bool(&mut attributes, "comma-separated", source.comma_separated);
        optional_locale_attribute(
            &mut attributes,
            FO,
            "language",
            source.language,
            LocaleLexical::LanguageCode,
        )?;
        optional_locale_attribute(
            &mut attributes,
            FO,
            "country",
            source.country,
            LocaleLexical::CountryOrScript,
        )?;
        optional_locale_attribute(
            &mut attributes,
            FO,
            "script",
            source.script,
            LocaleLexical::CountryOrScript,
        )?;
        optional_locale_attribute(
            &mut attributes,
            STYLE,
            "rfc-language-tag",
            source.rfc_language_tag,
            LocaleLexical::LanguageTag,
        )?;
        if let Some(sort_algorithm) = source.sort_algorithm {
            attributes.push(attr(TEXT, "sort-algorithm", sort_algorithm));
        }
        let mut content = Vec::new();
        if let Some(title) = source.title_template {
            content.push(element_content(title_template_element(title)?));
        }
        let mut token_count = 0usize;
        for template in source.entry_templates {
            required(&template.style_name, "alphabetical index entry style name")?;
            token_count = token_count
                .checked_add(template.tokens.len())
                .ok_or_else(|| {
                    Error::InvalidFormat("alphabetical index token count overflow".to_string())
                })?;
            if token_count > MAX_TOKENS {
                return invalid("alphabetical index contains too many tokens");
            }
            if template.tokens.iter().any(|token| {
                matches!(
                    token,
                    TextIndexEntryToken::LinkStart { .. } | TextIndexEntryToken::LinkEnd { .. }
                )
            }) {
                return invalid("alphabetical index templates do not permit link tokens");
            }
            content.push(element_content(element(
                TEXT,
                "alphabetical-index-entry-template",
                vec![
                    attr(TEXT, "outline-level", template.level.as_str()),
                    attr(TEXT, "style-name", template.style_name),
                ],
                template
                    .tokens
                    .into_iter()
                    .map(token_element)
                    .collect::<Result<Vec<_>>>()?
                    .into_iter()
                    .map(element_content)
                    .collect(),
            )));
        }
        finish_typed_index(
            TextIndexKind::Alphabetical,
            "alphabetical-index",
            "alphabetical-index-source",
            name,
            attributes,
            content,
            body,
        )
    }

    pub fn bibliography(
        name: impl Into<String>,
        source: BibliographyIndexSource,
        body: TextIndexBody,
    ) -> Result<Self> {
        let name = name.into();
        required(&name, "text index name")?;
        if source.entry_templates.len() > MAX_TEMPLATES {
            return invalid("bibliography has too many templates");
        }
        let mut content = Vec::new();
        if let Some(title) = source.title_template {
            content.push(element_content(title_template_element(title)?));
        }
        let mut token_count = 0usize;
        for template in source.entry_templates {
            required(&template.style_name, "bibliography entry style name")?;
            token_count = token_count
                .checked_add(template.tokens.len())
                .ok_or_else(|| {
                    Error::InvalidFormat("bibliography token count overflow".to_string())
                })?;
            if token_count > MAX_TOKENS {
                return invalid("bibliography contains too many tokens");
            }
            let tokens = template
                .tokens
                .into_iter()
                .map(bibliography_token_element)
                .collect::<Result<Vec<_>>>()?
                .into_iter()
                .map(element_content)
                .collect();
            content.push(element_content(element(
                TEXT,
                "bibliography-entry-template",
                vec![
                    attr(
                        TEXT,
                        "bibliography-type",
                        template.bibliography_type.as_str(),
                    ),
                    attr(TEXT, "style-name", template.style_name),
                ],
                tokens,
            )));
        }
        finish_typed_index(
            TextIndexKind::Bibliography,
            "bibliography",
            "bibliography-source",
            name,
            Vec::new(),
            content,
            body,
        )
    }

    pub fn with_protected(mut self, protected: bool) -> Self {
        set_attr(
            &mut self.root.attributes,
            TEXT,
            "protected",
            protected.to_string(),
        );
        self
    }

    pub fn with_style_name(mut self, style_name: impl Into<String>) -> Result<Self> {
        let style_name = style_name.into();
        required(&style_name, "text index style name")?;
        set_attr(&mut self.root.attributes, TEXT, "style-name", style_name);
        Ok(self)
    }

    /// Serialize one standalone, canonical XML fragment with owned declarations.
    pub fn to_xml_fragment(&self) -> Result<String> {
        validate_index(self)?;
        let mut namespaces = BTreeSet::new();
        collect_namespaces(&self.root, &mut namespaces);
        namespaces.insert(TEXT.to_string());
        let prefixes = namespace_prefixes(&namespaces);
        let mut output = String::new();
        write_element(&self.root, &prefixes, true, &mut output)?;
        if output.len() > MAX_FRAGMENT_BYTES {
            return invalid("serialized text index exceeds 16 MiB");
        }
        Ok(output)
    }
}

fn bibliography_token_element(token: TextBibliographyEntryToken) -> Result<TextIndexElement> {
    match token {
        TextBibliographyEntryToken::Field { field, style_name } => {
            let mut attributes = vec![attr(TEXT, "bibliography-data-field", field.as_str())];
            optional_name(&mut attributes, "style-name", style_name)?;
            Ok(element(
                TEXT,
                "index-entry-bibliography",
                attributes,
                Vec::new(),
            ))
        },
        TextBibliographyEntryToken::Span { style_name, text } => {
            token_element(TextIndexEntryToken::Span { style_name, text })
        },
        TextBibliographyEntryToken::TabStop(tab) => {
            token_element(TextIndexEntryToken::TabStop(tab))
        },
    }
}

#[derive(Clone, Copy)]
enum LocaleLexical {
    LanguageCode,
    CountryOrScript,
    LanguageTag,
}

fn optional_locale_attribute(
    attributes: &mut Vec<TextIndexAttribute>,
    namespace: &str,
    name: &str,
    value: Option<String>,
    lexical: LocaleLexical,
) -> Result<()> {
    if let Some(value) = value {
        validate_locale(&value, lexical, name)?;
        attributes.push(attr(namespace, name, value));
    }
    Ok(())
}

fn validate_locale(value: &str, lexical: LocaleLexical, context: &str) -> Result<()> {
    let valid = match lexical {
        LocaleLexical::LanguageCode => {
            (1..=8).contains(&value.len()) && value.bytes().all(|byte| byte.is_ascii_alphabetic())
        },
        LocaleLexical::CountryOrScript => {
            (1..=8).contains(&value.len()) && value.bytes().all(|byte| byte.is_ascii_alphanumeric())
        },
        LocaleLexical::LanguageTag => value.split('-').all(|part| {
            (1..=8).contains(&part.len()) && part.bytes().all(|byte| byte.is_ascii_alphanumeric())
        }),
    };
    if valid {
        Ok(())
    } else {
        invalid(format!("invalid {context} lexical {value:?}"))
    }
}

fn caption_index(
    kind: TextIndexKind,
    root_name: &str,
    source_name: &str,
    entry_name: &str,
    name: String,
    source: IllustrationIndexSource,
    body: TextIndexBody,
) -> Result<TextIndex> {
    required(&name, "text index name")?;
    let mut attributes = common_source_attributes(source.scope, source.relative_tab_stop_position);
    optional_bool(&mut attributes, "use-caption", source.use_caption);
    if let Some(sequence_name) = source.caption_sequence_name {
        attributes.push(attr(TEXT, "caption-sequence-name", sequence_name));
    }
    if let Some(format) = source.caption_sequence_format {
        attributes.push(attr(
            TEXT,
            "caption-sequence-format",
            match format {
                TextIndexCaptionSequenceFormat::Text => "text",
                TextIndexCaptionSequenceFormat::CategoryAndValue => "category-and-value",
                TextIndexCaptionSequenceFormat::Caption => "caption",
            },
        ));
    }
    let content =
        single_template_content(source.title_template, source.entry_template, entry_name)?;
    finish_typed_index(
        kind,
        root_name,
        source_name,
        name,
        attributes,
        content,
        body,
    )
}

fn common_source_attributes(
    scope: Option<TextIndexScope>,
    relative: Option<bool>,
) -> Vec<TextIndexAttribute> {
    let mut attributes = Vec::new();
    if let Some(scope) = scope {
        attributes.push(attr(
            TEXT,
            "index-scope",
            match scope {
                TextIndexScope::Document => "document",
                TextIndexScope::Chapter => "chapter",
            },
        ));
    }
    optional_bool(&mut attributes, "relative-tab-stop-position", relative);
    attributes
}

fn single_template_content(
    title: Option<TextIndexTitleTemplate>,
    entry: Option<TextIndexSimpleEntryTemplate>,
    entry_name: &str,
) -> Result<Vec<TextIndexContent>> {
    let mut content = Vec::new();
    if let Some(title) = title {
        content.push(element_content(title_template_element(title)?));
    }
    if let Some(entry) = entry {
        required(&entry.style_name, "index entry style name")?;
        if entry.tokens.len() > MAX_TOKENS {
            return invalid("index contains too many entry tokens");
        }
        validate_links(&entry.tokens)?;
        content.push(element_content(element(
            TEXT,
            entry_name,
            vec![attr(TEXT, "style-name", entry.style_name)],
            entry
                .tokens
                .into_iter()
                .map(token_element)
                .collect::<Result<Vec<_>>>()?
                .into_iter()
                .map(element_content)
                .collect(),
        )));
    }
    Ok(content)
}

fn title_template_element(title: TextIndexTitleTemplate) -> Result<TextIndexElement> {
    let mut attributes = Vec::new();
    optional_name(&mut attributes, "style-name", title.style_name)?;
    Ok(element(
        TEXT,
        "index-title-template",
        attributes,
        vec![TextIndexContent::Text(title.text)],
    ))
}

fn source_styles_element(styles: TextIndexSourceStyles) -> Result<TextIndexElement> {
    positive(styles.outline_level, "source styles outline level")?;
    let mut children = Vec::with_capacity(styles.style_names.len());
    for style_name in styles.style_names {
        required(&style_name, "source style name")?;
        children.push(element_content(element(
            TEXT,
            "index-source-style",
            vec![attr(TEXT, "style-name", style_name)],
            Vec::new(),
        )));
    }
    Ok(element(
        TEXT,
        "index-source-styles",
        vec![attr(
            TEXT,
            "outline-level",
            styles.outline_level.to_string(),
        )],
        children,
    ))
}

fn finish_typed_index(
    kind: TextIndexKind,
    root_name: &str,
    source_name: &str,
    name: String,
    source_attributes: Vec<TextIndexAttribute>,
    source_content: Vec<TextIndexContent>,
    body: TextIndexBody,
) -> Result<TextIndex> {
    if body.paragraphs.len() > MAX_BODY_PARAGRAPHS {
        return invalid("text index body exceeds configured paragraph limit");
    }
    let index = TextIndex {
        kind,
        root: element(
            TEXT,
            root_name,
            vec![attr(TEXT, "name", name)],
            vec![
                element_content(element(
                    TEXT,
                    source_name,
                    source_attributes,
                    source_content,
                )),
                element_content(body_element(body)?),
            ],
        ),
    };
    validate_index(&index)?;
    Ok(index)
}

fn body_element(body: TextIndexBody) -> Result<TextIndexElement> {
    let mut content = Vec::new();
    if let Some(title) = body.title {
        required(&title.name, "index body title name")?;
        let mut title_attrs = vec![attr(TEXT, "name", title.name)];
        optional_name(&mut title_attrs, "style-name", title.section_style_name)?;
        let mut paragraph_attrs = Vec::new();
        optional_name(
            &mut paragraph_attrs,
            "style-name",
            title.paragraph_style_name,
        )?;
        content.push(element_content(element(
            TEXT,
            "index-title",
            title_attrs,
            vec![element_content(element(
                TEXT,
                "p",
                paragraph_attrs,
                vec![TextIndexContent::Text(title.text)],
            ))],
        )));
    }
    for paragraph in body.paragraphs {
        let mut attributes = Vec::new();
        optional_name(&mut attributes, "style-name", paragraph.style_name)?;
        content.push(element_content(element(
            TEXT,
            "p",
            attributes,
            vec![TextIndexContent::Text(paragraph.text)],
        )));
    }
    Ok(element(TEXT, "index-body", Vec::new(), content))
}

pub fn insert_text_index_xml(xml: &str, index: &TextIndex) -> Result<String> {
    let current = validated_indexes(xml)?;
    if current
        .iter()
        .any(|candidate| candidate.name() == index.name())
    {
        return invalid(format!("duplicate text index name {:?}", index.name()));
    }
    let fragment = index.to_xml_fragment()?;
    let scan = scan_xml(xml)?;
    let mut output = String::with_capacity(xml.len() + fragment.len() + 32);
    match scan.text_container {
        TextContainer::Paired { close_start } => {
            output.push_str(&xml[..close_start]);
            output.push_str(&fragment);
            output.push_str(&xml[close_start..]);
        },
        TextContainer::Empty { start, end, qname } => {
            let raw = &xml[start..end];
            let slash = raw
                .rfind("/>")
                .ok_or_else(|| Error::InvalidFormat("malformed empty office:text".to_string()))?;
            output.push_str(&xml[..start]);
            output.push_str(&raw[..slash]);
            output.push('>');
            output.push_str(&fragment);
            output.push_str("</");
            output.push_str(&qname);
            output.push('>');
            output.push_str(&xml[end..]);
        },
    }
    validated_indexes(&output)?;
    Ok(output)
}

pub fn replace_text_index_xml(xml: &str, name: &str, replacement: &TextIndex) -> Result<String> {
    let current = validated_indexes(xml)?;
    let ordinal = unique_ordinal(&current, name)?;
    if replacement.name() != name
        && current
            .iter()
            .any(|candidate| candidate.name() == replacement.name())
    {
        return invalid(format!(
            "duplicate text index name {:?}",
            replacement.name()
        ));
    }
    let fragment = replacement.to_xml_fragment()?;
    let scan = scan_xml(xml)?;
    let span = scan.index_spans.get(ordinal).ok_or_else(|| {
        Error::InvalidFormat("text index parser/scanner order mismatch".to_string())
    })?;
    let mut output = String::with_capacity(xml.len() - (span.end - span.start) + fragment.len());
    output.push_str(&xml[..span.start]);
    output.push_str(&fragment);
    output.push_str(&xml[span.end..]);
    validated_indexes(&output)?;
    Ok(output)
}

pub fn remove_text_index_xml(xml: &str, name: &str) -> Result<String> {
    let current = validated_indexes(xml)?;
    let ordinal = unique_ordinal(&current, name)?;
    let scan = scan_xml(xml)?;
    let span = scan.index_spans.get(ordinal).ok_or_else(|| {
        Error::InvalidFormat("text index parser/scanner order mismatch".to_string())
    })?;
    let mut output = String::with_capacity(xml.len() - (span.end - span.start));
    output.push_str(&xml[..span.start]);
    output.push_str(&xml[span.end..]);
    validated_indexes(&output)?;
    Ok(output)
}

fn token_element(token: TextIndexEntryToken) -> Result<TextIndexElement> {
    let (name, attributes, text) = match token {
        TextIndexEntryToken::Chapter {
            style_name,
            display,
            outline_level,
        } => {
            let mut attrs = Vec::new();
            optional_name(&mut attrs, "style-name", style_name)?;
            if let Some(display) = display {
                attrs.push(attr(
                    TEXT,
                    "display",
                    match display {
                        TextIndexChapterDisplay::Name => "name",
                        TextIndexChapterDisplay::Number => "number",
                        TextIndexChapterDisplay::NumberAndName => "number-and-name",
                        TextIndexChapterDisplay::PlainNumber => "plain-number",
                        TextIndexChapterDisplay::PlainNumberAndName => "plain-number-and-name",
                    },
                ));
            }
            optional_positive(&mut attrs, "outline-level", outline_level)?;
            ("index-entry-chapter", attrs, None)
        },
        TextIndexEntryToken::PageNumber { style_name } => {
            let mut attrs = Vec::new();
            optional_name(&mut attrs, "style-name", style_name)?;
            ("index-entry-page-number", attrs, None)
        },
        TextIndexEntryToken::Text { style_name } => {
            let mut attrs = Vec::new();
            optional_name(&mut attrs, "style-name", style_name)?;
            ("index-entry-text", attrs, None)
        },
        TextIndexEntryToken::Span { style_name, text } => {
            let mut attrs = Vec::new();
            optional_name(&mut attrs, "style-name", style_name)?;
            ("index-entry-span", attrs, Some(text))
        },
        TextIndexEntryToken::LinkStart { style_name } => {
            let mut attrs = Vec::new();
            optional_name(&mut attrs, "style-name", style_name)?;
            ("index-entry-link-start", attrs, None)
        },
        TextIndexEntryToken::LinkEnd { style_name } => {
            let mut attrs = Vec::new();
            optional_name(&mut attrs, "style-name", style_name)?;
            ("index-entry-link-end", attrs, None)
        },
        TextIndexEntryToken::TabStop(tab) => {
            let mut attrs = Vec::new();
            match tab {
                TextIndexTabStop::Right { leader, style_name } => {
                    optional_name(&mut attrs, "style-name", style_name)?;
                    attrs.push(attr(STYLE, "type", "right"));
                    if let Some(leader) = leader {
                        attrs.push(attr(STYLE, "leader-char", leader.to_string()));
                    }
                },
                TextIndexTabStop::Left {
                    position,
                    leader,
                    style_name,
                } => {
                    optional_name(&mut attrs, "style-name", style_name)?;
                    validate_length(&position)?;
                    attrs.push(attr(STYLE, "type", "left"));
                    attrs.push(attr(STYLE, "position", position));
                    if let Some(leader) = leader {
                        attrs.push(attr(STYLE, "leader-char", leader.to_string()));
                    }
                },
            }
            ("index-entry-tab-stop", attrs, None)
        },
    };
    Ok(element(
        TEXT,
        name,
        attributes,
        text.map(TextIndexContent::Text).into_iter().collect(),
    ))
}

fn validate_index(index: &TextIndex) -> Result<()> {
    match index.kind {
        TextIndexKind::Illustration => {
            return validate_caption_index(
                index,
                "illustration-index",
                "illustration-index-source",
                "illustration-index-entry-template",
            );
        },
        TextIndexKind::Table => {
            return validate_caption_index(
                index,
                "table-index",
                "table-index-source",
                "table-index-entry-template",
            );
        },
        TextIndexKind::Object => return validate_object_index(index),
        TextIndexKind::User => return validate_user_index(index),
        TextIndexKind::Alphabetical => return validate_alphabetical_index(index),
        TextIndexKind::Bibliography => return validate_bibliography_index(index),
        _ => {},
    }
    if index.kind != TextIndexKind::TableOfContents {
        return Ok(());
    }
    named(&index.root, TEXT, "table-of-content")?;
    validate_root_attrs(&index.root)?;
    required(index.name(), "text index name")?;
    if let Some(value) = index.root.attribute(Some(TEXT), "protected") {
        boolean(value, "text:protected")?;
    }
    let children = structural_children(&index.root)?;
    if children.len() != 2 {
        return invalid("table-of-content requires exactly one source followed by one index-body");
    }
    let source = children[0];
    named(source, TEXT, "table-of-content-source")?;
    named(children[1], TEXT, "index-body")?;
    validate_attrs(
        source,
        &[
            (TEXT, "outline-level"),
            (TEXT, "use-outline-level"),
            (TEXT, "use-index-marks"),
            (TEXT, "use-index-source-styles"),
            (TEXT, "index-scope"),
            (TEXT, "relative-tab-stop-position"),
        ],
    )?;
    if let Some(value) = source.attribute(Some(TEXT), "outline-level") {
        parse_positive(value, "source outline level")?;
    }
    for name in [
        "use-outline-level",
        "use-index-marks",
        "use-index-source-styles",
        "relative-tab-stop-position",
    ] {
        if let Some(value) = source.attribute(Some(TEXT), name) {
            boolean(value, name)?;
        }
    }
    if let Some(value) = source.attribute(Some(TEXT), "index-scope")
        && value != "document"
        && value != "chapter"
    {
        return invalid("invalid text:index-scope");
    }
    let source_children = structural_children(source)?;
    let mut phase = 0u8;
    let mut templates = 0usize;
    let mut tokens = 0usize;
    for child in source_children {
        let next = match child.local_name.as_str() {
            "index-title-template" => 0,
            "table-of-content-entry-template" => 1,
            "index-source-styles" => 2,
            other => {
                return invalid(format!(
                    "unexpected table-of-content source child text:{other}"
                ));
            },
        };
        if next < phase || (next == 0 && phase == 0 && templates != 0) {
            return invalid("table-of-content source children are out of ODF order");
        }
        phase = next;
        templates += 1;
        if templates > MAX_TEMPLATES {
            return invalid("too many table-of-content templates");
        }
        match next {
            0 => {
                validate_attrs(child, &[(TEXT, "style-name")])?;
                text_only(child)?;
            },
            1 => {
                validate_attrs(child, &[(TEXT, "outline-level"), (TEXT, "style-name")])?;
                parse_positive(
                    required_attr(child, TEXT, "outline-level")?,
                    "entry outline level",
                )?;
                required(
                    required_attr(child, TEXT, "style-name")?,
                    "entry style name",
                )?;
                let mut link_depth = 0usize;
                for token in structural_children(child)? {
                    tokens += 1;
                    if tokens > MAX_TOKENS {
                        return invalid("too many table-of-content entry tokens");
                    }
                    validate_token(token, &mut link_depth)?;
                }
                if link_depth != 0 {
                    return invalid("unclosed index-entry-link-start");
                }
            },
            _ => {
                validate_attrs(child, &[(TEXT, "outline-level")])?;
                parse_positive(
                    required_attr(child, TEXT, "outline-level")?,
                    "source styles outline level",
                )?;
                for style in structural_children(child)? {
                    named(style, TEXT, "index-source-style")?;
                    validate_attrs(style, &[(TEXT, "style-name")])?;
                    required(
                        required_attr(style, TEXT, "style-name")?,
                        "source style name",
                    )?;
                    empty(style)?;
                }
            },
        }
    }
    Ok(())
}

fn validate_alphabetical_index(index: &TextIndex) -> Result<()> {
    let source = validate_index_shell(index, "alphabetical-index", "alphabetical-index-source")?;
    validate_attrs(
        source,
        &[
            (TEXT, "index-scope"),
            (TEXT, "relative-tab-stop-position"),
            (TEXT, "ignore-case"),
            (TEXT, "main-entry-style-name"),
            (TEXT, "alphabetical-separators"),
            (TEXT, "combine-entries"),
            (TEXT, "combine-entries-with-dash"),
            (TEXT, "combine-entries-with-pp"),
            (TEXT, "use-keys-as-entries"),
            (TEXT, "capitalize-entries"),
            (TEXT, "comma-separated"),
            (FO, "language"),
            (FO, "country"),
            (FO, "script"),
            (STYLE, "rfc-language-tag"),
            (TEXT, "sort-algorithm"),
        ],
    )?;
    validate_scope_and_relative(source)?;
    for name in [
        "ignore-case",
        "alphabetical-separators",
        "combine-entries",
        "combine-entries-with-dash",
        "combine-entries-with-pp",
        "use-keys-as-entries",
        "capitalize-entries",
        "comma-separated",
    ] {
        if let Some(value) = source.attribute(Some(TEXT), name) {
            boolean(value, name)?;
        }
    }
    validate_optional_locale(source, FO, "language", LocaleLexical::LanguageCode)?;
    validate_optional_locale(source, FO, "country", LocaleLexical::CountryOrScript)?;
    validate_optional_locale(source, FO, "script", LocaleLexical::CountryOrScript)?;
    validate_optional_locale(
        source,
        STYLE,
        "rfc-language-tag",
        LocaleLexical::LanguageTag,
    )?;
    let children = structural_children(source)?;
    let mut position = 0usize;
    if children
        .first()
        .is_some_and(|child| child.local_name == "index-title-template")
    {
        validate_attrs(children[0], &[(TEXT, "style-name")])?;
        text_only(children[0])?;
        position = 1;
    }
    if children.len().saturating_sub(position) > MAX_TEMPLATES {
        return invalid("too many alphabetical index templates");
    }
    for template in &children[position..] {
        named(template, TEXT, "alphabetical-index-entry-template")?;
        validate_attrs(template, &[(TEXT, "outline-level"), (TEXT, "style-name")])?;
        let level = required_attr(template, TEXT, "outline-level")?;
        if !["1", "2", "3", "separator"].contains(&level) {
            return invalid("invalid alphabetical index outline level");
        }
        required(
            required_attr(template, TEXT, "style-name")?,
            "alphabetical index entry style name",
        )?;
        let tokens = structural_children(template)?;
        if tokens.len() > MAX_TOKENS {
            return invalid("too many alphabetical index tokens");
        }
        let mut unused_link_depth = 0usize;
        for token in tokens {
            if !matches!(
                token.local_name.as_str(),
                "index-entry-chapter"
                    | "index-entry-page-number"
                    | "index-entry-text"
                    | "index-entry-span"
                    | "index-entry-tab-stop"
            ) {
                return invalid(format!(
                    "unexpected alphabetical index token text:{}",
                    token.local_name
                ));
            }
            validate_token(token, &mut unused_link_depth)?;
        }
    }
    Ok(())
}

fn validate_bibliography_index(index: &TextIndex) -> Result<()> {
    let source = validate_index_shell(index, "bibliography", "bibliography-source")?;
    validate_attrs(source, &[])?;
    let children = structural_children(source)?;
    let mut position = 0usize;
    if children
        .first()
        .is_some_and(|child| child.local_name == "index-title-template")
    {
        validate_attrs(children[0], &[(TEXT, "style-name")])?;
        text_only(children[0])?;
        position = 1;
    }
    if children.len().saturating_sub(position) > MAX_TEMPLATES {
        return invalid("too many bibliography templates");
    }
    let mut token_count = 0usize;
    for template in &children[position..] {
        named(template, TEXT, "bibliography-entry-template")?;
        validate_attrs(
            template,
            &[(TEXT, "bibliography-type"), (TEXT, "style-name")],
        )?;
        if bibliography_type(required_attr(template, TEXT, "bibliography-type")?).is_none() {
            return invalid("invalid text:bibliography-type");
        }
        required(
            required_attr(template, TEXT, "style-name")?,
            "bibliography entry style name",
        )?;
        for token in structural_children(template)? {
            token_count += 1;
            if token_count > MAX_TOKENS {
                return invalid("too many bibliography entry tokens");
            }
            match token.local_name.as_str() {
                "index-entry-span" => {
                    validate_attrs(token, &[(TEXT, "style-name")])?;
                    text_only(token)?;
                },
                "index-entry-tab-stop" => {
                    let mut depth = 0;
                    validate_token(token, &mut depth)?;
                },
                "index-entry-bibliography" => {
                    validate_attrs(
                        token,
                        &[(TEXT, "bibliography-data-field"), (TEXT, "style-name")],
                    )?;
                    if bibliography_field(required_attr(token, TEXT, "bibliography-data-field")?)
                        .is_none()
                    {
                        return invalid("invalid text:bibliography-data-field");
                    }
                    empty(token)?;
                },
                other => {
                    return invalid(format!("unexpected bibliography entry token text:{other}"));
                },
            }
        }
    }
    Ok(())
}

fn validate_optional_locale(
    element: &TextIndexElement,
    namespace: &str,
    name: &str,
    lexical: LocaleLexical,
) -> Result<()> {
    if let Some(value) = element.attribute(Some(namespace), name) {
        validate_locale(value, lexical, name)?;
    }
    Ok(())
}

fn bibliography_type(value: &str) -> Option<TextBibliographyType> {
    Some(match value {
        "article" => TextBibliographyType::Article,
        "book" => TextBibliographyType::Book,
        "booklet" => TextBibliographyType::Booklet,
        "conference" => TextBibliographyType::Conference,
        "custom1" => TextBibliographyType::Custom1,
        "custom2" => TextBibliographyType::Custom2,
        "custom3" => TextBibliographyType::Custom3,
        "custom4" => TextBibliographyType::Custom4,
        "custom5" => TextBibliographyType::Custom5,
        "email" => TextBibliographyType::Email,
        "inbook" => TextBibliographyType::InBook,
        "incollection" => TextBibliographyType::InCollection,
        "inproceedings" => TextBibliographyType::InProceedings,
        "journal" => TextBibliographyType::Journal,
        "manual" => TextBibliographyType::Manual,
        "mastersthesis" => TextBibliographyType::MastersThesis,
        "misc" => TextBibliographyType::Misc,
        "phdthesis" => TextBibliographyType::PhdThesis,
        "proceedings" => TextBibliographyType::Proceedings,
        "techreport" => TextBibliographyType::TechReport,
        "unpublished" => TextBibliographyType::Unpublished,
        "www" => TextBibliographyType::Www,
        _ => return None,
    })
}

fn bibliography_field(value: &str) -> Option<()> {
    if [
        "address",
        "annote",
        "author",
        "bibliography-type",
        "booktitle",
        "chapter",
        "custom1",
        "custom2",
        "custom3",
        "custom4",
        "custom5",
        "edition",
        "editor",
        "howpublished",
        "identifier",
        "institution",
        "isbn",
        "issn",
        "journal",
        "month",
        "note",
        "number",
        "organizations",
        "pages",
        "publisher",
        "report-type",
        "school",
        "series",
        "title",
        "url",
        "volume",
        "year",
    ]
    .contains(&value)
    {
        Some(())
    } else {
        None
    }
}

fn validate_caption_index(
    index: &TextIndex,
    root_name: &str,
    source_name: &str,
    entry_name: &str,
) -> Result<()> {
    let source = validate_index_shell(index, root_name, source_name)?;
    validate_attrs(
        source,
        &[
            (TEXT, "index-scope"),
            (TEXT, "relative-tab-stop-position"),
            (TEXT, "use-caption"),
            (TEXT, "caption-sequence-name"),
            (TEXT, "caption-sequence-format"),
        ],
    )?;
    validate_scope_and_relative(source)?;
    if let Some(value) = source.attribute(Some(TEXT), "use-caption") {
        boolean(value, "text:use-caption")?;
    }
    if let Some(value) = source.attribute(Some(TEXT), "caption-sequence-format")
        && !["text", "category-and-value", "caption"].contains(&value)
    {
        return invalid("invalid text:caption-sequence-format");
    }
    validate_single_template_children(source, entry_name)
}

fn validate_object_index(index: &TextIndex) -> Result<()> {
    let source = validate_index_shell(index, "object-index", "object-index-source")?;
    validate_attrs(
        source,
        &[
            (TEXT, "index-scope"),
            (TEXT, "relative-tab-stop-position"),
            (TEXT, "use-spreadsheet-objects"),
            (TEXT, "use-math-objects"),
            (TEXT, "use-draw-objects"),
            (TEXT, "use-chart-objects"),
            (TEXT, "use-other-objects"),
        ],
    )?;
    validate_scope_and_relative(source)?;
    for name in [
        "use-spreadsheet-objects",
        "use-math-objects",
        "use-draw-objects",
        "use-chart-objects",
        "use-other-objects",
    ] {
        if let Some(value) = source.attribute(Some(TEXT), name) {
            boolean(value, name)?;
        }
    }
    validate_single_template_children(source, "object-index-entry-template")
}

fn validate_user_index(index: &TextIndex) -> Result<()> {
    let source = validate_index_shell(index, "user-index", "user-index-source")?;
    validate_attrs(
        source,
        &[
            (TEXT, "index-name"),
            (TEXT, "index-scope"),
            (TEXT, "relative-tab-stop-position"),
            (TEXT, "use-index-marks"),
            (TEXT, "use-index-source-styles"),
            (TEXT, "use-graphics"),
            (TEXT, "use-tables"),
            (TEXT, "use-floating-frames"),
            (TEXT, "use-objects"),
            (TEXT, "copy-outline-levels"),
        ],
    )?;
    required_attr(source, TEXT, "index-name")?;
    validate_scope_and_relative(source)?;
    for name in [
        "use-index-marks",
        "use-index-source-styles",
        "use-graphics",
        "use-tables",
        "use-floating-frames",
        "use-objects",
        "copy-outline-levels",
    ] {
        if let Some(value) = source.attribute(Some(TEXT), name) {
            boolean(value, name)?;
        }
    }
    let mut phase = 0u8;
    let mut count = 0usize;
    for child in structural_children(source)? {
        let next = match child.local_name.as_str() {
            "index-title-template" => 0,
            "user-index-entry-template" => 1,
            "index-source-styles" => 2,
            other => return invalid(format!("unexpected user-index source child text:{other}")),
        };
        if next < phase || (next == 0 && count != 0) {
            return invalid("user-index source children are out of ODF order");
        }
        phase = next;
        count += 1;
        if count > MAX_TEMPLATES {
            return invalid("too many user-index templates");
        }
        match next {
            0 => {
                validate_attrs(child, &[(TEXT, "style-name")])?;
                text_only(child)?;
            },
            1 => validate_outline_template(child)?,
            _ => validate_source_styles(child)?,
        }
    }
    Ok(())
}

fn validate_index_shell<'a>(
    index: &'a TextIndex,
    root_name: &str,
    source_name: &str,
) -> Result<&'a TextIndexElement> {
    named(&index.root, TEXT, root_name)?;
    validate_root_attrs(&index.root)?;
    required(index.name(), "text index name")?;
    if let Some(value) = index.root.attribute(Some(TEXT), "protected") {
        boolean(value, "text:protected")?;
    }
    let children = structural_children(&index.root)?;
    if children.len() != 2 {
        return invalid(format!(
            "text:{root_name} requires exactly one source followed by one index-body"
        ));
    }
    named(children[0], TEXT, source_name)?;
    named(children[1], TEXT, "index-body")?;
    Ok(children[0])
}

fn validate_root_attrs(root: &TextIndexElement) -> Result<()> {
    validate_attrs(
        root,
        &[
            (TEXT, "name"),
            (TEXT, "protected"),
            (TEXT, "style-name"),
            (TEXT, "protection-key"),
            (TEXT, "protection-key-digest-algorithm"),
            ("http://www.w3.org/XML/1998/namespace", "id"),
        ],
    )
}

fn validate_scope_and_relative(source: &TextIndexElement) -> Result<()> {
    if let Some(value) = source.attribute(Some(TEXT), "index-scope")
        && value != "document"
        && value != "chapter"
    {
        return invalid("invalid text:index-scope");
    }
    if let Some(value) = source.attribute(Some(TEXT), "relative-tab-stop-position") {
        boolean(value, "text:relative-tab-stop-position")?;
    }
    Ok(())
}

fn validate_single_template_children(source: &TextIndexElement, entry_name: &str) -> Result<()> {
    let children = structural_children(source)?;
    if children.len() > 2 {
        return invalid("single-template index source has too many children");
    }
    let mut position = 0usize;
    if children
        .first()
        .is_some_and(|child| child.local_name == "index-title-template")
    {
        named(children[0], TEXT, "index-title-template")?;
        validate_attrs(children[0], &[(TEXT, "style-name")])?;
        text_only(children[0])?;
        position = 1;
    }
    if let Some(entry) = children.get(position) {
        named(entry, TEXT, entry_name)?;
        validate_simple_template(entry)?;
        position += 1;
    }
    if position != children.len() {
        return invalid("single-template index source children are out of ODF order");
    }
    Ok(())
}

fn validate_simple_template(template: &TextIndexElement) -> Result<()> {
    validate_attrs(template, &[(TEXT, "style-name")])?;
    required(
        required_attr(template, TEXT, "style-name")?,
        "index entry style name",
    )?;
    validate_template_tokens(template)
}

fn validate_outline_template(template: &TextIndexElement) -> Result<()> {
    validate_attrs(template, &[(TEXT, "outline-level"), (TEXT, "style-name")])?;
    parse_positive(
        required_attr(template, TEXT, "outline-level")?,
        "entry outline level",
    )?;
    required(
        required_attr(template, TEXT, "style-name")?,
        "index entry style name",
    )?;
    validate_template_tokens(template)
}

fn validate_template_tokens(template: &TextIndexElement) -> Result<()> {
    let mut link_depth = 0usize;
    let children = structural_children(template)?;
    if children.len() > MAX_TOKENS {
        return invalid("too many index entry tokens");
    }
    for token in children {
        validate_token(token, &mut link_depth)?;
    }
    if link_depth != 0 {
        return invalid("unclosed index-entry-link-start");
    }
    Ok(())
}

fn validate_source_styles(element: &TextIndexElement) -> Result<()> {
    named(element, TEXT, "index-source-styles")?;
    validate_attrs(element, &[(TEXT, "outline-level")])?;
    parse_positive(
        required_attr(element, TEXT, "outline-level")?,
        "source styles outline level",
    )?;
    for style in structural_children(element)? {
        named(style, TEXT, "index-source-style")?;
        validate_attrs(style, &[(TEXT, "style-name")])?;
        required(
            required_attr(style, TEXT, "style-name")?,
            "source style name",
        )?;
        empty(style)?;
    }
    Ok(())
}

fn validate_token(element: &TextIndexElement, link_depth: &mut usize) -> Result<()> {
    if element.namespace_uri.as_deref() != Some(TEXT) {
        return invalid("foreign element in TOC entry template");
    }
    match element.local_name.as_str() {
        "index-entry-chapter" => {
            validate_attrs(
                element,
                &[
                    (TEXT, "style-name"),
                    (TEXT, "display"),
                    (TEXT, "outline-level"),
                ],
            )?;
            if let Some(value) = element.attribute(Some(TEXT), "outline-level") {
                parse_positive(value, "chapter outline level")?;
            }
            if let Some(value) = element.attribute(Some(TEXT), "display")
                && ![
                    "name",
                    "number",
                    "number-and-name",
                    "plain-number",
                    "plain-number-and-name",
                ]
                .contains(&value)
            {
                return invalid("invalid text:index-entry-chapter display");
            }
            empty(element)?;
        },
        "index-entry-page-number" | "index-entry-text" => {
            validate_attrs(element, &[(TEXT, "style-name")])?;
            empty(element)?;
        },
        "index-entry-span" => {
            validate_attrs(element, &[(TEXT, "style-name")])?;
            text_only(element)?;
        },
        "index-entry-link-start" => {
            validate_attrs(element, &[(TEXT, "style-name")])?;
            empty(element)?;
            if *link_depth != 0 {
                return invalid("nested index entry links are not supported");
            }
            *link_depth = 1;
        },
        "index-entry-link-end" => {
            validate_attrs(element, &[(TEXT, "style-name")])?;
            empty(element)?;
            if *link_depth == 0 {
                return invalid("index-entry-link-end has no matching start");
            }
            *link_depth = 0;
        },
        "index-entry-tab-stop" => {
            validate_attrs(
                element,
                &[
                    (TEXT, "style-name"),
                    (STYLE, "type"),
                    (STYLE, "position"),
                    (STYLE, "leader-char"),
                ],
            )?;
            let tab_type = element.attribute(Some(STYLE), "type");
            match tab_type {
                Some("right") => {
                    if element.attribute(Some(STYLE), "position").is_some() {
                        return invalid("right tab stop cannot have style:position");
                    }
                },
                Some("left") => validate_length(required_attr(element, STYLE, "position")?)?,
                _ => return invalid("tab stop requires style:type left or right"),
            }
            if let Some(leader) = element.attribute(Some(STYLE), "leader-char")
                && leader.chars().count() != 1
            {
                return invalid("style:leader-char must be one character");
            }
            empty(element)?;
        },
        other => return invalid(format!("unexpected TOC entry token text:{other}")),
    }
    Ok(())
}

fn validated_indexes(xml: &str) -> Result<Vec<TextIndex>> {
    if xml.len() > MAX_XML_BYTES {
        return invalid("content XML exceeds 64 MiB");
    }
    let indexes = parse_text_indexes(xml)?;
    let mut names = BTreeSet::new();
    for index in &indexes {
        validate_index(index)?;
        if !names.insert(index.name().to_string()) {
            return invalid(format!("duplicate text index name {:?}", index.name()));
        }
    }
    Ok(indexes)
}

fn unique_ordinal(indexes: &[TextIndex], name: &str) -> Result<usize> {
    indexes
        .iter()
        .position(|index| index.name() == name)
        .ok_or_else(|| Error::InvalidFormat(format!("text index {name:?} was not found")))
}

#[derive(Debug)]
struct Span {
    start: usize,
    end: usize,
}

enum TextContainer {
    Paired {
        close_start: usize,
    },
    Empty {
        start: usize,
        end: usize,
        qname: String,
    },
}

struct XmlScan {
    text_container: TextContainer,
    index_spans: Vec<Span>,
}

fn scan_xml(xml: &str) -> Result<XmlScan> {
    if xml.len() > MAX_XML_BYTES {
        return invalid("content XML exceeds 64 MiB");
    }
    let mut reader = NsReader::from_str(xml);
    let mut buffer = Vec::new();
    let mut depth = 0usize;
    let mut text_open: Option<usize> = None;
    let mut text_container = None;
    let mut index_stack = Vec::<(usize, usize)>::new();
    let mut spans = Vec::new();
    loop {
        let start = reader.buffer_position() as usize;
        let (namespace, event) = reader
            .read_resolved_event_into(&mut buffer)
            .map_err(|error| {
                Error::InvalidFormat(format!(
                    "invalid content XML while locating text indexes: {error}"
                ))
            })?;
        let is_office = bound(&namespace, OFFICE);
        let is_text = bound(&namespace, super::TEXT_NAMESPACE);
        drop(namespace);
        let end = reader.buffer_position() as usize;
        match event {
            Event::Start(ref value) => {
                let local = value.local_name();
                if is_office && local.as_ref() == b"text" {
                    if text_open.is_some() || text_container.is_some() {
                        return invalid("content XML has multiple office:text elements");
                    }
                    text_open = Some(depth);
                }
                if is_text && kind_local(local.as_ref()) {
                    index_stack.push((depth, start));
                }
                depth += 1;
                if depth > MAX_DEPTH {
                    return invalid("content XML nesting exceeds 4096 elements");
                }
            },
            Event::Empty(ref value) => {
                let local = value.local_name();
                if is_office && local.as_ref() == b"text" {
                    if text_open.is_some() || text_container.is_some() {
                        return invalid("content XML has multiple office:text elements");
                    }
                    let qname = std::str::from_utf8(value.name().as_ref())
                        .map_err(|_| {
                            Error::InvalidFormat("non-UTF-8 office:text name".to_string())
                        })?
                        .to_string();
                    text_container = Some(TextContainer::Empty { start, end, qname });
                }
                if is_text && kind_local(local.as_ref()) {
                    spans.push(Span { start, end });
                }
            },
            Event::End(ref value) => {
                depth = depth
                    .checked_sub(1)
                    .ok_or_else(|| Error::InvalidFormat("unbalanced content XML".to_string()))?;
                if index_stack
                    .last()
                    .is_some_and(|(open_depth, _)| *open_depth == depth)
                    && kind_local(value.local_name().as_ref())
                {
                    let (_, open_start) = index_stack.pop().expect("checked");
                    spans.push(Span {
                        start: open_start,
                        end,
                    });
                }
                if text_open == Some(depth) && is_office && value.local_name().as_ref() == b"text" {
                    text_open = None;
                    if text_container.is_some() {
                        return invalid("content XML has multiple office:text elements");
                    }
                    text_container = Some(TextContainer::Paired { close_start: start });
                }
            },
            Event::DocType(_) => {
                return invalid("DOCTYPE is not allowed in mutable ODF content XML");
            },
            Event::Eof => break,
            _ => {},
        }
        buffer.clear();
    }
    if depth != 0 || !index_stack.is_empty() {
        return invalid("unbalanced content XML");
    }
    spans.sort_by_key(|span| span.start);
    Ok(XmlScan {
        text_container: text_container.ok_or_else(|| {
            Error::InvalidFormat("content XML has no office:text element".to_string())
        })?,
        index_spans: spans,
    })
}

fn kind_local(local: &[u8]) -> bool {
    matches!(
        local,
        b"table-of-content"
            | b"illustration-index"
            | b"table-index"
            | b"object-index"
            | b"user-index"
            | b"alphabetical-index"
            | b"bibliography"
    )
}

fn bound(namespace: &ResolveResult<'_>, expected: &[u8]) -> bool {
    matches!(namespace, ResolveResult::Bound(Namespace(value)) if *value == expected)
}

fn validate_links(tokens: &[TextIndexEntryToken]) -> Result<()> {
    let mut open = false;
    for token in tokens {
        match token {
            TextIndexEntryToken::LinkStart { .. } if open => {
                return invalid("nested index entry links are not supported");
            },
            TextIndexEntryToken::LinkStart { .. } => open = true,
            TextIndexEntryToken::LinkEnd { .. } if !open => {
                return invalid("index entry link end has no start");
            },
            TextIndexEntryToken::LinkEnd { .. } => open = false,
            _ => {},
        }
    }
    if open {
        return invalid("index entry link start has no end");
    }
    Ok(())
}

fn element(
    namespace: &str,
    local_name: &str,
    attributes: Vec<TextIndexAttribute>,
    content: Vec<TextIndexContent>,
) -> TextIndexElement {
    TextIndexElement {
        namespace_uri: Some(namespace.to_string()),
        local_name: local_name.to_string(),
        attributes,
        content,
    }
}

fn element_content(element: TextIndexElement) -> TextIndexContent {
    TextIndexContent::Element(element)
}

fn attr(namespace: &str, local_name: &str, value: impl Into<String>) -> TextIndexAttribute {
    TextIndexAttribute {
        namespace_uri: Some(namespace.to_string()),
        local_name: local_name.to_string(),
        value: value.into(),
    }
}

fn set_attr(
    attributes: &mut Vec<TextIndexAttribute>,
    namespace: &str,
    local_name: &str,
    value: String,
) {
    if let Some(attribute) = attributes.iter_mut().find(|attribute| {
        attribute.namespace_uri.as_deref() == Some(namespace) && attribute.local_name == local_name
    }) {
        attribute.value = value;
    } else {
        attributes.push(attr(namespace, local_name, value));
    }
}

fn optional_name(
    attributes: &mut Vec<TextIndexAttribute>,
    name: &str,
    value: Option<String>,
) -> Result<()> {
    if let Some(value) = value {
        required(&value, name)?;
        attributes.push(attr(TEXT, name, value));
    }
    Ok(())
}

fn optional_positive(
    attributes: &mut Vec<TextIndexAttribute>,
    name: &str,
    value: Option<u16>,
) -> Result<()> {
    if let Some(value) = value {
        positive(value, name)?;
        attributes.push(attr(TEXT, name, value.to_string()));
    }
    Ok(())
}

fn optional_bool(attributes: &mut Vec<TextIndexAttribute>, name: &str, value: Option<bool>) {
    if let Some(value) = value {
        attributes.push(attr(TEXT, name, value.to_string()));
    }
}

fn positive(value: u16, context: &str) -> Result<()> {
    if value == 0 {
        invalid(format!("{context} must be positive"))
    } else {
        Ok(())
    }
}

fn required(value: &str, context: &str) -> Result<()> {
    if value.is_empty() {
        invalid(format!("{context} cannot be empty"))
    } else {
        Ok(())
    }
}

fn boolean(value: &str, context: &str) -> Result<()> {
    if matches!(value, "true" | "false" | "1" | "0") {
        Ok(())
    } else {
        invalid(format!("{context} is not an XML boolean"))
    }
}

fn parse_positive(value: &str, context: &str) -> Result<u64> {
    let value = value
        .parse::<u64>()
        .map_err(|_| Error::InvalidFormat(format!("{context} is not a positive integer")))?;
    if value == 0 {
        return invalid(format!("{context} must be positive"));
    }
    Ok(value)
}

fn validate_length(value: &str) -> Result<()> {
    let unit = ["cm", "mm", "in", "pt", "pc", "px"]
        .into_iter()
        .find(|unit| value.ends_with(unit))
        .ok_or_else(|| Error::InvalidFormat(format!("invalid ODF length {value:?}")))?;
    let number = &value[..value.len() - unit.len()];
    number
        .parse::<f64>()
        .map_err(|_| Error::InvalidFormat(format!("invalid ODF length {value:?}")))?;
    Ok(())
}

fn named(element: &TextIndexElement, namespace: &str, local: &str) -> Result<()> {
    if element.namespace_uri.as_deref() == Some(namespace) && element.local_name == local {
        Ok(())
    } else {
        invalid(format!("expected {{{namespace}}}{local}"))
    }
}

fn validate_attrs(element: &TextIndexElement, allowed: &[(&str, &str)]) -> Result<()> {
    for attribute in &element.attributes {
        if !allowed.iter().any(|(namespace, local)| {
            attribute.namespace_uri.as_deref() == Some(*namespace) && attribute.local_name == *local
        }) {
            return invalid(format!(
                "unexpected attribute {{{}}}{} on text:{}",
                attribute.namespace_uri.as_deref().unwrap_or(""),
                attribute.local_name,
                element.local_name
            ));
        }
    }
    Ok(())
}

fn required_attr<'a>(
    element: &'a TextIndexElement,
    namespace: &str,
    local: &str,
) -> Result<&'a str> {
    element.attribute(Some(namespace), local).ok_or_else(|| {
        Error::InvalidFormat(format!(
            "text:{} requires attribute {local}",
            element.local_name
        ))
    })
}

fn structural_children(element: &TextIndexElement) -> Result<Vec<&TextIndexElement>> {
    let mut children = Vec::new();
    for content in &element.content {
        match content {
            TextIndexContent::Element(child) => children.push(child),
            TextIndexContent::Text(text) if text.trim().is_empty() => {},
            TextIndexContent::Text(_) => {
                return invalid(format!(
                    "text:{} cannot contain direct character data",
                    element.local_name
                ));
            },
        }
    }
    Ok(children)
}

fn empty(element: &TextIndexElement) -> Result<()> {
    if element
        .content
        .iter()
        .all(|content| matches!(content, TextIndexContent::Text(text) if text.is_empty()))
    {
        Ok(())
    } else {
        invalid(format!("text:{} must be empty", element.local_name))
    }
}

fn text_only(element: &TextIndexElement) -> Result<()> {
    if element
        .content
        .iter()
        .all(|content| matches!(content, TextIndexContent::Text(_)))
    {
        Ok(())
    } else {
        invalid(format!(
            "text:{} accepts character data only",
            element.local_name
        ))
    }
}

fn collect_namespaces(element: &TextIndexElement, output: &mut BTreeSet<String>) {
    if let Some(namespace) = &element.namespace_uri {
        output.insert(namespace.clone());
    }
    for attribute in &element.attributes {
        if let Some(namespace) = &attribute.namespace_uri {
            output.insert(namespace.clone());
        }
    }
    for content in &element.content {
        if let TextIndexContent::Element(child) = content {
            collect_namespaces(child, output);
        }
    }
}

fn namespace_prefixes(namespaces: &BTreeSet<String>) -> BTreeMap<String, String> {
    let known = [
        (TEXT, "text"),
        (STYLE, "style"),
        ("urn:oasis:names:tc:opendocument:xmlns:office:1.0", "office"),
        ("urn:oasis:names:tc:opendocument:xmlns:table:1.0", "table"),
        ("urn:oasis:names:tc:opendocument:xmlns:drawing:1.0", "draw"),
        (
            "urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0",
            "fo",
        ),
        ("http://www.w3.org/1999/xlink", "xlink"),
        ("http://www.w3.org/XML/1998/namespace", "xml"),
    ];
    let mut result = BTreeMap::new();
    let mut unknown = 0usize;
    for namespace in namespaces {
        let prefix = known
            .iter()
            .find(|(uri, _)| *uri == namespace)
            .map(|(_, prefix)| (*prefix).to_string())
            .unwrap_or_else(|| {
                let prefix = format!("ns{unknown}");
                unknown += 1;
                prefix
            });
        result.insert(namespace.clone(), prefix);
    }
    result
}

fn write_element(
    element: &TextIndexElement,
    prefixes: &BTreeMap<String, String>,
    root: bool,
    output: &mut String,
) -> Result<()> {
    output.push('<');
    write_name(
        element.namespace_uri.as_deref(),
        &element.local_name,
        prefixes,
        output,
    )?;
    if root {
        for (namespace, prefix) in prefixes {
            if prefix == "xml" {
                continue;
            }
            output.push_str(" xmlns:");
            output.push_str(prefix);
            output.push_str("=\"");
            escape_attr(namespace, output);
            output.push('"');
        }
    }
    let mut attributes: Vec<_> = element.attributes.iter().collect();
    attributes.sort_by(|left, right| {
        (left.namespace_uri.as_deref(), left.local_name.as_str())
            .cmp(&(right.namespace_uri.as_deref(), right.local_name.as_str()))
    });
    for attribute in attributes {
        output.push(' ');
        write_name(
            attribute.namespace_uri.as_deref(),
            &attribute.local_name,
            prefixes,
            output,
        )?;
        output.push_str("=\"");
        escape_attr(&attribute.value, output);
        output.push('"');
    }
    if element.content.is_empty() {
        output.push_str("/>");
        return Ok(());
    }
    output.push('>');
    for content in &element.content {
        match content {
            TextIndexContent::Text(text) => escape_text(text, output),
            TextIndexContent::Element(child) => write_element(child, prefixes, false, output)?,
        }
    }
    output.push_str("</");
    write_name(
        element.namespace_uri.as_deref(),
        &element.local_name,
        prefixes,
        output,
    )?;
    output.push('>');
    Ok(())
}

fn write_name(
    namespace: Option<&str>,
    local: &str,
    prefixes: &BTreeMap<String, String>,
    output: &mut String,
) -> Result<()> {
    if let Some(namespace) = namespace {
        let prefix = prefixes.get(namespace).ok_or_else(|| {
            Error::InvalidFormat(format!("missing prefix for namespace {namespace}"))
        })?;
        output.push_str(prefix);
        output.push(':');
    }
    output.push_str(local);
    Ok(())
}

fn escape_text(value: &str, output: &mut String) {
    for character in value.chars() {
        match character {
            '&' => output.push_str("&amp;"),
            '<' => output.push_str("&lt;"),
            '>' => output.push_str("&gt;"),
            _ => output.push(character),
        }
    }
}

fn escape_attr(value: &str, output: &mut String) {
    for character in value.chars() {
        match character {
            '&' => output.push_str("&amp;"),
            '<' => output.push_str("&lt;"),
            '"' => output.push_str("&quot;"),
            '\r' => output.push_str("&#13;"),
            '\n' => output.push_str("&#10;"),
            '\t' => output.push_str("&#9;"),
            _ => output.push(character),
        }
    }
}

fn invalid<T>(message: impl Into<String>) -> Result<T> {
    Err(Error::InvalidFormat(message.into()))
}

#[cfg(test)]
mod tests {
    use super::*;

    fn toc(name: &str) -> TextIndex {
        let mut source = TableOfContentsSource::new();
        source.outline_level = Some(10);
        source.title_template =
            Some(TextIndexTitleTemplate::new("Contents").with_style_name("Contents_20_Heading"));
        let mut entry = TextIndexEntryTemplate::new(1, "Contents_20_1");
        entry
            .push(TextIndexEntryToken::LinkStart { style_name: None })
            .push(TextIndexEntryToken::Text { style_name: None })
            .push(TextIndexEntryToken::TabStop(TextIndexTabStop::Right {
                leader: Some('.'),
                style_name: None,
            }))
            .push(TextIndexEntryToken::PageNumber { style_name: None })
            .push(TextIndexEntryToken::LinkEnd { style_name: None });
        source.push_entry_template(entry);
        let mut body = TextIndexBody::new();
        body.title = Some(TextIndexBodyTitle::new("Contents_Head", "Contents"));
        body.push_paragraph(TextIndexBodyParagraph::new("Cached heading  7"));
        TextIndex::table_of_contents(name, source, body).unwrap()
    }

    #[test]
    fn typed_toc_is_canonical_and_round_trips() {
        let xml = toc("TOC1").with_protected(true).to_xml_fragment().unwrap();
        assert!(xml.starts_with("<text:table-of-content xmlns:style="));
        let document = format!(
            "<office:document-content xmlns:office=\"{}\"><office:body><office:text>{xml}</office:text></office:body></office:document-content>",
            std::str::from_utf8(OFFICE).unwrap()
        );
        let parsed = validated_indexes(&document).unwrap();
        assert_eq!(parsed[0].name(), "TOC1");
        assert_eq!(
            parsed[0].body().unwrap().all_text(),
            "ContentsCached heading  7"
        );
    }

    #[test]
    fn mutation_preserves_unselected_xml_bytes() {
        let one = toc("one").to_xml_fragment().unwrap();
        let xml = format!(
            "<?xml version=\"1.0\"?><o:document-content xmlns:o=\"{}\"><o:body><o:text><!--keep--><text:p xmlns:text=\"{TEXT}\">before</text:p>{one}<text:p xmlns:text=\"{TEXT}\">after</text:p></o:text></o:body></o:document-content>",
            std::str::from_utf8(OFFICE).unwrap()
        );
        let replaced = replace_text_index_xml(&xml, "one", &toc("two")).unwrap();
        assert!(replaced.contains("<!--keep--><text:p xmlns:text="));
        assert!(replaced.contains(">before</text:p>"));
        assert!(replaced.contains(">after</text:p>"));
        let removed = remove_text_index_xml(&replaced, "two").unwrap();
        assert!(!removed.contains("table-of-content"));
    }

    #[test]
    fn hostile_template_order_and_foreign_children_are_rejected() {
        let office = std::str::from_utf8(OFFICE).unwrap();
        let xml = format!(
            "<office:document-content xmlns:office=\"{office}\" xmlns:text=\"{TEXT}\" xmlns:u=\"urn:bad\"><office:body><office:text><text:table-of-content text:name=\"x\"><text:table-of-content-source><text:index-source-styles text:outline-level=\"1\"/><text:index-title-template>x</text:index-title-template></text:table-of-content-source><text:index-body/></text:table-of-content></office:text></office:body></office:document-content>"
        );
        assert!(validated_indexes(&xml).is_err());
        let foreign = xml
            .replace(
                "<text:index-source-styles text:outline-level=\"1\"/>",
                "<u:extension/>",
            )
            .replace(
                "<text:index-title-template>x</text:index-title-template>",
                "",
            );
        assert!(validated_indexes(&foreign).is_err());
    }

    #[test]
    fn libreoffice_toc_fixture_round_trips_as_inert_markup() {
        let path = std::path::Path::new(env!("CARGO_MANIFEST_DIR")).join("../../test-data/libreoffice-core/sw/qa/extras/layout/data/toc-inline-heading-order.fodt");
        let xml = std::fs::read_to_string(path).unwrap();
        let indexes = validated_indexes(&xml).unwrap();
        assert!(!indexes.is_empty());
        let fragment = indexes[0].to_xml_fragment().unwrap();
        assert!(fragment.contains("text:index-body"));
    }

    fn simple_template(style: &str) -> TextIndexSimpleEntryTemplate {
        let mut template = TextIndexSimpleEntryTemplate::new(style);
        template
            .push(TextIndexEntryToken::LinkStart { style_name: None })
            .push(TextIndexEntryToken::Text { style_name: None })
            .push(TextIndexEntryToken::PageNumber { style_name: None })
            .push(TextIndexEntryToken::LinkEnd { style_name: None });
        template
    }

    #[test]
    fn shared_index_families_write_and_mutate_together() {
        let mut caption = IllustrationIndexSource::new();
        caption.use_caption = Some(true);
        caption.caption_sequence_name = Some("Figure".to_string());
        caption.caption_sequence_format = Some(TextIndexCaptionSequenceFormat::CategoryAndValue);
        caption.entry_template = Some(simple_template("Illustration_20_Index_20_1"));
        let illustration =
            TextIndex::illustration_index("figures", caption.clone(), TextIndexBody::new())
                .unwrap();
        let table = TextIndex::table_index("tables", caption, TextIndexBody::new()).unwrap();

        let mut objects = ObjectIndexSource::new();
        objects.use_chart_objects = Some(true);
        objects.use_math_objects = Some(false);
        objects.entry_template = Some(simple_template("Object_20_Index_20_1"));
        let object = TextIndex::object_index("objects", objects, TextIndexBody::new()).unwrap();

        let mut user = UserIndexSource::new("Glossary");
        user.use_index_marks = Some(true);
        let mut entry = TextIndexEntryTemplate::new(1, "User_20_Index_20_1");
        entry.push(TextIndexEntryToken::Text { style_name: None });
        user.push_entry_template(entry)
            .push_source_styles(TextIndexSourceStyles::new(1, vec!["Glossary".to_string()]));
        let user = TextIndex::user_index("user", user, TextIndexBody::new()).unwrap();

        let office = std::str::from_utf8(OFFICE).unwrap();
        let mut xml = format!(
            "<office:document-content xmlns:office=\"{office}\"><office:body><office:text/></office:body></office:document-content>"
        );
        for index in [&illustration, &table, &object, &user] {
            xml = insert_text_index_xml(&xml, index).unwrap();
        }
        let indexes = validated_indexes(&xml).unwrap();
        assert_eq!(
            indexes.iter().map(TextIndex::kind).collect::<Vec<_>>(),
            vec![
                TextIndexKind::Illustration,
                TextIndexKind::Table,
                TextIndexKind::Object,
                TextIndexKind::User
            ]
        );
        xml = replace_text_index_xml(
            &xml,
            "tables",
            &TextIndex::table_index(
                "new-tables",
                IllustrationIndexSource::new(),
                TextIndexBody::new(),
            )
            .unwrap(),
        )
        .unwrap();
        xml = remove_text_index_xml(&xml, "objects").unwrap();
        assert_eq!(validated_indexes(&xml).unwrap().len(), 3);
    }

    #[test]
    fn hostile_shared_sources_are_rejected() {
        let mut source = IllustrationIndexSource::new();
        source.entry_template = Some(simple_template("Index"));
        let fragment = TextIndex::illustration_index("figures", source, TextIndexBody::new())
            .unwrap()
            .to_xml_fragment()
            .unwrap();
        let wrong_order = fragment.replace("<text:index-body/>", "").replace(
            "</text:illustration-index>",
            "<text:index-body/><text:illustration-index-source/></text:illustration-index>",
        );
        let office = std::str::from_utf8(OFFICE).unwrap();
        let document = format!(
            "<office:document-content xmlns:office=\"{office}\"><office:body><office:text>{wrong_order}</office:text></office:body></office:document-content>"
        );
        assert!(validated_indexes(&document).is_err());

        let mut user = UserIndexSource::new("Glossary");
        user.push_source_styles(TextIndexSourceStyles::new(1, vec![]));
        let user = TextIndex::user_index("user", user, TextIndexBody::new())
            .unwrap()
            .to_xml_fragment()
            .unwrap();
        let missing_name = user.replace(" text:index-name=\"Glossary\"", "");
        let document = format!(
            "<office:document-content xmlns:office=\"{office}\"><office:body><office:text>{missing_name}</office:text></office:body></office:document-content>"
        );
        assert!(validated_indexes(&document).is_err());
    }

    #[test]
    fn libreoffice_table_index_source_round_trips_after_required_body_is_supplied() {
        let path = std::path::Path::new(env!("CARGO_MANIFEST_DIR")).join("../../test-data/libreoffice-core/sw/qa/extras/uiwriter/data/IndexElementsInHiddenSections.fodt");
        let producer_xml = std::fs::read_to_string(path).unwrap();
        assert!(producer_xml.contains("<text:table-index-source text:use-caption=\"false\""));
        let start = producer_xml.find("<text:table-index>").unwrap();
        let closing = "</text:table-index>";
        let end = producer_xml[start..].find(closing).unwrap() + start + closing.len();
        let schema_complete_index = producer_xml[start..end]
            .replace(
                "<text:table-index>",
                "<text:table-index text:name=\"TableIndex1\">",
            )
            .replace(
                "</text:table-index>",
                "<text:index-body/></text:table-index>",
            );
        let office = std::str::from_utf8(OFFICE).unwrap();
        let schema_complete = format!(
            "<office:document-content xmlns:office=\"{office}\" xmlns:text=\"{TEXT}\" xmlns:style=\"{STYLE}\"><office:body><office:text>{schema_complete_index}</office:text></office:body></office:document-content>"
        );
        let indexes = validated_indexes(&schema_complete).unwrap();
        let table = indexes
            .iter()
            .find(|index| index.kind() == TextIndexKind::Table)
            .unwrap();
        let fragment = table.to_xml_fragment().unwrap();
        assert!(fragment.contains("text:caption-sequence-name=\"Table\""));
        assert!(fragment.contains("text:table-index-entry-template"));
    }

    #[test]
    fn alphabetical_and_bibliography_write_and_mutate() {
        let mut alphabetical = AlphabeticalIndexSource::new();
        alphabetical.alphabetical_separators = Some(true);
        alphabetical.combine_entries = Some(true);
        alphabetical.language = Some("en".to_string());
        alphabetical.country = Some("US".to_string());
        alphabetical.rfc_language_tag = Some("en-US".to_string());
        let mut separator = TextAlphabeticalIndexEntryTemplate::new(
            TextAlphabeticalIndexLevel::Separator,
            "Alphabetical_20_Separator",
        );
        separator.push(TextIndexEntryToken::Text { style_name: None });
        let mut level = TextAlphabeticalIndexEntryTemplate::new(
            TextAlphabeticalIndexLevel::Level1,
            "Alphabetical_20_1",
        );
        level
            .push(TextIndexEntryToken::Text { style_name: None })
            .push(TextIndexEntryToken::TabStop(TextIndexTabStop::Right {
                leader: Some('.'),
                style_name: None,
            }))
            .push(TextIndexEntryToken::PageNumber { style_name: None });
        alphabetical
            .push_entry_template(separator)
            .push_entry_template(level);
        let alphabetical =
            TextIndex::alphabetical_index("alphabetical", alphabetical, TextIndexBody::new())
                .unwrap();

        let mut bibliography = BibliographyIndexSource::new();
        bibliography.title_template = Some(TextIndexTitleTemplate::new("Bibliography"));
        let mut book =
            TextBibliographyEntryTemplate::new(TextBibliographyType::Book, "Bibliography_20_1");
        book.push(TextBibliographyEntryToken::Field {
            field: crate::BibliographyField::Identifier,
            style_name: None,
        })
        .push(TextBibliographyEntryToken::Span {
            style_name: None,
            text: ": ".to_string(),
        })
        .push(TextBibliographyEntryToken::Field {
            field: crate::BibliographyField::Issn,
            style_name: None,
        });
        bibliography.push_entry_template(book);
        let bibliography =
            TextIndex::bibliography("bibliography", bibliography, TextIndexBody::new()).unwrap();
        let office = std::str::from_utf8(OFFICE).unwrap();
        let xml = format!(
            "<office:document-content xmlns:office=\"{office}\"><office:body><office:text/></office:body></office:document-content>"
        );
        let xml = insert_text_index_xml(&xml, &alphabetical).unwrap();
        let xml = insert_text_index_xml(&xml, &bibliography).unwrap();
        let parsed = validated_indexes(&xml).unwrap();
        assert_eq!(
            parsed.iter().map(TextIndex::kind).collect::<Vec<_>>(),
            vec![TextIndexKind::Alphabetical, TextIndexKind::Bibliography]
        );
        assert!(
            bibliography
                .to_xml_fragment()
                .unwrap()
                .contains("text:bibliography-data-field=\"issn\"")
        );
    }

    #[test]
    fn hostile_distinct_index_grammars_are_rejected() {
        let mut alphabetical = AlphabeticalIndexSource::new();
        alphabetical.language = Some("bad_language".to_string());
        assert!(TextIndex::alphabetical_index("bad", alphabetical, TextIndexBody::new()).is_err());
        let mut alphabetical = AlphabeticalIndexSource::new();
        let mut template =
            TextAlphabeticalIndexEntryTemplate::new(TextAlphabeticalIndexLevel::Level1, "Index");
        template.push(TextIndexEntryToken::LinkStart { style_name: None });
        alphabetical.push_entry_template(template);
        assert!(TextIndex::alphabetical_index("bad", alphabetical, TextIndexBody::new()).is_err());

        let office = std::str::from_utf8(OFFICE).unwrap();
        let malformed = format!(
            "<office:document-content xmlns:office=\"{office}\" xmlns:text=\"{TEXT}\"><office:body><office:text><text:bibliography text:name=\"b\"><text:bibliography-source><text:bibliography-entry-template text:bibliography-type=\"invalid\" text:style-name=\"B\"><text:index-entry-text/></text:bibliography-entry-template></text:bibliography-source><text:index-body/></text:bibliography></office:text></office:body></office:document-content>"
        );
        assert!(validated_indexes(&malformed).is_err());
    }

    #[test]
    fn libreoffice_bibliography_and_configuration_round_trip_inertly() {
        let path = std::path::Path::new(env!("CARGO_MANIFEST_DIR")).join(
            "../../test-data/libreoffice-core/sw/qa/uibase/shells/data/protectedLinkCopy.fodt",
        );
        let xml = std::fs::read_to_string(path).unwrap();
        assert!(
            xml.contains("<text:bibliography-configuration text:prefix=\"[\" text:suffix=\"]\"")
        );
        let indexes = validated_indexes(&xml).unwrap();
        let bibliography = indexes
            .iter()
            .find(|index| index.kind() == TextIndexKind::Bibliography)
            .unwrap();
        let fragment = bibliography.to_xml_fragment().unwrap();
        assert!(fragment.contains("text:bibliography-type=\"article\""));
        assert!(fragment.contains("text:bibliography-data-field=\"author\""));
        assert!(fragment.contains("text:index-body"));
    }
}
