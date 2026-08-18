//! Semantic ODF metadata values and conversion to common document metadata.

use chrono::{DateTime, Utc};
use litchi_core::{Error, Metadata as CoreMetadata, Result};
use std::collections::HashMap;

/// Comprehensive ODF metadata
#[derive(Debug, Clone, Default)]
pub struct Metadata {
    /// Document title
    pub title: Option<String>,
    /// Document description
    pub description: Option<String>,
    /// Document subject
    pub subject: Option<String>,
    /// Document keywords
    pub keywords: Vec<String>,
    /// Document creator/author
    pub creator: Option<String>,
    /// Original document creator
    pub initial_creator: Option<String>,
    /// Person who last printed the document
    pub printed_by: Option<String>,
    /// Document language
    pub language: Option<String>,
    /// Entity responsible for making contributions to the document
    pub contributor: Option<String>,
    /// Entity responsible for making the document available
    pub publisher: Option<String>,
    /// Rights held in and over the document
    pub rights: Option<String>,
    /// Spatial or temporal topic coverage of the document
    pub coverage: Option<String>,
    /// File format or medium of the document
    pub format: Option<String>,
    /// Unambiguous reference to the document
    pub identifier: Option<String>,
    /// Related resource
    pub relation: Option<String>,
    /// Resource from which the document is derived
    pub source: Option<String>,
    /// Nature or genre of the document
    pub r#type: Option<String>,
    /// Creation date
    pub creation_date: Option<String>,
    /// Last modification date
    pub modification_date: Option<String>,
    /// Last print date
    pub print_date: Option<String>,
    /// Generator application
    pub generator: Option<String>,
    /// Exact non-negative editing-cycle count
    pub editing_cycles: Option<String>,
    /// Exact XML Schema duration spent editing
    pub editing_duration: Option<String>,
    /// Template reference, if present
    pub template: Option<TemplateMetadata>,
    /// Automatic reload behavior, if present
    pub auto_reload: Option<AutoReloadMetadata>,
    /// Default hyperlink behavior, if present
    pub hyperlink_behaviour: Option<HyperlinkBehaviourMetadata>,
    /// Document statistics
    pub statistics: DocumentStatistics,
    /// Custom properties
    pub custom_properties: HashMap<String, String>,
    /// Ordered, typed user-defined metadata, including duplicate names
    pub user_defined: Vec<UserDefinedMetadata>,
}

/// A `meta:user-defined` property with its exact lexical value.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct UserDefinedMetadata {
    /// Property name.
    pub name: String,
    /// Declared ODF value type. Missing declarations default to `string`.
    pub value_type: UserDefinedValueType,
    /// Exact decoded element text.
    pub value: String,
}

/// Standard ODF user-defined metadata value types.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum UserDefinedValueType {
    /// XML Schema double.
    Float,
    /// XML Schema date or dateTime.
    Date,
    /// XML Schema duration.
    Time,
    /// XML Schema boolean.
    Boolean,
    /// String value.
    String,
}

/// Metadata describing the template used to create a document.
#[derive(Debug, Clone, Default, PartialEq, Eq)]
pub struct TemplateMetadata {
    /// Template URI.
    pub href: Option<String>,
    /// Human-readable template title.
    pub title: Option<String>,
    /// Template `dateTime` lexical value.
    pub date: Option<String>,
    /// `XLink` activation behavior.
    pub actuate: Option<String>,
}

/// Metadata describing automatic document reload behavior.
#[derive(Debug, Clone, Default, PartialEq, Eq)]
pub struct AutoReloadMetadata {
    /// Reload URI.
    pub href: Option<String>,
    /// Exact XML Schema duration delay.
    pub delay: Option<String>,
    /// `XLink` show behavior.
    pub show: Option<String>,
    /// `XLink` activation behavior.
    pub actuate: Option<String>,
}

/// Metadata describing default hyperlink behavior.
#[derive(Debug, Clone, Default, PartialEq, Eq)]
pub struct HyperlinkBehaviourMetadata {
    /// Target frame name.
    pub target_frame_name: Option<String>,
    /// `XLink` show behavior.
    pub show: Option<String>,
}

/// Document statistics from metadata
#[derive(Debug, Clone, Default)]
pub struct DocumentStatistics {
    /// Number of pages
    pub page_count: Option<String>,
    /// Number of paragraphs
    pub paragraph_count: Option<String>,
    /// Number of words
    pub word_count: Option<String>,
    /// Number of characters
    pub character_count: Option<String>,
    /// Number of tables
    pub table_count: Option<String>,
    /// Number of drawing objects
    pub draw_count: Option<String>,
    /// Number of images
    pub image_count: Option<String>,
    /// Number of embedded OLE objects
    pub ole_object_count: Option<String>,
    /// Number of objects
    pub object_count: Option<String>,
    /// Number of frames
    pub frame_count: Option<String>,
    /// Number of sentences
    pub sentence_count: Option<String>,
    /// Number of syllables
    pub syllable_count: Option<String>,
    /// Number of non-whitespace characters
    pub non_whitespace_character_count: Option<String>,
    /// Number of spreadsheet rows
    pub row_count: Option<String>,
    /// Number of spreadsheet cells
    pub cell_count: Option<String>,
}

impl Metadata {
    /// Parse a date string into `DateTime<Utc>`.
    pub(crate) fn parse_date(date_str: Option<String>) -> Option<DateTime<Utc>> {
        date_str.and_then(|value| {
            crate::datatype::DateTime::decode(&value)
                .ok()
                .map(|date| date.with_timezone(&Utc))
        })
    }

    /// Convert to common metadata while reporting every required allocation.
    pub fn try_into_core(self) -> Result<CoreMetadata> {
        let Self {
            title,
            description,
            subject,
            keywords,
            creator,
            initial_creator,
            editing_cycles,
            creation_date,
            modification_date,
            print_date,
            template,
            statistics,
            generator,
            ..
        } = self;

        let (author, last_modified_by) = match (initial_creator, creator) {
            (Some(initial), creator) => (Some(initial), creator),
            (None, Some(creator)) => {
                let mut author = String::new();
                author
                    .try_reserve_exact(creator.len())
                    .map_err(|source| Error::Allocation {
                        resource: "ODF common metadata author",
                        source,
                    })?;
                author.push_str(&creator);
                (Some(author), Some(creator))
            },
            (None, None) => (None, None),
        };

        let keywords = if keywords.is_empty() {
            None
        } else if keywords.len() == 1 {
            keywords.into_iter().next()
        } else {
            let keyword_bytes = keywords
                .iter()
                .try_fold(0usize, |total, keyword| total.checked_add(keyword.len()))
                .ok_or_else(|| {
                    Error::InvalidFormat("ODF metadata keyword size overflow".to_string())
                })?;
            let bytes = keyword_bytes
                .checked_add((keywords.len() - 1).checked_mul(2).ok_or_else(|| {
                    Error::InvalidFormat("ODF metadata keyword size overflow".to_string())
                })?)
                .ok_or_else(|| {
                    Error::InvalidFormat("ODF metadata keyword size overflow".to_string())
                })?;
            let mut joined = String::new();
            joined
                .try_reserve_exact(bytes)
                .map_err(|source| Error::Allocation {
                    resource: "ODF common metadata keywords",
                    source,
                })?;
            for (index, keyword) in keywords.into_iter().enumerate() {
                if index != 0 {
                    joined.push_str(", ");
                }
                joined.push_str(&keyword);
            }
            Some(joined)
        };

        Ok(CoreMetadata {
            title,
            author,
            subject,
            keywords,
            description,
            template: template.and_then(|template| template.href),
            last_modified_by,
            revision: editing_cycles,
            created: Self::parse_date(creation_date),
            modified: Self::parse_date(modification_date),
            page_count: parse_u32_count(statistics.page_count),
            word_count: parse_u32_count(statistics.word_count),
            character_count: parse_u32_count(statistics.character_count),
            application: generator,
            last_printed_time: Self::parse_date(print_date),
            ..Default::default()
        })
    }
}

impl From<Metadata> for CoreMetadata {
    fn from(odf_meta: Metadata) -> Self {
        let author = odf_meta
            .initial_creator
            .clone()
            .or_else(|| odf_meta.creator.clone());
        CoreMetadata {
            title: odf_meta.title,
            author,
            subject: odf_meta.subject,
            keywords: if odf_meta.keywords.is_empty() {
                None
            } else {
                Some(odf_meta.keywords.join(", "))
            },
            description: odf_meta.description,
            template: odf_meta.template.and_then(|template| template.href),
            last_modified_by: odf_meta.creator,
            revision: odf_meta.editing_cycles,
            created: Metadata::parse_date(odf_meta.creation_date),
            modified: Metadata::parse_date(odf_meta.modification_date),
            page_count: parse_u32_count(odf_meta.statistics.page_count),
            word_count: parse_u32_count(odf_meta.statistics.word_count),
            character_count: parse_u32_count(odf_meta.statistics.character_count),
            application: odf_meta.generator,
            last_printed_time: Metadata::parse_date(odf_meta.print_date),
            ..Default::default()
        }
    }
}

fn parse_u32_count(source: Option<String>) -> Option<u32> {
    source.and_then(|text| text.parse().ok())
}
