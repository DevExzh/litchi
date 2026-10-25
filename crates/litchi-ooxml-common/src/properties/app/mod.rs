//! Typed application (extended) document properties: the package-level
//! `docProps/app.xml` part of ECMA-376 Part 1, §22.2. This is the
//! OOXML-only layer behind a host's `app_props()`; the fields that
//! [`Metadata`] also carries are merged into it by the metadata
//! transaction.
//!
//! [`Props`] models the 23 scalar properties of `CT_Properties`. The vector
//! and blob properties (`HeadingPairs`, `TitlesOfParts`, `HLinks`, `DigSig`)
//! and any extension content are not modelled: the semantic metadata
//! transaction ([`crate::properties::Edit::set_app_props`]) splices only the
//! scalar elements whose value changes, so everything else in the part,
//! including their bytes, survives an edit. `CT_Properties` is an
//! `xsd:all`, so a new element can be placed anywhere; it is appended at the
//! end of the root.
//!
//! Reads are strict: a value outside its lexical space (`xsd:int`,
//! `xsd:boolean`), a repeated scalar element, element content inside a
//! scalar element, or any markup-compatibility markup is an error rather
//! than a guess. The metadata transaction ([`super::Snapshot`]) decodes the
//! part only when a caller reads or edits this family.

use litchi_core::metadata::{Field as MetadataField, Metadata};

mod codec;
mod package;

#[cfg(test)]
mod tests;

#[cfg(test)]
pub(crate) use codec::splice;
pub(crate) use codec::{
    SplicePlan, decode, encode_new, encode_new_len, encode_new_preflighted, plan_splice,
    splice_planned, validate_target,
};
pub(crate) use package::{inspect, inspect_source};
pub use package::{read, read_source_backed};

/// Transitional extended-properties namespace.
pub(crate) const NAMESPACE: &str =
    "http://schemas.openxmlformats.org/officeDocument/2006/extended-properties";
/// Strict extended-properties namespace.
pub(crate) const STRICT_NAMESPACE: &str =
    "http://purl.oclc.org/ooxml/officeDocument/extendedProperties";
/// Transitional `docPropsVTypes` namespace, declared on new parts.
pub(crate) const VT_NAMESPACE: &str =
    "http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes";
/// Strict package-level extended-properties relationship type.
pub(crate) const STRICT_RELATIONSHIP: &str =
    "http://purl.oclc.org/ooxml/officeDocument/relationships/extendedProperties";
/// Largest extended-properties part read or written.
pub(crate) const MAX_XML_BYTES: usize = 8 * 1024 * 1024;
/// Largest text value of one scalar property.
pub(crate) const MAX_TEXT_BYTES: usize = 1024 * 1024;

/// Owned extended (application) document properties.
///
/// Every field is optional; `None` means the element is absent. Public
/// fields keep struct-update syntax natural, while writes validate every
/// value before changing a package.
#[derive(Debug, Clone, Default, PartialEq, Eq)]
pub struct Props {
    /// `Template`: the template the document is based on.
    pub template: Option<String>,
    /// `Manager`: the author's manager.
    pub manager: Option<String>,
    /// `Company`: the company the document belongs to.
    pub company: Option<String>,
    /// `Pages`: the producer's page count.
    pub pages: Option<i32>,
    /// `Words`: the producer's word count.
    pub words: Option<i32>,
    /// `Characters`: the producer's character count.
    pub characters: Option<i32>,
    /// `PresentationFormat`: the intended presentation format.
    pub presentation_format: Option<String>,
    /// `Lines`: the producer's line count.
    pub lines: Option<i32>,
    /// `Paragraphs`: the producer's paragraph count.
    pub paragraphs: Option<i32>,
    /// `Slides`: the slide count.
    pub slides: Option<i32>,
    /// `Notes`: the number of slides with notes.
    pub notes: Option<i32>,
    /// `TotalTime`: total editing time in minutes.
    pub total_time: Option<i32>,
    /// `HiddenSlides`: the hidden-slide count.
    pub hidden_slides: Option<i32>,
    /// `MMClips`: the multimedia-clip count.
    pub multimedia_clips: Option<i32>,
    /// `ScaleCrop`: thumbnail display mode.
    pub scale_crop: Option<bool>,
    /// `LinksUpToDate`: whether hyperlinks are up to date.
    pub links_up_to_date: Option<bool>,
    /// `CharactersWithSpaces`: the character count including spaces.
    pub characters_with_spaces: Option<i32>,
    /// `SharedDoc`: whether the document is shared.
    pub shared_doc: Option<bool>,
    /// `HyperlinkBase`: the base for relative hyperlinks. Metadata only;
    /// nothing resolves or follows it.
    pub hyperlink_base: Option<String>,
    /// `HyperlinksChanged`: whether hyperlinks changed.
    pub hyperlinks_changed: Option<bool>,
    /// `Application`: the producing application.
    pub application: Option<String>,
    /// `AppVersion`: the producing application's version.
    pub app_version: Option<String>,
    /// `DocSecurity`: the document security level.
    pub doc_security: Option<i32>,
}

impl Props {
    /// Creates a present but empty extended-properties value.
    #[must_use]
    pub fn new() -> Self {
        Self::default()
    }

    /// Sets the template name.
    #[must_use]
    pub fn template(mut self, template: impl Into<String>) -> Self {
        self.template = Some(template.into());
        self
    }

    /// Sets the manager.
    #[must_use]
    pub fn manager(mut self, manager: impl Into<String>) -> Self {
        self.manager = Some(manager.into());
        self
    }

    /// Sets the company.
    #[must_use]
    pub fn company(mut self, company: impl Into<String>) -> Self {
        self.company = Some(company.into());
        self
    }

    /// Sets the hyperlink base.
    #[must_use]
    pub fn hyperlink_base(mut self, base: impl Into<String>) -> Self {
        self.hyperlink_base = Some(base.into());
        self
    }

    /// Encodes a standalone Transitional extended-properties document.
    ///
    /// Prefer [`crate::properties::Edit::set_app_props`] when a package is
    /// available: it edits an existing part in place and keeps everything
    /// this model does not carry.
    ///
    /// # Errors
    ///
    /// Returns an error when a text value holds a character XML 1.0 cannot
    /// carry or exceeds the text bound.
    pub fn xml(&self) -> crate::Result<String> {
        String::from_utf8(encode_new(self)?).map_err(|error| {
            crate::Error::Invalid(format!("extended-properties XML is not UTF-8: {error}"))
        })
    }
}

/// The [`Metadata`] fields the application-properties part stores.
pub const METADATA_FIELDS: [MetadataField; 11] = [
    MetadataField::Template,
    MetadataField::Manager,
    MetadataField::Company,
    MetadataField::Application,
    MetadataField::PageCount,
    MetadataField::WordCount,
    MetadataField::CharacterCount,
    MetadataField::CharacterCountWithSpaces,
    MetadataField::EditingTimeMinutes,
    MetadataField::HyperlinkBase,
    MetadataField::Security,
];

impl Props {
    /// Overwrite the application fields of `metadata` with this value's.
    ///
    /// A negative count has no [`Metadata`] form and projects as `None`.
    pub fn project_into(&self, metadata: &mut Metadata) {
        let count = |value: Option<i32>| value.and_then(|value| u32::try_from(value).ok());
        metadata.template.clone_from(&self.template);
        metadata.manager.clone_from(&self.manager);
        metadata.company.clone_from(&self.company);
        metadata.application.clone_from(&self.application);
        metadata.page_count = count(self.pages);
        metadata.word_count = count(self.words);
        metadata.character_count = count(self.characters);
        metadata.character_count_with_spaces = count(self.characters_with_spaces);
        metadata.editing_time_minutes = count(self.total_time);
        metadata.hyperlink_base.clone_from(&self.hyperlink_base);
        metadata.security = count(self.doc_security);
    }

    /// Merge the application fields of `target` onto `base`.
    ///
    /// A field whose projection is unchanged keeps `base`'s value exactly
    /// (including a negative count); every element [`Metadata`] does not
    /// carry keeps `base`'s value.
    ///
    /// # Errors
    ///
    /// Returns [`crate::Error::Invalid`] naming a changed count above
    /// `xsd:int`'s maximum.
    pub fn merged(base: Option<&Self>, target: &Metadata) -> crate::Result<Self> {
        let mut props = base.cloned().unwrap_or_default();
        let mut before = Metadata::default();
        props.project_into(&mut before);
        fn count(field: MetadataField, value: Option<u32>) -> crate::Result<Option<i32>> {
            value
                .map(|value| {
                    i32::try_from(value).map_err(|_| {
                        crate::Error::Invalid(format!(
                            "metadata field `{field}` value {value} exceeds the xsd:int range of application properties"
                        ))
                    })
                })
                .transpose()
        }
        for field in before.changed_fields(target).collect::<Vec<_>>() {
            match field {
                MetadataField::Template => props.template.clone_from(&target.template),
                MetadataField::Manager => props.manager.clone_from(&target.manager),
                MetadataField::Company => props.company.clone_from(&target.company),
                MetadataField::Application => props.application.clone_from(&target.application),
                MetadataField::PageCount => props.pages = count(field, target.page_count)?,
                MetadataField::WordCount => props.words = count(field, target.word_count)?,
                MetadataField::CharacterCount => {
                    props.characters = count(field, target.character_count)?
                },
                MetadataField::CharacterCountWithSpaces => {
                    props.characters_with_spaces =
                        count(field, target.character_count_with_spaces)?;
                },
                MetadataField::EditingTimeMinutes => {
                    props.total_time = count(field, target.editing_time_minutes)?;
                },
                MetadataField::HyperlinkBase => {
                    props.hyperlink_base.clone_from(&target.hyperlink_base)
                },
                MetadataField::Security => props.doc_security = count(field, target.security)?,
                _ => {},
            }
        }
        Ok(props)
    }
}

/// The value kind of one scalar element.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub(crate) enum Kind {
    Text,
    Int,
    Bool,
}

/// One scalar element of `CT_Properties`, in the order the schema lists them
/// (the order new parts are written in).
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub(crate) enum Field {
    Template,
    Manager,
    Company,
    Pages,
    Words,
    Characters,
    PresentationFormat,
    Lines,
    Paragraphs,
    Slides,
    Notes,
    TotalTime,
    HiddenSlides,
    MultimediaClips,
    ScaleCrop,
    LinksUpToDate,
    CharactersWithSpaces,
    SharedDoc,
    HyperlinkBase,
    HyperlinksChanged,
    Application,
    AppVersion,
    DocSecurity,
}

/// A borrowed scalar value.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub(crate) enum Value<'a> {
    Text(&'a str),
    Int(i32),
    Bool(bool),
}

impl Field {
    pub(crate) const ALL: [Self; 23] = [
        Self::Template,
        Self::Manager,
        Self::Company,
        Self::Pages,
        Self::Words,
        Self::Characters,
        Self::PresentationFormat,
        Self::Lines,
        Self::Paragraphs,
        Self::Slides,
        Self::Notes,
        Self::TotalTime,
        Self::HiddenSlides,
        Self::MultimediaClips,
        Self::ScaleCrop,
        Self::LinksUpToDate,
        Self::CharactersWithSpaces,
        Self::SharedDoc,
        Self::HyperlinkBase,
        Self::HyperlinksChanged,
        Self::Application,
        Self::AppVersion,
        Self::DocSecurity,
    ];

    pub(crate) const fn local_name(self) -> &'static str {
        match self {
            Self::Template => "Template",
            Self::Manager => "Manager",
            Self::Company => "Company",
            Self::Pages => "Pages",
            Self::Words => "Words",
            Self::Characters => "Characters",
            Self::PresentationFormat => "PresentationFormat",
            Self::Lines => "Lines",
            Self::Paragraphs => "Paragraphs",
            Self::Slides => "Slides",
            Self::Notes => "Notes",
            Self::TotalTime => "TotalTime",
            Self::HiddenSlides => "HiddenSlides",
            Self::MultimediaClips => "MMClips",
            Self::ScaleCrop => "ScaleCrop",
            Self::LinksUpToDate => "LinksUpToDate",
            Self::CharactersWithSpaces => "CharactersWithSpaces",
            Self::SharedDoc => "SharedDoc",
            Self::HyperlinkBase => "HyperlinkBase",
            Self::HyperlinksChanged => "HyperlinksChanged",
            Self::Application => "Application",
            Self::AppVersion => "AppVersion",
            Self::DocSecurity => "DocSecurity",
        }
    }

    pub(crate) const fn kind(self) -> Kind {
        match self {
            Self::Template
            | Self::Manager
            | Self::Company
            | Self::PresentationFormat
            | Self::HyperlinkBase
            | Self::Application
            | Self::AppVersion => Kind::Text,
            Self::ScaleCrop | Self::LinksUpToDate | Self::SharedDoc | Self::HyperlinksChanged => {
                Kind::Bool
            },
            Self::Pages
            | Self::Words
            | Self::Characters
            | Self::Lines
            | Self::Paragraphs
            | Self::Slides
            | Self::Notes
            | Self::TotalTime
            | Self::HiddenSlides
            | Self::MultimediaClips
            | Self::CharactersWithSpaces
            | Self::DocSecurity => Kind::Int,
        }
    }

    pub(crate) fn from_local_name(name: &[u8]) -> Option<Self> {
        Self::ALL
            .into_iter()
            .find(|field| field.local_name().as_bytes() == name)
    }

    pub(crate) fn get(self, props: &Props) -> Option<Value<'_>> {
        fn text(value: Option<&String>) -> Option<Value<'_>> {
            value.map(|value| Value::Text(value))
        }
        match self {
            Self::Template => text(props.template.as_ref()),
            Self::Manager => text(props.manager.as_ref()),
            Self::Company => text(props.company.as_ref()),
            Self::Pages => props.pages.map(Value::Int),
            Self::Words => props.words.map(Value::Int),
            Self::Characters => props.characters.map(Value::Int),
            Self::PresentationFormat => text(props.presentation_format.as_ref()),
            Self::Lines => props.lines.map(Value::Int),
            Self::Paragraphs => props.paragraphs.map(Value::Int),
            Self::Slides => props.slides.map(Value::Int),
            Self::Notes => props.notes.map(Value::Int),
            Self::TotalTime => props.total_time.map(Value::Int),
            Self::HiddenSlides => props.hidden_slides.map(Value::Int),
            Self::MultimediaClips => props.multimedia_clips.map(Value::Int),
            Self::ScaleCrop => props.scale_crop.map(Value::Bool),
            Self::LinksUpToDate => props.links_up_to_date.map(Value::Bool),
            Self::CharactersWithSpaces => props.characters_with_spaces.map(Value::Int),
            Self::SharedDoc => props.shared_doc.map(Value::Bool),
            Self::HyperlinkBase => text(props.hyperlink_base.as_ref()),
            Self::HyperlinksChanged => props.hyperlinks_changed.map(Value::Bool),
            Self::Application => text(props.application.as_ref()),
            Self::AppVersion => text(props.app_version.as_ref()),
            Self::DocSecurity => props.doc_security.map(Value::Int),
        }
    }

    /// Store a decoded lexical value, refusing one outside the element's
    /// lexical space.
    pub(crate) fn set_lexical(self, props: &mut Props, lexical: String) -> crate::Result<()> {
        match self.kind() {
            Kind::Text => {
                let slot = match self {
                    Self::Template => &mut props.template,
                    Self::Manager => &mut props.manager,
                    Self::Company => &mut props.company,
                    Self::PresentationFormat => &mut props.presentation_format,
                    Self::HyperlinkBase => &mut props.hyperlink_base,
                    Self::Application => &mut props.application,
                    _ => &mut props.app_version,
                };
                *slot = Some(lexical);
            },
            Kind::Int => {
                let value = parse_int(&lexical).ok_or_else(|| {
                    crate::Error::Invalid(format!(
                        "extended property {} is not an xsd:int: {lexical:?}",
                        self.local_name()
                    ))
                })?;
                let slot = match self {
                    Self::Pages => &mut props.pages,
                    Self::Words => &mut props.words,
                    Self::Characters => &mut props.characters,
                    Self::Lines => &mut props.lines,
                    Self::Paragraphs => &mut props.paragraphs,
                    Self::Slides => &mut props.slides,
                    Self::Notes => &mut props.notes,
                    Self::TotalTime => &mut props.total_time,
                    Self::HiddenSlides => &mut props.hidden_slides,
                    Self::MultimediaClips => &mut props.multimedia_clips,
                    Self::CharactersWithSpaces => &mut props.characters_with_spaces,
                    _ => &mut props.doc_security,
                };
                *slot = Some(value);
            },
            Kind::Bool => {
                let value = parse_bool(&lexical).ok_or_else(|| {
                    crate::Error::Invalid(format!(
                        "extended property {} is not an xsd:boolean: {lexical:?}",
                        self.local_name()
                    ))
                })?;
                let slot = match self {
                    Self::ScaleCrop => &mut props.scale_crop,
                    Self::LinksUpToDate => &mut props.links_up_to_date,
                    Self::SharedDoc => &mut props.shared_doc,
                    _ => &mut props.hyperlinks_changed,
                };
                *slot = Some(value);
            },
        }
        Ok(())
    }
}

/// `xsd:int` after `xsd:whiteSpace="collapse"`: an optional sign and
/// decimal digits.
fn parse_int(lexical: &str) -> Option<i32> {
    let value = lexical.trim_matches([' ', '\t', '\n', '\r']);
    let digits = value
        .strip_prefix('+')
        .or_else(|| value.strip_prefix('-'))
        .unwrap_or(value);
    if digits.is_empty() || !digits.bytes().all(|byte| byte.is_ascii_digit()) {
        return None;
    }
    value.strip_prefix('+').unwrap_or(value).parse().ok()
}

/// `xsd:boolean` after `xsd:whiteSpace="collapse"`.
fn parse_bool(lexical: &str) -> Option<bool> {
    match lexical.trim_matches([' ', '\t', '\n', '\r']) {
        "true" | "1" => Some(true),
        "false" | "0" => Some(false),
        _ => None,
    }
}
