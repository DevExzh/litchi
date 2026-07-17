#![allow(missing_docs)]

//! RTF (Rich Text Format) parser module.
//!
//! This module provides high-performance parsing of RTF documents with support
//! for the RTF 1.9.1 specification. It uses arena allocation (bumpalo) for efficient
//! memory management during parsing and zero-copy patterns where possible.
//!
//! # Architecture
//!
//! The parser is organized into several components:
//! - **Lexer**: Tokenizes RTF input into control words, symbols, and text
//! - **Parser**: Builds a structured document from tokens
//! - **Document**: High-level document representation with paragraphs, runs, and tables
//!
//! # Example
//!
//! ```rust,no_run
//! use litchi_rtf::RtfDocument;
//!
//! # fn main() -> Result<(), litchi_rtf::RtfError> {
//! let rtf_text = r#"{\rtf1\ansi{\fonttbl\f0\fswiss Helvetica;}\f0\pard Hello World!\par}"#;
//! let doc = RtfDocument::parse(rtf_text)?;
//! let text = doc.text();
//! # Ok::<(), litchi_rtf::RtfError>(())
//! # }
//! ```

mod annotation;
mod bookmark;
mod border;
mod compressed;
mod document;
mod document_variable;
mod error;
mod field;
mod form_field;
mod generator;
mod generated_list_marker;
mod revision_save;
mod xml_namespace;
mod theme;
mod latent_style;
mod legacy_text_box;
mod legacy_numbering;
mod paragraph_group;
mod note_options;
mod note_separator;
mod file_table;
mod data_store;
mod info;
mod lexer;
mod list;
mod language;
mod math_properties;
mod navigation_entry;
mod object;
mod parser;
mod picture;
mod section;
mod shape;
mod stylesheet;
mod table;
mod types;
mod user_property;
mod writer;

// Re-exports
pub use annotation::{Annotation, AnnotationType, Revision, RevisionAuthor, RevisionType};
pub use bookmark::{Bookmark, BookmarkTable};
pub use border::{
    Border, BorderStyle, Borders, CharacterBorder, CharacterBorderStyle, CharacterShading,
    MAX_PARAGRAPH_TAB_STOPS, Shading, ShadingPattern, TabAlignment, TabLeader, TabStop, TabStops,
};
pub use compressed::{compress, decompress, is_compressed_rtf};
pub use document::RtfDocument;
pub use document_variable::DocumentVariable;
pub use error::{RtfError, RtfResult};
pub use field::{
    Field, FieldCodeError, FieldCodeToken, FieldSwitch, FieldType, HyperlinkCode, ParsedFieldCode,
    ReferenceCode, parse_field_code,
};
pub use form_field::{FormField, FormFieldType, FormTextType};
pub use generator::DocumentGenerator;
pub use generated_list_marker::{GeneratedListMarker, GeneratedListMarkerKind};
pub use revision_save::RevisionSaveMetadata;
pub use xml_namespace::XmlNamespace;
pub use theme::DocumentTheme;
pub use latent_style::{LatentStyleException, LatentStyles};
pub use legacy_text_box::{
    LegacyHorizontalAnchor, LegacyTextBox, LegacyTextDirection, LegacyVerticalAnchor,
};
pub use legacy_numbering::{
    LegacyNumberingAlignment, LegacyNumberingFormat, LegacySectionNumbering,
    LegacySectionNumberingLevel,
};
pub use paragraph_group::{ParagraphGroupProperty, ParagraphGroupPropertyTable};
pub use note_options::{
    EndnoteRestart, FootnoteRestart, NoteNumberingStyle, NoteOptions, NotePlacement,
    PresentNoteKinds,
};
pub use section::{SectionFootnotePlacement, SectionNoteOptions};
pub use note_separator::{
    NoteSeparator, NoteSeparatorElement, NoteSeparatorKind, NoteSeparatorTable,
};
pub use file_table::{FileLocation, FileSystemValidity, FileTable, FileTableEntry};
pub use data_store::DocumentDataStore;
pub use info::{
    DocumentInfo, DocumentProtection, ProtectionLevel, ProtectionType, RtfTimestamp,
};
pub use lexer::CharacterSet;
pub use user_property::{
    UserProperty, UserPropertyDateTime, UserPropertyType, UserPropertyValue,
};
pub use list::{
    List, ListFollow, ListJustification, ListLevel, ListLevelType, ListOverride, ListOverrideLevel,
    ListOverrideTable, ListTable,
};
pub use language::{DocumentLanguageDefaults, LanguageId};
pub use math_properties::{
    DocumentMathProperties, MathBinaryOperatorBreak, MathBinarySubtractionBreak, MathFlag,
    MathJustification, MathLimitPlacement,
};
pub use navigation_entry::{
    IndexEntry, IndexPageReference, NavigationEntry, TableOfContentsEntry,
};
pub use object::{EmbeddedObject, ObjectKind, OleObjectHeader};
pub use picture::{ImageType, Picture, PictureIdentity, detect_image_type};
pub use section::{
    HeaderFooter, HeaderFooterParagraph, HeaderFooterType, Note, PageNumberFormat, PageOrientation,
    Section, SectionBreakType, SectionProperties, VerticalAlignment,
};
pub use shape::{
    Fill, FillType, GradientDirection, OfficeArtColor, OfficeArtOpacity, Shape, ShapeGeometry,
    ShapeGroup, ShapeLine, ShapeProperty, ShapeType, WrapMode,
};
pub use stylesheet::{Style, StyleSheet, StyleType};
pub use table::{Cell, Row, Table};
pub use types::{
    Alignment, AssociatedCharacterFormatting, Color, ColorRef, ColorTable, DocumentElement, Font,
    FontFamily, FontPitch, FontRef, FontTable, Formatting, Indentation, Paragraph, ParagraphContent,
    Run, Spacing, StyleBlock, TextDirection, UnderlineStyle,
};
pub use writer::{RtfWriter, WriterOptions};
