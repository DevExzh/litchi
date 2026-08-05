//! OpenDocument Text (.odt) implementation.
//!
//! This module provides comprehensive support for parsing, creating, and manipulating
//! OpenDocument Text documents (.odt files), which are the open standard
//! equivalent of Microsoft Word documents.
//!
//! # Implementation Progress
//!
//! ## ✅ Reading (`document.rs`, `parser.rs`) - COMPLETE
//! - ✅ `Document::open()` - Load from file path
//! - ✅ `Document::from_bytes()` - Load from memory
//! - ✅ `text()` - Extract all text content
//! - ✅ `paragraphs()` - Parse all paragraphs with formatting
//! - ✅ `tables()` - Parse tables with nested structure support
//! - ✅ `lists()` - Parse ordered and unordered lists
//! - ✅ `headings()` - Extract heading hierarchy
//! - ✅ `metadata()` - Extract document metadata
//! - ✅ `hyperlinks()` - Extract all hyperlinks
//! - ✅ `bookmarks()` - Parse bookmarks and references
//! - ✅ `footnotes()` / `endnotes()` - Parse notes
//! - ✅ `comments()` - Parse document comments
//! - ✅ `track_changes()` - Parse tracked changes
//! - ✅ `sections()` - Parse document sections
//! - ✅ Section protection keys, visibility conditions, linked sources, and inert DDE
//! - ✅ `text_indexes()` - Parse all generated index sources and cached bodies
//! - ✅ `text_index_marks()` - Parse TOC, user, alphabetical, and bibliography marks
//! - ✅ `reference_marks()` - Parse point/range cross-reference targets
//! - ✅ `ruby_annotations()` / `rubies()` - Parse structure-preserving and simplified ruby pairs
//! - ✅ `forms()` - Inert classic form controls (`office:forms`): text, textarea,
//!   checkbox, button, combobox/listbox with options, radio, fixed-text, hidden,
//!   and number/date/time controls with ids, labels, current state, and inert
//!   event-listener metadata
//! - ✅ `page_sequence()` - Parse explicit `text:page-sequence` master-page assignments
//! - ✅ Style parsing and resolution with registry
//!
//! ## ✅ Writing (`builder.rs`, `mutable.rs`) - COMPLETE
//! - ✅ `Builder::new()` - Create new documents
//! - ✅ `add_paragraph()` - Add paragraphs with text
//! - ✅ `add_hyperlink()` / `add_hyperlink_element()` - Add inert simple hyperlinks
//! - ✅ `add_note()` - Author plain-text or validated structured footnotes and endnotes
//! - ✅ `add_ruby_annotation()` / `add_ruby_style()` - Author typed ruby annotations and styles
//! - ✅ `add_table()` - Add tables with rows/cells
//! - ✅ `add_list()` - Add lists
//! - ✅ `add_heading()` - Add headings with levels
//! - ✅ `add_control_form()` and family APIs - Author typed forms and controls
//! - ✅ `set_title()` / `set_author()` - Set metadata
//! - ✅ `save()` / `to_bytes()` - Write to file or bytes
//! - ✅ `MutableDocument` - Modify existing documents
//! - ✅ Insert, replace, and remove form controls and form properties
//! - ✅ `set_page_sequence()` - Author, replace, and remove explicit page sequences
//!
//! ## 🚧 TODO - Advanced Features
//! - ⚠️ Table of contents generation
//! - ⚠️ Index creation (alphabetical, figure, table indexes)
//! - ⚠️ Mail merge and field replacement
//! - ⚠️ Advanced drawing objects (complex shapes, connectors)
//! - ⚠️ Master page editing
//! - ⚠️ Header and footer manipulation
//! - ⚠️ Document protection
//!
//! # References
//! - ODF Specification: §4-5 (Text Content)
//! - odfpy: `odf/text.py`, `odf/table.py`
//! - ODF Toolkit: Simple API - Document class

pub use litchi_odf_common::core;
pub use litchi_odf_common::{constants, coordinates, datatype, namespace};

pub use litchi_odf_common::detect;
pub mod odc;
pub use litchi_odf_common::rdf;

/// Typed values used by ODT table cells.
#[derive(Debug, Clone, PartialEq)]
pub enum CellValue {
    Empty,
    Text(String),
    Number(f64),
    Boolean(bool),
    Date(String),
    Currency(f64, String),
    Percentage(f64),
    Time(String),
}

pub mod auto_mark_file;
pub mod bibliography_configuration;
pub mod chart_properties;
pub mod content_validation;
pub mod dde_connection;
pub mod document_scripts;
pub mod drawing;
pub mod drawing_page_properties;
pub mod embedded_object;
pub mod font_face;
pub mod footnote_separator;
pub mod form;
pub mod generic;
pub mod graphic_properties;
pub mod header_footer_properties;
pub mod image_map;
pub mod line_numbering;
pub mod list_label_alignment;
pub mod list_style;
pub mod master_page;
pub mod media;
pub mod notes_configuration;
pub mod outline_style;
pub mod package;
pub mod ruby_family;
pub mod section_properties;
pub mod settings;
pub mod style;
pub mod variable_declaration;

#[allow(
    unused_imports,
    reason = "ODT facade exposes inert concordance metadata"
)]
pub(crate) use auto_mark_file::AlphabeticalIndexAutoMarkFile;
#[allow(unused_imports, reason = "ODT facade exposes bibliography semantics")]
pub(crate) use bibliography_configuration::{
    BibliographyConfiguration, BibliographyField, BibliographySortKey,
};
#[allow(
    unused_imports,
    reason = "ODT facade exposes unambiguous chart property semantics"
)]
pub(crate) use chart_properties::{
    Angle, AxisLabelPosition, AxisPosition, DataLabelNumber, Direction, Double, EmptyCellTreatment,
    ErrorCategory, Integer, Interpolation, LabelArrangement, LabelPosition, LabelSeparator,
    PositiveInteger, RegressionType, SeriesSource, SolidType, StylePropertiesSet, StyleRecord,
    SymbolImage, SymbolName, SymbolType, TickMarkPosition, parse_chart_style_properties,
    set_chart_style_properties_xml,
};
#[allow(
    unused_imports,
    reason = "ODT facade exposes the canonical shared package types"
)]
pub(crate) use core::{
    AutoReloadMetadata, Cipher, DocumentStatistics, HyperlinkBehaviourMetadata, Kdf, Manifest,
    ManifestChecksum, ManifestChecksumAlgorithm, ManifestEncryption, ManifestEncryptionAlgorithm,
    ManifestEntry, ManifestKeyDerivation, ManifestStartKeyGeneration, Metadata, OwnedPackage,
    PackageWriter, Profile, StartKey, Structure, TemplateMetadata, UserDefinedMetadata,
    UserDefinedValueType,
};
#[allow(
    unused_imports,
    reason = "ODT facade exposes DDE metadata used by variables"
)]
pub(crate) use dde_connection::{DdeConnectionDeclaration, DdeConnectionUse};
#[allow(unused_imports, reason = "ODT facade exposes document script metadata")]
pub(crate) use document_scripts::{
    EmbeddedScript, EventListener, ScriptBinding, ScriptEventListener, Scripts, parse_scripts,
};
#[allow(unused_imports, reason = "ODT facade exposes drawing style resources")]
pub(crate) use drawing_page_properties::{
    BackgroundSize, Color, Duration, Fill, FillRule, ImageRefPoint, LengthOrPercent,
    NonNegativeInteger, Percent, Repeat, Sound, SoundShow, StyleNameRef, StyleProperties, Styles,
    TileDirection, TileRepeatOffset, TransitionDirection, TransitionSpeed, TransitionStyle,
    TransitionType, Visibility, parse_drawing_page_style_properties,
    set_drawing_page_style_properties_xml,
};
#[allow(
    unused_imports,
    reason = "ODT facade exposes bookmark mutation primitives"
)]
pub(crate) use elements::bookmark::{
    Bookmark, BookmarkEnd, BookmarkFragments, BookmarkParser, BookmarkRange, BookmarkStart,
    BookmarkTarget, insert_bookmark_xml, parse_bookmark_targets, remove_bookmark_xml,
    replace_bookmark_xml,
};
#[allow(
    unused_imports,
    reason = "ODT facade exposes inert embedded-resource models"
)]
pub(crate) use embedded_object::{
    EmbeddedObject, EmbeddedObjectKind, EmbeddedObjectPart, EmbeddedObjectSource, InlineObjectRoot,
};
#[allow(
    unused_imports,
    reason = "ODT facade exposes footnote separator semantics"
)]
pub(crate) use footnote_separator::{
    FootnoteSeparatorAdjustment, FootnoteSeparatorLength, FootnoteSeparatorLineStyle,
    FootnoteSeparatorPercent, StyleFootnoteSeparator, parse_style_footnote_separators,
};
#[allow(unused_imports, reason = "ODT facade exposes typed form models")]
pub(crate) use form::{
    ButtonControl, ButtonType, CheckboxControl, CheckboxState, ComboItem, ComboboxControl,
    ConnectionResourceForm, ControlForm, ControlRef, ControlShape, FileControl, FixedTextControl,
    Form, FormAttribute, FormConnectionResource, FormControl, FormControlKind, FormDate,
    FormDouble, FormGroup, FormNode, FormPart, FormProperty, FormPropertyValue, FormScalarValue,
    FormScope, Forms, FrameControl, GenericControl, GenericControlMetadata, GenericForm,
    GenericFormControl, GridColumn, GridColumnControl, GridColumnControlKind, GridControl,
    GridForm, GridNonNegativeInteger, HiddenControl, ImageButtonType, ImageControl,
    ImageFrameControl, ImageFrameForm, InteractiveControl, InteractiveForm, ListLinkageType,
    ListOption, ListSourceType, ListboxControl, OwnedFormConnectionResource, PasswordControl,
    PasswordFileControl, PasswordFileForm, PropertyForm, RadioControl, RadioVisualEffect,
    RelativeImageAlign, RelativeImagePosition, SelectionControl, SelectionForm, TextControl,
    TextControlKind, TypedValueBound, TypedValueControl, TypedValueControlKind, TypedValueDuration,
    TypedValueForm, TypedValueNonNegativeInteger, ValueRangeControl, ValueRangeDuration,
    ValueRangeForm, ValueRangeInteger, ValueRangeNonNegativeInteger, ValueRangeOrientation,
    ValueRangePositiveInteger, VisualControl, VisualForm, form_connection_resources,
    form_properties, generic_form_controls, grid_controls, image_frame_controls,
    insert_form_connection_resource_xml, insert_form_property_xml, insert_generic_form_control_xml,
    insert_grid_control_xml, insert_image_frame_control_xml, insert_interactive_control_xml,
    insert_password_file_control_xml, insert_selection_control_xml, insert_text_control_xml,
    insert_typed_value_control_xml, insert_value_range_control_xml, insert_visual_control_xml,
    interactive_controls, password_file_controls, remove_form_connection_resource_xml,
    remove_form_property_xml, remove_generic_form_control_xml, remove_grid_control_xml,
    remove_image_frame_control_xml, remove_interactive_control_xml,
    remove_password_file_control_xml, remove_selection_control_xml, remove_text_control_xml,
    remove_typed_value_control_xml, remove_value_range_control_xml, remove_visual_control_xml,
    replace_form_connection_resource_xml, replace_form_property_xml,
    replace_generic_form_control_xml, replace_grid_control_xml, replace_image_frame_control_xml,
    replace_interactive_control_xml, replace_password_file_control_xml,
    replace_selection_control_xml, replace_text_control_xml, replace_typed_value_control_xml,
    replace_value_range_control_xml, replace_visual_control_xml, selection_controls, text_controls,
    typed_value_controls, value_range_controls, visual_controls,
};
#[allow(
    unused_imports,
    reason = "ODT facade exposes the canonical document package types"
)]
pub(crate) use generic::{FlatOpenDocument, OpenDocumentFamily, OpenDocumentPackage};
#[allow(unused_imports, reason = "ODT facade exposes graphic style semantics")]
pub(crate) use graphic_properties::{
    Child, ChildKind, Kind, Namespace, Properties, Property, Value, parse_graphic_style_properties,
    set_graphic_style_properties_xml,
};
#[allow(
    unused_imports,
    reason = "ODT facade exposes header and footer style configurations"
)]
pub(crate) use header_footer_properties::{HeaderFooterStyleProperties, PageHeaderFooterRegion};
#[allow(unused_imports, reason = "ODT facade exposes line-numbering semantics")]
pub(crate) use line_numbering::{
    LineNumberFormat, LineNumberPosition, LineNumberingConfiguration, LineNumberingSeparator,
    NonNegativeLength, parse_line_numbering_configuration,
};
#[allow(
    unused_imports,
    reason = "ODT facade exposes list and outline semantics"
)]
pub(crate) use list_label_alignment::{
    LabelFollowedBy, ListLabelLength, ListLevelLabelAlignment, ListLevelLabelAlignmentSet,
    ListStyleKind, ListStyleLevelLabelAlignment, parse_list_level_label_alignments,
};
#[allow(unused_imports, reason = "ODT facade exposes list style semantics")]
pub(crate) use list_style::{
    BulletRelativeSize, ListLevelBulletStyle, ListLevelImageSource, ListLevelKind,
    ListLevelNumberStyle, ListLevelStyle, ListStyle, ListStyleSet, MAX_LIST_STYLE_LEVEL,
    parse_list_styles,
};
#[allow(unused_imports, reason = "ODT facade exposes annotation semantics")]
pub(crate) use litchi_odf_common::annotation::{AnnotationElement, AnnotationNode, CellAnnotation};
#[allow(
    unused_imports,
    reason = "ODT facade exposes master-page package mutations"
)]
pub(crate) use master_page::{
    insert_master_page_xml, remove_master_page_xml, replace_master_page_xml,
};
#[allow(unused_imports, reason = "ODT facade exposes package media models")]
pub(crate) use media::{Image, ImageFrame, ImagePart, ImageSource};
#[allow(
    unused_imports,
    reason = "ODT facade exposes ODF chart models for embedded charts"
)]
pub(crate) use odc::{
    Attribute, Axis, AxisSpec, AxisUpdate, CachedCell, CachedRow, CachedTable, CachedValue,
    DataLabelSpec, DataPoint, DataPointSpec, DataSourceLabels, Definition, Dimension, DomainSpec,
    Element, ElementKind, EquationSpec, ExtensionAttribute, ExtensionElement, Extensions, Grid,
    GridClass, GridSpec, Legend, LegendPosition, LegendSpec, PlotArea, PlotAreaSpec,
    RegressionSpec, Series, SeriesSpec, SeriesUpdate, StyleElement, Text, serialize_chart_content,
};
#[allow(unused_imports, reason = "ODT facade exposes outline style semantics")]
pub(crate) use outline_style::{
    ListLevelPositionMode, MAX_OUTLINE_LEVELS, OutlineAttribute, OutlineLevelStyle,
    OutlineListLevelProperties, OutlineNumberFormat, OutlinePositiveInteger, OutlineStyle,
    OutlineStyles, OutlineTextAlign, OutlineTextProperties, parse_outline_styles,
    remove_outline_style_xml, set_outline_style_xml,
};
#[allow(
    unused_imports,
    reason = "ODT facade exposes package annotation operations"
)]
pub(crate) use package::annotation::{
    Annotation, AnnotationAnchor, AnnotationInfo, AnnotationPosition, AnnotationUpdate,
};
#[allow(
    unused_imports,
    reason = "ODT facade exposes embedded-chart storage policy"
)]
pub(crate) use package::charts::EmbeddedChartStorage;
#[allow(
    unused_imports,
    reason = "ODT facade exposes inert embedded-resource updates"
)]
pub(crate) use package::embedded::{
    EmbeddedResource, EmbeddedResourceFile, EmbeddedResourceKind, EmbeddedResourceSource,
};
#[allow(unused_imports, reason = "ODT facade exposes authored form operations")]
pub(crate) use package::forms::{AuthoredForm, AuthoredFormControl, AuthoredFormNode};
#[allow(
    unused_imports,
    reason = "ODT facade exposes inert package script resources"
)]
pub(crate) use package::scripts::{ScriptResource, ScriptResourceKind, ScriptResourceSpec};
#[allow(unused_imports, reason = "ODT facade exposes ruby semantics")]
pub(crate) use ruby_family::{
    RubyAlignment, RubyAnnotation, RubyAnnotations, RubyBase, RubyPosition, RubyProperties,
    RubyStyle, RubyStyles, insert_ruby_annotation_xml, parse_ruby_annotations, parse_ruby_styles,
    remove_ruby_annotation_xml, remove_ruby_style_xml, replace_ruby_annotation_xml,
    set_ruby_style_xml, wrap_ruby_annotation_xml,
};
#[allow(
    unused_imports,
    reason = "ODT facade exposes common style configurations"
)]
pub(crate) use section_properties::{BackgroundRepeat, SectionBackgroundImage};
#[allow(unused_imports, reason = "ODT facade exposes section style semantics")]
pub(crate) use section_properties::{
    SectionStyleProperties, SectionStylePropertiesSet, parse_section_style_properties,
};
#[allow(unused_imports, reason = "ODT facade exposes semantic ODF settings")]
pub(crate) use settings::{
    ConfigItem, ConfigMap, ConfigMapEntry, ConfigNode, ConfigSet, ConfigValue, Settings,
};
#[allow(
    unused_imports,
    reason = "ODT facade exposes paragraph style semantics"
)]
pub(crate) use style::paragraph::border::{
    ParagraphBackgroundTransparency, ParagraphBorder, ParagraphBorderProperties,
    ParagraphBorderWidth, ParagraphBorderWidths, ParagraphStyleBorder, ParagraphStyleBorderSet,
    parse_paragraph_style_borders, set_paragraph_style_border_xml,
};
#[allow(
    unused_imports,
    reason = "ODT facade exposes paragraph break semantics"
)]
pub(crate) use style::paragraph::breaks::{
    ParagraphBreak, ParagraphBreaks, ParagraphPageNumber, ParagraphStyleBreaks,
    ParagraphStyleBreaksSet, parse_paragraph_style_breaks, set_paragraph_style_breaks_xml,
};
#[allow(
    unused_imports,
    reason = "ODT facade exposes paragraph drop-cap semantics"
)]
pub(crate) use style::paragraph::drop_cap::{
    DropCapDistance, DropCapLength, ParagraphDropCap, ParagraphStyleDropCap,
    ParagraphStyleDropCapSet, parse_paragraph_style_drop_caps,
};
#[allow(unused_imports, reason = "ODT facade exposes paragraph flow semantics")]
pub(crate) use style::paragraph::flow::{
    HyphenationKeep, HyphenationLadder, Keep, LineBreak, ParagraphFlowProperties,
    ParagraphStyleFlow, ParagraphStyleFlowSet, PunctuationWrap, parse_paragraph_style_flows,
};
#[allow(
    unused_imports,
    reason = "ODT facade exposes paragraph spacing semantics"
)]
pub(crate) use style::paragraph::line_spacing::{
    LineHeight, LineHeightPercent, LineSpacingLength, ParagraphLineSpacing,
    ParagraphStyleLineSpacing, ParagraphStyleLineSpacingSet, TextAlignLast, TextAutospace,
    parse_paragraph_style_line_spacings,
};
#[allow(
    unused_imports,
    reason = "ODT facade exposes paragraph margin semantics"
)]
pub(crate) use style::paragraph::margin::{
    ParagraphHorizontalMargin, ParagraphMargins, ParagraphStyleMargins, ParagraphStyleMarginsSet,
    ParagraphTextIndent, ParagraphVerticalMargin, parse_paragraph_style_margins,
    set_paragraph_style_margins_xml,
};
#[allow(unused_imports, reason = "ODT facade exposes paragraph tab semantics")]
pub(crate) use style::paragraph::tab_stop::{
    MAX_PARAGRAPH_TAB_STOPS, ParagraphStyleTabStopSet, ParagraphStyleTabStops,
    ParagraphTabLeaderColor, ParagraphTabLeaderStyle, ParagraphTabLeaderType,
    ParagraphTabLeaderWidth, ParagraphTabStop, ParagraphTabStopType, ParagraphTabStops,
    TabStopPosition, parse_paragraph_style_tab_stops,
};
#[allow(
    unused_imports,
    reason = "ODT facade exposes paragraph writing-mode semantics"
)]
pub(crate) use style::paragraph::writing_mode::{
    ParagraphStyleWritingMode, ParagraphStyleWritingModeSet, ParagraphWritingMode,
    ParagraphWritingModeProperties, parse_paragraph_style_writing_modes,
    set_paragraph_style_writing_mode_xml,
};
#[allow(unused_imports, reason = "ODT facade exposes text style semantics")]
pub(crate) use style::text::{
    TextProperty, TextPropertyKind, TextPropertyNamespace, TextPropertyValue, TextStyleProperties,
    TextStylePropertiesSet, TextStyleRecord, parse_text_style_properties,
    set_text_style_properties_xml,
};
pub mod elements;
#[allow(
    unused_imports,
    reason = "ODT facade exposes rich field and text element models"
)]
pub(crate) use elements::field::{
    MetaFieldAttribute, MetaFieldContent, MetaFieldElement, MetaFieldNode, NoteBodyContent,
};
#[allow(
    unused_imports,
    reason = "ODT facade exposes hyperlink element vocabulary"
)]
pub(crate) use elements::text::{Hyperlink, TextHyperlinkActuate, TextHyperlinkShow};

mod builder;
mod document;
mod dynamic_text;
pub mod frame;
pub mod header_footer;
pub mod header_footer_content;
pub mod index;
#[allow(
    unused_imports,
    reason = "internal XML helper used by moved ODF modules"
)]
pub(crate) use index::expanded_attributes;
pub mod index_mark;
pub mod mutable;
pub mod note;
pub(crate) use note::parse_notes;
pub mod page_layout;
pub mod page_sequence;
pub mod parser;
pub mod reference_mark;
pub mod ruby;
pub mod section;
pub mod tracked_changes;

pub use builder::Builder;
pub use document::Document;
#[allow(unused_imports)]
pub use dynamic_text::{
    insert_database_field_xml, insert_dynamic_text_field_xml, remove_database_field_xml,
    remove_dynamic_text_field_xml, replace_database_field_xml, replace_dynamic_text_field_xml,
};
pub(crate) use frame::FrameAnchor;
pub(crate) use header_footer::{
    HeaderFooterKind, MasterPage, MasterPageChild, MasterPageChildKind,
};
#[allow(unused_imports)] // Library public API
pub(crate) use header_footer_content::{
    HeaderFooterBlock, HeaderFooterColumnRegion, HeaderFooterField, HeaderFooterFieldKind,
    HeaderFooterInline, HeaderFooterSenderFieldKind,
};
pub(crate) use index::{
    TextBibliographyType, TextIndex, insert_text_index_xml, remove_text_index_xml,
    replace_text_index_xml,
};
pub(crate) use index_mark::{
    TextIndexMark, insert_text_index_mark_xml, remove_text_index_mark_xml,
    replace_text_index_mark_xml,
};
pub(crate) use note::{Note, NoteClass, insert_note_xml, remove_note_xml, replace_note_xml};
pub(crate) use page_layout::PageLayout;
pub(crate) use page_sequence::Sequence;
pub(crate) use reference_mark::{
    ReferenceMark, insert_reference_mark_xml, remove_reference_mark_xml, replace_reference_mark_xml,
};
pub(crate) use ruby::Ruby;
pub(crate) use section::{
    Block, add_section_xml, clear_sections_xml, remove_section_xml, unwrap_section_xml,
    update_section_xml, wrap_section_xml,
};
pub(crate) use tracked_changes::{
    Position, mark_tracked_change_range_xml, mark_tracked_deletion_xml, set_tracked_changes_xml,
    unmark_tracked_change_xml,
};

// Re-export ODT-specific types for external use
#[allow(unused_imports)] // Library public API
pub(crate) use parser::{
    ChangeType, Comment, Parser, Section, SectionDdeSource, SectionDisplay, SectionSource,
    TrackChange, TrackedChanges,
};
