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
//! - ✅ `DocumentBuilder::new()` - Create new documents
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
pub mod style_columns;
pub mod table_cell_properties;
pub mod table_column_properties;
pub mod table_properties;
pub mod table_row_properties;
pub mod text_properties;
pub mod variable_declaration;

#[allow(
    unused_imports,
    reason = "ODT facade exposes inert concordance metadata"
)]
pub use auto_mark_file::AlphabeticalIndexAutoMarkFile;
#[allow(unused_imports, reason = "ODT facade exposes bibliography semantics")]
pub use bibliography_configuration::{
    BibliographyConfiguration, BibliographyField, BibliographySortKey,
};
#[allow(unused_imports, reason = "ODT facade exposes style geometry semantics")]
pub use chart_properties::{
    ChartAngle, ChartAxisLabelPosition, ChartAxisPosition, ChartDataLabelNumber, ChartDirection,
    ChartDouble, ChartEmptyCellTreatment, ChartErrorCategory, ChartInteger, ChartInterpolation,
    ChartLabelArrangement, ChartLabelPosition, ChartLabelSeparator, ChartNonNegativeInteger,
    ChartNonNegativeLength, ChartPercent, ChartPositiveInteger, ChartRegressionType,
    ChartSeriesSource, ChartSolidType, ChartStyleProperties, ChartStylePropertiesSet,
    ChartStyleRecord, ChartSymbolImage, ChartSymbolName, ChartSymbolType, ChartTickMarkPosition,
    parse_chart_style_properties, set_chart_style_properties_xml,
};
#[allow(
    unused_imports,
    reason = "ODT facade exposes the canonical shared package types"
)]
pub use core::{
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
pub use dde_connection::{DdeConnectionDeclaration, DdeConnectionUse};
#[allow(unused_imports, reason = "ODT facade exposes document script metadata")]
pub use document_scripts::{
    DocumentEventListener, DocumentScripts, EmbeddedScript, ScriptBinding, ScriptEventListener,
    parse_document_scripts,
};
#[allow(
    unused_imports,
    reason = "ODT facade exposes drawing resource semantics"
)]
pub use drawing::resources::{
    FillImage, FillImageActuate, FillImageLength, FillImageLengthUnit, FillImageLink,
    FillImageLinkKind, FillImageShow, FillImageSource, FillImages, Gradient, GradientAngle,
    GradientCoordinate, GradientCoordinateUnit, GradientIntensity, GradientPercent,
    GradientSpreadMethod, GradientStopOffset, Gradients, Hatch, HatchLength, HatchLengthUnit,
    HatchRotation, HatchStyle, Hatches, LegacyGradient, LegacyGradientStyle,
    LibreOfficeGradientColorType, LibreOfficeGradientStop, Marker, MarkerPathData, MarkerViewBox,
    Markers, Opacities, Opacity, OpacityAngle, OpacityGeometryPercent, OpacityPercent, OpacityStop,
    OpacityStopValue, OpacityStyle, RgbColor, StrokeDash, StrokeDashMeasure, StrokeDashMeasureUnit,
    StrokeDashStyle, StrokeDashes, SvgGradientCommon, SvgGradientStop, SvgLinearGradient,
    SvgRadialGradient, parse_drawing_fill_images, parse_drawing_gradients, parse_drawing_hatches,
    parse_drawing_markers, parse_drawing_opacities, parse_drawing_stroke_dashes,
};
#[allow(unused_imports, reason = "ODT facade exposes drawing style resources")]
pub use drawing_page_properties::{
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
pub use elements::bookmark::{
    Bookmark, BookmarkEnd, BookmarkFragments, BookmarkParser, BookmarkRange, BookmarkStart,
    BookmarkTarget, insert_bookmark_xml, parse_bookmark_targets, remove_bookmark_xml,
    replace_bookmark_xml,
};
#[allow(
    unused_imports,
    reason = "ODT facade exposes inert embedded-resource models"
)]
pub use embedded_object::{
    EmbeddedObject, EmbeddedObjectKind, EmbeddedObjectPart, EmbeddedObjectSource, InlineObjectRoot,
};
#[allow(unused_imports, reason = "ODT facade exposes font-face semantics")]
pub use font_face::{
    Face, Faces, GenericFamily, Link, Metric, MetricKind, Pitch, PositiveLength, Source, Stretch,
    Style, Variant, Weight, parse_font_face_declarations,
};
#[allow(
    unused_imports,
    reason = "ODT facade exposes footnote separator semantics"
)]
pub use footnote_separator::{
    FootnoteSeparatorAdjustment, FootnoteSeparatorLength, FootnoteSeparatorLineStyle,
    FootnoteSeparatorPercent, StyleFootnoteSeparator, parse_style_footnote_separators,
};
#[allow(unused_imports, reason = "ODT facade exposes typed form models")]
pub use form::{
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
pub use generic::{FlatOpenDocument, OpenDocumentFamily, OpenDocumentPackage};
#[allow(unused_imports, reason = "ODT facade exposes graphic style semantics")]
pub use graphic_properties::{
    GraphicProperty, GraphicPropertyChild, GraphicPropertyChildKind, GraphicPropertyKind,
    GraphicPropertyNamespace, GraphicPropertyValue, GraphicStyleProperties,
    GraphicStylePropertiesSet, GraphicStyleRecord, parse_graphic_style_properties,
    set_graphic_style_properties_xml,
};
#[allow(
    unused_imports,
    reason = "ODT facade exposes header and footer style configurations"
)]
pub use header_footer_properties::{HeaderFooterStyleProperties, PageHeaderFooterRegion};
#[allow(unused_imports, reason = "ODT facade exposes line-numbering semantics")]
pub use line_numbering::{
    LineNumberFormat, LineNumberPosition, LineNumberingConfiguration, LineNumberingSeparator,
    NonNegativeLength, parse_line_numbering_configuration,
};
#[allow(
    unused_imports,
    reason = "ODT facade exposes list and outline semantics"
)]
pub use list_label_alignment::{
    LabelFollowedBy, ListLabelLength, ListLevelLabelAlignment, ListLevelLabelAlignmentSet,
    ListStyleKind, ListStyleLevelLabelAlignment, parse_list_level_label_alignments,
};
#[allow(unused_imports, reason = "ODT facade exposes list style semantics")]
pub use list_style::{
    BulletRelativeSize, ListLevelBulletStyle, ListLevelImageSource, ListLevelKind,
    ListLevelNumberStyle, ListLevelStyle, ListStyle, ListStyleSet, MAX_LIST_STYLE_LEVEL,
    parse_list_styles,
};
#[allow(unused_imports, reason = "ODT facade exposes annotation semantics")]
pub use litchi_odf_common::annotation::{AnnotationElement, AnnotationNode, CellAnnotation};
#[allow(
    unused_imports,
    reason = "ODT facade exposes master-page package mutations"
)]
pub use master_page::{insert_master_page_xml, remove_master_page_xml, replace_master_page_xml};
#[allow(unused_imports, reason = "ODT facade exposes package media models")]
pub use media::{Image, ImageFrame, ImagePart, ImageSource};
#[allow(
    unused_imports,
    reason = "ODT facade exposes note configuration semantics"
)]
pub use notes_configuration::{
    FootnotePosition, NoteNumberingScope, NotesConfiguration, NotesConfigurations,
    parse_notes_configurations, remove_notes_configuration_xml, set_notes_configuration_xml,
};
#[allow(
    unused_imports,
    reason = "ODT facade exposes ODF chart models for embedded charts"
)]
pub use odc::{
    ChartAttribute, ChartAxis, ChartAxisDimension, ChartAxisSpec, ChartAxisUpdate, ChartCachedCell,
    ChartCachedRow, ChartCachedTable, ChartCachedValue, ChartDataLabelSpec, ChartDataPoint,
    ChartDataPointSpec, ChartDataSourceLabels, ChartDefinition, ChartDocument, ChartDomainSpec,
    ChartElement, ChartElementKind, ChartEquationSpec, ChartExtensionAttribute,
    ChartExtensionElement, ChartExtensions, ChartGrid, ChartGridClass, ChartGridSpec, ChartLegend,
    ChartLegendPosition, ChartLegendSpec, ChartPlotArea, ChartPlotAreaSpec, ChartRegressionSpec,
    ChartSeries, ChartSeriesSpec, ChartSeriesUpdate, ChartStyleElement, ChartText,
    serialize_chart_content,
};
#[allow(unused_imports, reason = "ODT facade exposes outline style semantics")]
pub use outline_style::{
    ListLevelPositionMode, MAX_OUTLINE_LEVELS, OutlineAttribute, OutlineLevelStyle,
    OutlineListLevelProperties, OutlineNumberFormat, OutlinePositiveInteger, OutlineStyle,
    OutlineStyles, OutlineTextAlign, OutlineTextProperties, parse_outline_styles,
    remove_outline_style_xml, set_outline_style_xml,
};
#[allow(
    unused_imports,
    reason = "ODT facade exposes package annotation operations"
)]
pub use package::annotation::{
    Annotation, AnnotationAnchor, AnnotationInfo, AnnotationPosition, AnnotationUpdate,
};
#[allow(
    unused_imports,
    reason = "ODT facade exposes embedded-chart storage policy"
)]
pub use package::charts::EmbeddedChartStorage;
#[allow(
    unused_imports,
    reason = "ODT facade exposes inert embedded-resource updates"
)]
pub use package::embedded::{
    EmbeddedResource, EmbeddedResourceFile, EmbeddedResourceKind, EmbeddedResourceSource,
};
#[allow(unused_imports, reason = "ODT facade exposes authored form operations")]
pub use package::forms::{AuthoredForm, AuthoredFormControl, AuthoredFormNode};
#[allow(
    unused_imports,
    reason = "ODT facade exposes inert package script resources"
)]
pub use package::scripts::{ScriptResource, ScriptResourceKind, ScriptResourceSpec};
#[allow(unused_imports, reason = "ODT facade exposes ruby semantics")]
pub use ruby_family::{
    RubyAlignment, RubyAnnotation, RubyAnnotations, RubyBase, RubyPosition, RubyProperties,
    RubyStyle, RubyStyles, insert_ruby_annotation_xml, parse_ruby_annotations, parse_ruby_styles,
    remove_ruby_annotation_xml, remove_ruby_style_xml, replace_ruby_annotation_xml,
    set_ruby_style_xml, wrap_ruby_annotation_xml,
};
#[allow(
    unused_imports,
    reason = "ODT facade exposes common style configurations"
)]
pub use section_properties::{BackgroundRepeat, SectionBackgroundImage};
#[allow(unused_imports, reason = "ODT facade exposes section style semantics")]
pub use section_properties::{
    SectionStyleProperties, SectionStylePropertiesSet, parse_section_style_properties,
};
#[allow(unused_imports, reason = "ODT facade exposes semantic ODF settings")]
pub use settings::{
    ConfigItem, ConfigMap, ConfigMapEntry, ConfigNode, ConfigSet, ConfigValue, Settings,
};
#[allow(
    unused_imports,
    reason = "ODT facade exposes paragraph style semantics"
)]
pub use style::paragraph::border::{
    ParagraphBackgroundTransparency, ParagraphBorder, ParagraphBorderProperties,
    ParagraphBorderWidth, ParagraphBorderWidths, ParagraphStyleBorder, ParagraphStyleBorderSet,
    parse_paragraph_style_borders, set_paragraph_style_border_xml,
};
#[allow(
    unused_imports,
    reason = "ODT facade exposes paragraph break semantics"
)]
pub use style::paragraph::breaks::{
    ParagraphBreak, ParagraphBreaks, ParagraphPageNumber, ParagraphStyleBreaks,
    ParagraphStyleBreaksSet, parse_paragraph_style_breaks, set_paragraph_style_breaks_xml,
};
#[allow(
    unused_imports,
    reason = "ODT facade exposes paragraph drop-cap semantics"
)]
pub use style::paragraph::drop_cap::{
    DropCapDistance, DropCapLength, ParagraphDropCap, ParagraphStyleDropCap,
    ParagraphStyleDropCapSet, parse_paragraph_style_drop_caps,
};
#[allow(unused_imports, reason = "ODT facade exposes paragraph flow semantics")]
pub use style::paragraph::flow::{
    HyphenationKeep, HyphenationLadder, Keep, LineBreak, ParagraphFlowProperties,
    ParagraphStyleFlow, ParagraphStyleFlowSet, PunctuationWrap, parse_paragraph_style_flows,
};
#[allow(
    unused_imports,
    reason = "ODT facade exposes paragraph spacing semantics"
)]
pub use style::paragraph::line_spacing::{
    LineHeight, LineHeightPercent, LineSpacingLength, ParagraphLineSpacing,
    ParagraphStyleLineSpacing, ParagraphStyleLineSpacingSet, TextAlignLast, TextAutospace,
    parse_paragraph_style_line_spacings,
};
#[allow(
    unused_imports,
    reason = "ODT facade exposes paragraph margin semantics"
)]
pub use style::paragraph::margin::{
    ParagraphHorizontalMargin, ParagraphMargins, ParagraphStyleMargins, ParagraphStyleMarginsSet,
    ParagraphTextIndent, ParagraphVerticalMargin, parse_paragraph_style_margins,
    set_paragraph_style_margins_xml,
};
#[allow(unused_imports, reason = "ODT facade exposes paragraph tab semantics")]
pub use style::paragraph::tab_stop::{
    MAX_PARAGRAPH_TAB_STOPS, ParagraphStyleTabStopSet, ParagraphStyleTabStops,
    ParagraphTabLeaderColor, ParagraphTabLeaderStyle, ParagraphTabLeaderType,
    ParagraphTabLeaderWidth, ParagraphTabStop, ParagraphTabStopType, ParagraphTabStops,
    TabStopPosition, parse_paragraph_style_tab_stops,
};
#[allow(
    unused_imports,
    reason = "ODT facade exposes paragraph writing-mode semantics"
)]
pub use style::paragraph::writing_mode::{
    ParagraphStyleWritingMode, ParagraphStyleWritingModeSet, ParagraphWritingMode,
    ParagraphWritingModeProperties, parse_paragraph_style_writing_modes,
    set_paragraph_style_writing_mode_xml,
};
#[allow(
    unused_imports,
    reason = "ODT facade exposes named page-layout style semantics"
)]
pub use style_columns::{
    MAX_STYLE_COLUMNS, StyleColumn, StyleColumnLength, StyleColumnSeparator,
    StyleColumnSeparatorAlignment, StyleColumnSeparatorStyle, StyleColumns, parse_style_columns,
};
#[allow(unused_imports, reason = "ODT facade exposes table style semantics")]
pub use table_cell_properties::{
    CellBorder, CellBorderWidths, CellDirection, CellGlyphOrientationVertical, CellLength,
    CellProtect, CellRotationAlign, CellRotationAngle, CellTextAlignSource, CellVerticalAlign,
    CellWrapOption, TableCellProperties, TableCellStyleProperties, TableCellStylePropertiesSet,
    parse_table_cell_style_properties, set_table_cell_style_properties_xml,
};
#[allow(
    unused_imports,
    reason = "ODT facade exposes table-column style semantics"
)]
pub use table_column_properties::{
    TableColumnLength, TableColumnProperties, TableColumnRelWidth, TableColumnStyleProperties,
    TableColumnStylePropertiesSet, parse_table_column_style_properties,
    set_table_column_style_properties_xml,
};
#[allow(unused_imports, reason = "ODT facade exposes table style semantics")]
pub use table_properties::{
    TableAlignment, TableBorderModel, TablePageNumber, TableProperties, TableShadow,
    TableStyleMeasure, TableStylePercent, TableStyleProperties, TableStylePropertiesSet,
    TableStyleWidth, TableWritingMode, parse_table_style_properties,
    set_table_style_properties_xml,
};
#[allow(
    unused_imports,
    reason = "ODT facade exposes table-row style semantics"
)]
pub use table_row_properties::{
    HorizontalBackgroundPosition, TableRowBackgroundColor, TableRowBackgroundImage,
    TableRowBackgroundPosition, TableRowBackgroundRepeat, TableRowBackgroundSource, TableRowBreak,
    TableRowKeepTogether, TableRowLength, TableRowOpacity, TableRowProperties,
    TableRowStyleProperties, TableRowStylePropertiesSet, VerticalBackgroundPosition,
    parse_table_row_style_properties, set_table_row_style_properties_xml,
};
#[allow(unused_imports, reason = "ODT facade exposes text style semantics")]
pub use text_properties::{
    TextProperty, TextPropertyKind, TextPropertyNamespace, TextPropertyValue, TextStyleProperties,
    TextStylePropertiesSet, TextStyleRecord, parse_text_style_properties,
    set_text_style_properties_xml,
};
#[allow(
    unused_imports,
    reason = "ODT facade exposes semantic ODF metadata models"
)]
pub use variable_declaration::{
    VariableBody, VariableDateValue, VariableDeclaration, VariableDeclarationGroup,
    VariableDeclarations, VariableHeaderFooter, VariableKind, VariablePart, VariableScope,
    VariableValue, VariableValueType, remove_variable_declaration_group_xml,
    set_variable_declaration_group_xml,
};

pub mod elements;
#[allow(
    unused_imports,
    reason = "ODT facade exposes rich field and text element models"
)]
pub use elements::field::{
    MetaFieldAttribute, MetaFieldContent, MetaFieldElement, MetaFieldNode, NoteBodyContent,
};
#[allow(
    unused_imports,
    reason = "ODT facade exposes hyperlink element vocabulary"
)]
pub use elements::text::{Hyperlink, TextHyperlinkActuate, TextHyperlinkShow};

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

pub use builder::DocumentBuilder;
pub use document::Document;
#[allow(unused_imports)]
pub use dynamic_text::{
    insert_database_field_xml, insert_dynamic_text_field_xml, remove_database_field_xml,
    remove_dynamic_text_field_xml, replace_database_field_xml, replace_dynamic_text_field_xml,
};
pub use frame::{FrameAnchor, ImageFormat, Length};
pub use header_footer::{
    HeaderFooter, HeaderFooterKind, MasterPage, MasterPageChild, MasterPageChildKind,
};
#[allow(unused_imports)] // Library public API
pub use header_footer_content::{
    HeaderFooterBlock, HeaderFooterColumnRegion, HeaderFooterField, HeaderFooterFieldKind,
    HeaderFooterInline, HeaderFooterSenderFieldKind,
};
pub use index::{
    AlphabeticalIndexSource, BibliographyIndexSource, IllustrationIndexSource, ObjectIndexSource,
    TableIndexSource, TableOfContentsSource, TextAlphabeticalIndexEntryTemplate,
    TextAlphabeticalIndexLevel, TextBibliographyEntryTemplate, TextBibliographyEntryToken,
    TextBibliographyType, TextIndex, TextIndexAttribute, TextIndexBody, TextIndexBodyParagraph,
    TextIndexBodyTitle, TextIndexCaptionSequenceFormat, TextIndexChapterDisplay, TextIndexContent,
    TextIndexElement, TextIndexEntryTemplate, TextIndexEntryToken, TextIndexKind, TextIndexScope,
    TextIndexSimpleEntryTemplate, TextIndexSourceStyles, TextIndexTabStop, TextIndexTitleTemplate,
    UserIndexSource, insert_text_index_xml, remove_text_index_xml, replace_text_index_xml,
};
pub use index_mark::{
    TextAlphabeticalMarkMetadata, TextIndexMark, TextIndexMarkFragments, TextIndexMarkKind,
    insert_text_index_mark_xml, remove_text_index_mark_xml, replace_text_index_mark_xml,
};
pub use mutable::MutableDocument;
pub use note::{Note, NoteClass, insert_note_xml, remove_note_xml, replace_note_xml};
pub use page_layout::{PageLayout, PageLayoutAttribute, PageLayoutProperties, PageUsage};
pub use page_sequence::Sequence;
pub use reference_mark::{
    ReferenceMark, ReferenceMarkFragments, insert_reference_mark_xml, remove_reference_mark_xml,
    replace_reference_mark_xml,
};
pub use ruby::Ruby;
pub use section::{
    Block, add_section_xml, clear_sections_xml, remove_section_xml, unwrap_section_xml,
    update_section_xml, wrap_section_xml,
};
pub use tracked_changes::{
    Position, Story, mark_tracked_change_range_xml, mark_tracked_deletion_xml,
    set_tracked_changes_xml, unmark_tracked_change_xml,
};

// Re-export ODT-specific types for external use
#[allow(unused_imports)] // Library public API
pub use parser::{
    ChangeType, Comment, Parser, Section, SectionDdeSource, SectionDisplay, SectionSource,
    TrackChange, TrackedChanges,
};
