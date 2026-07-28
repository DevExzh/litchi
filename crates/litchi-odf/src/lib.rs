//! OpenDocument Format (ODF) support.
//!
//! This module provides comprehensive support for parsing, creating, and manipulating OpenDocument
//! files conforming to ISO/IEC 26300 (ODF 1.2), including text documents (.odt), spreadsheets (.ods),
//! and presentations (.odp).
//!
//! # Implementation Progress
//!
//! This implementation is inspired by ODF Toolkit (Java) and odfpy (Python), aiming for a complete,
//! production-ready ODF reader/writer in Rust with high performance and memory efficiency.
//!
//! ## ✅ Core Infrastructure (COMPLETE)
//!
//! - **Package System** (`core/package.rs`)
//!   - ✅ ZIP archive reading with `Package<R>`
//!   - ✅ Manifest parsing and MIME type detection
//!   - ✅ File extraction and existence checking
//!   - ✅ Optimized zero-copy package with buffer pooling
//!   - ✅ PackageWriter for creating ODF files
//!
//! - **XML Processing** (`core/xml.rs`)
//!   - ✅ Content.xml parsing (main document content)
//!   - ✅ Styles.xml parsing (document styles)
//!   - ✅ Meta.xml parsing (document metadata)
//!   - ✅ Settings.xml support (document settings)
//!   - ✅ High-performance quick-xml based parsing
//!
//! - **Element Model** (`elements/`)
//!   - ✅ Text elements (paragraphs, spans, lists)
//!   - ✅ Table elements (tables, rows, cells)
//!   - ✅ Style elements (paragraph, text, table styles)
//!   - ✅ Draw elements (shapes, frames, images)
//!   - ✅ Field elements (date, time, page number)
//!   - ✅ Bookmark and reference support
//!   - ✅ Namespace handling for ODF XML
//!
//! - **Constants & Utilities** (`constants.rs`, `datatype.rs`)
//!   - ✅ MIME type constants and mappings
//!   - ✅ File extension detection
//!   - ✅ Standard ODF part paths
//!   - ✅ Data type conversions (Boolean, Date, DateTime, Duration)
//!   - ✅ A1 notation coordinate conversion
//!
//! ## ✅ ODT - Text Documents (COMPLETE for reading/writing)
//!
//! ### Reading (`odt/document.rs`, `odt/parser.rs`)
//! - ✅ Open from file path or bytes
//! - ✅ Full text extraction
//! - ✅ Paragraph and span parsing
//! - ✅ Table parsing (nested tables supported)
//! - ✅ List parsing (ordered and unordered)
//! - ✅ Heading hierarchy extraction
//! - ✅ Style registry and style resolution
//! - ✅ Metadata extraction
//! - ✅ Hyperlink extraction and inert `text:a` authoring
//! - ✅ Footnote and endnote support
//! - ✅ Exact note citations, fixed labels, nested bodies, and note classes
//! - ✅ Structure-preserving ruby base/pronunciation annotations and styles
//! - ✅ Bookmark and reference tracking
//! - ✅ Comment and change tracking parsing
//! - ✅ Section parsing
//! - ✅ Section protection, conditional visibility, linked sources, and inert DDE
//! - ✅ Generated indexes, source templates, and cached index bodies
//! - ✅ Point/range index source marks and inert bibliography records
//! - ✅ Point/range cross-reference targets with exact positions and text
//!
//! ### Writing (`odt/builder.rs`, `odt/mutable.rs`)
//! - ✅ DocumentBuilder for creating new ODT files
//! - ✅ Add paragraphs with text and styling
//! - ✅ Add inert simple hyperlinks with XLink metadata
//! - ✅ Add typed ruby annotations/styles and mutate them with CRUD APIs
//! - ✅ Add tables with rows and cells
//! - ✅ Add lists (ordered/unordered)
//! - ✅ Add headings with levels
//! - ✅ MutableDocument for modifying existing documents
//! - ✅ Set metadata (title, author, description, etc.)
//! - ✅ Save to file or bytes
//!
//! ### TODO - Advanced Features
//! - ⚠️ Table of contents generation
//! - ⚠️ Index generation
//! - ⚠️ Mail merge and field insertion
//! - ⚠️ Drawing object support (beyond basic shapes)
//! - ⚠️ Form controls
//! - ⚠️ Master page manipulation
//!
//! ## ✅ ODS - Spreadsheets (COMPLETE for reading/writing)
//!
//! ### Reading (`ods/spreadsheet.rs`, `ods/parser.rs`)
//! - ✅ Open from file path or bytes
//! - ✅ Sheet parsing (multiple sheets)
//! - ✅ Cell value extraction (String, Number, Boolean, Date, DateTime, Duration, Percentage, Currency)
//! - ✅ Formula representation
//! - ✅ Row and column operations
//! - ✅ Cell coordinate conversion (A1 notation)
//! - ✅ CSV export
//! - ✅ Metadata extraction
//! - ✅ Style parsing
//! - ✅ Repeated cells/rows expansion
//! - ✅ Merged cell handling
//! - ✅ Global and sheet-local named ranges and expressions
//! - ✅ Cell annotations with metadata, rich text/lists, extensions, and drawing geometry
//! - ✅ Database ranges, recursive filters, sort keys, and subtotal rules
//! - ✅ Inert database query/table/SQL source metadata
//! - ✅ Spreadsheet calculation settings, label ranges, and inert consolidations
//! - ✅ Inert DDE source declarations and document-stored cached tables
//! - ✅ Data-pilot (pivot-table) sources, fields, levels, references, and groups
//!
//! ### Writing (`ods/builder.rs`, `ods/mutable.rs`)
//! - ✅ SpreadsheetBuilder for creating new ODS files
//! - ✅ Add sheets with names
//! - ✅ Set cell values (all types)
//! - ✅ Preserve and author inert rich-text `text:a` hyperlink ranges with XLink metadata
//! - ✅ Set cell formulas
//! - ✅ Set cell styles
//! - ✅ MutableSpreadsheet for modifying existing spreadsheets
//! - ✅ Insert/delete rows and columns
//! - ✅ Set metadata
//! - ✅ Create, replace, edit, remove, and round-trip cell annotations
//! - ✅ Content-validation definitions, prompts, cell bindings, and inert event metadata
//! - ✅ Document/sheet keys, direct cell protection flags, and LibreOffice permissions
//! - ✅ Database ranges, filters, sorting, subtotals, and inert source metadata
//! - ✅ Calculation settings, row/column label ranges, and inert consolidations
//! - ✅ Create, edit, remove, and round-trip inert DDE caches
//! - ✅ Create, edit, remove, and round-trip data-pilot tables
//! - ✅ Save to file or bytes
//!
//! ### TODO - Advanced Features
//! - ⚠️ Chart creation and parsing
//! - ⚠️ Conditional formatting
//!
//! ## ✅ ODP - Presentations (COMPLETE for reading/writing)
//!
//! ### Reading (`odp/presentation.rs`, `odp/parser.rs`)
//! - ✅ Open from file path or bytes
//! - ✅ Slide parsing
//! - ✅ Shape extraction (text boxes, images, etc.)
//! - ✅ Slide layouts
//! - ✅ Master page parsing
//! - ✅ Text extraction from slides
//! - ✅ Metadata extraction
//! - ✅ Speaker notes extraction
//! - ✅ Slide transition and automatic-timing style resolution
//! - ✅ Inert ODF/SMIL timing trees and legacy presentation effects
//! - ✅ Inert audio/video plugin references and parameters
//! - ✅ Inert shape hyperlinks, presentation actions, and script bindings
//!
//! ### Writing (`odp/builder.rs`, `odp/mutable.rs`)
//! - ✅ PresentationBuilder for creating new ODP files
//! - ✅ Add slides
//! - ✅ Add shapes (text boxes, rectangles, etc.)
//! - ✅ Set slide layouts
//! - ✅ MutablePresentation for modifying existing presentations
//! - ✅ Set metadata
//! - ✅ Speaker notes
//! - ✅ Slide transitions, timings, and transition sounds
//! - ✅ Modern and legacy animation-tree creation and namespace-preserving round trips
//! - ✅ Package-contained audio/video embedding and mutable preservation
//! - ✅ Shape hyperlink/action creation and inert round trips
//! - ✅ Save to file or bytes
//!
//! ### TODO - Advanced Features
//! - ⚠️ Custom slide layouts
//! - ⚠️ Advanced shape properties
//! - ⚠️ Slide master manipulation
//!
//! ## ✅ ODG - Drawings (READ-ONLY SEMANTIC MODEL)
//!
//! - ✅ Open drawings and drawing templates from paths, readers, or bytes
//! - ✅ Namespace-aware page and standard 2D shape parsing
//! - ✅ Nested groups, exact geometry attributes, and inert enhanced geometry
//! - ✅ Text extraction, metadata, and package-contained media access
//! - ✅ Lossless unmodified saving with exact original-byte retention
//!
//! ## ✅ ODC - Standalone Charts (READ-ONLY SEMANTIC MODEL)
//!
//! - ✅ Open charts and chart templates from paths, readers, or bytes
//! - ✅ Namespace-aware complete chart subtree and expanded attributes
//! - ✅ Titles, legends, plot areas, axes, series, domains, data points, and analytics nodes
//! - ✅ Embedded cached tables and vendor extension elements
//! - ✅ Metadata and lossless exact original-byte saving
//!
//! ## ✅ ODF - Formula Documents (SEMANTIC MODEL AND VALIDATED PACKAGE CONSTRUCTION)
//!
//! - ✅ Open formulas and formula templates from paths, readers, or bytes
//! - ✅ Create formula and formula-template packages from validated MathML roots
//! - ✅ Namespace-aware direct MathML root and complete mixed-content subtree
//! - ✅ Expanded attributes, common presentation element kinds, and vendor elements
//! - ✅ MathML annotations and inert StarMath source extraction
//! - ✅ Typed MathML tree editing: validated mutation, well-formed serialization,
//!   typed schemata builders, and atomic `set_math` package replacement
//! - ✅ Metadata and lossless exact original-byte saving
//! - ✅ Formula markup remains inert and is never evaluated
//!
//! ## ✅ ODI - Image Documents (READ-ONLY SEMANTIC MODEL)
//!
//! - ✅ Open images and image templates from paths, readers, or bytes
//! - ✅ Namespace-aware required frame and complete mixed-content subtree
//! - ✅ Linked, package-local, and embedded base64 image payloads
//! - ✅ Text boxes, objects, tables, events, image maps, contours, and vendor elements
//! - ✅ External references remain inert and are never fetched
//! - ✅ Metadata and lossless exact original-byte saving
//!
//! ## ✅ ODM - Master Documents (READ-ONLY SEMANTIC MODEL)
//!
//! - ✅ Open master documents and master templates (`.odm`/`.otm`)
//! - ✅ Reuse the complete text-document semantic reader for cached content
//! - ✅ Namespace-aware linked sections, display/protection metadata, and XLink behavior
//! - ✅ Linked external documents remain inert and are never refreshed automatically
//! - ✅ Metadata and lossless exact original-byte saving
//!
//! ## ✅ OTH - Writer/Web Templates (READ-ONLY COMPATIBILITY MODEL)
//!
//! - ✅ Open LibreOffice/odfpy/odfdo `.oth` templates from paths, readers, or bytes
//! - ✅ Namespace-aware `office:text` package validation
//! - ✅ Reuse the complete text-document semantic reader
//! - ✅ Metadata and lossless exact original-byte saving
//! - ℹ️ `text-web` is a legacy producer MIME type, not an ODF 1.3/1.4 conformance MIME
//!
//! ## ✅ ODB - Database Front Ends (READ-ONLY SEMANTIC MODEL)
//!
//! - ✅ Open database front ends from paths, readers, or bytes
//! - ✅ Namespace-aware complete database configuration subtree and expanded attributes
//! - ✅ Connections, driver settings, forms, reports, queries, tables, schemas, keys, and indices
//! - ✅ Opaque access to package-contained embedded database engine resources
//! - ✅ Metadata and lossless exact original-byte saving
//! - ✅ Connections, SQL commands, embedded engines, and macros always remain inert
//!
//! Standard flat XML documents (`.fodt`, `.fods`, `.fodp`, `.fodg`, `.fodc`,
//! and `.fodi`) are validated and saved losslessly through [`FlatOpenDocument`].
//! The odfdo-compatible extended `.fodf` convention is accepted as well;
//! conforming packaged `.odf` formulas use a direct MathML `math:math` root.
//!
//! ## 🚧 Advanced Features (PLANNED)
//!
//! ### Embedded Objects
//! - 🔲 Embedded images (basic support exists, advanced features needed)
//! - 🔲 Embedded charts in documents/spreadsheets
//! - 🔲 Embedded objects (OLE)
//! - 🔲 Embedded videos and audio
//!
//! ### Collaboration Features
//! - 🔲 Change tracking (basic parsing exists, manipulation needed)
//! - 🔲 Comments and annotations (basic parsing exists)
//! - 🔲 Version control
//! - 🔲 Document comparison
//!
//! ### Performance Optimizations
//! - 🔲 Parallel sheet parsing for large ODS files
//! - 🔲 Streaming API for memory-constrained environments
//! - 🔲 Incremental parsing for large documents
//! - 🔲 Background saving
//!
//! ### Advanced Styling
//! - 🔲 Custom style creation
//! - 🔲 Style inheritance and cascading
//! - 🔲 Page layout manipulation
//! - 🔲 Header and footer customization
//!
//! ## References
//!
//! - **ODF Toolkit** (Java) - ODFDOM framework and validation tools
//! - **odfpy** (Python) - Pure Python ODF manipulation library
//! - **ODF Specification**: ISO/IEC 26300:2015 (ODF 1.2)
//! - **calamine** (Rust): Spreadsheet parsing patterns
//!
//! # Examples
//!
//! ```no_run
//! use litchi_odf::{Document, Spreadsheet, Presentation};
//!
//! # fn main() -> litchi_core::Result<()> {
//! // Open a text document
//! let mut doc = Document::open("document.odt")?;
//! let text = doc.text()?;
//!
//! // Open a spreadsheet
//! let mut sheet = Spreadsheet::open("data.ods")?;
//! let csv = sheet.to_csv()?;
//!
//! // Open a presentation
//! let mut pres = Presentation::open("slides.odp")?;
//! let slide_count = pres.slide_count()?;
//!
//! # Ok(())
//! # }
//! ```

/// ODF constants, MIME types, and XML tags
pub mod constants;
pub mod content_validation;
/// Cell coordinate conversion utilities (A1 notation)
pub mod coordinates;
/// Core ODF parsing functionality
mod core;
/// ODF data type conversions (Boolean, Date, DateTime, Duration)
pub mod datatype;
pub mod drawing_fill_image;
pub mod drawing_gradient;
pub mod drawing_hatch;
pub mod drawing_layer;
pub mod drawing_marker;
pub mod drawing_opacity;
mod drawing_style_resources;
pub mod drawing_stroke_dash;
/// ODF XML element classes
pub mod elements;
mod footnote_separator;
mod list_label_alignment;
pub mod named_expression;
mod outline_style;
mod paragraph_drop_cap;
mod paragraph_flow;
mod paragraph_tab_stop;
mod style_columns;
mod table_properties;
mod table_row_properties;
pub use elements::bookmark::{
    Bookmark, BookmarkEnd, BookmarkFragments, BookmarkParser, BookmarkRange, BookmarkStart,
    BookmarkTarget, insert_bookmark_xml, parse_bookmark_targets, remove_bookmark_xml,
    replace_bookmark_xml,
};
pub use elements::field::{
    OdfDatabaseConnectionResource, OdfDatabaseField, OdfDatabaseFieldKind, OdfDatabaseSource,
    OdfDatabaseTableType, OdfMetaFieldAttribute, OdfMetaFieldContent, OdfMetaFieldElement,
    OdfMetaFieldNode, OdfNoteBodyContent,
};
mod chart_properties;
mod annotation_package;
mod data_pilot_package;
mod ods_definition_package;
mod dde_connection;
mod digital_signature;
mod signature_crypto;
mod document_scripts;
mod drawing_page_properties;
mod font_face;
/// Inert semantic discovery of classic ODF forms and control shapes.
mod form;
mod form_package;
/// OpenDocument formula (.odf/.otf) support.
mod formula;
/// Format-neutral package access for every OpenDocument family.
mod generic;
/// Semantic family readers for flat OpenDocument XML documents.
mod flat;
mod graphic_properties;
mod handout_master;
mod image_map;
mod line_numbering;
mod master_page;
/// Shared semantic discovery of images in OpenDocument XML and packages.
mod media;
mod notes_configuration;
mod ruby_family;
mod rdf_package;
mod settings;
mod text_properties;
mod variable_declaration;
pub use dde_connection::{OdfDdeConnectionDeclaration, OdfDdeConnectionUse};
mod bibliography_configuration;
pub use bibliography_configuration::{
    OdfBibliographyConfiguration, OdfBibliographyField, OdfBibliographySortKey,
};
mod data_styles;
pub use data_styles::*;
/// Inert semantic discovery of embedded OpenDocument and OLE objects.
mod embedded_object;
mod embedded_chart;
mod embedded_package;
/// OpenDocument database front-end (.odb) support.
mod odb;
/// OpenDocument standalone chart (.odc/.otc) support.
mod odc;
/// OpenDocument drawing (.odg/.otg) support.
mod odg;
/// OpenDocument image (.odi/.oti) support.
mod odi;
/// OpenDocument master (.odm/.otm) support.
mod odm;
/// ODF presentation (.odp) support
mod odp;
/// ODF spreadsheet (.ods) support
mod ods;
/// ODF text document (.odt) support
mod odt;
/// LibreOffice-compatible OpenDocument web template (.oth) support.
mod oth;

// Re-export common utilities for convenience
// These are used across all Office formats, not ODF-specific
pub use chart_properties::{
    ChartAngle, ChartAxisLabelPosition, ChartAxisPosition, ChartDataLabelNumber, ChartDirection,
    ChartDouble, ChartEmptyCellTreatment, ChartErrorCategory, ChartInteger, ChartInterpolation,
    ChartLabelArrangement, ChartLabelPosition, ChartLabelSeparator, ChartNonNegativeInteger,
    ChartNonNegativeLength, ChartPercent, ChartPositiveInteger, ChartRegressionType,
    ChartSeriesSource, ChartSolidType, ChartStyleProperties, ChartStylePropertiesSet,
    ChartStyleRecord, ChartSymbolImage, ChartSymbolName, ChartSymbolType, ChartTickMarkPosition,
    parse_chart_style_properties, set_chart_style_properties_xml,
};
pub use core::{
    AutoReloadMetadata, DocumentStatistics, HyperlinkBehaviourMetadata, Manifest, ManifestChecksum,
    ManifestChecksumAlgorithm, ManifestEncryption, ManifestEncryptionAlgorithm, ManifestEntry,
    ManifestKeyDerivation, ManifestStartKeyGeneration, OdfEncryptionCipher, OdfEncryptionKdf,
    OdfEncryptionProfile, OdfEncryptionStartKey, OdfMetadata, OdfStructure, OwnedPackage,
    PackageWriter, TemplateMetadata, UserDefinedMetadata, UserDefinedValueType,
};
pub use digital_signature::{OdfDigitalSignature, OdfDigitalSignatures, OdfSignatureReference};
pub use signature_crypto::{
    OdfCanonicalizationAlgorithm, OdfDocumentSigner, OdfSignatureAlgorithm,
    OdfSignatureValidity, OdfSignatureVerification,
};
pub use document_scripts::{
    OdfDocumentEventListener, OdfDocumentScripts, OdfEmbeddedScript, OdfScriptBinding,
    OdfScriptEventListener, parse_document_scripts,
};
mod script_package;
pub use script_package::{OdfScriptResource, OdfScriptResourceKind, OdfScriptResourceSpec};
pub use annotation_package::{
    OdfAnnotation, OdfAnnotationAnchor, OdfAnnotationInfo, OdfAnnotationPosition,
    OdfAnnotationUpdate,
};
pub use data_pilot_package::DataPilotTableUpdate;
pub use ods_definition_package::{DatabaseRangeUpdate, NamedDefinitionUpdate};
pub use drawing_page_properties::{
    DrawingPageBackgroundSize, DrawingPageColor, DrawingPageDuration, DrawingPageFill,
    DrawingPageFillRule, DrawingPageImageRefPoint, DrawingPageLengthOrPercent,
    DrawingPageNonNegativeInteger, DrawingPagePercent, DrawingPageRepeat, DrawingPageSound,
    DrawingPageSoundShow, DrawingPageStyle, DrawingPageStyleNameRef, DrawingPageStyleProperties,
    DrawingPageStyleSet, DrawingPageTileDirection, DrawingPageTileRepeatOffset,
    DrawingPageTransitionDirection, DrawingPageTransitionSpeed, DrawingPageTransitionStyle,
    DrawingPageTransitionType, DrawingPageVisibility, parse_drawing_page_style_properties,
    set_drawing_page_style_properties_xml,
};
pub use embedded_object::{
    OdfEmbeddedObject, OdfEmbeddedObjectKind, OdfEmbeddedObjectPart, OdfEmbeddedObjectSource,
    OdfInlineObjectRoot,
};
pub use embedded_chart::OdfEmbeddedChartStorage;
pub use embedded_package::{
    OdfEmbeddedResource, OdfEmbeddedResourceFile, OdfEmbeddedResourceKind,
    OdfEmbeddedResourceSource,
};
pub use font_face::{
    OdfFontFace, OdfFontFaceDeclarations, OdfFontFaceLink, OdfFontFaceSource, OdfFontMetric,
    OdfFontMetricKind, OdfFontPitch, OdfFontStretch, OdfFontStyle, OdfFontVariant, OdfFontWeight,
    OdfGenericFontFamily, OdfPositiveLength, parse_font_face_declarations,
};
pub use footnote_separator::{
    FootnoteSeparatorAdjustment, FootnoteSeparatorLength, FootnoteSeparatorLineStyle,
    FootnoteSeparatorPercent, StyleFootnoteSeparator, parse_style_footnote_separators,
};
pub use form::{
    OdfButtonControl, OdfButtonType, OdfCheckboxControl, OdfCheckboxState, OdfComboItem,
    OdfComboboxControl, OdfControlForm, OdfControlRef, OdfControlShape, OdfFileControl,
    OdfFixedTextControl, OdfForm, OdfFormAttribute, OdfFormControl, OdfFormControlKind,
    OdfFormGroup, OdfFormNode, OdfFormPart, OdfFormProperty, OdfFormPropertyValue,
    OdfFormScalarValue, OdfFormScope, OdfForms, OdfFrameControl, OdfGenericControl,
    OdfGenericControlMetadata, OdfGenericForm, OdfGenericFormControl, OdfHiddenControl,
    OdfImageButtonType, OdfImageControl, OdfImageFrameControl, OdfImageFrameForm,
    OdfInteractiveControl, OdfInteractiveForm, OdfListLinkageType, OdfListOption,
    OdfListSourceType, OdfListboxControl, OdfPasswordControl, OdfPasswordFileControl,
    OdfPasswordFileForm, OdfPropertyForm, OdfRadioControl, OdfRadioVisualEffect,
    OdfRelativeImageAlign, OdfRelativeImagePosition, OdfSelectionControl, OdfSelectionForm,
    OdfTextControl, OdfTextControlKind, OdfVisualControl, OdfVisualForm, form_properties,
    generic_form_controls, image_frame_controls, insert_form_property_xml,
    insert_generic_form_control_xml, insert_image_frame_control_xml,
    insert_interactive_control_xml, insert_password_file_control_xml, insert_selection_control_xml,
    insert_text_control_xml, insert_visual_control_xml, interactive_controls,
    password_file_controls, remove_form_property_xml, remove_generic_form_control_xml,
    remove_image_frame_control_xml, remove_interactive_control_xml,
    remove_password_file_control_xml, remove_selection_control_xml, remove_text_control_xml,
    remove_visual_control_xml, replace_form_property_xml, replace_generic_form_control_xml,
    replace_image_frame_control_xml, replace_interactive_control_xml,
    replace_password_file_control_xml, replace_selection_control_xml, replace_text_control_xml,
    replace_visual_control_xml, selection_controls, text_controls, visual_controls,
};
pub use form_package::{OdfAuthoredForm, OdfAuthoredFormControl, OdfAuthoredFormNode};
pub use formula::builder as mathml;
pub use formula::{FormulaDocument, MathAttribute, MathContent, MathElement, MathElementKind};
pub use formula::builder::{MathDisplay, MathVariant};
pub use generic::{FlatOpenDocument, OpenDocumentFamily, OpenDocumentPackage};
pub use flat::{
    FlatChartDocument, FlatDrawingDocument, FlatImageDocument, FlatMutableChartDocument,
    FlatMutableDrawingDocument, FlatMutablePresentation, FlatMutableSpreadsheet,
    FlatMutableTextDocument, FlatPresentation, FlatSpreadsheet, FlatTextDocument,
};
pub use graphic_properties::{
    GraphicProperty, GraphicPropertyChild, GraphicPropertyChildKind, GraphicPropertyKind,
    GraphicPropertyNamespace, GraphicPropertyValue, GraphicStyleProperties,
    GraphicStylePropertiesSet, GraphicStyleRecord, parse_graphic_style_properties,
    set_graphic_style_properties_xml,
};
pub use line_numbering::{
    OdfLineNumberFormat, OdfLineNumberPosition, OdfLineNumberingConfiguration,
    OdfLineNumberingSeparator, OdfNonNegativeLength, parse_line_numbering_configuration,
};
pub use list_label_alignment::{
    LabelFollowedBy, ListLabelLength, ListLevelLabelAlignment, ListLevelLabelAlignmentSet,
    ListStyleKind, ListStyleLevelLabelAlignment, parse_list_level_label_alignments,
};
pub use litchi_core::RGBColor as Color;
pub use litchi_core::unit::{Length, LengthUnit};
pub use media::{OdfImage, OdfImageFrame, OdfImagePart, OdfImageSource};
pub use notes_configuration::{
    OdfFootnotePosition, OdfNoteClass, OdfNoteNumberingScope, OdfNotesConfiguration,
    OdfNotesConfigurations, parse_notes_configurations, remove_notes_configuration_xml,
    set_notes_configuration_xml,
};
pub use odb::{
    DatabaseAttribute, DatabaseContent, DatabaseDocument, DatabaseElement, DatabaseElementKind,
};
pub use odb::{
    OdfDatabaseApplicationConnectionSettings, OdfDatabaseAutoIncrement,
    OdfDatabaseBooleanComparisonMode, OdfDatabaseCharacterSet, OdfDatabaseDataSourceSetting,
    OdfDatabaseDelimiter, OdfDatabaseDriverSettings, OdfDatabaseInteger, OdfDatabaseSettingType,
    OdfDatabaseTableFilter, OdfDatabaseTableSetting, OdfDatabaseTrailingSettings,
    parse_database_trailing_settings_xml, set_database_application_connection_settings_xml,
    set_database_driver_settings_xml,
};
pub use odb::{
    OdfDatabaseColumn, OdfDatabaseColumnValue, OdfDatabaseQueries, OdfDatabaseQuery,
    OdfDatabaseQueryCollection, OdfDatabaseQueryItem, OdfDatabaseQueryModel, OdfDatabaseStatement,
    OdfDatabaseTableRepresentation, OdfDatabaseTableRepresentations, OdfDatabaseUpdateTable,
    parse_database_queries_xml, parse_database_query_model_xml,
    parse_database_table_representations_xml, set_database_queries_xml,
    set_database_table_representations_xml,
};
pub use odb::{
    OdfDatabaseColumnDefinition, OdfDatabaseDataType, OdfDatabaseIndex, OdfDatabaseIndexColumn,
    OdfDatabaseKey, OdfDatabaseKeyColumn, OdfDatabaseKeyType, OdfDatabaseNullable,
    OdfDatabaseReferentialRule, OdfDatabaseSchemaDefinition, OdfDatabaseSchemaPositiveInteger,
    OdfDatabaseTableDefinition, parse_database_schema_definition_xml,
    set_database_schema_definition_xml,
};
pub use odb::{
    OdfDatabaseComponent, OdfDatabaseComponentCollection, OdfDatabaseComponentItem,
    OdfDatabaseComponentLink, OdfDatabaseComponentModel, OdfDatabaseComponentPayload,
    OdfDatabaseForms, OdfDatabaseInertAttribute, OdfDatabaseInertContent, OdfDatabaseInertElement,
    OdfDatabaseReports, parse_database_components_xml, parse_database_forms_xml,
    parse_database_reports_xml, set_database_forms_xml, set_database_reports_xml,
};
pub use odb::{
    OdfDatabaseConnectionData, OdfDatabaseConnectionSource, OdfDatabaseFileSource,
    OdfDatabaseLogin, OdfDatabaseLoginIdentity, OdfDatabasePositiveInteger,
    OdfDatabaseServerLocation, OdfDatabaseServerSource, OdfOdbConnectionResource,
    parse_database_connection_data_xml, replace_database_connection_data_xml,
};
pub use odc::{
    ChartAttribute, ChartAxis, ChartAxisDimension, ChartAxisSpec, ChartAxisUpdate, ChartCachedCell, ChartCachedRow,
    ChartCachedTable, ChartCachedValue, ChartDataLabelSpec, ChartDataPoint, ChartDataPointSpec,
    ChartDataSourceLabels, ChartDefinition, ChartDocument, ChartDomainSpec, ChartElement,
    ChartElementKind, ChartEquationSpec, ChartExtensionAttribute, ChartExtensionElement,
    ChartExtensions, ChartGrid, ChartGridClass, ChartGridSpec, ChartLegend, ChartLegendPosition,
    ChartLegendSpec, ChartPlotArea, ChartPlotAreaSpec, ChartRegressionSpec, ChartSeries,
    ChartSeriesSpec, ChartSeriesUpdate, ChartStyleElement, ChartText, serialize_chart_content,
};
pub use odg::{
    DrawingBuilder, DrawingDocument, DrawingLayer, DrawingLayerDisplay, DrawingPage,
    DrawingPageProperties, MutableDrawing,
};
pub use odi::{ImageAttribute, ImageContent, ImageDocument, ImageElement, ImageElementKind};
pub use odm::{
    MasterDocument, MasterDocumentBuilder, MasterDocumentElement, MasterSection,
    MasterSubdocument, MutableMasterDocument,
};
pub use oth::{MutableWebDocument, WebDocument, WebDocumentBuilder};
pub use outline_style::{
    MAX_OUTLINE_LEVELS, OdfListLevelPositionMode, OdfOutlineAttribute, OdfOutlineLevelStyle,
    OdfOutlineListLevelProperties, OdfOutlineNumberFormat, OdfOutlinePositiveInteger,
    OdfOutlineStyle, OdfOutlineStyles, OdfOutlineTextAlign, OdfOutlineTextProperties,
    parse_outline_styles, remove_outline_style_xml, set_outline_style_xml,
};
pub use paragraph_drop_cap::{
    DropCapDistance, DropCapLength, ParagraphDropCap, ParagraphStyleDropCap,
    ParagraphStyleDropCapSet, parse_paragraph_style_drop_caps,
};
pub use paragraph_flow::{
    HyphenationKeep, HyphenationLadder, Keep, LineBreak, ParagraphFlowProperties,
    ParagraphStyleFlow, ParagraphStyleFlowSet, PunctuationWrap, parse_paragraph_style_flows,
};
pub use paragraph_tab_stop::{
    MAX_PARAGRAPH_TAB_STOPS, OdfTabStopPosition, ParagraphStyleTabStopSet, ParagraphStyleTabStops,
    ParagraphTabLeaderColor, ParagraphTabLeaderStyle, ParagraphTabLeaderType,
    ParagraphTabLeaderWidth, ParagraphTabStop, ParagraphTabStopType, ParagraphTabStops,
    parse_paragraph_style_tab_stops,
};
pub use ruby_family::{
    RubyAlignment, RubyAnnotation, RubyAnnotations, RubyBase, RubyPosition, RubyProperties,
    RubyStyle, RubyStyles, insert_ruby_annotation_xml, parse_ruby_annotations, parse_ruby_styles,
    remove_ruby_annotation_xml, remove_ruby_style_xml, replace_ruby_annotation_xml,
    set_ruby_style_xml, wrap_ruby_annotation_xml,
};
pub use rdf_package::{OdfRdfGraph, OdfRdfObject, OdfRdfSubject, OdfRdfTriple};
pub use settings::{
    OdfConfigItem, OdfConfigMap, OdfConfigMapEntry, OdfConfigNode, OdfConfigSet, OdfConfigValue,
    OdfSettings,
};
pub use style_columns::{
    MAX_STYLE_COLUMNS, StyleColumn, StyleColumnLength, StyleColumnSeparator,
    StyleColumnSeparatorAlignment, StyleColumnSeparatorStyle, StyleColumns, parse_style_columns,
};
pub use table_properties::{
    TableAlignment, TableBorderModel, TablePageNumber, TableProperties, TableShadow,
    TableStyleMeasure, TableStylePercent, TableStyleProperties, TableStylePropertiesSet,
    TableStyleWidth, TableWritingMode, parse_table_style_properties,
    set_table_style_properties_xml,
};
pub use table_row_properties::{
    HorizontalBackgroundPosition, TableRowBackgroundColor, TableRowBackgroundImage,
    TableRowBackgroundPosition, TableRowBackgroundRepeat, TableRowBackgroundSource, TableRowBreak,
    TableRowKeepTogether, TableRowLength, TableRowOpacity, TableRowProperties,
    TableRowStyleProperties, TableRowStylePropertiesSet, VerticalBackgroundPosition,
    parse_table_row_style_properties, set_table_row_style_properties_xml,
};
pub use text_properties::{
    TextProperty, TextPropertyKind, TextPropertyNamespace, TextPropertyValue, TextStyleProperties,
    TextStylePropertiesSet, TextStyleRecord, parse_text_style_properties,
    set_text_style_properties_xml,
};
pub use variable_declaration::{
    OdfVariableBody, OdfVariableDateValue, OdfVariableDeclaration, OdfVariableDeclarationGroup,
    OdfVariableDeclarations, OdfVariableHeaderFooter, OdfVariableKind, OdfVariablePart,
    OdfVariableScope, OdfVariableValue, OdfVariableValueType,
    remove_variable_declaration_group_xml, set_variable_declaration_group_xml,
};

// Re-export main types for convenience
pub use handout_master::HandoutMaster;
pub use image_map::{ImageMap, ImageMapArea, ImageMapAreaShape};
pub use master_page::{insert_master_page_xml, remove_master_page_xml, replace_master_page_xml};
pub use odp::{
    AnimationAttribute, AnimationAttributeNamespace, AnimationKind, AnimationNode,
    CustomPresentationShow, DrawingAttribute, DrawingAttributeNamespace, DrawingHyperlink,
    DrawingShapeKind, EnhancedGeometry, EnhancedGeometryChild, EnhancedGeometryChildKind,
    HyperlinkShow, LegacyAnimationKind, LegacyAnimationNode, MediaActuate, MediaParameter,
    MediaReference, MediaShow, MutablePresentation, Presentation, PresentationAction,
    PresentationBuilder, PresentationDateTimeDeclaration, PresentationDateTimeSource,
    PresentationDeclarationBinding, PresentationDeclarationTarget, PresentationDeclarations,
    PresentationEffect, PresentationEffectDirection, PresentationEventListener,
    PresentationFeatureState, PresentationMasterPage, PresentationMeasure,
    PresentationMeasureUnit, PresentationPageLayout, PresentationPageLayouts,
    PresentationPageMetadata, PresentationPageMetadataCollection,
    PresentationPlaceholder, PresentationPlaceholderClass, PresentationSettings,
    PresentationTextDeclaration, ScriptEventListener, ShapeEventListener, SlideTransition,
    TransitionDirection, TransitionSound, TransitionSoundShow, TransitionSpeed, TransitionStyle,
    TransitionType, parse_presentation_page_layouts, parse_presentation_settings,
    remove_presentation_page_layout_xml, set_presentation_page_layout_xml,
};
pub use ods::{
    AnnotationElement, AnnotationNode, CalculationIteration, CalculationNullDate,
    CalculationSettings, Cell as SCell, CellAnnotation, CellDetective, CellHyperlink,
    CellMatrixSpan, CellMerge, CellRangeSource, CellStyleProtection, CellTextContent, CellValue,
    Column as SColumn, ConditionalCellStyle,
    ConditionalCellStyleRule, Consolidation, ConsolidationUseLabels, ContentValidation,
    DataPilotDisplayInfo, DataPilotDisplayMemberMode, DataPilotField, DataPilotFieldReference,
    DataPilotGrandTotal, DataPilotGrandTotalElement, DataPilotGrandTotalOrientation,
    DataPilotGroup, DataPilotGroupBoundary, DataPilotGroupBy, DataPilotGroups,
    DataPilotLayoutInfo, DataPilotLayoutMode, DataPilotLevel, DataPilotMember,
    DataPilotOrientation, DataPilotReferenceMemberType, DataPilotReferenceType, DataPilotSortInfo,
    DataPilotSortMode, DataPilotSortOrder, DataPilotSource, DataPilotTable, DatabaseFilter,
    DatabaseOrientation, DatabaseRange, DatabaseSort, DatabaseSortKey, DatabaseSource,
    DdeConversionMode, DdeLink, DdeSource, DetectiveDirection, DetectiveHighlightedRange,
    DetectiveOperation, DetectiveOperationKind, EmbeddedNumberBehavior, FilterCondition,
    FilterConditionSource, FilterDataType, FilterExpression, FormulaNamespace, IterationStatus,
    LabelRange, LabelRangeOrientation, MutableSpreadsheet, NamedDefinition, NamedDefinitionScope,
    NamedExpression, NamedRange, NamedRangeUsage, ProtectionKey, Row as SRow, Sheet,
    SheetPrintSettings, SheetProtection, SheetProtectionOptions, SheetScenario, SheetShape,
    SheetShapeAnchor, SheetStyle,
    SheetStyleUsage, SheetTableSource, SortOrder, Spreadsheet, SpreadsheetBuilder, OdsWorkbook,
    TableCellProtectionStyle,
    SpreadsheetCellContentChange, SpreadsheetChangeAcceptance, SpreadsheetChangeCutOff,
    SpreadsheetChangeDimension, SpreadsheetChangeInfo, SpreadsheetChangeMetadata,
    SpreadsheetDeletion, SpreadsheetInsertion, SpreadsheetMovement, SpreadsheetNestedDeletion,
    SpreadsheetProtection, SpreadsheetTrackedCell, SpreadsheetTrackedCellAddress,
    SpreadsheetTrackedCellValue, SpreadsheetTrackedChange, SpreadsheetTrackedChanges,
    SpreadsheetTrackedRangeAddress, SubtotalField, SubtotalRule, SubtotalRules, SubtotalSortGroups,
    TableGroup, TableRange, TableSourceMode, TableStructure, TableTemplate, TableTemplateAxis,
    TableTemplateStyle, TableVisibility, ValidationDisplayList, ValidationErrorMacro,
    ValidationErrorMessage, ValidationEventListener, ValidationMessage, ValidationMessageType,
    ValidationPresentationEventListener, ValidationPresentationSound,
    ValidationScriptEventListener, normalize_open_formula,
};
pub use odt::{
    AlphabeticalIndexSource, BibliographyIndexSource, ChangeType, Document, DocumentBuilder,
    HeaderFooter,
    HeaderFooterKind, IllustrationIndexSource, MasterPage, MasterPageChild, MasterPageChildKind,
    MutableDocument, Note, NoteClass, ObjectIndexSource, OdfFrameAnchor, OdfImageFormat, OdfLength,
    OdtSectionBlock,
    OdtTrackedPosition, OdtTrackedStory,
    PageLayout, PageLayoutAttribute,
    PageLayoutProperties, PageUsage, ReferenceMark, ReferenceMarkFragments, Ruby, SectionDdeSource,
    Section, SectionDisplay, SectionSource, TableIndexSource, TableOfContentsSource,
    TextAlphabeticalIndexEntryTemplate, TextAlphabeticalIndexLevel, TextAlphabeticalMarkMetadata,
    TextBibliographyEntryTemplate, TextBibliographyEntryToken, TextBibliographyType, TextIndex,
    TextIndexAttribute, TextIndexBody, TextIndexBodyParagraph, TextIndexBodyTitle,
    TextIndexCaptionSequenceFormat, TextIndexChapterDisplay, TextIndexContent, TextIndexElement,
    TextIndexEntryTemplate, TextIndexEntryToken, TextIndexKind, TextIndexMark,
    TextIndexMarkFragments, TextIndexMarkKind, TextIndexScope, TextIndexSimpleEntryTemplate,
    TrackChange, TrackedChanges,
    TextIndexSourceStyles, TextIndexTabStop, TextIndexTitleTemplate, UserIndexSource,
    insert_database_field_xml, insert_note_xml, insert_reference_mark_xml, insert_text_index_mark_xml,
    insert_text_index_xml, remove_database_field_xml, remove_reference_mark_xml,
    remove_note_xml, remove_text_index_mark_xml, remove_text_index_xml, replace_database_field_xml,
    replace_note_xml, replace_reference_mark_xml, replace_text_index_mark_xml, replace_text_index_xml,
    mark_tracked_change_range_xml, mark_tracked_deletion_xml, set_tracked_changes_xml,
    unmark_tracked_change_xml,
    add_section_xml, clear_sections_xml, remove_section_xml, unwrap_section_xml,
    update_section_xml, wrap_section_xml,
};

// Re-export shapes for presentations
pub use odp::{Shape, Slide};

// Re-export document element types for unified API (for ODT tables)
pub use elements::table::{Table, TableCell as Cell, TableRow as Row};
pub use elements::text::Span as Run;
pub use elements::text::{
    Heading, Hyperlink, List, ListHeader, ListItem, Paragraph, TextHyperlinkActuate,
    TextHyperlinkShow,
}; // Span is equivalent to Run in ODF

// Re-export parser types for document element iteration
pub use elements::parser::{DocumentOrderElement, DocumentParser};

pub mod section_properties;
pub use section_properties::*;
mod header_footer_properties;
pub use header_footer_properties::*;

pub use form::{
    OdfValueRangeControl, OdfValueRangeDuration, OdfValueRangeForm, OdfValueRangeInteger,
    OdfValueRangeNonNegativeInteger, OdfValueRangeOrientation, OdfValueRangePositiveInteger,
    insert_value_range_control_xml, remove_value_range_control_xml,
    replace_value_range_control_xml, value_range_controls,
};

pub use form::{
    OdfFormDate, OdfFormDouble, OdfTypedValueBound, OdfTypedValueControl, OdfTypedValueControlKind,
    OdfTypedValueDuration, OdfTypedValueForm, OdfTypedValueNonNegativeInteger,
    insert_typed_value_control_xml, remove_typed_value_control_xml,
    replace_typed_value_control_xml, typed_value_controls,
};

pub use form::{
    OdfConnectionResourceForm, OdfFormConnectionResource, OdfOwnedFormConnectionResource,
    form_connection_resources, insert_form_connection_resource_xml,
    remove_form_connection_resource_xml, replace_form_connection_resource_xml,
};
pub use form::{
    OdfGridColumn, OdfGridColumnControl, OdfGridColumnControlKind, OdfGridControl, OdfGridForm,
    OdfGridNonNegativeInteger, grid_controls, insert_grid_control_xml, remove_grid_control_xml,
    replace_grid_control_xml,
};
