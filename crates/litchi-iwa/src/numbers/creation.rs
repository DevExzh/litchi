//! Construction of independent Numbers packages without bundled templates.

use crate::archive::{Archive, ArchiveObject, RawMessage};
use crate::identity::IWorkDocumentIdentity;
use crate::protobuf::{tn, tsa, tsce, tsd, tsk, tsp, tss, tst, tswp};
use crate::{IWorkPackage, IWorkThemeArchive, IWorkThemeExtensions, Result};
use plist::Value;
use prost::Message;

use super::editor::NumbersEditor;
use super::formula_owner::{empty_table_formula_owner, uuid_as_cfuuid};

const DOCUMENT_ARCHIVE_ENTRY: &str = "Index/Document.iwa";
const CALCULATION_ARCHIVE_ENTRY: &str = "Index/CalculationEngine.iwa";
const STYLESHEET_ARCHIVE_ENTRY: &str = "Index/DocumentStylesheet.iwa";
const VIEW_STATE_ARCHIVE_ENTRY: &str = "Index/ViewState.iwa";
const ANNOTATION_ARCHIVE_ENTRY: &str = "Index/AnnotationAuthorStorage.iwa";
const DOCUMENT_METADATA_ARCHIVE_ENTRY: &str = "Index/DocumentMetadata.iwa";
const PACKAGE_METADATA_ARCHIVE_ENTRY: &str = "Index/Metadata.iwa";
const DEFAULT_LANGUAGE: &str = "en";
const DEFAULT_LOCALE: &str = "en_US";
const DEFAULT_SHEET_NAME: &str = "Sheet 1";
const DEFAULT_TABLE_NAME: &str = "Table 1";
const DEFAULT_ROWS: usize = 10;
const DEFAULT_COLUMNS: usize = 5;
const MAX_TABLE_UIDS: usize = 1_100_000;

const DOCUMENT: u64 = 1;
const METADATA: u64 = 2;
// Numbers reserves identifier 3 for its lazily-created TSCKDocumentSupport root.
// Keeping generated objects above that slot allows an opened document to save.
const STYLESHEET: u64 = 40;
const THEME: u64 = 4;
const SIDEBAR_ROOT: u64 = 5;
const SIDEBAR_SHEET: u64 = 6;
const SIDEBAR_TABLE: u64 = 7;
const SHEET: u64 = 8;
const TABLE_INFO: u64 = 9;
const TABLE_MODEL: u64 = 10;
const LIST_STYLE: u64 = 11;
const PARAGRAPH_STYLE: u64 = 12;
const CHARACTER_STYLE: u64 = 13;
const SHAPE_STYLE: u64 = 14;
const MEDIA_STYLE: u64 = 15;
const DROP_CAP_STYLE: u64 = 16;
const SHEET_STYLE: u64 = 17;
const TABLE_STYLE: u64 = 18;
const CELL_STYLE: u64 = 19;
const TABLE_PRESET: u64 = 20;
const TILE: u64 = 22;
const ROW_HEADERS: u64 = 23;
const COLUMN_HEADERS: u64 = 24;
const STRING_LIST: u64 = 25;
const STYLE_LIST: u64 = 26;
const FORMULA_LIST: u64 = 27;
const FORMAT_LIST: u64 = 28;
const UID_MAP: u64 = 29;
const STROKE_SIDECAR: u64 = 30;
const CALCULATION_ENGINE: u64 = 31;
const TABLE_STYLE_NETWORK: u64 = 32;
const FUNCTION_BROWSER_STATE: u64 = 33;
const CUSTOM_FORMAT_LIST: u64 = 34;
const ANNOTATION_AUTHOR_STORAGE: u64 = 35;
const VIEW_STATE: u64 = 36;
const UI_STATE: u64 = 37;
const DOCUMENT_METADATA: u64 = 38;
const FORMULA_OWNER: u64 = 39;

const TABLE_FORMULA_OWNER_INTERNAL_ID: u32 = 6;

#[derive(Debug, Clone, Copy)]
#[repr(u32)]
enum NumbersMessageType {
    Document = 1,
    Sheet = 2,
    ViewState = 210,
    AnnotationAuthorStorage = 213,
    CustomFormatList = 222,
    TreeNode = 205,
    Stylesheet = 401,
    FunctionBrowserState = 601,
    CharacterStyle = 2_021,
    ParagraphStyle = 2_022,
    ListStyle = 2_023,
    ShapeStyle = 2_025,
    MediaStyle = 3_016,
    CalculationEngine = 4_000,
    FormulaOwnerDependencies = 4_008,
    TableInfo = 6_000,
    TableModel = 6_001,
    Tile = 6_002,
    TableStyle = 6_003,
    CellStyle = 6_004,
    DataList = 6_005,
    HeaderStorageBucket = 6_006,
    TablePreset = 6_008,
    TableAuxiliary = 6_200,
    TableStyleNetwork = 6_247,
    StrokeSidecar = 6_305,
    DropCapStyle = 10_024,
    PackageMetadata = 11_006,
    DocumentMetadata = 11_011,
    Theme = 12_009,
    UiState = 12_026,
    SheetStyle = 12_050,
}

impl NumbersMessageType {
    const fn value(self) -> u32 {
        self as u32
    }
}

const STYLESHEET_OBJECTS: &[u64] = &[
    STYLESHEET,
    LIST_STYLE,
    PARAGRAPH_STYLE,
    CHARACTER_STYLE,
    SHAPE_STYLE,
    MEDIA_STYLE,
    DROP_CAP_STYLE,
    SHEET_STYLE,
    TABLE_STYLE,
    CELL_STYLE,
];
const DOCUMENT_OBJECTS: &[u64] = &[
    DOCUMENT,
    THEME,
    SIDEBAR_ROOT,
    SIDEBAR_SHEET,
    SIDEBAR_TABLE,
    SHEET,
    TABLE_INFO,
    TABLE_MODEL,
    TABLE_PRESET,
    TILE,
    ROW_HEADERS,
    COLUMN_HEADERS,
    STRING_LIST,
    STYLE_LIST,
    FORMULA_LIST,
    FORMAT_LIST,
    UID_MAP,
    STROKE_SIDECAR,
    CALCULATION_ENGINE,
    TABLE_STYLE_NETWORK,
    FUNCTION_BROWSER_STATE,
    CUSTOM_FORMAT_LIST,
];

/// Builder for a new, independent Numbers spreadsheet.
///
/// Every archive and identity is generated from typed values. No Apple template
/// or prebuilt blank package is copied into the result.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct NumbersDocumentBuilder {
    sheet_name: String,
    table_name: String,
    rows: usize,
    columns: usize,
    language: String,
    locale: String,
}

impl Default for NumbersDocumentBuilder {
    fn default() -> Self {
        Self {
            sheet_name: DEFAULT_SHEET_NAME.to_owned(),
            table_name: DEFAULT_TABLE_NAME.to_owned(),
            rows: DEFAULT_ROWS,
            columns: DEFAULT_COLUMNS,
            language: DEFAULT_LANGUAGE.to_owned(),
            locale: DEFAULT_LOCALE.to_owned(),
        }
    }
}

impl NumbersDocumentBuilder {
    /// Start a spreadsheet containing one sheet and one empty table.
    pub fn new() -> Self {
        Self::default()
    }

    pub fn sheet_name(mut self, name: impl Into<String>) -> Self {
        self.sheet_name = name.into();
        self
    }

    pub fn table_name(mut self, name: impl Into<String>) -> Self {
        self.table_name = name.into();
        self
    }

    pub fn table_dimensions(mut self, rows: usize, columns: usize) -> Self {
        self.rows = rows;
        self.columns = columns;
        self
    }

    pub fn language(mut self, language: impl Into<String>) -> Self {
        self.language = language.into();
        self
    }

    pub fn locale(mut self, locale: impl Into<String>) -> Self {
        self.locale = locale.into();
        self
    }

    /// Build a mutable editor for the generated spreadsheet.
    pub fn build(self) -> Result<NumbersEditor> {
        NumbersEditor::from_package(self.build_package()?)
    }

    /// Build the underlying package for lower-level IWA manipulation.
    pub fn build_package(self) -> Result<IWorkPackage> {
        self.validate()?;
        let identity = IWorkDocumentIdentity::generate();
        let table_uuid = fresh_tsp_uuid();
        let mut package = IWorkPackage::new();
        package.replace_archive(
            DOCUMENT_ARCHIVE_ENTRY,
            &document_archive(&self, &table_uuid)?,
        )?;
        package.replace_archive(
            CALCULATION_ARCHIVE_ENTRY,
            &calculation_archive(&self.locale, &table_uuid)?,
        )?;
        package.replace_archive(STYLESHEET_ARCHIVE_ENTRY, &stylesheet_archive()?)?;
        package.replace_archive(VIEW_STATE_ARCHIVE_ENTRY, &view_state_archive()?)?;
        package.replace_archive(
            ANNOTATION_ARCHIVE_ENTRY,
            &annotation_author_storage_archive()?,
        )?;
        package.replace_archive(
            DOCUMENT_METADATA_ARCHIVE_ENTRY,
            &document_metadata_archive()?,
        )?;
        package.replace_archive(
            PACKAGE_METADATA_ARCHIVE_ENTRY,
            &metadata_archive(&identity)?,
        )?;
        add_plists(&mut package, &identity)?;
        Ok(package)
    }

    fn validate(&self) -> Result<()> {
        crate::text::TextLanguageTag::new(self.language.as_str())?;
        for (value, kind) in [(&self.sheet_name, "sheet"), (&self.table_name, "table")] {
            if value.trim().is_empty() {
                return Err(crate::Error::InvalidFormat(format!(
                    "Numbers {kind} name cannot be empty"
                )));
            }
        }
        if self.locale.trim().is_empty() {
            return Err(crate::Error::InvalidFormat(
                "Numbers document locale cannot be empty".to_owned(),
            ));
        }
        if self.rows == 0 || self.columns == 0 {
            return Err(crate::Error::InvalidFormat(
                "Numbers table dimensions must be non-zero".to_owned(),
            ));
        }
        let total_uids = self.rows.checked_add(self.columns).ok_or_else(|| {
            crate::Error::InvalidFormat("Numbers table dimensions overflow usize".to_owned())
        })?;
        if total_uids > MAX_TABLE_UIDS {
            return Err(crate::Error::InvalidFormat(format!(
                "Numbers table dimensions require {total_uids} UIDs, exceeding the safety limit {MAX_TABLE_UIDS}"
            )));
        }
        Ok(())
    }
}

impl NumbersEditor {
    pub fn builder() -> NumbersDocumentBuilder {
        NumbersDocumentBuilder::new()
    }

    pub fn create() -> Result<Self> {
        NumbersDocumentBuilder::new().build()
    }
}

fn document_archive(builder: &NumbersDocumentBuilder, table_uuid: &tsp::Uuid) -> Result<Archive> {
    let document = tn::DocumentArchive {
        sheets: vec![reference(SHEET)],
        super_: tsa::DocumentArchive {
            super_: tsk::DocumentArchive {
                locale_identifier: Some(builder.locale.clone()),
                annotation_author_storage: Some(reference(ANNOTATION_AUTHOR_STORAGE)),
                creation_locale_identifier: Some(builder.locale.clone()),
                prevent_image_conversion_on_open: Some(true),
                has_user_defined_locale: Some(false),
                should_measure_negatively_tracked_text_correctly: Some(true),
                use_optimized_text_vertical_alignment: Some(true),
                should_allow_ligatures_in_minimally_tracked_text: Some(true),
                formatting_symbols: Some(formatting_symbols()),
                ..Default::default()
            },
            document_language: Some(builder.language.clone()),
            calculation_engine: Some(reference(CALCULATION_ENGINE)),
            view_state: Some(reference(VIEW_STATE)),
            function_browser_state: Some(reference(FUNCTION_BROWSER_STATE)),
            needs_media_compatibility_upgrade: Some(false),
            template_identifier: Some("Application/Blank/Traditional".to_owned()),
            custom_format_list: Some(reference(CUSTOM_FORMAT_LIST)),
            collaborative_media_compatibility_upgrade_did_fail: Some(false),
            can_use_hevc: Some(false),
            is_content_source: Some(false),
            ..Default::default()
        },
        stylesheet: reference(STYLESHEET),
        sidebar_order: reference(SIDEBAR_ROOT),
        theme: reference(THEME),
        paper_id: Some("na-letter".to_owned()),
        page_size: Some(tsp::Size {
            width: 612.0,
            height: 792.0,
        }),
        ..Default::default()
    };

    let mut common_theme = IWorkThemeArchive::new(
        tss::ThemeArchive {
            theme_identifier: Some("Litchi Blank".to_owned()),
            document_stylesheet: Some(reference(STYLESHEET)),
            color_presets: repeated(30, black),
            ..Default::default()
        },
        IWorkThemeExtensions {
            drawing: Some(tsd::ThemePresetsArchive {
                gradient_fill_presets: repeated(6, solid_fill),
                image_fill_presets: repeated(6, solid_fill),
                shadow_presets: repeated(8, tsd::ShadowArchive::default),
                line_style_presets: repeated_reference(6, SHAPE_STYLE),
                shape_style_presets: repeated_reference(6, SHAPE_STYLE),
                textbox_style_presets: repeated_reference(1, SHAPE_STYLE),
                image_style_presets: repeated_reference(6, MEDIA_STYLE),
                movie_style_presets: repeated_reference(6, MEDIA_STYLE),
                drawing_line_style_presets: repeated_reference(1, SHAPE_STYLE),
            }),
            text: Some(tswp::ThemePresetsArchive {
                list_style_presets: repeated_reference(5, LIST_STYLE),
                character_style_presets: repeated_reference(7, CHARACTER_STYLE),
                paragraph_style_presets: repeated_reference(6, PARAGRAPH_STYLE),
                dropcap_style_presets: repeated_reference(6, DROP_CAP_STYLE),
                ..Default::default()
            }),
            chart: None,
            table: Some(tst::ThemePresetsArchive {
                table_style_presets: repeated_reference(6, TABLE_PRESET),
                ..Default::default()
            }),
            application: Some(tsa::ThemePresetsArchive {
                caption_style_presets: repeated_reference(2, PARAGRAPH_STYLE),
                svg_import_style_presets: repeated_reference(1, SHAPE_STYLE),
            }),
        },
    )
    .encode()?;
    // TN.ThemeArchive field 2 (prototypes) is empty, so the common wrapper is complete.
    let theme = std::mem::take(&mut common_theme);

    let rows = builder.rows as u32;
    let columns = builder.columns as u32;
    let table_info = tst::TableInfoArchive {
        super_: tsd::DrawableArchive {
            geometry: Some(tsd::GeometryArchive {
                position: Some(tsp::Point { x: 0.0, y: 0.0 }),
                size: Some(tsp::Size {
                    width: 490.0,
                    height: 200.0,
                }),
                flags: Some(3),
                angle: Some(0.0),
            }),
            parent: Some(reference(SHEET)),
            locked: Some(false),
            aspect_ratio_locked: Some(false),
            title_hidden: Some(true),
            caption_hidden: Some(true),
            ..Default::default()
        },
        table_model: reference(TABLE_MODEL),
        formula_coord_space: Some(0),
        ..Default::default()
    };
    let model = tst::TableModelArchive {
        table_id: format_tsp_uuid(table_uuid),
        table_style: reference(TABLE_STYLE),
        body_text_style: reference(PARAGRAPH_STYLE),
        header_row_text_style: reference(PARAGRAPH_STYLE),
        header_column_text_style: reference(PARAGRAPH_STYLE),
        footer_row_text_style: reference(PARAGRAPH_STYLE),
        body_cell_style: reference(CELL_STYLE),
        header_row_style: reference(CELL_STYLE),
        header_column_style: reference(CELL_STYLE),
        footer_row_style: reference(CELL_STYLE),
        table_style_preset: Some(reference(TABLE_PRESET)),
        preset_index: Some(0),
        base_data_store: tst::DataStore {
            row_headers: tst::HeaderStorage {
                bucket_hash_function: 1,
                buckets: vec![reference(ROW_HEADERS)],
            },
            column_headers: reference(COLUMN_HEADERS),
            tiles: tst::TileStorage {
                tiles: vec![tst::tile_storage::Tile {
                    tileid: 0,
                    tile: reference(TILE),
                }],
                tile_size: Some(256),
                ..Default::default()
            },
            string_table: reference(STRING_LIST),
            style_table: reference(STYLE_LIST),
            formula_table: reference(FORMULA_LIST),
            format_table_pre_bnc: reference(FORMAT_LIST),
            next_row_strip_id: 1,
            next_column_strip_id: 0,
            row_tile_tree: tst::TableRbTree {
                nodes: vec![tst::table_rb_tree::Node { key: 0, value: 0 }],
            },
            column_tile_tree: tst::TableRbTree::default(),
            storage_version_pre_bnc: Some(4),
            ..Default::default()
        },
        number_of_rows: rows,
        number_of_columns: columns,
        table_name: builder.table_name.clone(),
        table_name_enabled: Some(false),
        number_of_header_rows: Some(1),
        number_of_header_columns: Some(1),
        header_rows_frozen: Some(true),
        header_columns_frozen: Some(true),
        default_row_height: 20.0,
        default_column_width: 98.0,
        repeating_header_rows_enabled: Some(true),
        repeating_header_columns_enabled: Some(true),
        style_apply_clears_all: Some(false),
        base_column_row_uids: Some(reference(UID_MAP)),
        stroke_sidecar: Some(reference(STROKE_SIDECAR)),
        ..Default::default()
    };

    let objects = vec![
        object(
            DOCUMENT,
            NumbersMessageType::Document,
            document,
            &[
                STYLESHEET,
                SIDEBAR_ROOT,
                THEME,
                SHEET,
                CALCULATION_ENGINE,
                VIEW_STATE,
                FUNCTION_BROWSER_STATE,
                CUSTOM_FORMAT_LIST,
                ANNOTATION_AUTHOR_STORAGE,
            ],
        )?,
        raw_object(
            THEME,
            NumbersMessageType::Theme,
            theme,
            &[
                STYLESHEET,
                LIST_STYLE,
                PARAGRAPH_STYLE,
                CHARACTER_STYLE,
                SHAPE_STYLE,
                MEDIA_STYLE,
                DROP_CAP_STYLE,
                TABLE_PRESET,
            ],
        )?,
        object(
            SIDEBAR_ROOT,
            NumbersMessageType::TreeNode,
            tsk::TreeNode {
                name: None,
                children: vec![reference(SIDEBAR_SHEET)],
                object: None,
            },
            &[SIDEBAR_SHEET],
        )?,
        object(
            SIDEBAR_SHEET,
            NumbersMessageType::TreeNode,
            tsk::TreeNode {
                name: None,
                children: vec![reference(SIDEBAR_TABLE)],
                object: Some(reference(SHEET)),
            },
            &[SIDEBAR_TABLE, SHEET],
        )?,
        object(
            SIDEBAR_TABLE,
            NumbersMessageType::TreeNode,
            tsk::TreeNode {
                name: None,
                children: Vec::new(),
                object: Some(reference(TABLE_INFO)),
            },
            &[TABLE_INFO],
        )?,
        object(
            SHEET,
            NumbersMessageType::Sheet,
            tn::SheetArchive {
                name: builder.sheet_name.clone(),
                drawable_infos: vec![reference(TABLE_INFO)],
                in_portrait_page_orientation: Some(true),
                show_page_numbers: Some(true),
                is_autofit_on: Some(false),
                content_scale: Some(0.72),
                page_order: Some(tn::PageOrder::TopToBottom as i32),
                print_margins: Some(tsd::EdgeInsetsArchive {
                    top: 54.0,
                    left: 36.0,
                    bottom: 54.0,
                    right: 36.0,
                }),
                using_start_page_number: Some(false),
                start_page_number: Some(1),
                page_header_inset: Some(20.0),
                page_footer_inset: Some(20.0),
                uses_single_header_footer: Some(false),
                layout_direction: Some(tn::PageLayoutDirection::LeftToRight as i32),
                style: Some(reference(SHEET_STYLE)),
                print_backgrounds: Some(true),
                should_print_comments: Some(false),
                ..Default::default()
            },
            &[TABLE_INFO, SHEET_STYLE],
        )?,
        object(
            TABLE_INFO,
            NumbersMessageType::TableInfo,
            table_info,
            &[SHEET, TABLE_MODEL],
        )?,
        object(
            TABLE_MODEL,
            NumbersMessageType::TableModel,
            model,
            &[
                TABLE_STYLE,
                PARAGRAPH_STYLE,
                CELL_STYLE,
                TABLE_PRESET,
                TILE,
                ROW_HEADERS,
                COLUMN_HEADERS,
                STRING_LIST,
                STYLE_LIST,
                FORMULA_LIST,
                FORMAT_LIST,
                UID_MAP,
                STROKE_SIDECAR,
            ],
        )?,
        object(
            TABLE_PRESET,
            NumbersMessageType::TablePreset,
            tst::TableStylePresetArchive {
                index: Some(0),
                style_network: Some(reference(TABLE_STYLE_NETWORK)),
                ..Default::default()
            },
            &[TABLE_STYLE_NETWORK],
        )?,
        object(
            TILE,
            NumbersMessageType::Tile,
            tst::Tile {
                max_column: columns - 1,
                max_row: rows - 1,
                num_cells: 0,
                numrows: 0,
                row_infos: Vec::new(),
                storage_version: Some(5),
                last_saved_in_bnc: Some(true),
                should_use_wide_rows: None,
            },
            &[],
        )?,
        object(
            ROW_HEADERS,
            NumbersMessageType::HeaderStorageBucket,
            header_bucket(),
            &[],
        )?,
        object(
            COLUMN_HEADERS,
            NumbersMessageType::HeaderStorageBucket,
            header_bucket(),
            &[],
        )?,
        object(
            STRING_LIST,
            NumbersMessageType::DataList,
            data_list(tst::table_data_list::ListType::String),
            &[],
        )?,
        object(
            STYLE_LIST,
            NumbersMessageType::DataList,
            data_list(tst::table_data_list::ListType::Style),
            &[],
        )?,
        object(
            FORMULA_LIST,
            NumbersMessageType::DataList,
            data_list(tst::table_data_list::ListType::Formula),
            &[],
        )?,
        object(
            FORMAT_LIST,
            NumbersMessageType::DataList,
            data_list(tst::table_data_list::ListType::Format),
            &[],
        )?,
        object(
            UID_MAP,
            NumbersMessageType::TableAuxiliary,
            uid_map(rows, columns),
            &[],
        )?,
        object(
            STROKE_SIDECAR,
            NumbersMessageType::StrokeSidecar,
            tst::StrokeSidecarArchive {
                row_count: Some(rows),
                column_count: Some(columns),
                ..Default::default()
            },
            &[],
        )?,
        object(
            TABLE_STYLE_NETWORK,
            NumbersMessageType::TableStyleNetwork,
            tst::TableStyleNetworkArchive {
                body_text_style: reference(PARAGRAPH_STYLE),
                header_row_text_style: reference(PARAGRAPH_STYLE),
                header_column_text_style: reference(PARAGRAPH_STYLE),
                footer_row_text_style: reference(PARAGRAPH_STYLE),
                body_cell_style: reference(CELL_STYLE),
                header_row_style: reference(CELL_STYLE),
                header_column_style: reference(CELL_STYLE),
                footer_row_style: reference(CELL_STYLE),
                table_style: reference(TABLE_STYLE),
                preset_id: Some(0),
                ..Default::default()
            },
            &[PARAGRAPH_STYLE, CELL_STYLE, TABLE_STYLE],
        )?,
        object(
            FUNCTION_BROWSER_STATE,
            NumbersMessageType::FunctionBrowserState,
            tsa::FunctionBrowserStateArchive {
                current_function: Some(0),
                ..Default::default()
            },
            &[],
        )?,
        object(
            CUSTOM_FORMAT_LIST,
            NumbersMessageType::CustomFormatList,
            tsk::CustomFormatListArchive::default(),
            &[],
        )?,
    ];
    Ok(Archive { objects })
}

#[allow(deprecated)]
fn view_state_archive() -> Result<Archive> {
    Ok(Archive {
        objects: vec![
            object(
                VIEW_STATE,
                NumbersMessageType::ViewState,
                tsk::ViewStateArchive {
                    view_state_root: reference(UI_STATE),
                    ..Default::default()
                },
                &[UI_STATE],
            )?,
            object(
                UI_STATE,
                NumbersMessageType::UiState,
                tn::UiStateArchive {
                    active_sheet_index: 0,
                    editing_sheet_index: Some(0),
                    document_mode: Some(0),
                    in_chart_mode: Some(false),
                    inspector_pane_visible: Some(true),
                    show_canvas_guides: Some(true),
                    shows_comments: Some(true),
                    shows_rulers: Some(true),
                    editing_disabled: Some(false),
                    sidebar_visible: Some(true),
                    sidebar_width: Some(128.0),
                    ..Default::default()
                },
                &[],
            )?,
        ],
    })
}

fn annotation_author_storage_archive() -> Result<Archive> {
    Ok(Archive {
        objects: vec![object(
            ANNOTATION_AUTHOR_STORAGE,
            NumbersMessageType::AnnotationAuthorStorage,
            tsk::AnnotationAuthorStorageArchive::default(),
            &[],
        )?],
    })
}

fn document_metadata_archive() -> Result<Archive> {
    Ok(Archive {
        objects: vec![object(
            DOCUMENT_METADATA,
            NumbersMessageType::DocumentMetadata,
            tsp::DocumentMetadata {
                is_in_collaboration_mode: Some(false),
                ..Default::default()
            },
            &[],
        )?],
    })
}

fn calculation_archive(locale: &str, table_uuid: &tsp::Uuid) -> Result<Archive> {
    let formula_owner =
        empty_table_formula_owner(table_uuid, TABLE_INFO, TABLE_FORMULA_OWNER_INTERNAL_ID);
    Ok(Archive {
        objects: vec![
            object(
                CALCULATION_ENGINE,
                NumbersMessageType::CalculationEngine,
                tsce::CalculationEngineArchive {
                    dependency_tracker: tsce::DependencyTrackerArchive {
                        owner_id_map: Some(tsce::OwnerIdMapArchive {
                            map_entry: vec![tsce::owner_id_map_archive::OwnerIdMapArchiveEntry {
                                internal_owner_id: TABLE_FORMULA_OWNER_INTERNAL_ID,
                                owner_id: uuid_as_cfuuid(&formula_owner.formula_owner_uid),
                            }],
                            ..Default::default()
                        }),
                        number_of_formulas: Some(0),
                        formula_owner_dependencies: vec![reference(FORMULA_OWNER)],
                        ..Default::default()
                    },
                    saved_locale_identifier: Some(locale.to_owned()),
                    ..Default::default()
                },
                &[FORMULA_OWNER],
            )?,
            object(
                FORMULA_OWNER,
                NumbersMessageType::FormulaOwnerDependencies,
                formula_owner,
                &[],
            )?,
        ],
    })
}

fn stylesheet_archive() -> Result<Archive> {
    let styles = &STYLESHEET_OBJECTS[1..];
    Ok(Archive {
        objects: vec![
            object(
                STYLESHEET,
                NumbersMessageType::Stylesheet,
                tss::StylesheetArchive {
                    styles: styles.iter().copied().map(reference).collect(),
                    is_locked: Some(false),
                    can_cull_styles: Some(true),
                    ..Default::default()
                },
                styles,
            )?,
            object(
                LIST_STYLE,
                NumbersMessageType::ListStyle,
                tswp::ListStyleArchive {
                    super_: style("None", "litchi-list-none"),
                    override_count: Some(0),
                    ..Default::default()
                },
                &[STYLESHEET],
            )?,
            object(
                PARAGRAPH_STYLE,
                NumbersMessageType::ParagraphStyle,
                tswp::ParagraphStyleArchive {
                    super_: style("Body", "litchi-table-body"),
                    override_count: Some(0),
                    ..Default::default()
                },
                &[STYLESHEET],
            )?,
            object(
                CHARACTER_STYLE,
                NumbersMessageType::CharacterStyle,
                tswp::CharacterStyleArchive {
                    super_: style("Default", "litchi-character-default"),
                    override_count: Some(0),
                    ..Default::default()
                },
                &[STYLESHEET],
            )?,
            object(
                SHAPE_STYLE,
                NumbersMessageType::ShapeStyle,
                tswp::ShapeStyleArchive {
                    super_: tsd::ShapeStyleArchive {
                        super_: style("Shape", "litchi-shape-default"),
                        override_count: Some(0),
                        ..Default::default()
                    },
                    override_count: Some(0),
                    ..Default::default()
                },
                &[STYLESHEET],
            )?,
            object(
                MEDIA_STYLE,
                NumbersMessageType::MediaStyle,
                tsd::MediaStyleArchive {
                    super_: style("Media", "litchi-media-default"),
                    override_count: Some(0),
                    ..Default::default()
                },
                &[STYLESHEET],
            )?,
            object(
                DROP_CAP_STYLE,
                NumbersMessageType::DropCapStyle,
                tswp::DropCapStyleArchive {
                    super_: style("Drop Cap", "litchi-dropcap"),
                    override_count: Some(0),
                    ..Default::default()
                },
                &[STYLESHEET],
            )?,
            object(
                SHEET_STYLE,
                NumbersMessageType::SheetStyle,
                tn::SheetStyleArchive {
                    super_: style("Sheet", "litchi-sheet-default"),
                    override_count: Some(0),
                    ..Default::default()
                },
                &[STYLESHEET],
            )?,
            object(
                TABLE_STYLE,
                NumbersMessageType::TableStyle,
                tst::TableStyleArchive {
                    super_: style("Table", "litchi-table-default"),
                    override_count: Some(0),
                    table_properties: Some(tst::TableStylePropertiesArchive {
                        behaves_like_spreadsheet: Some(true),
                        auto_resize: Some(false),
                        ..Default::default()
                    }),
                },
                &[STYLESHEET],
            )?,
            object(
                CELL_STYLE,
                NumbersMessageType::CellStyle,
                tst::CellStyleArchive {
                    super_: style("Cell", "litchi-cell-default"),
                    override_count: Some(0),
                    cell_properties: Some(tst::CellStylePropertiesArchive {
                        text_wrap: Some(false),
                        ..Default::default()
                    }),
                },
                &[STYLESHEET],
            )?,
        ],
    })
}

fn metadata_archive(identity: &IWorkDocumentIdentity) -> Result<Archive> {
    let component = |identifier, locator: &str, version: &[u32]| tsp::ComponentInfo {
        identifier,
        preferred_locator: locator.to_owned(),
        document_read_version: version.to_vec(),
        document_write_version: version.to_vec(),
        save_token: Some(1),
        ..Default::default()
    };
    let mut stylesheet = component(STYLESHEET, "DocumentStylesheet", &[2, 0, 0]);
    stylesheet.object_uuid_map_entries = STYLESHEET_OBJECTS
        .iter()
        .copied()
        .map(object_uuid)
        .collect();
    let mut document = component(DOCUMENT, "Document", &[2, 0, 0]);
    document.object_uuid_map_entries = DOCUMENT_OBJECTS
        .iter()
        .copied()
        .filter(|identifier| *identifier != CALCULATION_ENGINE)
        .map(object_uuid)
        .collect();
    document.external_references = std::iter::once(None)
        .chain(STYLESHEET_OBJECTS.iter().copied().map(Some))
        .map(|object_identifier| tsp::ComponentExternalReference {
            component_identifier: STYLESHEET,
            object_identifier,
            is_weak: None,
        })
        .collect();
    document
        .external_references
        .push(tsp::ComponentExternalReference {
            component_identifier: CALCULATION_ENGINE,
            object_identifier: None,
            is_weak: None,
        });
    let mut calculation = component(CALCULATION_ENGINE, "CalculationEngine", &[3, 2, 10]);
    calculation.object_uuid_map_entries = [CALCULATION_ENGINE, FORMULA_OWNER]
        .into_iter()
        .map(object_uuid)
        .collect();
    calculation.external_references = vec![tsp::ComponentExternalReference {
        component_identifier: DOCUMENT,
        object_identifier: Some(TABLE_INFO),
        is_weak: None,
    }];
    let view_state = component(VIEW_STATE, "ViewState", &[2, 0, 0]);
    let annotation = component(
        ANNOTATION_AUTHOR_STORAGE,
        "AnnotationAuthorStorage",
        &[2, 0, 0],
    );
    let document_metadata = component(DOCUMENT_METADATA, "DocumentMetadata", &[2, 0, 0]);
    document.external_references.extend([
        tsp::ComponentExternalReference {
            component_identifier: VIEW_STATE,
            object_identifier: None,
            is_weak: None,
        },
        tsp::ComponentExternalReference {
            component_identifier: ANNOTATION_AUTHOR_STORAGE,
            object_identifier: None,
            is_weak: None,
        },
    ]);
    let metadata = tsp::PackageMetadata {
        last_object_identifier: STYLESHEET,
        revision: Some(tsp::DocumentRevision {
            sequence_32: Some(0),
            identifier: Some(identity.version_uuid().to_owned()),
            sequence_64: None,
        }),
        components: vec![
            document_metadata,
            stylesheet,
            calculation,
            view_state,
            annotation,
            document,
        ],
        read_version: vec![3, 2, 10],
        write_version: vec![3, 2, 10],
        file_format_version: vec![14, 4, 1],
        save_token: Some(1),
        ..Default::default()
    };
    Ok(Archive {
        objects: vec![object(
            METADATA,
            NumbersMessageType::PackageMetadata,
            metadata,
            &[],
        )?],
    })
}

fn add_plists(package: &mut IWorkPackage, identity: &IWorkDocumentIdentity) -> Result<()> {
    let mut properties = plist::Dictionary::new();
    for key in ["documentUUID", "stableDocumentUUID", "shareUUID"] {
        properties.insert(
            key.to_owned(),
            Value::String(identity.document_uuid().to_owned()),
        );
    }
    properties.insert(
        "fileFormatVersion".to_owned(),
        Value::String("14.4.1".to_owned()),
    );
    properties.insert(
        "hasExternalReferenceOrMissingOrUnmaterializedRemoteData".to_owned(),
        Value::Boolean(false),
    );
    properties.insert("isMultiPage".to_owned(), Value::Boolean(false));
    properties.insert(
        "privateUUID".to_owned(),
        Value::String(identity.private_uuid().to_owned()),
    );
    properties.insert(
        "versionUUID".to_owned(),
        Value::String(identity.version_uuid().to_owned()),
    );
    properties.insert(
        "revision".to_owned(),
        Value::String(format!("0::{}", identity.version_uuid())),
    );
    let mut encoded = Vec::new();
    Value::Dictionary(properties)
        .to_writer_binary(&mut encoded)
        .map_err(|error| {
            crate::Error::InvalidFormat(format!(
                "failed to encode generated Numbers properties: {error}"
            ))
        })?;
    package.insert_entry("Metadata/Properties.plist", encoded)?;
    package.insert_entry(
        "Metadata/DocumentIdentifier",
        identity.document_uuid().as_bytes().to_vec(),
    )?;
    let mut history = Vec::new();
    Value::Array(vec![Value::String("Created by litchi-iwa".to_owned())])
        .to_writer_binary(&mut history)
        .map_err(|error| {
            crate::Error::InvalidFormat(format!(
                "failed to encode generated Numbers build history: {error}"
            ))
        })?;
    package.insert_entry("Metadata/BuildVersionHistory.plist", history)?;
    Ok(())
}

fn header_bucket() -> tst::HeaderStorageBucket {
    tst::HeaderStorageBucket {
        bucket_hash_function: 1,
        headers: Vec::new(),
    }
}

fn data_list(list_type: tst::table_data_list::ListType) -> tst::TableDataList {
    tst::TableDataList {
        list_type: list_type as i32,
        next_list_id: 1,
        entries: Vec::new(),
        segments: Vec::new(),
        is_new_for_bnc: Some(true),
    }
}

fn formatting_symbols() -> tsk::FormattingSymbolsArchive {
    tsk::FormattingSymbolsArchive {
        version: Some("4302.00*".to_owned()),
        calendar: Some("gregorian".to_owned()),
        numbering_system: Some("latn".to_owned()),
        am_symbol: Some("AM".to_owned()),
        pm_symbol: Some("PM".to_owned()),
        short_date_pattern: Some("M/d/yy".to_owned()),
        medium_date_pattern: Some("MMM d, y".to_owned()),
        long_date_pattern: Some("MMMM d, y".to_owned()),
        full_date_pattern: Some("EEEE, MMMM d, y".to_owned()),
        short_time_pattern: Some("HH:mm".to_owned()),
        medium_time_pattern: Some("HH:mm:ss".to_owned()),
        long_time_pattern: Some("HH:mm:ss z".to_owned()),
        full_time_pattern: Some("HH:mm:ss zzzz".to_owned()),
        decimal_separator: Some(".".to_owned()),
        grouping_separator: Some(",".to_owned()),
        currency_decimal_separator: Some(".".to_owned()),
        currency_grouping_separator: Some(",".to_owned()),
        plus_sign: Some("+".to_owned()),
        minus_sign: Some("-".to_owned()),
        exponential_symbol: Some("E".to_owned()),
        percent_symbol: Some("%".to_owned()),
        per_mille_symbol: Some("‰".to_owned()),
        infinity_symbol: Some("+∞".to_owned()),
        nan_symbol: Some("NaN".to_owned()),
        decimal_pattern: Some("#,##0.###".to_owned()),
        scientific_pattern: Some("#E0".to_owned()),
        percent_pattern: Some("#,##0%".to_owned()),
        currency_pattern: Some("¤#,##0.00".to_owned()),
        currency_code: Some("USD".to_owned()),
        ..Default::default()
    }
}

fn uid_map(rows: u32, columns: u32) -> tst::ColumnRowUidMapArchive {
    tst::ColumnRowUidMapArchive {
        sorted_column_uids: (0..columns).map(|_| fresh_tsp_uuid()).collect(),
        column_index_for_uid: (0..columns).collect(),
        column_uid_for_index: (0..columns).collect(),
        sorted_row_uids: (0..rows).map(|_| fresh_tsp_uuid()).collect(),
        row_index_for_uid: (0..rows).collect(),
        row_uid_for_index: (0..rows).collect(),
    }
}

fn fresh_tsp_uuid() -> tsp::Uuid {
    let bytes = litchi_core::id::generate_guid_bytes();
    tsp::Uuid {
        upper: u64::from_be_bytes([
            bytes[0], bytes[1], bytes[2], bytes[3], bytes[4], bytes[5], bytes[6], bytes[7],
        ]),
        lower: u64::from_be_bytes([
            bytes[8], bytes[9], bytes[10], bytes[11], bytes[12], bytes[13], bytes[14], bytes[15],
        ]),
    }
}

fn object_uuid(identifier: u64) -> tsp::ObjectUuidMapEntry {
    tsp::ObjectUuidMapEntry {
        identifier,
        uuid: fresh_tsp_uuid(),
    }
}

fn format_tsp_uuid(uuid: &tsp::Uuid) -> String {
    format!(
        "{:08X}-{:04X}-{:04X}-{:04X}-{:012X}",
        uuid.upper >> 32,
        (uuid.upper >> 16) & 0xffff,
        uuid.upper & 0xffff,
        uuid.lower >> 48,
        uuid.lower & 0xffff_ffff_ffff,
    )
}

fn object(
    identifier: u64,
    message_type: NumbersMessageType,
    message: impl Message,
    references: &[u64],
) -> Result<ArchiveObject> {
    raw_object(
        identifier,
        message_type,
        message.encode_to_vec(),
        references,
    )
}

fn raw_object(
    identifier: u64,
    message_type: NumbersMessageType,
    data: Vec<u8>,
    references: &[u64],
) -> Result<ArchiveObject> {
    let mut object = ArchiveObject::new(
        identifier,
        vec![RawMessage {
            type_: message_type.value(),
            data,
        }],
    )?;
    object.archive_info.message_infos[0].versions = vec![1, 0, 5];
    object.archive_info.message_infos[0].object_references = references.to_vec();
    Ok(object)
}

fn style(name: &str, identifier: &str) -> tss::StyleArchive {
    tss::StyleArchive {
        name: Some(name.to_owned()),
        style_identifier: Some(identifier.to_owned()),
        stylesheet: Some(reference(STYLESHEET)),
        ..Default::default()
    }
}

fn repeated<T>(count: usize, make: impl Fn() -> T) -> Vec<T> {
    std::iter::repeat_with(make).take(count).collect()
}

fn repeated_reference(count: usize, identifier: u64) -> Vec<tsp::Reference> {
    repeated(count, || reference(identifier))
}

fn reference(identifier: u64) -> tsp::Reference {
    tsp::Reference {
        identifier,
        ..Default::default()
    }
}

fn black() -> tsp::Color {
    tsp::Color {
        model: tsp::color::ColorModel::Rgb as i32,
        r: Some(0.0),
        g: Some(0.0),
        b: Some(0.0),
        rgbspace: Some(tsp::color::RgbColorSpace::Srgb as i32),
        a: Some(1.0),
        ..Default::default()
    }
}

fn solid_fill() -> tsd::FillArchive {
    tsd::FillArchive {
        color: Some(black()),
        ..Default::default()
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::numbers::{
        CellValue, FormulaBinaryOperator, FormulaCellReference, FormulaExpression, NumbersDocument,
    };

    #[test]
    fn creates_and_reopens_independent_spreadsheet() {
        let editor = NumbersDocumentBuilder::new()
            .sheet_name("Inventory")
            .table_name("Stock")
            .table_dimensions(12, 4)
            .build()
            .unwrap();

        assert_eq!(editor.sheets().unwrap()[0].name, "Inventory");
        let table = &editor.tables().unwrap()[0];
        assert_eq!(table.name, "Stock");
        assert_eq!((table.rows, table.columns), (12, 4));

        let bytes = editor.to_bytes().unwrap();
        let reopened = NumbersEditor::from_bytes(&bytes).unwrap();
        assert_eq!(reopened.sheets().unwrap()[0].name, "Inventory");
        assert_eq!(reopened.tables().unwrap()[0].name, "Stock");
    }

    #[test]
    fn generated_package_contains_only_synthetic_required_entries() {
        const DOCUMENT_SUPPORT_RESERVED_IDENTIFIER: u64 = 3;
        const EXPECTED_ENTRIES: [&str; 10] = [
            DOCUMENT_ARCHIVE_ENTRY,
            CALCULATION_ARCHIVE_ENTRY,
            STYLESHEET_ARCHIVE_ENTRY,
            VIEW_STATE_ARCHIVE_ENTRY,
            ANNOTATION_ARCHIVE_ENTRY,
            DOCUMENT_METADATA_ARCHIVE_ENTRY,
            PACKAGE_METADATA_ARCHIVE_ENTRY,
            "Metadata/Properties.plist",
            "Metadata/DocumentIdentifier",
            "Metadata/BuildVersionHistory.plist",
        ];

        let first = NumbersDocumentBuilder::new().build_package().unwrap();
        let second = NumbersDocumentBuilder::new().build_package().unwrap();
        assert_eq!(first.entry_names().collect::<Vec<_>>(), EXPECTED_ENTRIES);
        assert_eq!(first.len(), EXPECTED_ENTRIES.len());
        assert!(first.entry_names().all(|name| !name.starts_with("Data/")));
        assert!(first.entry_names().all(|name| !name.starts_with("preview")));
        assert!(
            first
                .archive(STYLESHEET_ARCHIVE_ENTRY)
                .unwrap()
                .object(DOCUMENT_SUPPORT_RESERVED_IDENTIFIER)
                .is_none(),
            "object identifier 3 must remain available for Numbers document support"
        );
        assert_ne!(
            first.entry("Metadata/DocumentIdentifier"),
            second.entry("Metadata/DocumentIdentifier")
        );
    }

    #[test]
    fn generated_table_supports_cell_crud() {
        let mut editor = NumbersEditor::create().unwrap();
        let table_id = editor.tables().unwrap()[0].object_id;
        editor
            .set_cell(table_id, 1, 1, CellValue::Text("Litchi".to_owned()))
            .unwrap();
        editor
            .set_cell(table_id, 2, 2, CellValue::Number(42.5))
            .unwrap();
        editor
            .set_cell(table_id, 3, 3, CellValue::Boolean(true))
            .unwrap();
        editor.clear_cell(table_id, 3, 3).unwrap();

        let bytes = editor.to_bytes().unwrap();
        let document = NumbersDocument::from_bytes(&bytes).unwrap();
        let sheets = document.sheets().unwrap();
        let table = &sheets[0].tables[0];
        assert_eq!(
            table.get_cell(1, 1),
            Some(&CellValue::Text("Litchi".to_owned()))
        );
        assert_eq!(table.get_cell(2, 2), Some(&CellValue::Number(42.5)));
        assert!(table.get_cell(3, 3).is_none_or(CellValue::is_empty));
    }

    #[test]
    fn generated_table_supports_formula_crud() {
        let mut editor = NumbersEditor::create().unwrap();
        let table_id = editor.tables().unwrap()[0].object_id;
        editor
            .set_cell(table_id, 0, 0, CellValue::Number(40.0))
            .unwrap();
        editor
            .set_cell(table_id, 0, 1, CellValue::Number(2.0))
            .unwrap();
        let baseline = editor.to_bytes().unwrap();

        editor
            .set_formula(
                table_id,
                0,
                2,
                FormulaExpression::binary(
                    FormulaBinaryOperator::Divide,
                    FormulaExpression::cell(FormulaCellReference::relative(0, 0)),
                    FormulaExpression::cell(FormulaCellReference::relative(0, 1)),
                ),
            )
            .unwrap();
        let document = NumbersDocument::from_bytes(&editor.to_bytes().unwrap()).unwrap();
        assert_eq!(
            document.sheets().unwrap()[0].tables[0].get_cell(0, 2),
            Some(&CellValue::Formula("=(A1/B1)".to_owned()))
        );

        editor.clear_cell(table_id, 0, 2).unwrap();
        assert_eq!(editor.to_bytes().unwrap(), baseline);
    }

    #[test]
    fn generated_added_table_supports_formula_crud() {
        let mut editor = NumbersEditor::create().unwrap();
        let sheet_id = editor.sheets().unwrap()[0].object_id;
        let table = editor
            .add_empty_table(sheet_id, "Calculated", 4, 3)
            .unwrap();
        editor
            .set_cell(table.object_id, 0, 0, CellValue::Number(21.0))
            .unwrap();
        editor
            .set_cell(table.object_id, 0, 1, CellValue::Number(2.0))
            .unwrap();
        let baseline = editor.to_bytes().unwrap();

        editor
            .set_formula(
                table.object_id,
                0,
                2,
                FormulaExpression::binary(
                    FormulaBinaryOperator::Multiply,
                    FormulaExpression::cell(FormulaCellReference::relative(0, 0)),
                    FormulaExpression::cell(FormulaCellReference::relative(0, 1)),
                ),
            )
            .unwrap();
        let document = NumbersDocument::from_bytes(&editor.to_bytes().unwrap()).unwrap();
        let sheets = document.sheets().unwrap();
        let added = sheets[0]
            .tables
            .iter()
            .find(|candidate| candidate.name == table.name)
            .unwrap();
        assert_eq!(
            added.get_cell(0, 2),
            Some(&CellValue::Formula("=(A1*B1)".to_owned()))
        );

        editor.clear_cell(table.object_id, 0, 2).unwrap();
        assert_eq!(editor.to_bytes().unwrap(), baseline);
    }

    #[test]
    fn generated_added_table_removes_its_formula_owner() {
        let mut editor = NumbersEditor::create().unwrap();
        let sheet_id = editor.sheets().unwrap()[0].object_id;
        let table = editor
            .add_empty_table(sheet_id, "Disposable", 3, 2)
            .unwrap();
        editor
            .set_formula(
                table.object_id,
                0,
                0,
                FormulaExpression::binary(
                    FormulaBinaryOperator::Add,
                    FormulaExpression::Number(1.0),
                    FormulaExpression::Number(2.0),
                ),
            )
            .unwrap();
        let before = editor.package().archive(CALCULATION_ARCHIVE_ENTRY).unwrap();
        assert_eq!(
            before
                .objects
                .iter()
                .flat_map(|object| &object.messages)
                .filter(
                    |message| message.type_ == NumbersMessageType::FormulaOwnerDependencies.value()
                )
                .count(),
            2
        );

        editor.remove_table(table.object_id).unwrap();
        let calculation = editor.package().archive(CALCULATION_ARCHIVE_ENTRY).unwrap();
        let engine = calculation
            .objects
            .iter()
            .flat_map(|object| &object.messages)
            .find(|message| message.type_ == NumbersMessageType::CalculationEngine.value())
            .map(|message| tsce::CalculationEngineArchive::decode(message.data.as_slice()).unwrap())
            .unwrap();
        assert_eq!(engine.dependency_tracker.number_of_formulas, Some(0));
        assert_eq!(
            engine.dependency_tracker.formula_owner_dependencies.len(),
            1
        );
        assert_eq!(
            calculation
                .objects
                .iter()
                .flat_map(|object| &object.messages)
                .filter(
                    |message| message.type_ == NumbersMessageType::FormulaOwnerDependencies.value()
                )
                .count(),
            1
        );
        let package = editor.package();
        let maximum_identifier = package
            .iwa_entry_names()
            .flat_map(|name| package.archive(name).unwrap().objects)
            .filter_map(|object| object.archive_info.identifier)
            .max()
            .unwrap();
        assert_eq!(
            crate::package_metadata::package_last_object_identifier(package).unwrap(),
            Some(maximum_identifier)
        );
    }

    #[test]
    fn generated_table_removal_rejects_incoming_formula_references() {
        let mut editor = NumbersEditor::create().unwrap();
        let sheet_id = editor.sheets().unwrap()[0].object_id;
        let source_table_id = editor.tables().unwrap()[0].object_id;
        let target = editor
            .add_empty_table(sheet_id, "Referenced", 2, 2)
            .unwrap();
        editor
            .set_cell(target.object_id, 0, 0, CellValue::Number(7.0))
            .unwrap();
        editor
            .set_formula(
                source_table_id,
                0,
                0,
                FormulaExpression::table_cell(
                    target.object_id,
                    FormulaCellReference::relative(0, 0),
                ),
            )
            .unwrap();
        let baseline = editor.to_bytes().unwrap();

        assert!(editor.remove_table(target.object_id).is_err());
        assert_eq!(editor.to_bytes().unwrap(), baseline);
    }

    #[test]
    fn generated_spreadsheet_supports_sheet_and_table_lifecycle() {
        let mut editor = NumbersEditor::create().unwrap();
        let sheet = editor.add_empty_sheet("Archive").unwrap();
        let table = editor
            .add_empty_table(sheet.object_id, "History", 3, 2)
            .unwrap();
        assert!(editor.tables().unwrap().iter().any(|item| item == &table));

        assert_eq!(editor.remove_table(table.object_id).unwrap(), table);
        assert_eq!(editor.remove_sheet(sheet.object_id).unwrap(), sheet);

        let reopened = NumbersEditor::from_bytes(&editor.to_bytes().unwrap()).unwrap();
        assert_eq!(reopened.sheets().unwrap().len(), 1);
        assert_eq!(reopened.tables().unwrap().len(), 1);
    }

    #[test]
    fn rejects_invalid_creation_parameters() {
        assert!(
            NumbersDocumentBuilder::new()
                .sheet_name(" ")
                .build()
                .is_err()
        );
        assert!(
            NumbersDocumentBuilder::new()
                .table_dimensions(0, 1)
                .build()
                .is_err()
        );
        assert!(NumbersDocumentBuilder::new().locale("").build().is_err());
    }
}
