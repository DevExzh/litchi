//! Read-only projection of ODF text tables.

use litchi_odf_common::datatype::DurationValue;

use crate::paragraph::Paragraph;

/// Table-column or table-row visibility.
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
#[non_exhaustive]
pub enum Visibility {
    /// The declaration is displayed normally.
    Visible,
    /// The declaration is collapsed.
    Collapse,
    /// The declaration is filtered.
    Filter,
}

/// Inert linked-table refresh mode.
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub enum TableSourceMode {
    /// Copy the complete source table.
    CopyAll,
    /// Copy only source results.
    CopyResultsOnly,
}

/// Inert linked-table activation policy.
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub enum TableSourceActuate {
    /// Refresh only when explicitly requested by the host application.
    OnRequest,
}

/// Optional metadata attributes on a `table:table` element.
#[derive(Clone, Debug, PartialEq, Eq)]
pub struct TableProperties {
    template_name: Option<String>,
    use_first_row_styles: Option<bool>,
    use_last_row_styles: Option<bool>,
    use_first_column_styles: Option<bool>,
    use_last_column_styles: Option<bool>,
    use_banding_rows_styles: Option<bool>,
    use_banding_columns_styles: Option<bool>,
    protected: Option<bool>,
    protection_key: Option<String>,
    protection_key_digest_algorithm: Option<String>,
    print: Option<bool>,
    print_ranges: Option<String>,
    xml_id: Option<String>,
    is_sub_table: Option<bool>,
}

impl TableProperties {
    #[allow(
        clippy::too_many_arguments,
        reason = "projection fields mirror ODF table attributes"
    )]
    pub(crate) const fn projected(
        template_name: Option<String>,
        use_first_row_styles: Option<bool>,
        use_last_row_styles: Option<bool>,
        use_first_column_styles: Option<bool>,
        use_last_column_styles: Option<bool>,
        use_banding_rows_styles: Option<bool>,
        use_banding_columns_styles: Option<bool>,
        protected: Option<bool>,
        protection_key: Option<String>,
        protection_key_digest_algorithm: Option<String>,
        print: Option<bool>,
        print_ranges: Option<String>,
        xml_id: Option<String>,
        is_sub_table: Option<bool>,
    ) -> Self {
        Self {
            template_name,
            use_first_row_styles,
            use_last_row_styles,
            use_first_column_styles,
            use_last_column_styles,
            use_banding_rows_styles,
            use_banding_columns_styles,
            protected,
            protection_key,
            protection_key_digest_algorithm,
            print,
            print_ranges,
            xml_id,
            is_sub_table,
        }
    }

    /// Optional table template name.
    #[must_use]
    pub fn template_name(&self) -> Option<&str> {
        self.template_name.as_deref()
    }
    /// Optional first-row style flag.
    #[must_use]
    pub const fn use_first_row_styles(&self) -> Option<bool> {
        self.use_first_row_styles
    }
    /// Optional last-row style flag.
    #[must_use]
    pub const fn use_last_row_styles(&self) -> Option<bool> {
        self.use_last_row_styles
    }
    /// Optional first-column style flag.
    #[must_use]
    pub const fn use_first_column_styles(&self) -> Option<bool> {
        self.use_first_column_styles
    }
    /// Optional last-column style flag.
    #[must_use]
    pub const fn use_last_column_styles(&self) -> Option<bool> {
        self.use_last_column_styles
    }
    /// Optional banded-row style flag.
    #[must_use]
    pub const fn use_banding_rows_styles(&self) -> Option<bool> {
        self.use_banding_rows_styles
    }
    /// Optional banded-column style flag.
    #[must_use]
    pub const fn use_banding_columns_styles(&self) -> Option<bool> {
        self.use_banding_columns_styles
    }
    /// Optional protection flag, preserving source presence.
    #[must_use]
    pub const fn protected(&self) -> Option<bool> {
        self.protected
    }
    /// Optional protection key.
    #[must_use]
    pub fn protection_key(&self) -> Option<&str> {
        self.protection_key.as_deref()
    }
    /// Optional protection-key digest algorithm IRI.
    #[must_use]
    pub fn protection_key_digest_algorithm(&self) -> Option<&str> {
        self.protection_key_digest_algorithm.as_deref()
    }
    /// Optional print flag.
    #[must_use]
    pub const fn print(&self) -> Option<bool> {
        self.print
    }
    /// Optional exact print-range lexical value.
    #[must_use]
    pub fn print_ranges(&self) -> Option<&str> {
        self.print_ranges.as_deref()
    }
    /// Optional XML identity.
    #[must_use]
    pub fn xml_id(&self) -> Option<&str> {
        self.xml_id.as_deref()
    }
    /// Optional sub-table flag.
    #[must_use]
    pub const fn is_sub_table(&self) -> Option<bool> {
        self.is_sub_table
    }
}

/// Inert linked-table source metadata.
#[derive(Clone, Debug, PartialEq, Eq)]
pub struct TableSource {
    mode: Option<TableSourceMode>,
    table_name: Option<String>,
    href: String,
    actuate: Option<TableSourceActuate>,
    filter_name: Option<String>,
    filter_options: Option<String>,
    refresh_delay: Option<DurationValue>,
}

impl TableSource {
    pub(crate) const fn projected(
        mode: Option<TableSourceMode>,
        table_name: Option<String>,
        href: String,
        actuate: Option<TableSourceActuate>,
        filter_name: Option<String>,
        filter_options: Option<String>,
        refresh_delay: Option<DurationValue>,
    ) -> Self {
        Self {
            mode,
            table_name,
            href,
            actuate,
            filter_name,
            filter_options,
            refresh_delay,
        }
    }

    /// Optional source copy mode.
    #[must_use]
    pub const fn mode(&self) -> Option<TableSourceMode> {
        self.mode
    }
    /// Optional source table name.
    #[must_use]
    pub fn table_name(&self) -> Option<&str> {
        self.table_name.as_deref()
    }
    /// Required decoded lexical source IRI.
    #[must_use]
    pub fn href(&self) -> &str {
        &self.href
    }
    /// Optional source activation policy.
    #[must_use]
    pub const fn actuate(&self) -> Option<TableSourceActuate> {
        self.actuate
    }
    /// Optional source filter name.
    #[must_use]
    pub fn filter_name(&self) -> Option<&str> {
        self.filter_name.as_deref()
    }
    /// Optional source filter options.
    #[must_use]
    pub fn filter_options(&self) -> Option<&str> {
        self.filter_options.as_deref()
    }
    /// Optional exact ODF duration.
    #[must_use]
    pub fn refresh_delay(&self) -> Option<&DurationValue> {
        self.refresh_delay.as_ref()
    }
}

/// In-content RDFa metadata attached to a table cell.
#[derive(Clone, Debug, PartialEq, Eq)]
pub struct InContentMeta {
    about: String,
    property: String,
    datatype: Option<String>,
    content: Option<String>,
}

impl InContentMeta {
    pub(crate) const fn projected(
        about: String,
        property: String,
        datatype: Option<String>,
        content: Option<String>,
    ) -> Self {
        Self {
            about,
            property,
            datatype,
            content,
        }
    }

    /// RDFa subject lexical value.
    #[must_use]
    pub fn about(&self) -> &str {
        &self.about
    }
    /// RDFa property lexical value.
    #[must_use]
    pub fn property(&self) -> &str {
        &self.property
    }
    /// Optional RDFa datatype lexical value.
    #[must_use]
    pub fn datatype(&self) -> Option<&str> {
        self.datatype.as_deref()
    }
    /// Optional RDFa content lexical value.
    #[must_use]
    pub fn content(&self) -> Option<&str> {
        self.content.as_deref()
    }
}

/// Lossless, typed projection of ODF table-cell values.
#[derive(Clone, Debug, PartialEq, Eq)]
#[non_exhaustive]
pub enum CellValue {
    /// A floating-point value retaining its source lexical spelling.
    Float { lexical: String },
    /// A percentage retaining its source lexical spelling.
    Percentage { lexical: String },
    /// A currency value retaining both value and currency lexicals.
    Currency {
        lexical: String,
        currency: Option<String>,
    },
    /// An ODF date lexical value.
    Date { lexical: String },
    /// An exact ODF duration value.
    Time { value: DurationValue },
    /// A boolean retaining its lexical spelling and parsed value.
    Boolean { lexical: String, value: bool },
    /// A string value, including absent `office:string-value` presence.
    String { value: Option<String> },
    /// An error value, including absent `office:string-value` presence.
    Error { value: Option<String> },
}

impl CellValue {
    /// Lexical scalar value, when this value family has one.
    #[must_use]
    pub fn lexical(&self) -> Option<&str> {
        match self {
            Self::Float { lexical }
            | Self::Percentage { lexical }
            | Self::Currency { lexical, .. }
            | Self::Date { lexical }
            | Self::Boolean { lexical, .. } => Some(lexical),
            Self::Time { value } => Some(value.as_str()),
            Self::String { .. } | Self::Error { .. } => None,
        }
    }
    /// Parsed boolean value, for boolean cells.
    #[must_use]
    pub const fn boolean_value(&self) -> Option<bool> {
        match self {
            Self::Boolean { value, .. } => Some(*value),
            _ => None,
        }
    }
    /// Optional currency lexical value.
    #[must_use]
    pub fn currency(&self) -> Option<&str> {
        match self {
            Self::Currency { currency, .. } => currency.as_deref(),
            _ => None,
        }
    }
    /// Exact duration value, for time cells.
    #[must_use]
    pub fn duration(&self) -> Option<&DurationValue> {
        match self {
            Self::Time { value } => Some(value),
            _ => None,
        }
    }
    /// Optional string cell value.
    #[must_use]
    pub fn string_value(&self) -> Option<&str> {
        match self {
            Self::String { value } => value.as_deref(),
            _ => None,
        }
    }
    /// Optional error cell value.
    #[must_use]
    pub fn error_value(&self) -> Option<&str> {
        match self {
            Self::Error { value } => value.as_deref(),
            _ => None,
        }
    }
}

/// The semantic kind of a table cell.
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
#[non_exhaustive]
pub enum CellKind {
    /// A normal `table:table-cell`.
    Cell,
    /// A covered cell represented by `table:covered-table-cell`.
    Covered,
}

/// A projected table column declaration.
#[derive(Clone, Debug, PartialEq, Eq)]
pub struct Column {
    default_cell_style_name: Option<String>,
    declared_repeated: Option<usize>,
    repeated: usize,
    style_name: Option<String>,
    visibility: Option<Visibility>,
    xml_id: Option<String>,
}

impl Column {
    pub(crate) const fn projected(
        style_name: Option<String>,
        default_cell_style_name: Option<String>,
        declared_repeated: Option<usize>,
        repeated: usize,
        visibility: Option<Visibility>,
        xml_id: Option<String>,
    ) -> Self {
        Self {
            default_cell_style_name,
            declared_repeated,
            repeated,
            style_name,
            visibility,
            xml_id,
        }
    }

    /// Column style reference.
    #[must_use]
    pub fn style_name(&self) -> Option<&str> {
        self.style_name.as_deref()
    }

    /// Default cell style reference.
    #[must_use]
    pub fn default_cell_style_name(&self) -> Option<&str> {
        self.default_cell_style_name.as_deref()
    }

    /// Number of physical columns represented by this declaration.
    #[must_use]
    pub const fn repeat_count(&self) -> usize {
        self.repeated
    }

    /// Explicit physical repetition, preserving absent-versus-one presence.
    #[must_use]
    pub const fn declared_repeat_count(&self) -> Option<usize> {
        self.declared_repeated
    }

    /// Optional column visibility.
    #[must_use]
    pub const fn visibility(&self) -> Option<Visibility> {
        self.visibility
    }

    /// Optional XML identity.
    #[must_use]
    pub fn xml_id(&self) -> Option<&str> {
        self.xml_id.as_deref()
    }
}

/// A projected table cell.
#[derive(Clone, Debug, PartialEq, Eq)]
pub struct Cell {
    columns_spanned: usize,
    declared_columns_spanned: Option<usize>,
    declared_matrix_columns_spanned: Option<usize>,
    declared_matrix_rows_spanned: Option<usize>,
    formula: Option<String>,
    content_validation_name: Option<String>,
    kind: CellKind,
    in_content_meta: Option<InContentMeta>,
    paragraphs: Vec<Paragraph>,
    protected: Option<bool>,
    protect: Option<bool>,
    repeated: usize,
    declared_repeated: Option<usize>,
    rows_spanned: usize,
    declared_rows_spanned: Option<usize>,
    style_name: Option<String>,
    text: String,
    typed_value: Option<CellValue>,
    value: Option<String>,
    value_type: Option<String>,
    xml_id: Option<String>,
}

impl Cell {
    #[allow(
        clippy::too_many_arguments,
        reason = "projection fields mirror ODF cell attributes"
    )]
    pub(crate) fn projected(
        kind: CellKind,
        style_name: Option<String>,
        declared_repeated: Option<usize>,
        repeated: usize,
        declared_columns_spanned: Option<usize>,
        columns_spanned: usize,
        declared_rows_spanned: Option<usize>,
        rows_spanned: usize,
        declared_matrix_columns_spanned: Option<usize>,
        declared_matrix_rows_spanned: Option<usize>,
        formula: Option<String>,
        content_validation_name: Option<String>,
        protect: Option<bool>,
        protected: Option<bool>,
        xml_id: Option<String>,
        in_content_meta: Option<InContentMeta>,
        typed_value: Option<CellValue>,
        value_type: Option<String>,
        value: Option<String>,
        text: String,
        paragraphs: Vec<Paragraph>,
    ) -> Self {
        Self {
            columns_spanned,
            declared_columns_spanned,
            declared_matrix_columns_spanned,
            declared_matrix_rows_spanned,
            formula,
            content_validation_name,
            kind,
            in_content_meta,
            paragraphs,
            protected,
            protect,
            repeated,
            declared_repeated,
            rows_spanned,
            declared_rows_spanned,
            style_name,
            text,
            typed_value,
            value,
            value_type,
            xml_id,
        }
    }

    /// Whether this is a normal or covered cell.
    #[must_use]
    pub const fn kind(&self) -> CellKind {
        self.kind
    }

    /// Cell style reference.
    #[must_use]
    pub fn style_name(&self) -> Option<&str> {
        self.style_name.as_deref()
    }

    /// Number of repeated physical cells represented by this node.
    #[must_use]
    pub const fn repeat_count(&self) -> usize {
        self.repeated
    }

    /// Explicit physical repetition, preserving absent-versus-one presence.
    #[must_use]
    pub const fn declared_repeat_count(&self) -> Option<usize> {
        self.declared_repeated
    }

    /// Number of columns spanned by the cell, including the first column.
    #[must_use]
    pub const fn columns_spanned(&self) -> usize {
        self.columns_spanned
    }

    /// Explicit normal column span, preserving source presence.
    #[must_use]
    pub const fn declared_columns_spanned(&self) -> Option<usize> {
        self.declared_columns_spanned
    }

    /// Number of rows spanned by the cell, including the first row.
    #[must_use]
    pub const fn rows_spanned(&self) -> usize {
        self.rows_spanned
    }

    /// Explicit normal row span, preserving source presence.
    #[must_use]
    pub const fn declared_rows_spanned(&self) -> Option<usize> {
        self.declared_rows_spanned
    }

    /// Explicit matrix column span, when present.
    #[must_use]
    pub const fn matrix_columns_spanned(&self) -> Option<usize> {
        self.declared_matrix_columns_spanned
    }

    /// Explicit matrix row span, when present.
    #[must_use]
    pub const fn matrix_rows_spanned(&self) -> Option<usize> {
        self.declared_matrix_rows_spanned
    }

    /// Inert table formula, if present.
    #[must_use]
    pub fn formula(&self) -> Option<&str> {
        self.formula.as_deref()
    }

    /// Optional named content-validation rule.
    #[must_use]
    pub fn content_validation_name(&self) -> Option<&str> {
        self.content_validation_name.as_deref()
    }

    /// Optional `table:protect` flag.
    #[must_use]
    pub const fn protect(&self) -> Option<bool> {
        self.protect
    }

    /// Optional `table:protected` flag.
    #[must_use]
    pub const fn protected(&self) -> Option<bool> {
        self.protected
    }

    /// Optional XML identity.
    #[must_use]
    pub fn xml_id(&self) -> Option<&str> {
        self.xml_id.as_deref()
    }

    /// Optional RDFa metadata.
    #[must_use]
    pub fn in_content_meta(&self) -> Option<&InContentMeta> {
        self.in_content_meta.as_ref()
    }

    /// Complete typed cell value, if the source declared a value type.
    #[must_use]
    pub fn typed_value(&self) -> Option<&CellValue> {
        self.typed_value.as_ref()
    }

    /// ODF value type, if present.
    #[must_use]
    pub fn value_type(&self) -> Option<&str> {
        self.value_type.as_deref()
    }

    /// Stored lexical value, if present.
    #[must_use]
    pub fn value(&self) -> Option<&str> {
        self.value.as_deref()
    }

    /// Projected visible cell text. Formulas are never evaluated.
    #[must_use]
    pub fn text(&self) -> &str {
        &self.text
    }

    /// Direct paragraphs in source order.
    #[must_use]
    pub fn paragraphs(&self) -> &[Paragraph] {
        &self.paragraphs
    }
}

/// A projected table row.
#[derive(Clone, Debug, PartialEq, Eq)]
pub struct Row {
    cells: Vec<Cell>,
    declared_repeated: Option<usize>,
    default_cell_style_name: Option<String>,
    repeated: usize,
    style_name: Option<String>,
    visibility: Option<Visibility>,
    xml_id: Option<String>,
}

impl Row {
    pub(crate) const fn projected(
        style_name: Option<String>,
        default_cell_style_name: Option<String>,
        declared_repeated: Option<usize>,
        repeated: usize,
        visibility: Option<Visibility>,
        xml_id: Option<String>,
        cells: Vec<Cell>,
    ) -> Self {
        Self {
            cells,
            declared_repeated,
            default_cell_style_name,
            repeated,
            style_name,
            visibility,
            xml_id,
        }
    }

    /// Row style reference.
    #[must_use]
    pub fn style_name(&self) -> Option<&str> {
        self.style_name.as_deref()
    }

    /// Default cell style reference.
    #[must_use]
    pub fn default_cell_style_name(&self) -> Option<&str> {
        self.default_cell_style_name.as_deref()
    }

    /// Number of physical rows represented by this node.
    #[must_use]
    pub const fn repeat_count(&self) -> usize {
        self.repeated
    }

    /// Explicit physical repetition, preserving absent-versus-one presence.
    #[must_use]
    pub const fn declared_repeat_count(&self) -> Option<usize> {
        self.declared_repeated
    }

    /// Optional row visibility.
    #[must_use]
    pub const fn visibility(&self) -> Option<Visibility> {
        self.visibility
    }

    /// Optional XML identity.
    #[must_use]
    pub fn xml_id(&self) -> Option<&str> {
        self.xml_id.as_deref()
    }

    /// Cells in source order.
    #[must_use]
    pub fn cells(&self) -> &[Cell] {
        &self.cells
    }
}

/// A projected `table:table` body structure.
#[derive(Clone, Debug, PartialEq, Eq)]
pub struct Table {
    columns: Vec<Column>,
    name: Option<String>,
    properties: TableProperties,
    rows: Vec<Row>,
    source: Option<TableSource>,
    style_name: Option<String>,
    declared_column_count: usize,
    column_count: usize,
}

impl Table {
    pub(crate) const fn projected(
        name: Option<String>,
        style_name: Option<String>,
        properties: TableProperties,
        source: Option<TableSource>,
        columns: Vec<Column>,
        rows: Vec<Row>,
        declared_column_count: usize,
        column_count: usize,
    ) -> Self {
        Self {
            columns,
            name,
            properties,
            rows,
            source,
            style_name,
            declared_column_count,
            column_count,
        }
    }

    /// Producer-visible table name.
    #[must_use]
    pub fn name(&self) -> Option<&str> {
        self.name.as_deref()
    }

    /// Table style reference.
    #[must_use]
    pub fn style_name(&self) -> Option<&str> {
        self.style_name.as_deref()
    }

    /// Optional table metadata attributes.
    #[must_use]
    pub const fn properties(&self) -> &TableProperties {
        &self.properties
    }

    /// Optional inert linked-table source.
    #[must_use]
    pub const fn source(&self) -> Option<&TableSource> {
        self.source.as_ref()
    }

    /// Explicit column declarations.
    #[must_use]
    pub fn columns(&self) -> &[Column] {
        &self.columns
    }

    /// Rows in source order.
    #[must_use]
    pub fn rows(&self) -> &[Row] {
        &self.rows
    }

    /// Number of physical columns declared by `table:table-column` nodes.
    /// Repeated declarations stay compact in [`Column::repeat_count`].
    #[must_use]
    pub const fn declared_column_count(&self) -> usize {
        self.declared_column_count
    }

    /// Logical table width after applying repeated cells and spans.
    ///
    /// This is the maximum of the declared column width and every row's
    /// checked expanded cell width. Repetition is retained as a count and is
    /// never materialized into additional cells.
    #[must_use]
    pub const fn column_count(&self) -> usize {
        self.column_count
    }
}
