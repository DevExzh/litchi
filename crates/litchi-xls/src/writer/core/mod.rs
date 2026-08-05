//! XLS file writer implementation.
//!
//! This module provides functionality to create and modify Microsoft Excel files
//! in the legacy binary format (.xls files) using the BIFF (Binary Interchange Format).
//!
//! The public surface is kept here as a semantic facade. Workbook state and caller-facing
//! configuration live in [`model`], BIFF-facing writer operations live in [`codec`], and OLE
//! package assembly lives in [`package`]. The established child modules remain the owners of
//! their specialized record families.
//!
//! # Example
//!
//! ```rust,no_run
//! use litchi_xls::Writer;
//!
//! let mut writer = Writer::new();
//! let sheet = writer.add_worksheet("Sheet1")?;
//! writer.write_string(sheet, 0, 0, "Hello")?;
//! writer.write_number(sheet, 0, 1, 42.0)?;
//! writer.save("output.xls")?;
//! # Ok::<(), Box<dyn std::error::Error>>(())
//! ```

mod codec;
mod model;
mod package;
#[cfg(test)]
mod tests;

mod comment;
mod conditional_format;
mod data_validation;
mod named_range;
pub mod shape;
mod shape_group;
mod stream;
mod worksheet;

use self::data_validation::DataValidationBiffPayload;
use self::model::{FileSharing, VbaWriteMetadata, WorkbookProtection};
use self::worksheet::{
    AutoFilterColumnDef, AutoFilterRange, CellPos, HorizontalPageBreak, Hyperlink, MergedRange,
    PivotCellXfRole, SheetProtection, SortConfig, VerticalPageBreak, WritableCell,
    WritablePivotDataItem, WritablePivotField, WritablePivotItem, WritablePivotTable,
    WritableWorksheet,
};

pub use self::comment::{CommentTextRunWrite, CommentWriteOptions};
pub use self::conditional_format::{
    ConditionalFormat, ConditionalFormat12Group, ConditionalFormat12Rule, ConditionalFormat12Type,
    ConditionalFormatGroup, ConditionalFormatOperator, ConditionalFormatRange,
    ConditionalFormatRule, ConditionalFormatType, ConditionalPattern,
};
pub use self::data_validation::{
    DataValidation, DataValidationErrorStyle, DataValidationFormulaKind, DataValidationImeMode,
    DataValidationOperator, DataValidationOptions, DataValidationRange, DataValidationTableOptions,
    DataValidationType,
};
pub use self::model::{
    AddInFunctionOptions, CalculationSettings, CellValue, CustomTableStyles, DdeOrOleItemOptions,
    DdeOrOleLinkOptions, ExternalCacheRowOptions, ExternalDefinedNameOptions, ExternalSheetOptions,
    ExternalWorkbookOptions, FunctionGroupOptions, PageSetupOptions, PivotCacheValue,
    PivotDataItemConfig, PivotFieldConfig, PivotItemConfig, PivotTableConfig,
    WorkbookEnvironmentOptions, WorkbookWindowOptions, WorksheetLayoutOptions, Writer,
};
pub use self::named_range::{DefinedName, DefinedNameRecordOptions};
pub use self::shape::{
    ShapeColor, ShapeFill, ShapeKind, ShapeLine, ShapeText, ShapeTextRun, ShapeWrite,
};
pub use self::shape_group::{ShapeGroupChild, ShapeGroupWrite};

fn validate_list_object_relationships(
    worksheets: &[WritableWorksheet],
    custom: Option<&CustomTableStyles>,
    defined_names: &[DefinedName],
    defined_name_records: &[(DefinedNameRecordOptions, crate::DefinedNameFutureRecords)],
) -> crate::Result<()> {
    model::validate_list_object_relationships(
        worksheets,
        custom,
        defined_names,
        defined_name_records,
    )
}
