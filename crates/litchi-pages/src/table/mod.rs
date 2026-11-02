//! Archive-free semantic Pages table values.

/// The ordered catalog of rooted body tables.
pub mod catalog {
    pub use crate::package::body_table_catalog::{Catalog, Error, LimitKind, Snapshot};
}
/// Typed failures of the body-table cell reader.
pub mod cells {
    pub use crate::package::table_cells::{Error, LimitKind};
}
/// Formula values of [`values::Edit::set_formula_cached`] (record 0376a).
pub mod formula {
    pub use litchi_iwa_common::formula::{
        FormulaAxisReference, FormulaBinaryOperator, FormulaCachedValue, FormulaCellReference,
        FormulaExpression,
    };
}

/// Cell values of one body table (record 0309).
pub mod values {
    pub use crate::package::body_table_values::{
        Commit, Diagnostics, Edit, Error, LimitKind, Patch, Refusal, Value,
    };
}
/// Exact-source creation of a body table cloned from the document's first
/// body table (record 0329), or of a populated copy of an existing body
/// table (record 0371).
pub mod creation {
    pub use crate::package::body_table_creation::{
        Commit, Diagnostics, Edit, Error, LimitKind, NewTable, Patch, Refusal,
    };
}
/// Display formats of stored body-table cells: one family over the neutral
/// `DataFormat` (record 0379).
pub mod format {
    pub use crate::package::body_table_cell_format::{
        Commit, Diagnostics, Edit, Error, LimitKind, Patch, Refusal,
    };
    pub use litchi_iwa_common::table::cell::data_format::DataFormat;
    pub use litchi_iwa_common::table::coordinate::At;
}
/// One character or paragraph property of whole body-table cells' text,
/// started by [`crate::Package::edit_body_table_cell_text`] and read by
/// [`crate::Package::body_table_cell_text`] (record 0412).
pub mod cell_text {
    pub use crate::package::body_table_cell_text::{
        Commit, Diagnostics, Edit, Error, LimitKind, Patch, Refusal,
    };
    pub use litchi_iwa_common::table::coordinate::At;
    pub use litchi_iwa_text_wire::paragraph_style::{Property, PropertyKind};
}
/// Paragraph lists of body-table cells' text, started by
/// [`crate::Package::edit_body_table_cell_lists`]: a plain cell is promoted
/// to rich text on its first list change (record 0412).
pub mod cell_list {
    pub use crate::package::body_table_cell_list::{
        Commit, Diagnostics, Edit, Error, LimitKind, Patch, Refusal,
    };
    pub use litchi_iwa_common::table::coordinate::At;
    pub use litchi_iwa_text::paragraph::list::{
        ParagraphList, ParagraphListBullet, ParagraphListBulletBaselineOffset,
        ParagraphListBulletGeometry, ParagraphListBulletScale, ParagraphListIndentation,
        ParagraphListLabelColor, ParagraphListLabelIndent, ParagraphListLevel,
        ParagraphListLevelPlacement, ParagraphListNumberFormat, ParagraphListNumberPunctuation,
        ParagraphListNumberScale, ParagraphListNumberSequence, ParagraphListNumberTiering,
        ParagraphListNumbering, ParagraphListPlacement, ParagraphListStart, ParagraphListTextGap,
    };
    pub use litchi_iwa_text::position::TextPosition;
    pub use litchi_iwa_text_wire::list::{Property, PropertyKind};
}
/// Exact-source resize of one rooted body table to a new row and column
/// count (record 0371).
pub mod size {
    pub use crate::package::body_table_resize::{
        Commit, Diagnostics, Edit, Error, LimitKind, Patch, Refusal,
    };
}
/// Exact-source deletion of one rooted body table.
pub mod deletion {
    pub use crate::package::body_table_deletion::{
        Commit, Diagnostics, Edit, Error, LimitKind, Patch, Refusal,
    };
}
/// Effective style, gridline, banding, and row-sizing settings for a body table.
pub mod appearance;
/// Cell fill and text layout of one body-table cell (record 0371).
pub mod cell_appearance;
/// Checked, archive-free row and column dimension values.
pub mod dimension;
/// Lossless header, footer, freeze, and print-repetition settings.
pub mod headers;
/// Archive-free hidden-row and hidden-column values.
pub mod hidden_axes;
/// Interactive lock state for a body-attached table.
pub mod lock;
/// Archive-free merged-cell geometry and topology algebra.
pub mod merge;
/// Archive-free table values together with sparse cell comments.
pub mod read {
    pub use litchi_iwa_common::table::cell::value::Value;
    pub use litchi_iwa_common::table::read::{
        Builder, CellComment, Comment, CommentAuthor, CommentReply, CommentTimestamp, Result,
        TableRead,
    };
}
/// Exact-source table-cell comment and direct-reply transactions.
pub mod comments {
    pub use crate::package::body_table_comments::{
        BodyTableCommentCommit as Commit, BodyTableCommentDiagnostics as Diagnostics,
        BodyTableCommentEdit as Edit, BodyTableCommentError as Error,
        BodyTableCommentLimitKind as LimitKind, BodyTableCommentPatch as Patch,
        BodyTableCommentPath as Path,
    };
    pub use litchi_iwa_common::table::read::{
        Comment, CommentAuthor, CommentReply, CommentTimestamp,
    };
}
/// Validated, archive-free names for Pages body tables.
pub mod name;
/// Persisted row-order rules for a body-attached table.
pub mod sort;
pub mod title;
