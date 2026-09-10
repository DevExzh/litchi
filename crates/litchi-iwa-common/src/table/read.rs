//! Immutable, archive-free table reads.
//!
//! A [`crate::table::read::TableRead`] owns one shared
//! [`crate::table::model::Table`] and the sparse
//! comments resolved for that table.  The comment records deliberately carry
//! only typed semantic positions and display metadata.  Package object
//! identifiers, storage identifiers, UUIDs, and wire values stay in the
//! concrete iWork adapter that produced the read.

use super::cell::value::{FiniteF64, FiniteF64Error, Value};
use super::coordinate::{CellPosition, CellRange};
use super::model::{self, Cell, Dimensions, Error, Grid, GridBudget, InsertError, Table, View};
use std::fmt;

/// Result type for checked table-read construction and A1 lookup.
pub type Result<T> = std::result::Result<T, Error>;

/// A finite comment timestamp measured in seconds from Apple's 2001-01-01
/// reference date.
///
/// The value is stored as canonical `f64` bits so signed zero compares
/// consistently while NaN and either infinity are rejected at construction.
/// The timestamp contains no source or package identifier.
#[repr(transparent)]
#[derive(Clone, Copy, Eq, Hash, PartialEq)]
pub struct CommentTimestamp(u64);

impl CommentTimestamp {
    /// Creates a timestamp from Apple-epoch seconds.
    #[must_use]
    pub const fn new(seconds: f64) -> Option<Self> {
        if !seconds.is_finite() {
            return None;
        }
        let seconds = if seconds == 0.0 { 0.0 } else { seconds };
        Some(Self(seconds.to_bits()))
    }

    /// Returns Apple-epoch seconds.
    #[must_use]
    pub const fn as_f64(self) -> f64 {
        f64::from_bits(self.0)
    }

    /// Returns Apple-epoch seconds.
    #[must_use]
    pub const fn get(self) -> f64 {
        self.as_f64()
    }
}

impl fmt::Debug for CommentTimestamp {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter
            .debug_tuple("CommentTimestamp")
            .field(&self.as_f64())
            .finish()
    }
}

impl fmt::Display for CommentTimestamp {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        self.as_f64().fmt(formatter)
    }
}

impl PartialOrd for CommentTimestamp {
    fn partial_cmp(&self, other: &Self) -> Option<std::cmp::Ordering> {
        self.as_f64().partial_cmp(&other.as_f64())
    }
}

impl TryFrom<f64> for CommentTimestamp {
    type Error = FiniteF64Error;

    fn try_from(seconds: f64) -> std::result::Result<Self, Self::Error> {
        Self::new(seconds).ok_or(FiniteF64Error)
    }
}

impl From<CommentTimestamp> for f64 {
    fn from(timestamp: CommentTimestamp) -> Self {
        timestamp.as_f64()
    }
}

impl TryFrom<FiniteF64> for CommentTimestamp {
    type Error = FiniteF64Error;

    fn try_from(seconds: FiniteF64) -> std::result::Result<Self, Self::Error> {
        Self::try_from(seconds.get())
    }
}

/// Optional display metadata for a comment author.
///
/// `public_id` is the semantic public identifier supplied by the source
/// author payload.  It is not a native object identity and is retained only
/// when the source exposes it.
#[derive(Clone, Debug, Default, Eq, PartialEq)]
pub struct CommentAuthor {
    display_name: Option<Box<str>>,
    public_id: Option<Box<str>>,
}

impl CommentAuthor {
    /// Creates author metadata from already-owned strings.
    #[must_use]
    pub fn new(display_name: Option<Box<str>>, public_id: Option<Box<str>>) -> Self {
        Self {
            display_name,
            public_id,
        }
    }

    /// Creates author metadata by copying borrowed display fields.
    ///
    /// # Errors
    ///
    /// Returns [`Error::Allocation`] if either field cannot be copied.
    pub fn try_new(display_name: Option<&str>, public_id: Option<&str>) -> Result<Self> {
        let display_name = display_name
            .map(|value| try_boxed_str(value, "table comment author display name"))
            .transpose()?;
        let public_id = public_id
            .map(|value| try_boxed_str(value, "table comment author public id"))
            .transpose()?;
        Ok(Self::new(display_name, public_id))
    }

    /// Returns the optional display name.
    #[must_use]
    pub fn display_name(&self) -> Option<&str> {
        self.display_name.as_deref()
    }

    /// Returns the optional semantic public author identifier.
    #[must_use]
    pub fn public_id(&self) -> Option<&str> {
        self.public_id.as_deref()
    }

    /// Returns the owned display fields.
    #[must_use]
    pub fn into_parts(self) -> (Option<Box<str>>, Option<Box<str>>) {
        (self.display_name, self.public_id)
    }
}

/// One archive-free semantic reply attached to a table-cell comment.
#[derive(Clone, Debug, Eq, PartialEq)]
pub struct CommentReply {
    text: Box<str>,
    timestamp: Option<CommentTimestamp>,
    author: Option<CommentAuthor>,
}

impl CommentReply {
    /// Adopts already retained reply text without allocating.
    #[must_use]
    pub fn from_owned_parts(
        text: Box<str>,
        timestamp: Option<CommentTimestamp>,
        author: Option<CommentAuthor>,
    ) -> Self {
        Self {
            text,
            timestamp,
            author,
        }
    }

    /// Creates a reply with no optional metadata.
    #[must_use]
    pub fn new(text: impl Into<String>) -> Self {
        Self::with_metadata(text, None, None)
    }

    /// Creates a reply with optional timestamp and author metadata.
    #[must_use]
    pub fn with_metadata(
        text: impl Into<String>,
        timestamp: Option<CommentTimestamp>,
        author: Option<CommentAuthor>,
    ) -> Self {
        Self {
            text: text.into().into_boxed_str(),
            timestamp,
            author,
        }
    }

    /// Fallibly copies reply text with the common allocation error.
    ///
    /// # Errors
    ///
    /// Returns [`Error::Allocation`] if the text cannot be retained.
    pub fn try_new(text: impl AsRef<str>) -> Result<Self> {
        Self::try_with_metadata(text, None, None)
    }

    /// Fallibly copies reply text with optional metadata.
    ///
    /// # Errors
    ///
    /// Returns [`Error::Allocation`] if the text cannot be retained.
    pub fn try_with_metadata(
        text: impl AsRef<str>,
        timestamp: Option<CommentTimestamp>,
        author: Option<CommentAuthor>,
    ) -> Result<Self> {
        Ok(Self {
            text: try_boxed_str(text.as_ref(), "table comment reply text")?,
            timestamp,
            author,
        })
    }

    /// Borrows the reply text.
    #[must_use]
    pub fn text(&self) -> &str {
        &self.text
    }

    /// Returns the optional creation timestamp.
    #[must_use]
    pub const fn timestamp(&self) -> Option<CommentTimestamp> {
        self.timestamp
    }

    /// Borrows optional author metadata.
    #[must_use]
    pub fn author(&self) -> Option<&CommentAuthor> {
        self.author.as_ref()
    }

    /// Consumes the reply and returns its semantic fields.
    #[must_use]
    pub fn into_parts(self) -> (Box<str>, Option<CommentTimestamp>, Option<CommentAuthor>) {
        (self.text, self.timestamp, self.author)
    }
}

/// One archive-free semantic comment attached to a table cell.
///
/// Reply absence is represented separately from an explicitly read empty
/// reply list: [`Self::replies`] returns `None` when the source did not read
/// replies and `Some(&[])` when it did and found no replies.
#[derive(Clone, Debug, Eq, PartialEq)]
pub struct Comment {
    text: Box<str>,
    timestamp: Option<CommentTimestamp>,
    author: Option<CommentAuthor>,
    replies: Option<Box<[CommentReply]>>,
}

impl Comment {
    /// Adopts already retained text and an optional resolved reply list.
    ///
    /// This constructor performs no allocation, allowing bounded readers to
    /// transfer storage they have already reserved and charged.
    #[must_use]
    pub fn from_owned_parts(
        text: Box<str>,
        timestamp: Option<CommentTimestamp>,
        author: Option<CommentAuthor>,
        replies: Option<Box<[CommentReply]>>,
    ) -> Self {
        Self {
            text,
            timestamp,
            author,
            replies,
        }
    }

    /// Creates a comment with no optional metadata or read reply list.
    #[must_use]
    pub fn new(text: impl Into<String>) -> Self {
        Self::with_metadata(text, None, None)
    }

    /// Creates a comment with optional timestamp and author metadata.
    #[must_use]
    pub fn with_metadata(
        text: impl Into<String>,
        timestamp: Option<CommentTimestamp>,
        author: Option<CommentAuthor>,
    ) -> Self {
        Self {
            text: text.into().into_boxed_str(),
            timestamp,
            author,
            replies: None,
        }
    }

    /// Creates a comment with an explicitly read reply list.
    ///
    /// An empty iterator produces `Some(&[])` from [`Self::replies`], which
    /// distinguishes a read empty list from replies that were not requested.
    #[must_use]
    pub fn with_replies<I>(
        text: impl Into<String>,
        timestamp: Option<CommentTimestamp>,
        author: Option<CommentAuthor>,
        replies: I,
    ) -> Self
    where
        I: IntoIterator<Item = CommentReply>,
    {
        Self {
            text: text.into().into_boxed_str(),
            timestamp,
            author,
            replies: Some(replies.into_iter().collect::<Vec<_>>().into_boxed_slice()),
        }
    }

    /// Fallibly copies comment text with no optional metadata.
    ///
    /// # Errors
    ///
    /// Returns [`Error::Allocation`] if the text cannot be retained.
    pub fn try_new(text: impl AsRef<str>) -> Result<Self> {
        Self::try_with_metadata(text, None, None)
    }

    /// Fallibly copies comment text with optional metadata.
    ///
    /// # Errors
    ///
    /// Returns [`Error::Allocation`] if the text cannot be retained.
    pub fn try_with_metadata(
        text: impl AsRef<str>,
        timestamp: Option<CommentTimestamp>,
        author: Option<CommentAuthor>,
    ) -> Result<Self> {
        Ok(Self {
            text: try_boxed_str(text.as_ref(), "table comment text")?,
            timestamp,
            author,
            replies: None,
        })
    }

    /// Fallibly copies comment text and retains an explicitly read reply
    /// list.
    ///
    /// # Errors
    ///
    /// Returns [`Error::Allocation`] if the text or reply list cannot be
    /// retained.
    pub fn try_with_replies<I>(
        text: impl AsRef<str>,
        timestamp: Option<CommentTimestamp>,
        author: Option<CommentAuthor>,
        replies: I,
    ) -> Result<Self>
    where
        I: IntoIterator<Item = CommentReply>,
    {
        let text = try_boxed_str(text.as_ref(), "table comment text")?;
        let iterator = replies.into_iter();
        let (lower_bound, _) = iterator.size_hint();
        let mut retained = Vec::new();
        retained
            .try_reserve(lower_bound)
            .map_err(|_allocation| Error::Allocation {
                resource: "table comment replies",
                amount: lower_bound,
            })?;
        for reply in iterator {
            retained
                .try_reserve(1)
                .map_err(|_allocation| Error::Allocation {
                    resource: "table comment replies",
                    amount: retained.len().saturating_add(1),
                })?;
            retained.push(reply);
        }
        Ok(Self {
            text,
            timestamp,
            author,
            replies: Some(retained.into_boxed_slice()),
        })
    }

    /// Borrows the comment text.
    #[must_use]
    pub fn text(&self) -> &str {
        &self.text
    }

    /// Returns the optional creation timestamp.
    #[must_use]
    pub const fn timestamp(&self) -> Option<CommentTimestamp> {
        self.timestamp
    }

    /// Borrows optional author metadata.
    #[must_use]
    pub fn author(&self) -> Option<&CommentAuthor> {
        self.author.as_ref()
    }

    /// Borrows the reply list when replies were read.
    #[must_use]
    pub fn replies(&self) -> Option<&[CommentReply]> {
        self.replies.as_deref()
    }

    /// Iterates over replies, yielding no items when replies were not read.
    #[must_use]
    pub fn iter_replies(&self) -> impl ExactSizeIterator<Item = &CommentReply> + '_ {
        self.replies.as_deref().unwrap_or(&[]).iter()
    }

    /// Returns the number of retained replies.
    #[must_use]
    pub fn reply_count(&self) -> usize {
        self.replies.as_deref().map_or(0, <[CommentReply]>::len)
    }

    /// Returns whether the source supplied a reply list, including an empty
    /// list.
    #[must_use]
    pub const fn replies_were_read(&self) -> bool {
        self.replies.is_some()
    }

    /// Consumes the comment and returns its semantic fields.
    #[must_use]
    pub fn into_parts(
        self,
    ) -> (
        Box<str>,
        Option<CommentTimestamp>,
        Option<CommentAuthor>,
        Option<Box<[CommentReply]>>,
    ) {
        (self.text, self.timestamp, self.author, self.replies)
    }
}

/// A typed semantic position and its cell comment.
#[derive(Clone, Debug, Eq, PartialEq)]
pub struct CellComment {
    position: CellPosition,
    comment: Comment,
}

impl CellComment {
    /// Creates a comment record at a zero-based table position.
    #[must_use]
    pub const fn new(position: CellPosition, comment: Comment) -> Self {
        Self { position, comment }
    }

    /// Returns the typed zero-based position.
    #[must_use]
    pub const fn position(&self) -> CellPosition {
        self.position
    }

    /// Borrows the comment.
    #[must_use]
    pub const fn comment(&self) -> &Comment {
        &self.comment
    }

    /// Consumes the record and returns its position and comment.
    #[must_use]
    pub fn into_parts(self) -> (CellPosition, Comment) {
        (self.position, self.comment)
    }
}

impl From<(CellPosition, Comment)> for CellComment {
    fn from((position, comment): (CellPosition, Comment)) -> Self {
        Self::new(position, comment)
    }
}

/// An immutable table and the sparse comments read for it.
///
/// The table remains the shared [`super::model::Table`] value.  Comments are
/// stored in row-major order and can be queried without materializing a dense
/// comment grid.
#[derive(Clone, Debug, PartialEq)]
pub struct TableRead {
    table: Table,
    comments: Box<[CellComment]>,
}

impl From<Table> for TableRead {
    fn from(table: Table) -> Self {
        Self::new(table)
    }
}

impl TableRead {
    /// Creates an immutable read with no comments.
    #[must_use]
    pub fn new(table: Table) -> Self {
        Self {
            table,
            comments: Box::new([]),
        }
    }

    /// Creates a fallible builder over an existing semantic table.
    #[must_use]
    pub fn builder(table: Table) -> Builder {
        Builder::new(table)
    }

    /// Builds a read from an iterable of sparse comments.
    ///
    /// # Errors
    ///
    /// Returns [`Error::OutOfBounds`] for a comment position outside the
    /// table, [`Error::DuplicatePosition`] for repeated positions, or
    /// [`Error::Allocation`] when comment storage cannot be reserved.
    pub fn try_from_comments<I>(table: Table, comments: I) -> Result<Self>
    where
        I: IntoIterator<Item = CellComment>,
    {
        let mut builder = Builder::new(table);
        for comment in comments {
            builder
                .push_comment(comment)
                .map_err(|error| error.into_parts().0)?;
        }
        builder.finish()
    }

    /// Builds a read from an iterable of owned sparse comments.
    ///
    /// Use [`Self::try_from_owned_parts`] to reuse an existing vector.
    ///
    /// # Errors
    ///
    /// Returns the same bounds, duplicate-position, and allocation errors as
    /// [`Self::try_from_comments`].
    pub fn try_from_parts<I>(table: Table, comments: I) -> Result<Self>
    where
        I: IntoIterator<Item = CellComment>,
    {
        Self::try_from_comments(table, comments)
    }

    /// Validates and seals an existing sparse comment vector without rebuilding it.
    ///
    /// Comments are sorted in place. Callers that already reserved and accounted
    /// for their records can transfer that storage directly.
    ///
    /// # Errors
    ///
    /// Returns [`Error::OutOfBounds`] for a position outside the table or
    /// [`Error::DuplicatePosition`] for repeated positions.
    pub fn try_from_owned_parts(table: Table, comments: Vec<CellComment>) -> Result<Self> {
        for comment in &comments {
            ensure_position(table.dimensions(), comment.position())?;
        }
        Builder {
            table,
            comments,
            sorted: false,
        }
        .finish()
    }

    /// Borrows the shared semantic table.
    #[must_use]
    pub const fn table(&self) -> &Table {
        &self.table
    }

    /// Consumes the read and returns its shared semantic table.
    #[must_use]
    pub fn into_table(self) -> Table {
        self.table
    }

    /// Borrows sparse comments in row-major position order.
    #[must_use]
    pub fn comments(&self) -> impl ExactSizeIterator<Item = &CellComment> + '_ {
        self.comments.iter()
    }

    /// Iterates over sparse comments in row-major position order.
    #[must_use]
    pub fn iter_comments(&self) -> impl ExactSizeIterator<Item = &CellComment> + '_ {
        self.comments.iter()
    }

    /// Returns the number of retained comments.
    #[must_use]
    pub fn comment_count(&self) -> usize {
        self.comments.len()
    }

    /// Returns the full positioned comment record at `position`.
    #[must_use]
    pub fn comment_at(&self, position: CellPosition) -> Option<&CellComment> {
        self.comments
            .binary_search_by_key(&position, CellComment::position)
            .ok()
            .map(|index| &self.comments[index])
    }

    /// Returns only the comment value at `position`.
    #[must_use]
    pub fn get_comment(&self, position: CellPosition) -> Option<&Comment> {
        self.comment_at(position).map(CellComment::comment)
    }

    /// Looks up a comment by a checked A1 address.
    ///
    /// # Errors
    ///
    /// Returns a typed address or bounds error.  An in-grid position with no
    /// comment returns `Ok(None)`.
    pub fn get_comment_a1(&self, address: &str) -> Result<Option<&Comment>> {
        let position = CellPosition::from_a1(address)?;
        ensure_position(self.table.dimensions(), position)?;
        Ok(self.get_comment(position))
    }

    /// Looks up the full positioned comment record by a checked A1 address.
    ///
    /// # Errors
    ///
    /// Returns a typed address or bounds error.  An in-grid position with no
    /// comment returns `Ok(None)`.
    pub fn comment_at_a1(&self, address: &str) -> Result<Option<&CellComment>> {
        let position = CellPosition::from_a1(address)?;
        ensure_position(self.table.dimensions(), position)?;
        Ok(self.comment_at(position))
    }

    /// Borrows a materialized value at a typed position.
    #[must_use]
    pub fn get(&self, position: CellPosition) -> Option<&Value> {
        self.table.get(position)
    }

    /// Looks up a materialized value by a checked A1 address.
    pub fn get_a1(&self, address: &str) -> model::Result<Option<&Value>> {
        self.table.get_a1(address)
    }

    /// Returns the stored or missing view at a typed position.
    #[must_use]
    pub fn view(&self, position: CellPosition) -> View<'_> {
        self.table.view(position)
    }

    /// Looks up a stored or missing view by a checked A1 address.
    pub fn view_a1(&self, address: &str) -> model::Result<View<'_>> {
        self.table.view_a1(address)
    }

    /// Iterates sparse cells in a checked range.
    pub fn cells(&self, range: CellRange) -> model::Result<impl Iterator<Item = &Cell> + '_> {
        self.table.cells(range)
    }

    /// Iterates sparse cells selected by a checked A1 range.
    pub fn cells_a1(&self, address: &str) -> model::Result<impl Iterator<Item = &Cell> + '_> {
        self.table.cells_a1(address)
    }

    /// Iterates all materialized sparse cells in row-major order.
    #[must_use]
    pub fn iter_cells(&self) -> impl ExactSizeIterator<Item = &Cell> + '_ {
        self.table.iter_cells()
    }

    /// Returns the declared table extent.
    #[must_use]
    pub const fn dimensions(&self) -> Dimensions {
        self.table.dimensions()
    }

    /// Returns the semantic table name.
    #[must_use]
    pub fn name(&self) -> &str {
        self.table.name()
    }

    /// Returns the declared row count.
    #[must_use]
    pub const fn row_count(&self) -> u32 {
        self.table.row_count()
    }

    /// Returns the declared column count.
    #[must_use]
    pub const fn column_count(&self) -> u32 {
        self.table.column_count()
    }

    /// Returns the number of materialized sparse cells.
    #[must_use]
    pub fn cell_count(&self) -> usize {
        self.table.cell_count()
    }

    /// Returns the number of materialized non-empty values.
    #[must_use]
    pub fn non_empty_cell_count(&self) -> usize {
        self.table.non_empty_cell_count()
    }

    /// Iterates column headers in native order.
    #[must_use]
    pub fn column_headers(&self) -> impl ExactSizeIterator<Item = &str> + '_ {
        self.table.column_headers()
    }

    /// Iterates row headers in native order.
    #[must_use]
    pub fn row_headers(&self) -> impl ExactSizeIterator<Item = &str> + '_ {
        self.table.row_headers()
    }

    /// Projects the table values to RFC 4180-compatible CSV text.
    #[must_use]
    pub fn to_csv(&self) -> String {
        self.table.to_csv()
    }

    /// Creates a bounded dense view over the table values.
    pub fn grid(&self, range: CellRange, budget: GridBudget) -> model::Result<Grid<'_>> {
        self.table.grid(range, budget)
    }

    /// Consumes the read and returns its table and sparse comments.
    #[must_use]
    pub fn into_parts(self) -> (Table, Box<[CellComment]>) {
        (self.table, self.comments)
    }
}

/// A fallible builder for an immutable [`TableRead`].
#[derive(Clone, Debug, PartialEq)]
pub struct Builder {
    table: Table,
    comments: Vec<CellComment>,
    sorted: bool,
}

impl Builder {
    /// Creates an empty read builder over `table`.
    #[must_use]
    pub fn new(table: Table) -> Self {
        Self {
            table,
            comments: Vec::new(),
            sorted: true,
        }
    }

    /// Borrows the table being decorated.
    #[must_use]
    pub const fn table(&self) -> &Table {
        &self.table
    }

    /// Returns the declared table extent.
    #[must_use]
    pub const fn dimensions(&self) -> Dimensions {
        self.table.dimensions()
    }

    /// Returns the number of staged comments.
    #[must_use]
    pub fn comment_count(&self) -> usize {
        self.comments.len()
    }

    /// Iterates staged comments in their current storage order.
    #[must_use]
    pub fn comments(&self) -> impl ExactSizeIterator<Item = &CellComment> + '_ {
        self.comments.iter()
    }

    /// Returns the staged comment at a typed position.
    #[must_use]
    pub fn get_comment(&self, position: CellPosition) -> Option<&Comment> {
        if self.sorted {
            self.comments
                .binary_search_by_key(&position, CellComment::position)
                .ok()
                .map(|index| self.comments[index].comment())
        } else {
            self.comments
                .iter()
                .find(|comment| comment.position() == position)
                .map(CellComment::comment)
        }
    }

    /// Inserts a comment and rejects duplicate positions.
    ///
    /// The rejected comment is returned for every bounds, duplicate, and
    /// allocation error, preserving ownership at the read boundary.
    pub fn insert_comment(
        &mut self,
        position: CellPosition,
        comment: Comment,
    ) -> std::result::Result<(), InsertError<Comment>> {
        if let Err(error) = ensure_position(self.table.dimensions(), position) {
            return Err(InsertError::new(error, comment));
        }
        self.sort_comments();
        let index = match self
            .comments
            .binary_search_by_key(&position, CellComment::position)
        {
            Ok(_index) => {
                return Err(InsertError::new(
                    Error::DuplicatePosition { position },
                    comment,
                ));
            },
            Err(index) => index,
        };
        if self.comments.try_reserve(1).is_err() {
            return Err(InsertError::new(
                Error::Allocation {
                    resource: "table comments",
                    amount: 1,
                },
                comment,
            ));
        }
        self.comments
            .insert(index, CellComment::new(position, comment));
        Ok(())
    }

    /// Alias for [`Self::insert_comment`] using the shorter builder verb.
    pub fn insert(
        &mut self,
        position: CellPosition,
        comment: Comment,
    ) -> std::result::Result<(), InsertError<Comment>> {
        self.insert_comment(position, comment)
    }

    /// Replaces or inserts a comment at a typed position.
    ///
    /// The rejected comment is returned for bounds or allocation failures.
    pub fn set_comment(
        &mut self,
        position: CellPosition,
        comment: Comment,
    ) -> std::result::Result<(), InsertError<Comment>> {
        if let Err(error) = ensure_position(self.table.dimensions(), position) {
            return Err(InsertError::new(error, comment));
        }
        self.sort_comments();
        match self
            .comments
            .binary_search_by_key(&position, CellComment::position)
        {
            Ok(index) => {
                self.comments[index] = CellComment::new(position, comment);
            },
            Err(index) => {
                if self.comments.try_reserve(1).is_err() {
                    return Err(InsertError::new(
                        Error::Allocation {
                            resource: "table comments",
                            amount: 1,
                        },
                        comment,
                    ));
                }
                self.comments
                    .insert(index, CellComment::new(position, comment));
            },
        }
        Ok(())
    }

    /// Alias for [`Self::set_comment`] using the shorter builder verb.
    pub fn set(
        &mut self,
        position: CellPosition,
        comment: Comment,
    ) -> std::result::Result<(), InsertError<Comment>> {
        self.set_comment(position, comment)
    }

    /// Appends a positioned comment for high-throughput ingestion.
    ///
    /// Appended records are sorted once by [`Self::finish`]. Duplicate
    /// positions are reported by that finish operation.
    #[allow(
        clippy::result_large_err,
        reason = "Returning the rejected positioned comment preserves the fallible builder's ownership contract"
    )]
    pub fn push_comment(
        &mut self,
        cell_comment: CellComment,
    ) -> std::result::Result<(), InsertError<CellComment>> {
        if let Err(error) = ensure_position(self.table.dimensions(), cell_comment.position()) {
            return Err(InsertError::new(error, cell_comment));
        }
        if let Some(last) = self.comments.last().map(CellComment::position) {
            self.sorted &= cell_comment.position() >= last;
        }
        if self.comments.try_reserve(1).is_err() {
            return Err(InsertError::new(
                Error::Allocation {
                    resource: "table comments",
                    amount: 1,
                },
                cell_comment,
            ));
        }
        self.comments.push(cell_comment);
        Ok(())
    }

    /// Alias for [`Self::push_comment`] using the shorter builder verb.
    #[allow(
        clippy::result_large_err,
        reason = "The alias retains push_comment's ownership-preserving error type"
    )]
    pub fn push(
        &mut self,
        cell_comment: CellComment,
    ) -> std::result::Result<(), InsertError<CellComment>> {
        self.push_comment(cell_comment)
    }

    /// Consumes the builder and returns its sorted sparse comments.
    #[must_use]
    pub fn into_comments(mut self) -> Box<[CellComment]> {
        self.sort_comments();
        self.comments.into_boxed_slice()
    }

    /// Sorts and seals the builder into an immutable table read.
    ///
    /// # Errors
    ///
    /// Returns [`Error::DuplicatePosition`] when records appended with
    /// [`Self::push_comment`] contain the same position.
    pub fn finish(mut self) -> Result<TableRead> {
        self.sort_comments();
        for pair in self.comments.windows(2) {
            if pair[0].position() == pair[1].position() {
                return Err(Error::DuplicatePosition {
                    position: pair[0].position(),
                });
            }
        }
        Ok(TableRead {
            table: self.table,
            comments: self.comments.into_boxed_slice(),
        })
    }

    /// Alias for [`Self::finish`].
    pub fn build(self) -> Result<TableRead> {
        self.finish()
    }

    fn sort_comments(&mut self) {
        if !self.sorted {
            self.comments.sort_unstable_by_key(CellComment::position);
            self.sorted = true;
        }
    }
}

fn ensure_position(dimensions: Dimensions, position: CellPosition) -> Result<()> {
    if position.row() >= dimensions.rows() || position.column() >= dimensions.columns() {
        return Err(Error::OutOfBounds {
            position,
            dimensions,
        });
    }
    Ok(())
}

fn try_boxed_str(value: &str, resource: &'static str) -> Result<Box<str>> {
    let mut retained = String::new();
    retained
        .try_reserve_exact(value.len())
        .map_err(|_allocation| Error::Allocation {
            resource,
            amount: value.len(),
        })?;
    retained.push_str(value);
    Ok(retained.into_boxed_str())
}

#[cfg(test)]
mod tests {
    use super::*;

    fn table() -> Table {
        let mut builder = Table::builder("Summary", Dimensions::new(3, 3));
        builder
            .set(CellPosition::new(0, 0), Value::Text("value".to_owned()))
            .expect("test cell is in bounds");
        builder.finish().expect("test table is valid")
    }

    #[test]
    fn timestamp_rejects_non_finite_values_and_canonicalizes_zero() {
        assert_eq!(CommentTimestamp::new(f64::NAN), None);
        assert_eq!(CommentTimestamp::new(f64::INFINITY), None);
        assert_eq!(CommentTimestamp::new(f64::NEG_INFINITY), None);
        assert_eq!(CommentTimestamp::new(-0.0), CommentTimestamp::new(0.0));
        assert_eq!(
            CommentTimestamp::new(42.5).map(|value| value.get()),
            Some(42.5)
        );
    }

    #[test]
    fn owned_comments_are_sorted_and_validated() {
        let comment =
            |row, column| CellComment::new(CellPosition::new(row, column), Comment::new("note"));
        let read = TableRead::try_from_owned_parts(table(), vec![comment(2, 1), comment(0, 0)])
            .expect("valid comments");
        assert_eq!(
            read.comments().next().unwrap().position(),
            CellPosition::new(0, 0)
        );
        assert!(matches!(
            TableRead::try_from_owned_parts(table(), vec![comment(3, 0)]),
            Err(Error::OutOfBounds { .. })
        ));
        assert!(matches!(
            TableRead::try_from_owned_parts(table(), vec![comment(0, 0), comment(0, 0)]),
            Err(Error::DuplicatePosition { .. })
        ));
    }

    #[test]
    fn owned_comment_metadata_reuses_retained_storage() {
        let text: Box<str> = "root".into();
        let text_pointer = text.as_ptr();
        let reply_text: Box<str> = "reply".into();
        let reply_pointer = reply_text.as_ptr();
        let replies =
            vec![CommentReply::from_owned_parts(reply_text, None, None)].into_boxed_slice();
        let replies_pointer = replies.as_ptr();
        let comment = Comment::from_owned_parts(text, None, None, Some(replies));
        assert_eq!(comment.text().as_ptr(), text_pointer);
        let replies = comment.replies().expect("resolved replies");
        assert_eq!(replies.as_ptr(), replies_pointer);
        assert_eq!(replies[0].text().as_ptr(), reply_pointer);
    }

    #[test]
    fn read_keeps_values_and_sparse_comments_in_position_order() {
        let author = CommentAuthor::new(Some("Ada".into()), Some("public-1".into()));
        let comment = Comment::with_replies(
            "root",
            CommentTimestamp::new(1.5),
            Some(author.clone()),
            [CommentReply::with_metadata(
                "reply",
                CommentTimestamp::new(2.5),
                Some(author),
            )],
        );
        let mut builder = TableRead::builder(table());
        builder
            .push_comment(CellComment::new(
                CellPosition::new(2, 2),
                Comment::new("late"),
            ))
            .expect("late comment is in bounds");
        builder
            .insert_comment(CellPosition::new(0, 1), comment)
            .expect("first comment is unique");
        let read = builder.finish().expect("comments are valid");

        assert_eq!(
            read.get(CellPosition::new(0, 0)),
            Some(&Value::Text("value".to_owned()))
        );
        assert_eq!(
            read.get_comment(CellPosition::new(0, 1)).map(Comment::text),
            Some("root")
        );
        assert_eq!(
            read.get_comment_a1("C3")
                .map(|value| value.map(Comment::text)),
            Ok(Some("late"))
        );
        assert_eq!(
            read.iter_comments()
                .map(CellComment::position)
                .collect::<Vec<_>>(),
            [CellPosition::new(0, 1), CellPosition::new(2, 2),]
        );
        let root = read
            .get_comment(CellPosition::new(0, 1))
            .expect("root comment");
        assert_eq!(root.replies().map(<[CommentReply]>::len), Some(1));
        assert_eq!(
            root.author().and_then(CommentAuthor::display_name),
            Some("Ada")
        );
    }

    #[test]
    fn strict_insert_and_finish_reject_duplicate_positions() {
        let mut builder = TableRead::builder(table());
        builder
            .push(CellComment::new(
                CellPosition::new(1, 1),
                Comment::new("one"),
            ))
            .expect("first comment is in bounds");
        let duplicate = builder.push(CellComment::new(
            CellPosition::new(1, 1),
            Comment::new("two"),
        ));
        assert!(duplicate.is_ok());
        assert!(matches!(
            builder.finish(),
            Err(Error::DuplicatePosition { position }) if position == CellPosition::new(1, 1)
        ));

        let mut builder = TableRead::builder(table());
        builder
            .insert(CellPosition::new(1, 1), Comment::new("one"))
            .expect("first comment is unique");
        let duplicate = builder
            .insert(CellPosition::new(1, 1), Comment::new("two"))
            .expect_err("strict insert rejects duplicates");
        let (error, comment) = duplicate.into_parts();
        assert!(matches!(error, Error::DuplicatePosition { .. }));
        assert_eq!(comment.text(), "two");
    }

    #[test]
    fn builder_rejects_out_of_bounds_comments_without_publishing_them() {
        let mut builder = TableRead::builder(table());
        let error = builder
            .insert_comment(CellPosition::new(3, 0), Comment::new("outside"))
            .expect_err("row three is outside a three-row table");
        let (error, comment) = error.into_parts();
        assert!(matches!(error, Error::OutOfBounds { .. }));
        assert_eq!(comment.text(), "outside");
        assert_eq!(builder.comment_count(), 0);
    }
}
