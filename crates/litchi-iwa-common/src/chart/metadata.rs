//! Archive-free chart metadata shared by concrete iWork format owners.

use super::kind::Kind;

/// Owned, format-neutral metadata describing one chart.
///
/// The value deliberately contains no native object identifier, archive
/// reference, or protobuf type. Format adapters resolve those details while
/// reading a package and precharge any required allocations before passing
/// their owned values to [`Self::from_owned`].
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct ChartMetadata {
    kind: Kind,
    title: Option<String>,
    row_names: Vec<String>,
    column_names: Vec<String>,
    series_count: usize,
    contains_default_data: bool,
}

impl ChartMetadata {
    /// Build chart metadata from already-owned, validated values.
    ///
    /// This constructor does not copy any of its string or vector arguments.
    /// Callers that decode a bounded format should charge those allocations
    /// before constructing the value.
    #[must_use]
    pub fn from_owned(
        kind: Kind,
        title: Option<String>,
        row_names: Vec<String>,
        column_names: Vec<String>,
        series_count: usize,
        contains_default_data: bool,
    ) -> Self {
        Self {
            kind,
            title,
            row_names,
            column_names,
            series_count,
            contains_default_data,
        }
    }

    /// Return the chart's lossless native kind.
    #[must_use]
    pub const fn kind(&self) -> Kind {
        self.kind
    }

    /// Return the optional chart title.
    #[must_use]
    pub fn title(&self) -> Option<&str> {
        self.title.as_deref()
    }

    /// Return the chart's row labels.
    #[must_use]
    pub fn row_names(&self) -> &[String] {
        &self.row_names
    }

    /// Return the chart's column labels.
    #[must_use]
    pub fn column_names(&self) -> &[String] {
        &self.column_names
    }

    /// Return the number of data series represented by the chart.
    #[must_use]
    pub const fn series_count(&self) -> usize {
        self.series_count
    }

    /// Return whether the chart still contains the producer's default data.
    #[must_use]
    pub const fn contains_default_data(&self) -> bool {
        self.contains_default_data
    }

    /// Borrow all textual chart metadata in title, row, then column order.
    pub fn all_text(&self) -> impl Iterator<Item = &str> {
        self.title
            .iter()
            .map(String::as_str)
            .chain(self.row_names.iter().map(String::as_str))
            .chain(self.column_names.iter().map(String::as_str))
    }

    /// Return whether any title or axis-label text is present.
    #[must_use]
    pub fn has_content(&self) -> bool {
        self.title.is_some() || !self.row_names.is_empty() || !self.column_names.is_empty()
    }
}

#[cfg(test)]
mod tests {
    use super::ChartMetadata;
    use crate::chart::kind::Kind;

    #[test]
    fn empty_metadata_preserves_scalar_state_and_has_no_content() {
        let metadata = ChartMetadata::from_owned(
            Kind::from_native(9_001),
            None,
            Vec::new(),
            Vec::new(),
            7,
            true,
        );

        assert_eq!(metadata.kind(), Kind::from_native(9_001));
        assert_eq!(metadata.title(), None);
        assert!(metadata.row_names().is_empty());
        assert!(metadata.column_names().is_empty());
        assert_eq!(metadata.series_count(), 7);
        assert!(metadata.contains_default_data());
        assert!(!metadata.has_content());
        assert!(metadata.all_text().next().is_none());
    }

    #[test]
    fn metadata_borrows_owned_text_in_stable_order() {
        let title = String::from("Revenue");
        let rows = vec![String::from("Q1"), String::from("Q2")];
        let columns = vec![String::from("North"), String::from("South")];
        let title_ptr = title.as_ptr();
        let row_ptr = rows[0].as_ptr();
        let column_ptr = columns[0].as_ptr();
        let metadata =
            ChartMetadata::from_owned(Kind::Column2d, Some(title), rows, columns, 2, false);

        assert_eq!(
            metadata.all_text().collect::<Vec<_>>(),
            ["Revenue", "Q1", "Q2", "North", "South",]
        );
        assert_eq!(metadata.title().expect("title").as_ptr(), title_ptr);
        assert_eq!(metadata.row_names()[0].as_ptr(), row_ptr);
        assert_eq!(metadata.column_names()[0].as_ptr(), column_ptr);
        assert!(metadata.has_content());
        assert!(!metadata.contains_default_data());
    }

    #[test]
    fn empty_title_is_content_and_unknown_kind_is_lossless() {
        let metadata = ChartMetadata::from_owned(
            Kind::from_native(i32::MAX),
            Some(String::new()),
            Vec::new(),
            Vec::new(),
            0,
            false,
        );

        assert_eq!(metadata.title(), Some(""));
        assert!(metadata.has_content());
        assert_eq!(metadata.kind().native_value(), i32::MAX);
    }
}
