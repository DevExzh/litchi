//! The document's font table (M1 L-DOC-R4).

use super::prelude::*;
use crate::parts::fonts::FontTable;

impl Document {
    /// The document's font table (`SttbfFfn`), which character properties
    /// index by `ftc` ([`crate::parts::chp::CharacterProperties::font_index`]
    /// and its siblings); `None` when the document has none.
    ///
    /// The table is parsed on first use and kept; a malformed table is
    /// reported here, every time, and never affects opening the document.
    ///
    /// ```no_run
    /// use litchi_doc::Package;
    ///
    /// let package = Package::open("report.doc")?;
    /// let document = package.document();
    /// if let Some(fonts) = document.font_table()? {
    ///     for font in fonts.iter() {
    ///         println!("{} ({:?})", font.name(), font.family());
    ///     }
    /// }
    /// # Ok::<(), Box<dyn std::error::Error>>(())
    /// ```
    ///
    /// # Errors
    ///
    /// Returns [`PackageError::Corrupted`] when the table's FIB range does
    /// not fit the table stream or the table itself is malformed.
    pub fn font_table(&self) -> Result<Option<&FontTable>> {
        match self
            .0
            .font_table
            .get_or_init(|| match &self.0.font_table_source {
                Ok(Some(source)) => FontTable::parse(source)
                    .map(Some)
                    .map_err(|error| error.to_string()),
                Ok(None) => Ok(None),
                Err(error) => Err(error.clone()),
            }) {
            Ok(table) => Ok(table.as_ref()),
            Err(error) => Err(PackageError::Corrupted(format!(
                "invalid font table: {error}"
            ))),
        }
    }

    /// The name of the font at index `ftc` of [`Self::font_table`]: `None`
    /// when the document has no font table or the table has no such font.
    ///
    /// ```no_run
    /// use litchi_doc::Package;
    ///
    /// let package = Package::open("report.doc")?;
    /// let document = package.document();
    /// for paragraph in document.paragraphs()? {
    ///     for run in paragraph.runs()? {
    ///         if let Some(ftc) = run.font_index() {
    ///             println!("{:?}: {:?}", run.text()?, document.font_name(ftc)?);
    ///         }
    ///     }
    /// }
    /// # Ok::<(), Box<dyn std::error::Error>>(())
    /// ```
    ///
    /// # Errors
    ///
    /// Returns the errors of [`Self::font_table`].
    pub fn font_name(&self, ftc: u16) -> Result<Option<&str>> {
        Ok(self.font_table()?.and_then(|fonts| fonts.name(ftc)))
    }
}
