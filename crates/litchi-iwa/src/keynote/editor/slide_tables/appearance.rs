//! Copy-on-write appearance CRUD for Keynote slide tables.

use super::*;
use crate::table_appearance::TableAppearance;
use litchi_core::Position;
use litchi_keynote::TableSelector;

impl KeynoteEditor {
    /// Read the effective alternating-row and automatic-sizing settings.
    ///
    /// The raw model identifier is retained for the migration-host boundary;
    /// it is resolved to the checked slide/table position before entering the
    /// focused Keynote package API. Legacy native decoding is deliberately not
    /// used as a fallback for malformed, locked, or dependency-bearing input.
    #[deprecated(
        since = "0.0.1",
        note = "use litchi_keynote::Package::slide_table_appearance with SlideSelector and TableSelector"
    )]
    pub fn slide_table_appearance(
        &self,
        slide_index: usize,
        model_object_id: u64,
    ) -> Result<TableAppearance> {
        let (slide, table) = focused_table_position(self, slide_index, model_object_id)?;
        focused_table_appearance_package(self)?
            .slide_table_appearance(slide, table)
            .map_err(map_focused_table_appearance_error)
    }

    /// Replace appearance settings without mutating styles shared by other tables.
    #[deprecated(
        since = "0.0.1",
        note = "use litchi_keynote::Package::edit_slide_table_appearance with SlideSelector and TableSelector"
    )]
    pub fn set_slide_table_appearance(
        &mut self,
        slide_index: usize,
        model_object_id: u64,
        appearance: TableAppearance,
    ) -> Result<()> {
        let (slide, table) = focused_table_position(self, slide_index, model_object_id)?;
        let package = focused_table_appearance_package(self)?;
        let commit = package
            .edit_slide_table_appearance(slide, table)
            .map_err(map_focused_table_appearance_error)?
            .set(appearance)
            .commit()
            .map_err(map_focused_table_appearance_error)?;
        if commit.patch().is_noop() {
            return Ok(());
        }
        replace_from_focused_table_appearance_commit(self, commit.package())?;
        let (verified_slide, verified_table) =
            focused_table_position(self, slide_index, model_object_id)?;
        let verified = focused_table_appearance_package(self)?
            .slide_table_appearance(verified_slide, verified_table)
            .map_err(map_focused_table_appearance_error)?;
        if verified != appearance {
            return Err(Error::InvalidFormat(
                "Keynote table appearance failed round-trip validation".to_owned(),
            ));
        }
        Ok(())
    }
}

fn focused_table_position(
    editor: &KeynoteEditor,
    slide_index: usize,
    model_object_id: u64,
) -> Result<(Position, TableSelector)> {
    let tables = editor.slide_tables(slide_index)?;
    let mut table_index = None;
    for (index, table) in tables.iter().enumerate() {
        if table.model_object_id != model_object_id {
            continue;
        }
        if table_index.replace(index).is_some() {
            return Err(Error::InvalidFormat(format!(
                "Keynote object {model_object_id} has ambiguous table ownership on slide {slide_index}"
            )));
        }
    }
    let table_index = table_index.ok_or_else(|| {
        Error::InvalidFormat(format!(
            "Keynote object {model_object_id} has no table on slide {slide_index}"
        ))
    })?;
    Ok((
        Position::new(slide_index),
        TableSelector::index(table_index),
    ))
}

fn focused_table_appearance_package(editor: &KeynoteEditor) -> Result<litchi_keynote::Package> {
    let bytes = editor.to_bytes()?;
    litchi_keynote::Package::from_bytes(&bytes).map_err(|error| {
        Error::InvalidFormat(format!(
            "focused Keynote table appearance source failed: {error}"
        ))
    })
}

fn replace_from_focused_table_appearance_commit(
    editor: &mut KeynoteEditor,
    package: &litchi_keynote::Package,
) -> Result<()> {
    let mut bytes = Vec::new();
    package.write_to(&mut bytes).map_err(|error| {
        Error::InvalidFormat(format!(
            "focused Keynote table appearance write failed: {error}"
        ))
    })?;
    *editor = KeynoteEditor::from_bytes(&bytes)?;
    Ok(())
}

fn map_focused_table_appearance_error<E: std::fmt::Display>(error: E) -> Error {
    Error::InvalidFormat(format!(
        "focused Keynote table appearance operation failed: {error}"
    ))
}

#[cfg(test)]
#[allow(deprecated)]
mod tests {
    use super::*;
    use crate::keynote::KeynoteDocumentBuilder;
    use crate::table_appearance::{
        TableGridlineVisibility, TableGridlines, TableRowBanding, TableRowSizing,
    };

    #[test]
    fn scratch_table_appearance_is_copy_on_write() {
        let mut editor = KeynoteDocumentBuilder::new().build().unwrap();
        let source = editor
            .add_slide_table(
                0,
                "Source",
                3,
                2,
                DrawablePoint { x: 40.0, y: 40.0 },
                DrawableSize {
                    width: 320.0,
                    height: 180.0,
                },
            )
            .unwrap();
        let duplicate = editor
            .duplicate_slide_table(0, source.drawable_object_id)
            .unwrap();
        let appearance = TableAppearance {
            row_banding: TableRowBanding::Enabled,
            row_sizing: TableRowSizing::FitCellContents,
            gridlines: TableGridlines {
                body_horizontal: TableGridlineVisibility::Hidden,
                header_columns_horizontal: TableGridlineVisibility::Visible,
                body_vertical: TableGridlineVisibility::Hidden,
                header_rows_vertical: TableGridlineVisibility::Visible,
                footer_rows_vertical: TableGridlineVisibility::Hidden,
            },
        };

        editor
            .set_slide_table_appearance(0, duplicate.model_object_id, appearance)
            .unwrap();

        assert_eq!(
            editor
                .slide_table_appearance(0, source.model_object_id)
                .unwrap(),
            TableAppearance::default()
        );
        assert_eq!(
            editor
                .slide_table_appearance(0, duplicate.model_object_id)
                .unwrap(),
            appearance
        );
        assert_eq!(
            editor
                .slide_tables(0)
                .unwrap()
                .into_iter()
                .find(|table| table.model_object_id == duplicate.model_object_id)
                .unwrap()
                .appearance,
            appearance
        );
    }
}
