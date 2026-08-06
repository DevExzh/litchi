//! Typed row-height and column-width editing for Numbers tables.

use super::*;
use litchi_numbers::TableSelector;
use litchi_numbers::table::dimension::{Dimension, Points, Size};

mod storage;

use storage::{read_dimension_size, write_dimension_size};

const DEFAULT_DIMENSION_POINTS: f32 = 0.0;

impl NumbersEditor {
    /// Read a row-height or column-width override.
    pub fn table_dimension_size(
        &self,
        selector: TableSelector<'_>,
        dimension: Dimension,
    ) -> Result<Size> {
        let table_id = super::selectors::table_id(self, selector)?;
        read_attached_table_dimension_size(&self.package, table_id, dimension)
    }

    /// Transactionally set or clear a row-height or column-width override.
    pub fn set_table_dimension_size(
        &mut self,
        selector: TableSelector<'_>,
        dimension: Dimension,
        size: Size,
    ) -> Result<()> {
        let table_id = super::selectors::table_id(self, selector)?;
        let mut staged = self.package.clone();
        set_attached_table_dimension_size(&mut staged, table_id, dimension, size)?;

        let verified = NumbersEditor::from_bytes(&staged.to_bytes()?)?;
        if verified.table_dimension_size(selector, dimension)? != size {
            return Err(Error::InvalidFormat(format!(
                "Numbers table {} size failed validation",
                dimension.noun()
            )));
        }
        self.package = staged;
        Ok(())
    }

    /// Read a row-height override by zero-based row index.
    pub fn table_row_height(&self, selector: TableSelector<'_>, row: usize) -> Result<Size> {
        self.table_dimension_size(selector, Dimension::Row(row))
    }

    /// Transactionally set or clear a row-height override.
    pub fn set_table_row_height(
        &mut self,
        selector: TableSelector<'_>,
        row: usize,
        size: Size,
    ) -> Result<()> {
        self.set_table_dimension_size(selector, Dimension::Row(row), size)
    }

    /// Read a column-width override by zero-based column index.
    pub fn table_column_width(&self, selector: TableSelector<'_>, column: usize) -> Result<Size> {
        self.table_dimension_size(selector, Dimension::Column(column))
    }

    /// Transactionally set or clear a column-width override.
    pub fn set_table_column_width(
        &mut self,
        selector: TableSelector<'_>,
        column: usize,
        size: Size,
    ) -> Result<()> {
        self.set_table_dimension_size(selector, Dimension::Column(column), size)
    }
}

fn validate_dimension_index(model: &TableModelArchive, dimension: Dimension) -> Result<()> {
    let limit = match dimension {
        Dimension::Row(_) => model.number_of_rows as usize,
        Dimension::Column(_) => model.number_of_columns as usize,
    };
    if dimension.index() >= limit {
        return Err(Error::ParseError(format!(
            "Numbers table {} {} is outside the table's {limit} {}s",
            dimension.noun(),
            dimension.index(),
            dimension.noun()
        )));
    }
    Ok(())
}

pub(super) fn set_attached_table_dimension_size(
    package: &mut IWorkPackage,
    table_id: u64,
    dimension: Dimension,
    size: Size,
) -> Result<()> {
    let descriptor = attached_table_descriptor(package, table_id)?;
    validate_dimension_index(&descriptor.model, dimension)?;
    let locations = object_locations(package)?;
    write_dimension_size(
        package,
        &locations,
        &descriptor.model,
        dimension,
        match size {
            Size::Default => DEFAULT_DIMENSION_POINTS,
            Size::Points(points) => points.value(),
        },
    )
}

pub(super) fn read_attached_table_dimension_size(
    package: &IWorkPackage,
    table_id: u64,
    dimension: Dimension,
) -> Result<Size> {
    let descriptor = attached_table_descriptor(package, table_id)?;
    validate_dimension_index(&descriptor.model, dimension)?;
    let locations = object_locations(package)?;
    match read_dimension_size(package, &locations, &descriptor.model, dimension)? {
        None => Ok(Size::Default),
        Some(points) => Ok(Size::Points(Points::new(points).map_err(|_| {
            Error::InvalidFormat(format!(
                "Numbers table {} {} has invalid size {points}",
                dimension.noun(),
                dimension.index()
            ))
        })?)),
    }
}

pub(super) fn attached_table_size_points(
    package: &IWorkPackage,
    table_id: u64,
) -> Result<(f32, f32)> {
    let descriptor = attached_table_descriptor(package, table_id)?;
    let locations = object_locations(package)?;
    let mut width = 0.0f64;
    for column in 0..descriptor.model.number_of_columns as usize {
        width += f64::from(effective_dimension_points(
            package,
            &locations,
            &descriptor.model,
            Dimension::Column(column),
        )?);
    }
    let mut height = 0.0f64;
    for row in 0..descriptor.model.number_of_rows as usize {
        height += f64::from(effective_dimension_points(
            package,
            &locations,
            &descriptor.model,
            Dimension::Row(row),
        )?);
    }
    let width = width as f32;
    let height = height as f32;
    if !width.is_finite() || !height.is_finite() || width <= 0.0 || height <= 0.0 {
        return Err(Error::InvalidFormat(format!(
            "Numbers table {table_id} has invalid rendered size {width}x{height}"
        )));
    }
    Ok((width, height))
}

fn effective_dimension_points(
    package: &IWorkPackage,
    locations: &HashMap<u64, String>,
    model: &TableModelArchive,
    dimension: Dimension,
) -> Result<f32> {
    let points =
        read_dimension_size(package, locations, model, dimension)?.unwrap_or(match dimension {
            Dimension::Row(_) => model.default_row_height as f32,
            Dimension::Column(_) => model.default_column_width as f32,
        });
    Points::new(points).map(Points::value).map_err(|_| {
        Error::InvalidFormat(format!(
            "Numbers table {} {} has invalid effective size {points}",
            dimension.noun(),
            dimension.index()
        ))
    })
}
