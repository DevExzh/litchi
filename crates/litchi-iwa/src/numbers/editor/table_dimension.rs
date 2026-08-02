//! Typed row-height and column-width editing for Numbers tables.

use super::*;

mod storage;

use storage::{read_dimension_size, write_dimension_size};

const DEFAULT_DIMENSION_POINTS: f32 = 0.0;

/// One physical table dimension addressed by zero-based index.
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub enum NumbersTableDimension {
    Row(usize),
    Column(usize),
}

impl NumbersTableDimension {
    fn index(self) -> usize {
        match self {
            Self::Row(index) | Self::Column(index) => index,
        }
    }

    fn noun(self) -> &'static str {
        match self {
            Self::Row(_) => "row",
            Self::Column(_) => "column",
        }
    }
}

/// A validated positive, finite table dimension measured in points.
#[derive(Clone, Copy, Debug, PartialEq, PartialOrd)]
pub struct NumbersTablePoints(f32);

impl NumbersTablePoints {
    /// Validate and construct a point measurement.
    pub fn new(points: f32) -> Result<Self> {
        if !points.is_finite() || points <= 0.0 {
            return Err(Error::ParseError(
                "Numbers table dimension points must be positive and finite".to_owned(),
            ));
        }
        Ok(Self(points))
    }

    /// Return the point measurement.
    pub fn get(self) -> f32 {
        self.0
    }
}

impl TryFrom<f32> for NumbersTablePoints {
    type Error = Error;

    fn try_from(points: f32) -> Result<Self> {
        Self::new(points)
    }
}

impl From<NumbersTablePoints> for f32 {
    fn from(points: NumbersTablePoints) -> Self {
        points.get()
    }
}

/// Either the table style's default size or an explicit point override.
#[derive(Clone, Copy, Debug, Default, PartialEq)]
pub enum NumbersTableDimensionSize {
    #[default]
    Default,
    Points(NumbersTablePoints),
}

impl NumbersTableDimensionSize {
    /// Validate an explicit point override.
    pub fn points(points: f32) -> Result<Self> {
        Ok(Self::Points(NumbersTablePoints::new(points)?))
    }

    fn stored_points(self) -> f32 {
        match self {
            Self::Default => DEFAULT_DIMENSION_POINTS,
            Self::Points(points) => points.get(),
        }
    }
}

impl NumbersEditor {
    /// Read a row-height or column-width override.
    pub fn table_dimension_size(
        &self,
        table_id: u64,
        dimension: NumbersTableDimension,
    ) -> Result<NumbersTableDimensionSize> {
        read_attached_table_dimension_size(&self.package, table_id, dimension)
    }

    /// Transactionally set or clear a row-height or column-width override.
    pub fn set_table_dimension_size(
        &mut self,
        table_id: u64,
        dimension: NumbersTableDimension,
        size: NumbersTableDimensionSize,
    ) -> Result<()> {
        let mut staged = self.package.clone();
        set_attached_table_dimension_size(&mut staged, table_id, dimension, size)?;

        let verified = NumbersEditor::from_bytes(&staged.to_bytes()?)?;
        if verified.table_dimension_size(table_id, dimension)? != size {
            return Err(Error::InvalidFormat(format!(
                "Numbers table {} size failed validation",
                dimension.noun()
            )));
        }
        self.package = staged;
        Ok(())
    }

    /// Read a row-height override by zero-based row index.
    pub fn table_row_height(&self, table_id: u64, row: usize) -> Result<NumbersTableDimensionSize> {
        self.table_dimension_size(table_id, NumbersTableDimension::Row(row))
    }

    /// Transactionally set or clear a row-height override.
    pub fn set_table_row_height(
        &mut self,
        table_id: u64,
        row: usize,
        size: NumbersTableDimensionSize,
    ) -> Result<()> {
        self.set_table_dimension_size(table_id, NumbersTableDimension::Row(row), size)
    }

    /// Read a column-width override by zero-based column index.
    pub fn table_column_width(
        &self,
        table_id: u64,
        column: usize,
    ) -> Result<NumbersTableDimensionSize> {
        self.table_dimension_size(table_id, NumbersTableDimension::Column(column))
    }

    /// Transactionally set or clear a column-width override.
    pub fn set_table_column_width(
        &mut self,
        table_id: u64,
        column: usize,
        size: NumbersTableDimensionSize,
    ) -> Result<()> {
        self.set_table_dimension_size(table_id, NumbersTableDimension::Column(column), size)
    }
}

fn validate_dimension_index(
    model: &TableModelArchive,
    dimension: NumbersTableDimension,
) -> Result<()> {
    let limit = match dimension {
        NumbersTableDimension::Row(_) => model.number_of_rows as usize,
        NumbersTableDimension::Column(_) => model.number_of_columns as usize,
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
    dimension: NumbersTableDimension,
    size: NumbersTableDimensionSize,
) -> Result<()> {
    let descriptor = attached_table_descriptor(package, table_id)?;
    validate_dimension_index(&descriptor.model, dimension)?;
    let locations = object_locations(package)?;
    write_dimension_size(
        package,
        &locations,
        &descriptor.model,
        dimension,
        size.stored_points(),
    )
}

pub(super) fn read_attached_table_dimension_size(
    package: &IWorkPackage,
    table_id: u64,
    dimension: NumbersTableDimension,
) -> Result<NumbersTableDimensionSize> {
    let descriptor = attached_table_descriptor(package, table_id)?;
    validate_dimension_index(&descriptor.model, dimension)?;
    let locations = object_locations(package)?;
    match read_dimension_size(package, &locations, &descriptor.model, dimension)? {
        None => Ok(NumbersTableDimensionSize::Default),
        Some(points) => Ok(NumbersTableDimensionSize::Points(
            NumbersTablePoints::new(points).map_err(|_| {
                Error::InvalidFormat(format!(
                    "Numbers table {} {} has invalid size {points}",
                    dimension.noun(),
                    dimension.index()
                ))
            })?,
        )),
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
            NumbersTableDimension::Column(column),
        )?);
    }
    let mut height = 0.0f64;
    for row in 0..descriptor.model.number_of_rows as usize {
        height += f64::from(effective_dimension_points(
            package,
            &locations,
            &descriptor.model,
            NumbersTableDimension::Row(row),
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
    dimension: NumbersTableDimension,
) -> Result<f32> {
    let points =
        read_dimension_size(package, locations, model, dimension)?.unwrap_or(match dimension {
            NumbersTableDimension::Row(_) => model.default_row_height as f32,
            NumbersTableDimension::Column(_) => model.default_column_width as f32,
        });
    NumbersTablePoints::new(points)
        .map(NumbersTablePoints::get)
        .map_err(|_| {
            Error::InvalidFormat(format!(
                "Numbers table {} {} has invalid effective size {points}",
                dimension.noun(),
                dimension.index()
            ))
        })
}
