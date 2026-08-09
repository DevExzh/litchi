use super::super::comment;
use super::super::{CommentWriteOptions, ShapeGroupWrite, ShapeWrite, WritableWorksheet, Writer};
use crate::error::{Error, Result};
use std::collections::HashSet;

impl Writer {
    /// Add a canonical, macro-inert BIFF8 comment to a cell.
    /// # Errors
    ///
    /// Returns an error if validation, decoding, encoding, or the requested operation fails.
    pub fn add_comment(
        &mut self,
        sheet: usize,
        row: u32,
        col: u16,
        author: &str,
        text: &str,
    ) -> Result<()> {
        self.add_comment_with_options(
            sheet,
            row,
            col,
            author,
            text,
            CommentWriteOptions::default(),
        )
    }

    /// Add a canonical BIFF8 comment with explicit visibility, anchor, rich runs, and GUID options.
    /// # Errors
    ///
    /// Returns an error if validation, decoding, encoding, or the requested operation fails.
    pub fn add_comment_with_options(
        &mut self,
        sheet: usize,
        row: u32,
        col: u16,
        author: &str,
        text: &str,
        options: CommentWriteOptions,
    ) -> Result<()> {
        let (row, column) = comment::validate_comment(row, col, author, text, &options)?;
        let worksheet = self
            .worksheets
            .get_mut(sheet)
            .ok_or_else(|| Error::WorksheetNotFound(format!("Sheet {sheet}")))?;
        if worksheet.comments.len() >= 1022 {
            return Err(Error::InvalidData(
                "a worksheet cannot contain more than 1022 canonical comment shapes".to_string(),
            ));
        }
        if worksheet
            .comments
            .iter()
            .any(|comment| comment.row == row && comment.column == column)
        {
            return Err(Error::InvalidData(
                "a cell cannot contain more than one comment".to_string(),
            ));
        }
        if let Some(guid) = options.guid
            && worksheet
                .comments
                .iter()
                .any(|comment| comment.options.guid == Some(guid))
        {
            return Err(Error::InvalidData(
                "comment GUID override is duplicated on the worksheet".to_string(),
            ));
        }
        let comment = comment::WritableComment::try_new(row, column, author, text, options)?;
        worksheet.add_comment(comment)
    }

    /// Add a validated, macro-inert primitive shape and return its worksheet OBJ identifier.
    /// # Errors
    ///
    /// Returns an error if validation, decoding, encoding, or the requested operation fails.
    pub fn add_shape(&mut self, sheet: usize, mut shape: ShapeWrite) -> Result<u16> {
        shape.validate()?;
        let worksheet = self
            .worksheets
            .get_mut(sheet)
            .ok_or_else(|| Error::WorksheetNotFound(format!("Sheet {sheet}")))?;
        let reserved = collect_reserved_object_ids(worksheet, 0)?;
        let object_count =
            reserved
                .len()
                .checked_add(worksheet.comments.len())
                .ok_or(Error::Allocation(
                    "computing the worksheet drawing-object count",
                ))?;
        if object_count >= 1022 {
            return Err(Error::InvalidData(
                "a worksheet cannot contain more than 1022 drawing objects".to_string(),
            ));
        }
        let object_id = if let Some(requested) = shape.object_id {
            if reserved.contains(&requested) {
                return Err(Error::InvalidData(
                    "shape object ID collides with another worksheet object".to_string(),
                ));
            }
            requested
        } else {
            (1..u16::MAX)
                .find(|candidate| !reserved.contains(candidate))
                .ok_or_else(|| {
                    Error::InvalidData("worksheet object IDs are exhausted".to_string())
                })?
        };
        worksheet
            .shapes
            .try_reserve(1)
            .map_err(|_error| Error::Allocation("reserving worksheet shape storage"))?;
        shape.object_id = Some(object_id);
        worksheet.shapes.push(shape);
        Ok(object_id)
    }

    /// Remove a primitive by its assigned OBJ identifier.
    /// # Errors
    ///
    /// Returns an error if validation, decoding, encoding, or the requested operation fails.
    pub fn remove_shape(&mut self, sheet: usize, object_id: u16) -> Result<ShapeWrite> {
        if object_id == 0 || object_id == u16::MAX {
            return Err(Error::InvalidData(
                "shape object ID 0 and 65535 are reserved".to_string(),
            ));
        }
        let worksheet = self
            .worksheets
            .get_mut(sheet)
            .ok_or_else(|| Error::WorksheetNotFound(format!("Sheet {sheet}")))?;
        let index = worksheet
            .shapes
            .iter()
            .position(|shape| shape.object_id == Some(object_id))
            .ok_or_else(|| Error::InvalidData("shape object ID was not found".to_string()))?;
        Ok(worksheet.shapes.remove(index))
    }

    /// Remove all writable primitive shapes from a worksheet.
    /// # Errors
    ///
    /// Returns an error if validation, decoding, encoding, or the requested operation fails.
    pub fn clear_shapes(&mut self, sheet: usize) -> Result<usize> {
        let worksheet = self
            .worksheets
            .get_mut(sheet)
            .ok_or_else(|| Error::WorksheetNotFound(format!("Sheet {sheet}")))?;
        let count = worksheet.shapes.len();
        worksheet.shapes.clear();
        Ok(count)
    }

    /// Add a validated shape group and return the group's worksheet OBJ identifier.
    ///
    /// The group consumes one object ID for itself plus one per child; assigned
    /// child identifiers are stored back into the group before it is retained.
    /// # Errors
    ///
    /// Returns an error if validation, decoding, encoding, or the requested operation fails.
    pub fn add_shape_group(&mut self, sheet: usize, mut group: ShapeGroupWrite) -> Result<u16> {
        group.validate()?;
        let worksheet = self
            .worksheets
            .get_mut(sheet)
            .ok_or_else(|| Error::WorksheetNotFound(format!("Sheet {sheet}")))?;
        let group_count = group.object_count()?;
        let mut reserved = collect_reserved_object_ids(worksheet, group_count)?;
        let object_count = reserved
            .len()
            .checked_add(worksheet.comments.len())
            .and_then(|count| count.checked_add(group_count))
            .ok_or(Error::Allocation(
                "computing the worksheet drawing-object count",
            ))?;
        if object_count > 1022 {
            return Err(Error::InvalidData(
                "a worksheet cannot contain more than 1022 drawing objects".to_string(),
            ));
        }
        for requested in group_object_ids(&group) {
            if reserved.contains(&requested) {
                return Err(Error::InvalidData(
                    "shape object ID collides with another worksheet object".to_string(),
                ));
            }
        }
        for requested in group_object_ids(&group) {
            reserved.insert(requested);
        }
        worksheet
            .shape_groups
            .try_reserve(1)
            .map_err(|_error| Error::Allocation("reserving worksheet shape-group storage"))?;
        let group_id = assign_object_id(&mut reserved, group.object_id)?;
        group.object_id = Some(group_id);
        for child in &mut group.children {
            child.object_id = Some(assign_object_id(&mut reserved, child.object_id)?);
        }
        worksheet.shape_groups.push(group);
        Ok(group_id)
    }

    /// Remove a shape group by the group's assigned OBJ identifier.
    /// # Errors
    ///
    /// Returns an error if validation, decoding, encoding, or the requested operation fails.
    pub fn remove_shape_group(&mut self, sheet: usize, object_id: u16) -> Result<ShapeGroupWrite> {
        if object_id == 0 || object_id == u16::MAX {
            return Err(Error::InvalidData(
                "shape object ID 0 and 65535 are reserved".to_string(),
            ));
        }
        let worksheet = self
            .worksheets
            .get_mut(sheet)
            .ok_or_else(|| Error::WorksheetNotFound(format!("Sheet {sheet}")))?;
        let index = worksheet
            .shape_groups
            .iter()
            .position(|group| group.object_id == Some(object_id))
            .ok_or_else(|| Error::InvalidData("shape object ID was not found".to_string()))?;
        Ok(worksheet.shape_groups.remove(index))
    }
}

/// Iterate every OBJ identifier requested or assigned inside a shape group.
fn group_object_ids(group: &ShapeGroupWrite) -> impl Iterator<Item = u16> + '_ {
    group
        .object_id
        .into_iter()
        .chain(group.children.iter().filter_map(|child| child.object_id))
}

/// Collect every existing worksheet drawing-object identifier with fallible
/// capacity reservation for any IDs the caller will add before retaining data.
fn collect_reserved_object_ids(
    worksheet: &WritableWorksheet,
    additional: usize,
) -> Result<HashSet<u16>> {
    let pivot_capacity = worksheet
        .pivot_tables
        .iter()
        .try_fold(0usize, |count, table| {
            count
                .checked_add(table.page_entries.len())
                .ok_or(Error::Allocation(
                    "computing worksheet pivot-object ID capacity",
                ))
        })?;
    let group_capacity = worksheet
        .shape_groups
        .iter()
        .try_fold(0usize, |count, group| {
            count
                .checked_add(group.object_count()?)
                .ok_or(Error::Allocation(
                    "computing worksheet shape-group ID capacity",
                ))
        })?;
    let capacity = pivot_capacity
        .checked_add(worksheet.shapes.len())
        .and_then(|count| count.checked_add(group_capacity))
        .and_then(|count| count.checked_add(additional))
        .ok_or(Error::Allocation(
            "computing worksheet drawing-object ID capacity",
        ))?;
    let mut reserved = HashSet::new();
    reserved
        .try_reserve(capacity)
        .map_err(|_error| Error::Allocation("reserving worksheet drawing-object ID storage"))?;
    reserved.extend(
        worksheet
            .pivot_tables
            .iter()
            .flat_map(|table| table.page_entries.iter().map(|entry| entry.2))
            .filter(|id| *id != 0 && *id != u16::MAX),
    );
    reserved.extend(worksheet.shapes.iter().filter_map(|shape| shape.object_id));
    reserved.extend(worksheet.shape_groups.iter().flat_map(group_object_ids));
    Ok(reserved)
}

/// Reserve the requested OBJ identifier or the first free canonical one.
fn assign_object_id(reserved: &mut HashSet<u16>, requested: Option<u16>) -> Result<u16> {
    let object_id = match requested {
        Some(object_id) => object_id,
        None => (1..u16::MAX)
            .find(|candidate| !reserved.contains(candidate))
            .ok_or_else(|| Error::InvalidData("worksheet object IDs are exhausted".to_string()))?,
    };
    if requested.is_none() {
        reserved.insert(object_id);
    }
    Ok(object_id)
}
