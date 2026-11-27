//! Semantic invariants for the typed data-pilot model.

use std::collections::HashSet;

use crate::model::database_range::{self, validate_filter};
use litchi_core::{Error, Result};

use super::{
    MAX_DATA_PILOT_FIELDS, MAX_DATA_PILOT_ITEMS, MAX_DATA_PILOT_STRING, MAX_DATA_PILOT_TABLES,
    invalid_message,
    model::{Field, Groups, ReferenceMemberType, SortMode, Source, Table},
    range::{parse_data_pilot_range, ranges_overlap},
};

impl Field {
    ///
    /// # Errors
    /// Returns an error when the operation cannot be completed.
    pub fn validate(&self) -> Result<()> {
        if self.orientation == super::model::Orientation::Page && self.selected_page.is_none() {
            return Err(Error::InvalidFormat(
                "page-oriented data-pilot field requires table:selected-page".to_string(),
            ));
        }
        if self.orientation != super::model::Orientation::Page && self.selected_page.is_some() {
            return Err(Error::InvalidFormat(
                "table:selected-page is valid only for a page-oriented data-pilot field"
                    .to_string(),
            ));
        }
        if let Some(reference) = &self.reference {
            let named = reference.member_type == ReferenceMemberType::Named;
            if named != reference.member_name.is_some() {
                return Err(Error::InvalidFormat(
                    "named data-pilot field references require exactly one member name".to_string(),
                ));
            }
        }
        if let Some(level) = &self.level
            && let Some(sort) = &level.sort
            && (sort.mode == SortMode::Data) != sort.data_field.is_some()
        {
            return Err(Error::InvalidFormat(
                "data-pilot data sorting requires exactly one data field".to_string(),
            ));
        }
        if let Some(groups) = &self.groups {
            if !groups.step.is_finite() || groups.step <= 0.0 {
                return Err(Error::InvalidFormat(
                    "data-pilot grouping step must be finite and greater than zero".to_string(),
                ));
            }
            for boundary in [&groups.start, &groups.end] {
                if matches!(boundary, super::model::GroupBoundary::Number(value) if !value.is_finite())
                {
                    return Err(Error::InvalidFormat(
                        "data-pilot grouping boundaries must be finite".to_string(),
                    ));
                }
            }
            if groups.groups.is_empty()
                || groups.groups.iter().any(|group| group.members.is_empty())
            {
                return Err(Error::InvalidFormat(
                    "data-pilot groups and group member lists cannot be empty".to_string(),
                ));
            }
        }
        Ok(())
    }
}

impl Table {
    ///
    /// # Errors
    /// Returns an error when the operation cannot be completed.
    pub fn validate(&self) -> Result<()> {
        validate_string("data-pilot table name", &self.name, false)?;
        if self.target_range_address.is_empty() {
            return Err(Error::InvalidFormat(
                "data-pilot target range address cannot be empty".to_string(),
            ));
        }
        parse_data_pilot_range(&self.target_range_address)?;
        for value in [self.application_data.as_deref(), self.buttons.as_deref()]
            .into_iter()
            .flatten()
        {
            validate_string("data-pilot attribute", value, true)?;
        }
        let mut grand_orientations = HashSet::new();
        for total in &self.grand_totals {
            if !grand_orientations.insert(total.orientation) {
                return Err(invalid_message(
                    "duplicate data-pilot grand-total orientation",
                ));
            }
            if let Some(name) = &total.display_name {
                validate_string("data-pilot grand-total display name", name, true)?;
            }
        }
        if self.fields.is_empty() {
            return Err(Error::InvalidFormat(
                "data-pilot table requires at least one field".to_string(),
            ));
        }
        if self.fields.len() > MAX_DATA_PILOT_FIELDS {
            return Err(Error::InvalidFormat(format!(
                "data-pilot field count exceeds the {MAX_DATA_PILOT_FIELDS} field safety limit"
            )));
        }
        if let Some(Source::CellRange {
            name,
            cell_range_address,
            filter,
        }) = &self.source
        {
            if let Some(name) = name {
                validate_string("data-pilot named source", name, false)?;
            }
            validate_string(
                "data-pilot source cell range address",
                cell_range_address,
                false,
            )?;
            parse_data_pilot_range(cell_range_address)?;
            if let Some(filter) = filter {
                validate_filter(filter)?;
            }
        }
        if let Some(Source::Service {
            name,
            source_name,
            object_name,
            user_name,
            password,
        }) = &self.source
        {
            for value in [name.as_str(), source_name.as_str(), object_name.as_str()] {
                validate_string("data-pilot service source", value, false)?;
            }
            for value in [user_name.as_deref(), password.as_deref()]
                .into_iter()
                .flatten()
            {
                validate_string("data-pilot service source", value, true)?;
            }
        }
        if let Some(Source::Database(source)) = &self.source {
            match source {
                database_range::Source::Sql {
                    database_name,
                    statement,
                    ..
                } => {
                    validate_string("data-pilot database name", database_name, false)?;
                    validate_string("data-pilot SQL statement", statement, false)?;
                },
                database_range::Source::Table {
                    database_name,
                    table_name,
                } => {
                    validate_string("data-pilot database name", database_name, false)?;
                    validate_string("data-pilot database table", table_name, false)?;
                },
                database_range::Source::Query {
                    database_name,
                    query_name,
                } => {
                    validate_string("data-pilot database name", database_name, false)?;
                    validate_string("data-pilot database query", query_name, false)?;
                },
            }
        }
        self.fields.iter().try_for_each(Field::validate)?;
        let field_names: HashSet<&str> = self
            .fields
            .iter()
            .map(|field| field.source_field_name.as_str())
            .collect();
        let mut item_count = self.fields.len();
        for field in &self.fields {
            validate_string(
                "data-pilot source field name",
                &field.source_field_name,
                true,
            )?;
            if let Some(reference) = &field.reference
                && !field_names.contains(reference.field_name.as_str())
            {
                return Err(Error::InvalidFormat(format!(
                    "data-pilot field reference '{}' does not name a field",
                    reference.field_name
                )));
            }
            if let Some(groups) = &field.groups {
                if !field_names.contains(groups.source_field_name.as_str()) {
                    return Err(Error::InvalidFormat(format!(
                        "data-pilot grouping source '{}' does not name a field",
                        groups.source_field_name
                    )));
                }
                validate_groups(groups, &mut item_count)?;
            }
            if let Some(level) = &field.level {
                let mut members = HashSet::new();
                for member in &level.members {
                    validate_string("data-pilot member", &member.name, false)?;
                    if !members.insert(member.name.as_str()) {
                        return Err(Error::InvalidFormat(format!(
                            "duplicate data-pilot member '{}'",
                            member.name
                        )));
                    }
                }
                item_count = item_count
                    .checked_add(level.members.len() + level.subtotals.len())
                    .ok_or_else(|| invalid_message("data-pilot item count overflow"))?;
            }
        }
        if item_count > MAX_DATA_PILOT_ITEMS {
            return Err(invalid_message("data-pilot declaration exceeds item limit"));
        }
        Ok(())
    }
}

fn validate_groups(groups: &Groups, item_count: &mut usize) -> Result<()> {
    let mut names = HashSet::new();
    for group in &groups.groups {
        validate_string("data-pilot group name", &group.name, false)?;
        if !names.insert(group.name.as_str()) {
            return Err(Error::InvalidFormat(format!(
                "duplicate data-pilot group '{}'",
                group.name
            )));
        }
        let mut members = HashSet::new();
        for member in &group.members {
            validate_string("data-pilot group member", member, false)?;
            if !members.insert(member.as_str()) {
                return Err(Error::InvalidFormat(format!(
                    "duplicate member '{member}' in data-pilot group '{}'",
                    group.name
                )));
            }
        }
        *item_count = item_count
            .checked_add(group.members.len() + 1)
            .ok_or_else(|| invalid_message("data-pilot item count overflow"))?;
    }
    Ok(())
}

pub(super) fn validate_string(label: &str, value: &str, allow_empty: bool) -> Result<()> {
    if !allow_empty && value.is_empty() {
        return Err(Error::InvalidFormat(format!("{label} cannot be empty")));
    }
    if value.len() > MAX_DATA_PILOT_STRING {
        return Err(Error::InvalidFormat(format!("{label} exceeds size limit")));
    }
    if value
        .chars()
        .any(|ch| matches!(u32::from(ch), 0..=8 | 11 | 12 | 14..=31 | 0xFFFE | 0xFFFF))
    {
        return Err(Error::InvalidFormat(format!(
            "{label} contains an XML-prohibited character"
        )));
    }
    Ok(())
}

pub(crate) fn validate_data_pilot_tables(tables: &[Table]) -> Result<()> {
    if tables.len() > MAX_DATA_PILOT_TABLES {
        return Err(invalid_message(
            "data-pilot table count exceeds safety limit",
        ));
    }
    let mut names = HashSet::new();
    let mut targets = Vec::with_capacity(tables.len());
    for table in tables {
        table.validate()?;
        if !names.insert(table.name.as_str()) {
            return Err(Error::InvalidFormat(format!(
                "duplicate data-pilot table name '{}'",
                table.name
            )));
        }
        let range = parse_data_pilot_range(&table.target_range_address)?;
        for (other_name, other) in &targets {
            if ranges_overlap(&range, other) {
                return Err(Error::InvalidFormat(format!(
                    "data-pilot target ranges for '{other_name}' and '{}' overlap",
                    table.name
                )));
            }
        }
        targets.push((table.name.as_str(), range));
    }
    Ok(())
}
