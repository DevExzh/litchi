//! Read-only detection of table topology that constrains structural edits.

use super::*;

const FORMULA_OWNER_DEPENDENCIES_MESSAGE_TYPE: u32 = 4_008;

pub(super) fn filter_has_row_state(
    package: &IWorkPackage,
    locations: &HashMap<u64, String>,
    reference: Option<&tsp::Reference>,
) -> Result<bool> {
    let Some(reference) = reference else {
        return Ok(false);
    };
    let archive_name = locations.get(&reference.identifier).ok_or_else(|| {
        Error::InvalidFormat(format!(
            "Numbers filter object {} is missing",
            reference.identifier
        ))
    })?;
    let archive = package.archive(archive_name)?;
    let object = archive.object(reference.identifier).ok_or_else(|| {
        Error::InvalidFormat(format!(
            "Numbers filter object {} is missing",
            reference.identifier
        ))
    })?;
    let filter = object
        .messages
        .iter()
        .find_map(|message| tst::FilterSetArchive::decode(message.data.as_slice()).ok())
        .ok_or_else(|| {
            Error::InvalidFormat(format!(
                "Object {} has no Numbers filter-set payload",
                reference.identifier
            ))
        })?;
    Ok(filter.is_enabled.unwrap_or(true)
        && (!filter.filter_rules_prepivot.is_empty()
            || !filter.filter_rules.is_empty()
            || filter.filter_enabled.iter().any(|enabled| *enabled)))
}

pub(super) fn category_grouping_is_enabled(
    package: &IWorkPackage,
    locations: &HashMap<u64, String>,
    reference: Option<&tsp::Reference>,
) -> Result<bool> {
    let Some(reference) = reference else {
        return Ok(false);
    };
    let archive_name = locations.get(&reference.identifier).ok_or_else(|| {
        Error::InvalidFormat(format!(
            "Numbers category owner {} is missing",
            reference.identifier
        ))
    })?;
    let archive = package.archive(archive_name)?;
    let object = archive.object(reference.identifier).ok_or_else(|| {
        Error::InvalidFormat(format!(
            "Numbers category owner {} is missing",
            reference.identifier
        ))
    })?;
    let references = object
        .messages
        .iter()
        .find_map(|message| tst::CategoryOwnerRefArchive::decode(message.data.as_slice()).ok())
        .map(|owner| owner.group_by)
        .unwrap_or_default();
    for group in references {
        let archive_name = locations.get(&group.identifier).ok_or_else(|| {
            Error::InvalidFormat(format!(
                "Numbers group-by object {} is missing",
                group.identifier
            ))
        })?;
        let archive = package.archive(archive_name)?;
        let object = archive.object(group.identifier).ok_or_else(|| {
            Error::InvalidFormat(format!(
                "Numbers group-by object {} is missing",
                group.identifier
            ))
        })?;
        let enabled = object.messages.iter().any(|message| {
            tst::GroupByArchive::decode(message.data.as_slice()).is_ok_and(|group| group.is_enabled)
        });
        if enabled {
            return Ok(true);
        }
    }
    Ok(false)
}

pub(super) fn deprecated_category_grouping_is_enabled(
    owner: Option<&tst::CategoryOwnerArchive>,
) -> bool {
    owner.is_some_and(|owner| owner.group_by.iter().any(|group| group.is_enabled))
}

pub(super) fn table_has_spill_state(package: &IWorkPackage, table_info_id: u64) -> Result<bool> {
    let Some(component) = package.calculation_engine_entry_name()? else {
        return Ok(false);
    };
    let archive = package.archive(component)?;
    Ok(archive
        .objects
        .iter()
        .flat_map(|object| &object.messages)
        .any(|message| {
            message.type_ == FORMULA_OWNER_DEPENDENCIES_MESSAGE_TYPE
                && tsce::FormulaOwnerDependenciesArchive::decode(message.data.as_slice()).is_ok_and(
                    |owner| {
                        owner
                            .formula_owner
                            .is_some_and(|reference| reference.identifier == table_info_id)
                            && owner
                                .spill_range_sizes
                                .is_some_and(|spills| !spills.spills.is_empty())
                    },
                )
        }))
}
