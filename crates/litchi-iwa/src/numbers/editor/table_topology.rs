//! Read-only detection of table topology that constrains structural edits.

use super::*;

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
