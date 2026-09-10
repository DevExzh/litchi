//! Source-bound proof adapters for workbook Data Model metadata.
//!
//! Workbook time-grouping records point at columns that are materialized by
//! the opaque XLDM part.  This adapter composes the neutral XLDM section
//! inspectors and admits a record only after the complete version-140
//! identity/dependency closure has been proven.

use litchi_xldm::{
    OlapProofLimits, Xldm140TimeGroupingBinding, Xldm140TimeGroupingContentType,
    prove_xldm140_closure, prove_xldm140_olap,
};

use super::model::TimeGrouping;
use crate::package::error::{Error, Result};

pub(crate) fn prove_time_grouping(
    payload: &[u8],
    grouping: &TimeGrouping,
) -> Result<Xldm140TimeGroupingBinding> {
    let storage = litchi_xldm::inspect(payload)
        .map_err(|error| invalid(format!("XLDM outer proof failed: {error}")))?;
    let metadata = litchi_xldm::metadata::inspect(&storage)
        .map_err(|error| invalid(format!("XLDM metadata proof failed: {error}")))?;
    let native = litchi_xldm::native::inspect(&storage, &metadata.native_parse_options())
        .map_err(|error| invalid(format!("XLDM native-data proof failed: {error}")))?;
    let generated = litchi_xldm::generated::inspect_system_generated(&storage)
        .map_err(|error| invalid(format!("XLDM generated-data proof failed: {error}")))?;
    let olap = litchi_xldm::olap::inspect(&storage, &metadata)
        .map_err(|error| invalid(format!("XLDM OLAP proof failed: {error}")))?;
    let olap_proof = prove_xldm140_olap(&storage, &metadata, &olap, OlapProofLimits::default())
        .map_err(|error| invalid(format!("XLDM complete OLAP proof failed: {error}")))?;
    if !olap_proof.is_complete() {
        return Err(invalid(
            "XLDM complete OLAP proof contains unknown or unlinked members".into(),
        ));
    }
    let closure = prove_xldm140_closure(&storage, &metadata, &olap, &native, &generated)
        .map_err(|error| invalid(format!("XLDM identity closure proof failed: {error}")))?;
    let calculated = grouping
        .columns
        .iter()
        .map(|column| {
            (
                column.column_name.as_str(),
                column.column_id.as_str(),
                match column.content_type {
                    super::model::TimeGroupingContentType::Years => {
                        Xldm140TimeGroupingContentType::Years
                    },
                    super::model::TimeGroupingContentType::Quarters => {
                        Xldm140TimeGroupingContentType::Quarters
                    },
                    super::model::TimeGroupingContentType::MonthsIndex => {
                        Xldm140TimeGroupingContentType::MonthsIndex
                    },
                    super::model::TimeGroupingContentType::Months => {
                        Xldm140TimeGroupingContentType::Months
                    },
                    super::model::TimeGroupingContentType::DaysIndex => {
                        Xldm140TimeGroupingContentType::DaysIndex
                    },
                    super::model::TimeGroupingContentType::Days => {
                        Xldm140TimeGroupingContentType::Days
                    },
                    super::model::TimeGroupingContentType::Hours => {
                        Xldm140TimeGroupingContentType::Hours
                    },
                    super::model::TimeGroupingContentType::Minutes => {
                        Xldm140TimeGroupingContentType::Minutes
                    },
                    super::model::TimeGroupingContentType::Seconds => {
                        Xldm140TimeGroupingContentType::Seconds
                    },
                },
            )
        })
        .collect::<Vec<_>>();
    closure
        .bind_time_grouping_with_content_types(
            &grouping.table_name,
            &grouping.column_name,
            &grouping.column_id,
            &calculated,
        )
        .map_err(|error| invalid(format!("XLDM time-grouping identity proof failed: {error}")))
}

fn invalid(message: String) -> Error {
    Error::InvalidFormat(message)
}
