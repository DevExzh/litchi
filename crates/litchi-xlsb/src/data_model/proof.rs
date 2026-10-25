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
    max_part_bytes: usize,
) -> Result<Xldm140TimeGroupingBinding> {
    let limits = litchi_xldm::Limits::new(max_part_bytes, max_part_bytes);
    let storage = litchi_xldm::inspect(payload, &limits).map_err(map_xldm_proof_error)?;
    let metadata =
        litchi_xldm::metadata::inspect(&storage, &limits).map_err(map_xldm_metadata_error)?;
    let native = litchi_xldm::native::inspect(&storage, &metadata.native_parse_options())
        .map_err(|_error| invalid("XLDM native-data proof failed"))?;
    let generated = litchi_xldm::generated::inspect_system_generated(&storage)
        .map_err(|_error| invalid("XLDM generated-data proof failed"))?;
    let olap = litchi_xldm::olap::inspect(&storage, &metadata)
        .map_err(|_error| invalid("XLDM OLAP proof failed"))?;
    let olap_proof = prove_xldm140_olap(&storage, &metadata, &olap, OlapProofLimits::default())
        .map_err(map_xldm_olap_proof_for_time_grouping)?;
    if !olap_proof.is_complete() {
        return Err(invalid(
            "XLDM complete OLAP proof contains unknown or unlinked members".into(),
        ));
    }
    let closure = prove_xldm140_closure(&storage, &metadata, &olap, &native, &generated)
        .map_err(|_error| invalid("XLDM identity closure proof failed"))?;
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
        .map_err(|_error| invalid("XLDM time-grouping identity proof failed"))
}

fn invalid(message: &'static str) -> Error {
    Error::InvalidFormat(message)
}

#[allow(clippy::wildcard_enum_match_arm)]
pub(crate) fn map_xldm_error(error: litchi_xldm::Error) -> Error {
    match error {
        litchi_xldm::Error::Limit {
            kind,
            observed,
            maximum,
        } => Error::limit(super::LimitKind::Model(kind), observed, maximum),
        litchi_xldm::Error::InvalidLimit {
            kind,
            value,
            maximum,
        } => Error::InvalidLimit {
            kind: crate::package::error::LimitKind::DataModel(crate::data_model::LimitKind::Model(
                kind,
            )),
            value,
            maximum,
        },
        litchi_xldm::Error::Allocation { resource, source } => {
            Error::Allocation { resource, source }
        },
        litchi_xldm::Error::Unsupported { feature: _ } => Error::refused(
            litchi_core::edit::RefusalKind::UnsupportedContent,
            "edit_data_model",
            "invalid value",
        ),
        litchi_xldm::Error::Invalid(_message) | litchi_xldm::Error::Xml(_message) => {
            invalid("invalid value")
        },
        _ => invalid("invalid value"),
    }
}

#[allow(clippy::wildcard_enum_match_arm)]
fn map_xldm_proof_error(error: litchi_xldm::Error) -> Error {
    match error {
        litchi_xldm::Error::Limit {
            kind,
            observed,
            maximum,
        } => Error::limit(super::LimitKind::Model(kind), observed, maximum),
        litchi_xldm::Error::InvalidLimit {
            kind,
            value,
            maximum,
        } => Error::InvalidLimit {
            kind: crate::package::error::LimitKind::DataModel(crate::data_model::LimitKind::Model(
                kind,
            )),
            value,
            maximum,
        },
        litchi_xldm::Error::Invalid(_)
        | litchi_xldm::Error::Unsupported { .. }
        | litchi_xldm::Error::Allocation { .. }
        | litchi_xldm::Error::Xml(_) => invalid("XLDM outer proof failed"),
        _ => invalid("XLDM outer proof failed"),
    }
}

pub(crate) fn map_xldm_metadata_error(error: litchi_xldm::metadata::MetadataError) -> Error {
    match error {
        litchi_xldm::metadata::MetadataError::Limit {
            kind,
            observed,
            maximum,
        } => Error::limit(super::LimitKind::Model(kind), observed, maximum),
        litchi_xldm::metadata::MetadataError::InvalidLimit {
            kind,
            value,
            maximum,
        } => Error::InvalidLimit {
            kind: crate::package::error::LimitKind::DataModel(crate::data_model::LimitKind::Model(
                kind,
            )),
            value,
            maximum,
        },
        litchi_xldm::metadata::MetadataError::Allocation { .. }
        | litchi_xldm::metadata::MetadataError::Invalid(_) => invalid("invalid value"),
    }
}

pub(crate) fn map_xldm_olap_proof_error(error: litchi_xldm::OlapProofError) -> Error {
    match error {
        litchi_xldm::OlapProofError::UnsupportedProfile => Error::refused(
            litchi_core::edit::RefusalKind::UnsupportedContent,
            "edit_data_model",
            "the XLDM OLAP profile is unsupported",
        ),
        litchi_xldm::OlapProofError::Limit {
            kind,
            observed,
            maximum,
        } => Error::limit(super::LimitKind::Model(kind), observed, maximum),
        litchi_xldm::OlapProofError::Invalid { .. }
        | litchi_xldm::OlapProofError::Unproven { .. } => {
            invalid("the XLDM OLAP source is invalid or unproven")
        },
        litchi_xldm::OlapProofError::Allocation { .. } => {
            invalid("could not reserve XLDM OLAP proof")
        },
    }
}

fn map_xldm_olap_proof_for_time_grouping(error: litchi_xldm::OlapProofError) -> Error {
    match error {
        litchi_xldm::OlapProofError::Limit {
            kind,
            observed,
            maximum,
        } => Error::limit(super::LimitKind::Model(kind), observed, maximum),
        litchi_xldm::OlapProofError::UnsupportedProfile
        | litchi_xldm::OlapProofError::Invalid { .. }
        | litchi_xldm::OlapProofError::Unproven { .. }
        | litchi_xldm::OlapProofError::Allocation { .. } => {
            invalid("XLDM complete OLAP proof failed")
        },
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn nested_xldm_limit_dimensions_survive_the_xlsb_adapter() {
        let raw = map_xldm_error(litchi_xldm::Error::Limit {
            kind: litchi_xldm::LimitKind::XldmSourceBytes,
            observed: 17,
            maximum: 16,
        });
        assert!(matches!(
            raw,
            Error::Limit {
                kind: crate::package::error::LimitKind::DataModel(
                    crate::data_model::LimitKind::Model(litchi_xldm::LimitKind::XldmSourceBytes)
                ),
                observed: 17,
                maximum: 16,
            }
        ));

        let metadata =
            map_xldm_metadata_error(litchi_xldm::metadata::MetadataError::InvalidLimit {
                kind: litchi_xldm::LimitKind::MetadataTextBytes,
                value: 65,
                maximum: 64,
            });
        assert!(matches!(
            metadata,
            Error::InvalidLimit {
                kind: crate::package::error::LimitKind::DataModel(
                    crate::data_model::LimitKind::Model(litchi_xldm::LimitKind::MetadataTextBytes)
                ),
                value: 65,
                maximum: 64,
            }
        ));

        let proof = map_xldm_olap_proof_error(litchi_xldm::OlapProofError::Limit {
            kind: litchi_xldm::LimitKind::Work,
            observed: 33,
            maximum: 32,
        });
        assert!(matches!(
            proof,
            Error::Limit {
                kind: crate::package::error::LimitKind::DataModel(
                    crate::data_model::LimitKind::Model(litchi_xldm::LimitKind::Work)
                ),
                observed: 33,
                maximum: 32,
            }
        ));
    }
}
