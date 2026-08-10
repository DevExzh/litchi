#![cfg_attr(
    test,
    allow(
        clippy::expect_used,
        reason = "workbook XML-map tests use panic-on-failure extraction for asserted transaction outcomes"
    )
)]

//! Workbook facade for read-only XLSB XML Maps snapshots.

use litchi_opc::PackURI;

use super::{Commit, Patch, ReadLimits, Snapshot};
use crate::package::error::Result;

impl crate::Workbook {
    /// Read the workbook Custom XML Maps catalog and inert BIFF12 bindings.
    pub fn xml_maps(&self) -> Result<Snapshot> {
        self.xml_maps_with_limits(ReadLimits::DEFAULT)
    }

    /// Read the workbook Custom XML Maps catalog and bindings with explicit limits.
    pub fn xml_maps_with_limits(&self, limits: ReadLimits) -> Result<Snapshot> {
        Snapshot::read_for_worksheets(&self.package, self.xml_maps_worksheet_parts()?, limits)
    }

    /// Atomically publish one detached, source-bound XML Maps commit.
    pub fn apply_xml_maps(&mut self, commit: &Commit) -> Result<Snapshot> {
        self.apply_xml_maps_patch(commit.patch())
    }

    /// Atomically publish one reversible XML Maps patch.
    pub fn apply_xml_maps_patch(&mut self, patch: &Patch) -> Result<Snapshot> {
        let worksheets = self.xml_maps_worksheet_parts()?;
        let current = patch.check_source(&self.package, worksheets.clone())?;
        if patch.is_empty() {
            return Ok(current);
        }
        self.edit_opc(|candidate| {
            patch.materialize(candidate)?;
            let resulting = Snapshot::read_for_worksheets(candidate, worksheets, patch.limits())?;
            if &resulting != patch.after() {
                return Err(crate::package::error::Error::InvalidFormat(
                    "XML Maps publication changed the planned semantic or owned graph".to_string(),
                ));
            }
            Ok(resulting)
        })
    }

    pub(crate) fn xml_maps_worksheet_parts(&self) -> Result<Vec<PackURI>> {
        (0..self.formula_context.worksheet_names.len())
            .map(|index| self.worksheet_uri(index))
            .collect()
    }
}

#[cfg(test)]
mod publication_tests {
    use crate::xml_maps::{XmlMap, XmlMapConformance, XmlMapInfo, XmlSchema};

    #[test]
    fn failed_candidate_postcondition_leaves_the_workbook_unchanged() {
        let mut workbook = crate::Package::create()
            .expect("empty package")
            .into_workbook()
            .expect("workbook");
        let before = workbook.xml_maps().expect("source snapshot");
        let before_parts = workbook.opc_package().part_count();
        let mut transaction = before.edit();
        transaction
            .set_catalog(fixture_catalog())
            .expect("stage catalog");
        let commit = transaction.commit().expect("commit");
        let inconsistent_after = commit
            .patch()
            .after()
            .clone()
            .with_conformance_fixture(XmlMapConformance::Strict);
        let inconsistent = commit
            .patch()
            .clone()
            .with_after_fixture(inconsistent_after);

        let error = workbook
            .apply_xml_maps_patch(&inconsistent)
            .expect_err("forced postcondition mismatch");
        assert!(
            error
                .to_string()
                .contains("planned semantic or owned graph")
        );
        assert_eq!(workbook.opc_package().part_count(), before_parts);
        assert_eq!(workbook.xml_maps().expect("unchanged snapshot"), before);
    }

    fn fixture_catalog() -> XmlMapInfo {
        XmlMapInfo {
            selection_namespaces: "xmlns:e='urn:test'".to_string(),
            schemas: vec![XmlSchema {
                id: "schema-7".to_string(),
                schema_reference: Some("urn:test".to_string()),
                namespace: Some("urn:test".to_string()),
                payload_xml: Some(b"<e:schema xmlns:e=\"urn:test\"/>".to_vec()),
            }],
            maps: vec![XmlMap {
                id: 7,
                name: "Map".to_string(),
                root_element: "root".to_string(),
                schema_id: "schema-7".to_string(),
                show_import_export_validation_errors: false,
                auto_fit: false,
                append: false,
                preserve_sort_auto_filter_layout: false,
                preserve_format: false,
                data_binding: None,
            }],
        }
    }
}
