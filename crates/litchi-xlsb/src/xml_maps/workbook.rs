#![cfg_attr(
    test,
    allow(
        clippy::expect_used,
        reason = "workbook XML-map tests use panic-on-failure extraction for asserted transaction outcomes"
    )
)]

//! Workbook facade for read-only XLSB XML Maps snapshots.

use litchi_opc::PackURI;

use super::{Commit, Edit, EditOptions, Patch, Snapshot};
use crate::package::error::Result;
use litchi_core::edit::Diagnostics;
use litchi_core::sheet::WorkbookTrait;

impl crate::Workbook {
    /// Read the workbook Custom XML Maps catalog and inert BIFF12 bindings
    /// under this projection's [`crate::SemanticLimits::xml_maps_value`].
    pub fn xml_maps(&self) -> Result<Snapshot> {
        Snapshot::read_for_worksheets(
            &self.package,
            self.xml_maps_worksheet_parts()?,
            self.semantic_limits().xml_maps_value(),
        )
    }

    pub(crate) fn xml_maps_worksheet_parts(&self) -> Result<Vec<PackURI>> {
        (0..self.worksheet_count())
            .map(|index| self.worksheet_uri(index))
            .collect()
    }
}

impl crate::Package {
    /// Read the workbook Custom XML Maps catalog and inert BIFF12 bindings
    /// under the package's [`crate::SemanticLimits::xml_maps_value`].
    pub fn xml_maps(&self) -> Result<Snapshot> {
        self.workbook()?.xml_maps()
    }

    /// Start a Custom XML Maps edit under the package's limits.
    pub fn edit_xml_maps(&self) -> Result<Edit> {
        self.edit_xml_maps_with_options(EditOptions::default())
    }

    /// Start a Custom XML Maps edit with explicit options.
    pub fn edit_xml_maps_with_options(&self, options: EditOptions) -> Result<Edit> {
        let limits = options
            .limits_value()
            .unwrap_or_else(|| self.semantic_limits().xml_maps_value());
        Ok(Edit::new(self.clone(), limits))
    }

    /// Apply one reversible, source-checked Custom XML Maps patch to a new
    /// package; this one is unchanged.
    ///
    /// A changed patch refuses a signed workbook (ADR 0006): remove the
    /// signature explicitly with [`crate::Package::unsign`] first. An empty
    /// patch publishes nothing and keeps the signature.
    pub fn apply_xml_maps(&self, patch: &Patch) -> Result<Commit> {
        let worksheets = self.workbook()?.xml_maps_worksheet_parts()?;
        patch.check_source(self.opc_package(), worksheets.clone())?;
        if patch.is_noop() {
            return Ok(Commit::new(
                self.clone(),
                patch.clone(),
                Diagnostics::UNCHANGED,
            ));
        }
        crate::signature_policy::ensure_unsigned(self.opc_package())?;
        let mut candidate = self.opc_package().clone();
        patch.materialize(&mut candidate)?;
        let resulting = Snapshot::read_for_worksheets(&candidate, worksheets, patch.limits())?;
        if &resulting != patch.after() {
            return Err(crate::package::error::Error::InvalidFormat(
                "XML Maps publication changed the planned semantic or owned graph",
            ));
        }
        crate::Workbook::validate_edit_candidate(&candidate)?;
        Ok(Commit::new(
            self.reparse_candidate(candidate)?,
            patch.clone(),
            Diagnostics::CHANGED,
        ))
    }
}

#[cfg(test)]
mod publication_tests {
    use crate::xml_maps::{XmlMap, XmlMapConformance, XmlMapInfo, XmlSchema};

    #[test]
    fn failed_candidate_postcondition_leaves_the_workbook_unchanged() {
        let package = crate::Package::new().expect("empty package");
        let before = package.xml_maps().expect("source snapshot");
        let before_parts = package.opc_package().part_count();
        let mut transaction = package.edit_xml_maps().expect("edit");
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

        let error = package
            .apply_xml_maps(&inconsistent)
            .expect_err("forced postcondition mismatch");
        assert!(
            error
                .to_string()
                .contains("planned semantic or owned graph")
        );
        assert_eq!(package.opc_package().part_count(), before_parts);
        assert_eq!(package.xml_maps().expect("unchanged snapshot"), before);
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
