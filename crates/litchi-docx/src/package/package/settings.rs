//! Web settings, attached templates, variables, and glossary ownership.

use super::super::model::*;

use super::parts::settings_part_from_snapshot;

impl Package {
    /// Read typed web-output settings and their conformance family.
    pub fn web(&self) -> Result<Option<(docx_web::Settings, docx_web::Conformance)>> {
        Ok(docx_web::load(&self.opc)?)
    }

    /// Move complete web-output settings into package ownership.
    pub fn put_web(
        &mut self,
        settings: docx_web::Settings,
        conformance: docx_web::Conformance,
    ) -> Result<bool> {
        Ok(docx_web::put(&mut self.opc, settings, conformance)?)
    }

    /// Remove the document-owned web-settings part.
    pub fn remove_web(&mut self) -> Result<bool> {
        Ok(docx_web::remove(&mut self.opc)?)
    }

    /// Inspect the external template associated with this document without dereferencing it.
    pub fn attached_template(&self) -> Result<Option<AttachedTemplate>> {
        let snapshot = self.settings_part_snapshot()?;
        if !snapshot.relationship_exists {
            return Ok(None);
        }
        let part = settings_part_from_snapshot(&snapshot, snapshot.xml.clone(), None);
        Ok(DocumentSettings::extract_from_part(&part)?
            .attached_template()
            .cloned())
    }

    /// Associate this document with an external template URI.
    ///
    /// The URI is recorded inertly and is never fetched or executed.
    pub fn set_attached_template_uri(
        &mut self,
        target_uri: impl Into<String>,
    ) -> Result<&mut Self> {
        use litchi_opc::part::Part;

        let target_uri = target_uri.into();
        validate_attached_template_target(&target_uri)?;
        let snapshot = self.settings_part_snapshot()?;
        let original = settings_part_from_snapshot(&snapshot, snapshot.xml.clone(), None);
        let settings = DocumentSettings::extract_from_part(&original)?;
        let old_id = settings
            .attached_template()
            .map(|template| template.relationship_id().to_owned());

        let mut replacement =
            settings_part_from_snapshot(&snapshot, snapshot.xml.clone(), old_id.as_deref());
        let relationship_id = if let Some(id) = old_id {
            replacement.rels_mut().add_relationship(
                ATTACHED_TEMPLATE_RELATIONSHIP.to_owned(),
                target_uri,
                id.clone(),
                true,
            );
            id
        } else {
            replacement.relate_to_ext(&target_uri, ATTACHED_TEMPLATE_RELATIONSHIP)
        };
        replacement.set_blob(patch_attached_template(
            &snapshot.xml,
            Some(&relationship_id),
        )?);
        DocumentSettings::extract_from_part(&replacement)?;
        self.commit_settings_part(&snapshot, replacement)?;
        Ok(self)
    }

    /// Remove the attached-template element and its referenced relationship.
    pub fn remove_attached_template(&mut self) -> Result<Option<AttachedTemplate>> {
        let snapshot = self.settings_part_snapshot()?;
        if !snapshot.relationship_exists {
            return Ok(None);
        }
        let original = settings_part_from_snapshot(&snapshot, snapshot.xml.clone(), None);
        let settings = DocumentSettings::extract_from_part(&original)?;
        let Some(attached_template) = settings.attached_template().cloned() else {
            return Ok(None);
        };
        let replacement = settings_part_from_snapshot(
            &snapshot,
            patch_attached_template(&snapshot.xml, None)?,
            Some(attached_template.relationship_id()),
        );
        DocumentSettings::extract_from_part(&replacement)?;
        self.commit_settings_part(&snapshot, replacement)?;
        Ok(Some(attached_template))
    }

    /// Read the document variables stored in `settings.xml`.
    pub fn document_variables(&self) -> Result<Option<Variables>> {
        let snapshot = self.settings_part_snapshot()?;
        if !snapshot.relationship_exists {
            return Ok(None);
        }
        let part = settings_part_from_snapshot(&snapshot, snapshot.xml.clone(), None);
        Ok(Some(extract_document_variables(&part)?))
    }

    /// Insert or replace one document variable atomically.
    pub fn set_document_variable(
        &mut self,
        name: impl Into<String>,
        value: impl Into<String>,
    ) -> Result<Option<String>> {
        let snapshot = self.settings_part_snapshot()?;
        let original = settings_part_from_snapshot(&snapshot, snapshot.xml.clone(), None);
        DocumentSettings::extract_from_part(&original)?;
        let mut variables = extract_document_variables(&original)?;
        let previous = variables.insert(name, value)?;
        let xml = patch_document_variables(&snapshot.xml, &variables)?;
        let replacement = settings_part_from_snapshot(&snapshot, xml, None);
        DocumentSettings::extract_from_part(&replacement)?;
        extract_document_variables(&replacement)?;
        self.commit_settings_part(&snapshot, replacement)?;
        Ok(previous)
    }

    /// Remove one document variable atomically.
    pub fn remove_document_variable(&mut self, name: &str) -> Result<Option<String>> {
        let snapshot = self.settings_part_snapshot()?;
        if !snapshot.relationship_exists {
            return Ok(None);
        }
        let original = settings_part_from_snapshot(&snapshot, snapshot.xml.clone(), None);
        DocumentSettings::extract_from_part(&original)?;
        let mut variables = extract_document_variables(&original)?;
        let Some(previous) = variables.remove(name) else {
            return Ok(None);
        };
        let xml = patch_document_variables(&snapshot.xml, &variables)?;
        let replacement = settings_part_from_snapshot(&snapshot, xml, None);
        DocumentSettings::extract_from_part(&replacement)?;
        extract_document_variables(&replacement)?;
        self.commit_settings_part(&snapshot, replacement)?;
        Ok(Some(previous))
    }

    /// Remove every document variable atomically and return the number removed.
    pub fn clear_document_variables(&mut self) -> Result<usize> {
        let snapshot = self.settings_part_snapshot()?;
        if !snapshot.relationship_exists {
            return Ok(0);
        }
        let original = settings_part_from_snapshot(&snapshot, snapshot.xml.clone(), None);
        DocumentSettings::extract_from_part(&original)?;
        let mut variables = extract_document_variables(&original)?;
        let count = variables.count();
        if count == 0 {
            return Ok(0);
        }
        variables.clear();
        let xml = patch_document_variables(&snapshot.xml, &variables)?;
        let replacement = settings_part_from_snapshot(&snapshot, xml, None);
        DocumentSettings::extract_from_part(&replacement)?;
        extract_document_variables(&replacement)?;
        self.commit_settings_part(&snapshot, replacement)?;
        Ok(count)
    }

    /// Load the typed glossary/building-block catalog and its dialect.
    pub fn glossary(&self) -> Result<Option<(glossary::Catalog, glossary::Conformance)>> {
        Ok(glossary::load(&self.opc)?)
    }

    /// Move a complete semantic catalog into the package.
    pub fn put_glossary(
        &mut self,
        catalog: glossary::Catalog,
        conformance: glossary::Conformance,
    ) -> Result<bool> {
        Ok(glossary::put(&mut self.opc, catalog, conformance)?)
    }

    /// Load the complete low-level glossary OPC graph without copying payloads.
    pub fn glossary_graph(&self) -> Result<Option<glossary::raw::Graph>> {
        Ok(glossary::raw::load(&self.opc)?)
    }

    /// Publish a complete low-level glossary OPC graph into the package.
    ///
    /// Returns `false` when the graph is already identical, preserving package
    /// bytes and digital signatures.
    pub fn put_glossary_graph(&mut self, graph: &glossary::raw::Graph) -> Result<bool> {
        Ok(glossary::raw::put(&mut self.opc, graph)?)
    }

    /// Remove and return the complete low-level glossary OPC graph.
    pub fn take_glossary_graph(&mut self) -> Result<Option<glossary::raw::Graph>> {
        Ok(glossary::raw::remove(&mut self.opc)?)
    }

    /// Remove the complete glossary-owned graph.
    pub fn remove_glossary(&mut self) -> Result<bool> {
        Ok(glossary::remove(&mut self.opc)?)
    }
}
