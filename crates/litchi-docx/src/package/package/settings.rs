#![expect(
    clippy::shadow_reuse,
    reason = "parser bindings are intentionally refined after validation"
)]
//! Web settings, attached templates, variables, and glossary ownership.

use super::super::model::{
    ATTACHED_TEMPLATE_RELATIONSHIP, AttachedTemplate, DocumentSettings, Package, Result, Variables,
    docx_web, extract_document_variables, glossary, patch_attached_template,
    patch_document_variables, validate_attached_template_target,
};

use super::parts::settings_part_from_snapshot;

impl Package {
    /// Read typed web-output settings and their conformance family.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation cannot be completed.
    pub fn web(&self) -> Result<Option<(docx_web::Settings, docx_web::Conformance)>> {
        docx_web::load(&self.opc, &docx_web::Limits::DEFAULT)
    }

    /// Move complete web-output settings into package ownership.
    ///
    /// The signature policy is [`docx_web::raw::put`]'s: equal settings publish
    /// nothing, and a change to a signed package is refused until
    /// [`Self::unsign`] removes the signature.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation cannot be completed.
    pub(crate) fn put_web(
        &mut self,
        settings: docx_web::Settings,
        conformance: docx_web::Conformance,
    ) -> Result<bool> {
        docx_web::package::put(&mut self.opc, settings, conformance)
    }

    /// Remove the document-owned web-settings part.
    ///
    /// The signature policy is [`docx_web::raw::remove`]'s: removing a part from a
    /// signed package is refused until [`Self::unsign`] removes the signature.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation cannot be completed.
    pub(crate) fn remove_web(&mut self) -> Result<bool> {
        docx_web::package::remove(&mut self.opc)
    }

    /// Inspect the external template associated with this document without dereferencing it.
    ///
    /// Settings repairs applied while reading (see
    /// [`crate::leniency::SettingsDefect`]) are enumerable through
    /// [`DocumentSettings::tolerance_report`] on [`crate::Document::settings`].
    ///
    /// # Errors
    ///
    /// Returns an error if the operation cannot be completed.
    pub fn attached_template(&self) -> Result<Option<AttachedTemplate>> {
        let snapshot = self.settings_part_snapshot()?;
        if !snapshot.relationship_exists {
            return Ok(None);
        }
        let part = settings_part_from_snapshot(&snapshot, snapshot.xml.clone(), None);
        let limits = crate::mail_merge::mce_limits_for_source(&self.opc.read_limits());
        Ok(
            DocumentSettings::extract_from_part_with_mce_limits(&part, &limits)?
                .attached_template()
                .cloned(),
        )
    }

    /// Associate this document with an external template URI.
    ///
    /// The URI is recorded inertly and is never fetched or executed.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation cannot be completed.
    pub(crate) fn set_attached_template_uri(
        &mut self,
        target_uri: impl Into<String>,
    ) -> Result<&mut Self> {
        use litchi_opc::part::Part;

        let target_uri = target_uri.into();
        validate_attached_template_target(&target_uri)?;
        let snapshot = self.settings_part_snapshot()?;
        let original = settings_part_from_snapshot(&snapshot, snapshot.xml.clone(), None);
        let limits = crate::mail_merge::mce_limits_for_source(&self.opc.read_limits());
        let settings = DocumentSettings::extract_from_part_with_mce_limits(&original, &limits)?;
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
            &limits,
        )?);
        DocumentSettings::extract_from_part_with_mce_limits(&replacement, &limits)?;
        self.commit_settings_part(&snapshot, replacement)?;
        Ok(self)
    }

    /// Remove the attached-template element and its referenced relationship.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation cannot be completed.
    pub(crate) fn remove_attached_template(&mut self) -> Result<Option<AttachedTemplate>> {
        let snapshot = self.settings_part_snapshot()?;
        if !snapshot.relationship_exists {
            return Ok(None);
        }
        let original = settings_part_from_snapshot(&snapshot, snapshot.xml.clone(), None);
        let limits = crate::mail_merge::mce_limits_for_source(&self.opc.read_limits());
        let settings = DocumentSettings::extract_from_part_with_mce_limits(&original, &limits)?;
        let Some(attached_template) = settings.attached_template().cloned() else {
            return Ok(None);
        };
        let replacement = settings_part_from_snapshot(
            &snapshot,
            patch_attached_template(&snapshot.xml, None, &limits)?,
            Some(attached_template.relationship_id()),
        );
        DocumentSettings::extract_from_part_with_mce_limits(&replacement, &limits)?;
        self.commit_settings_part(&snapshot, replacement)?;
        Ok(Some(attached_template))
    }

    /// Read the document variables stored in `settings.xml`.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation cannot be completed.
    pub fn document_variables(&self) -> Result<Option<Variables>> {
        let snapshot = self.settings_part_snapshot()?;
        if !snapshot.relationship_exists {
            return Ok(None);
        }
        let part = settings_part_from_snapshot(&snapshot, snapshot.xml.clone(), None);
        let limits = crate::mail_merge::mce_limits_for_source(&self.opc.read_limits());
        Ok(Some(extract_document_variables(&part, &limits)?))
    }

    /// Insert or replace one document variable atomically.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation cannot be completed.
    pub(crate) fn set_document_variable(
        &mut self,
        name: impl Into<String>,
        value: impl Into<String>,
    ) -> Result<Option<String>> {
        let snapshot = self.settings_part_snapshot()?;
        let original = settings_part_from_snapshot(&snapshot, snapshot.xml.clone(), None);
        let limits = crate::mail_merge::mce_limits_for_source(&self.opc.read_limits());
        DocumentSettings::extract_from_part_with_mce_limits(&original, &limits)?;
        let mut variables = extract_document_variables(&original, &limits)?;
        let previous = variables.insert(name, value)?;
        let xml = patch_document_variables(&snapshot.xml, &variables, &limits)?;
        let replacement = settings_part_from_snapshot(&snapshot, xml, None);
        DocumentSettings::extract_from_part_with_mce_limits(&replacement, &limits)?;
        extract_document_variables(&replacement, &limits)?;
        self.commit_settings_part(&snapshot, replacement)?;
        Ok(previous)
    }

    /// Remove one document variable atomically.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation cannot be completed.
    pub(crate) fn remove_document_variable(&mut self, name: &str) -> Result<Option<String>> {
        let snapshot = self.settings_part_snapshot()?;
        if !snapshot.relationship_exists {
            return Ok(None);
        }
        let original = settings_part_from_snapshot(&snapshot, snapshot.xml.clone(), None);
        let limits = crate::mail_merge::mce_limits_for_source(&self.opc.read_limits());
        DocumentSettings::extract_from_part_with_mce_limits(&original, &limits)?;
        let mut variables = extract_document_variables(&original, &limits)?;
        let Some(previous) = variables.remove(name) else {
            return Ok(None);
        };
        let xml = patch_document_variables(&snapshot.xml, &variables, &limits)?;
        let replacement = settings_part_from_snapshot(&snapshot, xml, None);
        DocumentSettings::extract_from_part_with_mce_limits(&replacement, &limits)?;
        extract_document_variables(&replacement, &limits)?;
        self.commit_settings_part(&snapshot, replacement)?;
        Ok(Some(previous))
    }

    /// Remove every document variable atomically and return the number removed.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation cannot be completed.
    pub(crate) fn clear_document_variables(&mut self) -> Result<usize> {
        let snapshot = self.settings_part_snapshot()?;
        if !snapshot.relationship_exists {
            return Ok(0);
        }
        let original = settings_part_from_snapshot(&snapshot, snapshot.xml.clone(), None);
        let limits = crate::mail_merge::mce_limits_for_source(&self.opc.read_limits());
        DocumentSettings::extract_from_part_with_mce_limits(&original, &limits)?;
        let mut variables = extract_document_variables(&original, &limits)?;
        let count = variables.count();
        if count == 0 {
            return Ok(0);
        }
        variables.clear();
        let xml = patch_document_variables(&snapshot.xml, &variables, &limits)?;
        let replacement = settings_part_from_snapshot(&snapshot, xml, None);
        DocumentSettings::extract_from_part_with_mce_limits(&replacement, &limits)?;
        extract_document_variables(&replacement, &limits)?;
        self.commit_settings_part(&snapshot, replacement)?;
        Ok(count)
    }

    /// Load the typed glossary/building-block catalog and its dialect.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation cannot be completed.
    pub fn glossary(&self) -> Result<Option<(glossary::Catalog, glossary::Conformance)>> {
        let limits = self.semantic_limits().glossary_value();
        glossary::load(&self.opc, &limits)
    }

    /// Move a complete semantic catalog into the package.
    ///
    /// The signature policy is [`glossary::raw::put`]'s: an unchanged catalog
    /// publishes nothing, and a change to a signed package is refused until
    /// [`Self::unsign`] removes the signature.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation cannot be completed.
    pub(crate) fn put_glossary(
        &mut self,
        catalog: glossary::Catalog,
        conformance: glossary::Conformance,
    ) -> Result<bool> {
        let limits = self.semantic_limits().glossary_value();
        glossary::package::put(&mut self.opc, catalog, conformance, &limits)
    }

    /// Load the complete low-level glossary OPC graph without copying payloads.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation cannot be completed.
    pub fn glossary_graph(&self) -> Result<Option<glossary::raw::Graph>> {
        let limits = self.semantic_limits().glossary_value();
        glossary::raw::load(&self.opc, &limits)
    }

    /// Publish a complete low-level glossary OPC graph into the package.
    ///
    /// Returns `false` when the graph is already identical, preserving package
    /// bytes and digital signatures. A different graph changes the package,
    /// so a signed package is refused with
    /// [`litchi_opc::OpcError::SignedSourceRequiresExplicitPolicy`] (ADR 0006)
    /// until [`Self::unsign`] removes the signature.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation cannot be completed.
    pub(crate) fn put_glossary_graph(&mut self, graph: &glossary::raw::Graph) -> Result<bool> {
        let limits = self.semantic_limits().glossary_value();
        glossary::raw::put(&mut self.opc, graph, &limits)
    }

    /// Remove and return the complete low-level glossary OPC graph.
    ///
    /// The signature policy is [`Self::remove_glossary`]'s.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation cannot be completed.
    pub(crate) fn take_glossary_graph(&mut self) -> Result<Option<glossary::raw::Graph>> {
        let limits = self.semantic_limits().glossary_value();
        glossary::raw::remove(&mut self.opc, &limits)
    }

    /// Remove the complete glossary-owned graph.
    ///
    /// A package without a glossary is left untouched, signed or not. Removing
    /// one from a signed package is refused with
    /// [`litchi_opc::OpcError::SignedSourceRequiresExplicitPolicy`] (ADR 0006)
    /// until [`Self::unsign`] removes the signature.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation cannot be completed.
    pub(crate) fn remove_glossary(&mut self) -> Result<bool> {
        let limits = self.semantic_limits().glossary_value();
        glossary::package::remove(&mut self.opc, &limits)
    }
}
