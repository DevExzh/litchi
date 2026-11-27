//! Format-neutral access to any packaged `OpenDocument` document family.

//!
//! [`Package::edit`] stages an [`Edit`] of in-content metadata (RDFa,
//! `text:meta`), XForms models and variable declarations; [`Edit::commit`]
//! publishes a new [`Package`] with its [`Patch`] and [`Diagnostics`], and
//! [`Package::apply`] replays a patch onto its exact source. The flat
//! single-file owner is [`flat::Document`] with the same family.

mod codec;
pub mod flat;
mod model;
mod package;
pub(crate) mod staged;

pub(crate) use flat::{ChargedXml, FlatMutationBudget, MemoryLease, allocate_xml};
pub(crate) use model::FlatDocument;

#[cfg(test)]
mod tests;

pub use model::{Family, Package};

staged::staged_family!(Package, "generic");

impl staged::Staged for Package {
    fn working_copy(&self) -> litchi_core::Result<Self> {
        Ok(self.clone())
    }

    fn reopen(&self, bytes: &[u8]) -> litchi_core::Result<Self> {
        let options = litchi_odf_common::OpenOptions::new()
            .source_limits(self.owned_package().limits())
            .xml_limits(self.xml_limits);
        Ok(Self::from_bytes_with_options(bytes, options)?)
    }

    fn staged_bytes(&self) -> &[u8] {
        self.as_bytes()
    }
}

impl Package {
    /// Stage an edit of this package; the package itself never changes.
    #[must_use]
    pub fn edit(&self) -> Edit<'_> {
        Edit::new(self)
    }

    /// Replay `patch` onto this exact package.
    ///
    /// # Errors
    ///
    /// Returns [`Error::PatchConflict`] when this package is not the patch
    /// source, or the reopen failure of the patched bytes.
    pub fn apply(&self, patch: &Patch) -> Result<Commit, Error> {
        patch.replay(self)
    }
}

impl Edit<'_> {
    /// Stages setting or clearing paragraph RDFa.
    pub fn set_paragraph_rdfa(
        &mut self,
        position: litchi_core::Position,
        value: &crate::RdfaAttributes,
    ) -> &mut Self {
        let value = value.clone();
        self.stage(move |document| document.set_paragraph_rdfa(position, &value))
    }

    /// Stages setting or clearing RDFa on a uniquely named bookmark start.
    pub fn set_bookmark_rdfa(
        &mut self,
        name: impl Into<String>,
        value: &crate::RdfaAttributes,
    ) -> &mut Self {
        let name = name.into();
        let value = value.clone();
        self.stage(move |document| document.set_bookmark_rdfa(&name, &value))
    }

    /// Stages setting or clearing RDFa on one inline `text:meta`.
    pub fn set_text_meta_rdfa(
        &mut self,
        position: litchi_core::Position,
        value: &crate::RdfaAttributes,
    ) -> &mut Self {
        let value = value.clone();
        self.stage(move |document| document.set_text_meta_rdfa(position, &value))
    }

    /// Stages inserting an inline `text:meta` into one paragraph.
    pub fn insert_text_meta(
        &mut self,
        paragraph: litchi_core::Position,
        value: &crate::TextMeta,
    ) -> &mut Self {
        let value = value.clone();
        self.stage(move |document| document.insert_text_meta(paragraph, &value))
    }

    /// Stages replacing one inline `text:meta`.
    pub fn replace_text_meta(
        &mut self,
        position: litchi_core::Position,
        value: &crate::TextMeta,
    ) -> &mut Self {
        let value = value.clone();
        self.stage(move |document| document.replace_text_meta(position, &value))
    }

    /// Stages removing one inline `text:meta`.
    pub fn remove_text_meta(&mut self, position: litchi_core::Position) -> &mut Self {
        self.stage(move |document| document.remove_text_meta(position))
    }

    /// Stages replacing one inert XForms model.
    pub fn replace_xforms_model(
        &mut self,
        position: litchi_core::Position,
        model: &crate::xforms::Model,
    ) -> &mut Self {
        let model = model.clone();
        self.stage(move |document| document.replace_xforms_model(position, &model))
    }

    /// Stages inserting an inert XForms model into `office:forms`.
    pub fn insert_xforms_model(&mut self, model: &crate::xforms::Model) -> &mut Self {
        let model = model.clone();
        self.stage(move |document| document.insert_xforms_model(&model))
    }

    /// Stages removing one inert XForms model.
    pub fn remove_xforms_model(&mut self, position: litchi_core::Position) -> &mut Self {
        self.stage(move |document| document.remove_xforms_model(position))
    }

    /// Stages inserting or replacing one variable declaration container.
    pub fn set_variable_declaration_group(
        &mut self,
        group: &crate::variable_declaration::Group,
    ) -> &mut Self {
        let group = group.clone();
        self.stage(move |document| document.set_variable_declaration_group(&group).map(drop))
    }

    /// Stages replacing the document interaction policy in `settings.xml`
    /// (inert metadata; never enforced).
    pub fn set_protection(&mut self, policy: &crate::protection::Policy) -> &mut Self {
        let policy = policy.clone();
        self.stage(move |document| document.set_protection(&policy).map(drop))
    }

    /// Stages removing one variable declaration container; the commit fails
    /// while a remaining field references one of its declarations.
    pub fn remove_variable_declaration_group(
        &mut self,
        part: crate::variable_declaration::Part,
        scope: &crate::variable_declaration::Scope,
        kind: crate::variable_declaration::Kind,
    ) -> &mut Self {
        let scope = scope.clone();
        self.stage(move |document| {
            document
                .remove_variable_declaration_group(part, &scope, kind)
                .map(drop)
        })
    }
}
