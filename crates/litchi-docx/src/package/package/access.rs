//! Package graph access, signatures, and transactional low-level edits.

use super::super::model::{
    CustomProps, CustomPropsHost, Error, OpcPackage, Package, Result, Slot, embedded,
    validate_document_main_content_type,
};

impl Package {
    /// Refuse a story snapshot while facade-owned state is not materialized.
    pub(crate) fn ensure_story_opc_current(&self, operation: &'static str) -> Result<()> {
        self.ensure_opc_current(operation)
    }

    /// Refuse a conflict-story snapshot when facade-owned state has not yet
    /// been materialized into the OPC graph.
    pub(crate) fn ensure_conflict_opc_current(&self, operation: &'static str) -> Result<()> {
        self.ensure_story_opc_current(operation)
    }

    /// The semantic read budgets this package keeps (record 0355).
    #[must_use]
    pub fn semantic_limits(&self) -> &crate::SemanticLimits {
        &self.semantic_limits
    }

    /// The same package under other semantic read budgets. O(1): every part
    /// is shared.
    #[must_use]
    pub fn with_semantic_limits(&self, limits: crate::SemanticLimits) -> Self {
        let mut next = self.clone();
        next.semantic_limits = std::sync::Arc::new(limits);
        next
    }

    /// Get the underlying OPC package.
    ///
    /// This provides access to lower-level package operations.
    #[inline]
    #[must_use]
    pub fn opc_package(&self) -> &OpcPackage {
        &self.opc
    }

    /// Retain the immutable graph without cloning its part and relationship maps.
    pub(crate) fn shared_opc_package(&self) -> std::sync::Arc<OpcPackage> {
        self.opc.shared()
    }

    /// Return whether this document contains package signatures.
    #[must_use]
    #[inline]
    pub fn is_signed(&self) -> bool {
        self.opc.is_signed()
    }

    /// Verify package signatures with the safe strict policy.
    #[cfg(feature = "sign")]
    ///
    /// # Errors
    ///
    /// Returns an error if the operation cannot be completed.
    pub fn signatures(&self) -> litchi_opc::sign::Result<Vec<litchi_opc::sign::Report>> {
        self.opc.signatures()
    }

    /// Verify package signatures with an explicit trust-neutral policy.
    #[cfg(feature = "sign")]
    ///
    /// # Errors
    ///
    /// Returns an error if the operation cannot be completed.
    pub fn signatures_with(
        &self,
        policy: &litchi_sign::Policy,
    ) -> litchi_opc::sign::Result<Vec<litchi_opc::sign::Report>> {
        self.opc.signatures_with(policy)
    }

    /// Add a signature while preserving every existing valid signature.
    #[cfg(feature = "sign")]
    ///
    /// # Errors
    ///
    /// Returns an error if the operation cannot be completed.
    pub fn sign(&self, signer: &litchi_sign::Signer) -> litchi_opc::sign::Result<Self> {
        let mut next = self.clone();
        next.opc.sign(signer)?;
        Ok(next)
    }

    /// Add a signature with explicit authoring resource bounds.
    #[cfg(feature = "sign")]
    ///
    /// # Errors
    ///
    /// Returns an error if the operation cannot be completed.
    pub fn sign_with(
        &self,
        signer: &litchi_sign::Signer,
        limits: &litchi_sign::Limits,
    ) -> litchi_opc::sign::Result<Self> {
        let mut next = self.clone();
        next.opc.sign_with(signer, limits)?;
        Ok(next)
    }

    /// Atomically replace all signatures with one signature.
    #[cfg(feature = "sign")]
    ///
    /// # Errors
    ///
    /// Returns an error if the operation cannot be completed.
    pub fn resign(&self, signer: &litchi_sign::Signer) -> litchi_opc::sign::Result<Self> {
        let mut next = self.clone();
        next.opc.resign(signer)?;
        Ok(next)
    }

    /// Atomically replace signatures with explicit authoring resource bounds.
    #[cfg(feature = "sign")]
    ///
    /// # Errors
    ///
    /// Returns an error if the operation cannot be completed.
    pub fn resign_with(
        &self,
        signer: &litchi_sign::Signer,
        limits: &litchi_sign::Limits,
    ) -> litchi_opc::sign::Result<Self> {
        let mut next = self.clone();
        next.opc.resign_with(signer, limits)?;
        Ok(next)
    }

    /// Remove all package signatures.
    ///
    /// This is the explicit authorization ADR 0006 requires: a package-level
    /// edit never removes a signature itself, so a change to a signed package
    /// (mail merge, glossary, web settings, fonts, Ribbon, document patches,
    /// and edited core or custom properties at a managed save) is refused
    /// until this is called.
    ///
    /// Returns the unsigned package; this package keeps its signature.
    #[must_use]
    pub fn unsign(&self) -> Self {
        let mut next = self.clone();
        next.opc.unsign();
        next
    }

    /// Discover inert embedded-object and embedded-package relationships
    /// using the shared safe default resource limits.
    ///
    /// Use [`embedded::scan`] with [`Self::opc_package`] when a lower
    /// layer needs explicitly tuned limits.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation cannot be completed.
    pub fn embedded(&self) -> Result<Vec<embedded::Entry<'_>>> {
        Ok(embedded::scan(&self.opc, &embedded::Limits::default())?)
    }

    /// Load the bounded, inert classic-chart graph owned by the main document.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation cannot be completed.
    pub fn chart_graph(&self) -> Result<crate::chart::Graph> {
        let document = self.opc.main_document_part()?.partname().clone();
        crate::chart::package::load_with_limits(
            &self.opc,
            &document,
            &crate::chart::Limits::DEFAULT,
        )
    }

    /// The chart part (a `c:chartSpace` or `cx:chartSpace`) the main
    /// document relates through `relationship_id` — an
    /// [`InlineChart`](crate::chart::InlineChart)'s — or `None` when the main
    /// document has no chart relationship of that ID.
    ///
    /// # Errors
    ///
    /// Returns an error for an external chart relationship or a target that
    /// is not a chart part.
    pub fn chart_xml(&self, relationship_id: &str) -> Result<Option<Vec<u8>>> {
        let main = self.opc.main_document_part()?;
        let Some(relationship) = main.rels().get(relationship_id) else {
            return Ok(None);
        };
        let Some(content_type) = crate::chart::inline::chart_content_type(relationship.reltype())
        else {
            return Ok(None);
        };
        if relationship.is_external() {
            return Err(Error::InvalidFormat(
                "a chart relationship cannot be external".into(),
            ));
        }
        let part = self.opc.get_part(&relationship.target_partname()?)?;
        if part.content_type() != content_type {
            return Err(Error::InvalidFormat(
                "a chart relationship targets a part that is not a chart".into(),
            ));
        }
        Ok(Some(part.blob().to_vec()))
    }

    /// Load the typed, inert `SmartArt` (`DrawingML` diagram) inventory anchored
    /// in the main document.
    ///
    /// Each returned [`crate::smartart::Diagram`] carries the
    /// parsed data-model node tree, the layout/quick-style/colors part
    /// metadata, and the diagram part names. Both transitional and Strict
    /// namespace dialects are supported.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation cannot be completed.
    pub fn smart_arts(&self) -> Result<Vec<crate::smartart::Diagram>> {
        let document = self.opc.main_document_part()?.partname().clone();
        crate::smartart::load_smart_arts(&self.opc, &document)
    }

    /// Load the typed, inert text-box and `WordArt` inventory anchored in the
    /// main document.
    ///
    /// Each returned [`crate::textbox::TextBox`] carries the shape
    /// identity, the `wps:bodyPr` text-body properties, the story as
    /// paragraphs with runs, and `WordArt` warp/styling presence flags. Both
    /// `DrawingML` shapes and legacy VML `w:pict` fallbacks are recognized, in
    /// both the transitional and Strict namespace dialects.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation cannot be completed.
    pub fn text_boxes(&self) -> Result<Vec<crate::textbox::TextBox>> {
        crate::textbox::load_text_boxes(self.opc.main_document_part()?.blob())
    }

    /// Transactionally edit the current plaintext OPC graph.
    ///
    /// The closure receives a structural candidate whose built-in part payloads
    /// share immutable `Arc` storage. Returning an error or unwinding leaves
    /// this package's graph unpublished; custom `Part` implementations retain
    /// their own clone and interior-mutability policy. Before a successful
    /// commit, the candidate's Word main relationship, content type, core
    /// properties, and custom properties are validated and facade-owned state
    /// is reloaded. Committing a raw edit disables the legacy document writer
    /// so it cannot later erase the edit.
    /// Signature state is preserved. Changed signed sources require an explicit
    /// [`Self::unsign`] or a signing/disposition operation inside the closure;
    /// deleting signature parts alone does not authorize publication.
    /// Replacing a signed candidate with a separately opened or constructed
    /// package requires calling [`Self::unsign`] before entering the closure.
    /// Opaque non-Part entries in the signature directory permit exact no-op
    /// publication only; `unsign` does not remove those retained archive entries.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation cannot be completed.
    pub(crate) fn edit_opc_in_place<T>(
        &mut self,
        edit: impl FnOnce(&mut OpcPackage) -> Result<T>,
    ) -> Result<T> {
        self.edit_opc_candidate("edit_opc", true, edit)
    }

    /// Atomically publish a crate-owned semantic OPC edit.
    ///
    /// Unlike [`Self::edit_opc`], this invalidates and later rebuilds the
    /// mutable document facade without classifying the coherent edit as raw.
    /// A prior raw edit remains raw and cannot be laundered through this path.
    pub(crate) fn edit_semantic_opc<T>(
        &mut self,
        operation: &'static str,
        edit: impl FnOnce(&mut OpcPackage) -> Result<T>,
    ) -> Result<T> {
        self.edit_opc_candidate(operation, false, edit)
    }

    fn edit_opc_candidate<T>(
        &mut self,
        operation: &'static str,
        raw: bool,
        edit: impl FnOnce(&mut OpcPackage) -> Result<T>,
    ) -> Result<T> {
        self.edit_opc_candidate_known(operation, raw, |opc| Ok((edit(opc)?, None)))
    }

    /// [`Self::edit_semantic_opc`] for an edit that knows the core-properties
    /// slot of its candidate and leaves the custom-properties part untouched
    /// (`Some(slot)`), so publication adopts the slot and keeps the custom
    /// properties instead of decoding both parts again (record 0377a).
    pub(crate) fn edit_semantic_opc_known<T>(
        &mut self,
        operation: &'static str,
        edit: impl FnOnce(&mut OpcPackage) -> Result<(T, Option<Slot>)>,
    ) -> Result<T> {
        self.edit_opc_candidate_known(operation, false, edit)
    }

    fn edit_opc_candidate_known<T>(
        &mut self,
        operation: &'static str,
        raw: bool,
        edit: impl FnOnce(&mut OpcPackage) -> Result<(T, Option<Slot>)>,
    ) -> Result<T> {
        self.ensure_opc_current(operation)?;

        if !raw && (self.opc.is_signed() || self.opc.requires_signature_edit_policy()) {
            return Err(Error::Opc(
                litchi_opc::OpcError::SignedSourceRequiresExplicitPolicy,
            ));
        }
        let mut candidate = self.opc.clone();
        let (value, known) = edit(&mut candidate)?;
        candidate.validate_signature_edit_from(&self.opc)?;

        let main_part = candidate
            .main_document_part()
            .map_err(|error| Error::PartNotFound(format!("main document part: {error}")))?;
        validate_document_main_content_type(main_part.content_type())?;
        match known {
            Some(properties) => {
                self.opc = candidate;
                self.properties = properties.into();
            },
            None => {
                let properties = Slot::load(&candidate)?;
                let custom_props = CustomProps::read_for(&candidate, CustomPropsHost::Word)?;
                self.opc = candidate;
                self.properties = properties.into();
                self.custom_props = custom_props.into();
            },
        }
        Ok(value)
    }
}
