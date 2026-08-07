//! Package graph access, signatures, and transactional low-level edits.

use super::super::model::*;

impl Package {
    /// Get the underlying OPC package.
    ///
    /// This provides access to lower-level package operations.
    #[inline]
    pub fn opc_package(&self) -> &OpcPackage {
        &self.opc
    }

    /// Return whether this document contains package signatures.
    #[must_use]
    #[inline]
    pub fn is_signed(&self) -> bool {
        self.opc.is_signed()
    }

    /// Verify package signatures with the safe strict policy.
    pub fn signatures(&self) -> litchi_opc::sign::Result<Vec<litchi_opc::sign::Report>> {
        self.opc.signatures()
    }

    /// Verify package signatures with an explicit trust-neutral policy.
    pub fn signatures_with(
        &self,
        policy: &litchi_sign::Policy,
    ) -> litchi_opc::sign::Result<Vec<litchi_opc::sign::Report>> {
        self.opc.signatures_with(policy)
    }

    /// Add a signature while preserving every existing valid signature.
    pub fn sign(&mut self, signer: &litchi_sign::Signer) -> litchi_opc::sign::Result<PackURI> {
        self.opc.sign(signer)
    }

    /// Add a signature with explicit authoring resource bounds.
    pub fn sign_with(
        &mut self,
        signer: &litchi_sign::Signer,
        limits: &litchi_sign::Limits,
    ) -> litchi_opc::sign::Result<PackURI> {
        self.opc.sign_with(signer, limits)
    }

    /// Atomically replace all signatures with one signature.
    pub fn resign(&mut self, signer: &litchi_sign::Signer) -> litchi_opc::sign::Result<PackURI> {
        self.opc.resign(signer)
    }

    /// Atomically replace signatures with explicit authoring resource bounds.
    pub fn resign_with(
        &mut self,
        signer: &litchi_sign::Signer,
        limits: &litchi_sign::Limits,
    ) -> litchi_opc::sign::Result<PackURI> {
        self.opc.resign_with(signer, limits)
    }

    /// Remove all package signatures.
    pub fn unsign(&mut self) -> &mut Self {
        self.opc.unsign();
        self
    }

    /// Discover inert embedded-object and embedded-package relationships
    /// using the shared safe default resource limits.
    ///
    /// Use [`embedded::scan_with`] with [`Self::opc_package`] when a lower
    /// layer needs explicitly tuned limits.
    pub fn embedded(&self) -> Result<Vec<embedded::Entry<'_>>> {
        Ok(embedded::scan(&self.opc)?)
    }

    /// Load the bounded, inert classic-chart graph owned by the main document.
    pub fn chart_graph(&self) -> Result<crate::chart::Graph> {
        let document = self.opc.main_document_part()?.partname().clone();
        crate::chart::load(&self.opc, &document)
    }

    /// Load the typed, inert SmartArt (DrawingML diagram) inventory anchored
    /// in the main document.
    ///
    /// Each returned [`crate::smartart::Diagram`] carries the
    /// parsed data-model node tree, the layout/quick-style/colors part
    /// metadata, and the diagram part names. Both transitional and Strict
    /// namespace dialects are supported.
    pub fn smart_arts(&self) -> Result<Vec<crate::smartart::Diagram>> {
        let document = self.opc.main_document_part()?.partname().clone();
        crate::smartart::load_smart_arts(&self.opc, &document)
    }

    /// Load the typed, inert text-box and WordArt inventory anchored in the
    /// main document.
    ///
    /// Each returned [`crate::textbox::TextBox`] carries the shape
    /// identity, the `wps:bodyPr` text-body properties, the story as
    /// paragraphs with runs, and WordArt warp/styling presence flags. Both
    /// DrawingML shapes and legacy VML `w:pict` fallbacks are recognized, in
    /// both the transitional and Strict namespace dialects.
    pub fn text_boxes(&self) -> Result<Vec<crate::textbox::TextBox>> {
        crate::textbox::load_text_boxes(self.opc.main_document_part()?.blob())
    }

    /// Deterministically store an already coherent classic-chart graph.
    pub fn store_chart_graph(&mut self, graph: &crate::chart::Graph) -> Result<()> {
        let document = self.opc.main_document_part()?.partname().clone();
        crate::chart::store(&mut self.opc, &document, graph)
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
    pub fn edit_opc<T>(&mut self, edit: impl FnOnce(&mut OpcPackage) -> Result<T>) -> Result<T> {
        self.ensure_opc_current("edit_opc")?;
        #[cfg(feature = "fonts")]
        if self.font_embedding.is_some() {
            return Err(Error::UnsafeEdit {
                format: "DOCX",
                operation: "edit_opc",
                reason: "raw OPC editing cannot honor an automatic font policy; use the managed font facade",
            });
        }

        let mut candidate = self.opc.clone();
        candidate.unsign();
        let value = edit(&mut candidate)?;

        let main_part = candidate
            .main_document_part()
            .map_err(|error| Error::PartNotFound(format!("main document part: {error}")))?;
        validate_document_main_content_type(main_part.content_type())?;
        let properties = Slot::load(&candidate)?;
        let custom_props = CustomProps::read(&candidate)?;

        self.opc = candidate;
        self.properties = properties;
        self.custom_props = custom_props;
        self.custom_props_dirty = false;
        self.mutable_doc = None;
        self.raw_edit_committed = true;
        Ok(value)
    }
}
