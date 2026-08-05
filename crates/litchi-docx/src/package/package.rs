//! Relationship-backed DOCX package graph orchestration.

use super::model::*;

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
        if self.opc.save_options().fonts != litchi_opc::FontEmbedding::None {
            return Err(Error::UnsafeEdit {
                format: "DOCX",
                operation: "edit_opc",
                reason: "raw OPC editing cannot honor an automatic font policy; use the managed font facade",
            });
        }

        let mut candidate = self.opc.clone();
        candidate.unsign();
        let value = edit(&mut candidate)?;

        if candidate.save_options().fonts != litchi_opc::FontEmbedding::None {
            return Err(Error::UnsafeEdit {
                format: "DOCX",
                operation: "edit_opc",
                reason: "raw OPC transactions cannot configure automatic font embedding; use the managed font facade",
            });
        }
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

    /// Get a mutable document for writing and modification.
    ///
    /// This returns a `MutableDocument` that allows you to add and modify
    /// paragraphs, tables, and other document elements.
    ///
    /// # Examples
    ///
    /// ```rust,no_run
    /// use litchi_docx::Package;
    ///
    /// let mut pkg = Package::new()?;
    /// let mut doc = pkg.document_mut()?;
    ///
    /// // Add content
    /// doc.add_paragraph_with_text("Hello, World!");
    /// let para = doc.add_paragraph();
    /// para.add_run_with_text("Bold text").bold(true);
    ///
    /// // Add a table
    /// let table = doc.add_table(3, 2);
    /// if let Some(cell) = table.cell(0, 0) {
    ///     cell.set_text("Header 1");
    /// }
    ///
    /// pkg.save("output.docx")?;
    /// # Ok::<(), Box<dyn std::error::Error + Send + Sync>>(())
    /// ```
    pub fn document_mut(&mut self) -> Result<&mut MutableDocument> {
        if self.raw_edit_committed {
            return Err(Error::UnsafeEdit {
                format: "DOCX",
                operation: "document_mut",
                reason: "a raw OPC edit committed; use edit_opc for further low-level changes",
            });
        }

        // If we don't have a mutable document, try to load it from the package
        if self.mutable_doc.is_none() {
            let doc_uri = PackURI::new("/word/document.xml")
                .map_err(|e| Error::InvalidUri(format!("document URI: {}", e)))?;

            // Try to get existing document content
            if let Ok(part) = self.opc.get_part(&doc_uri) {
                let xml = std::str::from_utf8(part.blob())
                    .map_err(|e| Error::InvalidFormat(format!("Invalid UTF-8: {}", e)))?;
                self.mutable_doc = Some(MutableDocument::from_xml(xml)?);
            } else {
                // Create a new empty document
                self.mutable_doc = Some(MutableDocument::new());
            }
        }

        self.mutable_doc.as_mut().ok_or_else(|| {
            Error::InvalidFormat("mutable document initialization did not complete".into())
        })
    }

    /// Append a package-backed alternative-format import to the document body.
    pub fn add_alt(&mut self, import: Import, match_source: Option<bool>) -> Result<Chunk> {
        let index = self.document_mut()?.alts().len();
        self.insert_alt(index, import, match_source)
    }

    /// Insert a package-backed alternative-format import by anchor-relative index.
    ///
    /// Part, relationship, and body mutations are rolled back together on error.
    pub fn insert_alt(
        &mut self,
        index: usize,
        import: Import,
        match_source: Option<bool>,
    ) -> Result<Chunk> {
        let count = self.document_mut()?.alts().len();
        if index > count {
            return Err(Error::InvalidFormat(format!(
                "altChunk index {index} is out of range"
            )));
        }
        if count >= MAX_CHUNKS {
            return Err(Error::InvalidFormat(format!(
                "alternative-format anchor limit of {MAX_CHUNKS} is exhausted"
            )));
        }
        let namespace = self.alt_chunk_namespace()?;
        let (chunk, installed_part) =
            self.install_alt_chunk_target(import, match_source, namespace)?;
        if let Err(error) = self
            .document_mut()?
            .insert_alt(index, chunk.clone(), namespace)
        {
            self.rollback_alt_chunk_target(chunk.relationship().as_str(), installed_part.as_ref())?;
            return Err(error);
        }
        Ok(chunk)
    }

    /// Replace an anchor and its relationship as one package mutation.
    pub fn replace_alt(
        &mut self,
        index: usize,
        import: Import,
        match_source: Option<bool>,
    ) -> Result<Chunk> {
        let old = self
            .document_mut()?
            .alts()
            .get(index)
            .cloned()
            .ok_or_else(|| {
                Error::InvalidFormat(format!("altChunk index {index} is out of range"))
            })?;
        let old_target = self.validate_alt_chunk_relationship(&old)?;
        let namespace = self.alt_chunk_namespace()?;
        let (new, installed_part) =
            self.install_alt_chunk_target(import, match_source, namespace)?;
        if let Err(error) = self
            .document_mut()?
            .replace_alt(index, new.clone(), namespace)
        {
            self.rollback_alt_chunk_target(new.relationship().as_str(), installed_part.as_ref())?;
            return Err(error);
        }
        self.remove_alt_relationship(&old, old_target.as_ref())?;
        Ok(old)
    }

    /// Remove an anchor, its relationship, and an unreachable internal payload.
    pub fn remove_alt(&mut self, index: usize) -> Result<Chunk> {
        let old = self
            .document_mut()?
            .alts()
            .get(index)
            .cloned()
            .ok_or_else(|| {
                Error::InvalidFormat(format!("altChunk index {index} is out of range"))
            })?;
        let old_target = self.validate_alt_chunk_relationship(&old)?;
        self.document_mut()?.remove_alt(index)?;
        self.remove_alt_relationship(&old, old_target.as_ref())?;
        Ok(old)
    }

    /// Reorder body anchors without changing their package relationships.
    pub fn move_alt(&mut self, from: usize, to: usize) -> Result<()> {
        self.document_mut()?.move_alt(from, to)
    }

    fn alt_chunk_namespace(&self) -> Result<Conformance> {
        let document_uri = PackURI::new("/word/document.xml")
            .map_err(|error| Error::InvalidUri(format!("document URI: {error}")))?;
        let strict = self
            .opc
            .get_part(&document_uri)
            .map(|part| {
                part.blob()
                    .windows(b"http://purl.oclc.org/ooxml/wordprocessingml/main".len())
                    .any(|window| window == b"http://purl.oclc.org/ooxml/wordprocessingml/main")
            })
            .unwrap_or(false);
        Ok(if strict {
            Conformance::Strict
        } else {
            Conformance::Transitional
        })
    }

    fn install_alt_chunk_target(
        &mut self,
        import: Import,
        match_source: Option<bool>,
        namespace: Conformance,
    ) -> Result<(Chunk, Option<PackURI>)> {
        let document_uri = PackURI::new("/word/document.xml")
            .map_err(|error| Error::InvalidUri(format!("document URI: {error}")))?;
        let document = self.opc.get_part(&document_uri)?;
        let relationship_id = (1usize..=MAX_CHUNKS)
            .map(|number| format!("rIdAltChunk{number}"))
            .find(|id| document.rels().get(id).is_none())
            .ok_or_else(|| {
                Error::InvalidFormat("altChunk relationship ID space is exhausted".into())
            })?;
        let relationship = Rel::new(relationship_id.clone())?;
        let relationship_type = namespace.relationship();
        let (target_ref, target_mode, installed_part) = match import {
            Import::Link(uri) => (uri.into_string(), TargetMode::External, None),
            Import::Data(data) => {
                data.validate()?;
                let media_type = data.media_type();
                let (uri, target_ref) = (1usize..=MAX_CHUNKS)
                    .find_map(|number| {
                        let target_ref = format!("afchunk{number}.{}", data.extension());
                        let uri = PackURI::new(format!("/word/{target_ref}")).ok()?;
                        self.opc
                            .get_part(&uri)
                            .is_err()
                            .then_some((uri, target_ref))
                    })
                    .ok_or_else(|| {
                        Error::InvalidFormat(
                            "alternative-format part-name space is exhausted".into(),
                        )
                    })?;
                self.opc.try_add_part(Box::new(BlobPart::new(
                    uri.clone(),
                    media_type.to_string(),
                    data.into_bytes(),
                )))?;
                (target_ref, TargetMode::Internal, Some(uri))
            },
        };
        let relation_result = self
            .opc
            .get_part_mut(&document_uri)?
            .rels_mut()
            .try_add_relationship(
                relationship_type.to_string(),
                target_ref,
                relationship_id.clone(),
                target_mode,
            );
        if let Err(error) = relation_result {
            if let Some(uri) = &installed_part {
                self.opc.remove_part(uri);
            }
            return Err(error.into());
        }
        Ok((Chunk::new(relationship, match_source), installed_part))
    }

    fn validate_alt_chunk_relationship(&self, chunk: &Chunk) -> Result<Option<PackURI>> {
        let document_uri = PackURI::new("/word/document.xml")
            .map_err(|error| Error::InvalidUri(format!("document URI: {error}")))?;
        let relationship = self
            .opc
            .get_part(&document_uri)?
            .rels()
            .get(chunk.relationship().as_str())
            .ok_or_else(|| {
                Error::InvalidFormat(format!(
                    "altChunk relationship {:?} is missing",
                    chunk.relationship().as_str()
                ))
            })?;
        if !is_relationship(relationship.reltype()) {
            return Err(Error::InvalidFormat(format!(
                "relationship {:?} is not an alternative-format import",
                chunk.relationship().as_str()
            )));
        }
        if relationship.is_external() {
            return Ok(None);
        }
        let target = relationship.target_partname().map_err(|error| {
            Error::InvalidFormat(format!(
                "invalid altChunk relationship {:?}: {error}",
                chunk.relationship().as_str()
            ))
        })?;
        self.opc.get_part(&target).map_err(|_| {
            Error::InvalidFormat(format!(
                "altChunk relationship {:?} targets a missing part",
                chunk.relationship().as_str()
            ))
        })?;
        Ok(Some(target))
    }

    fn rollback_alt_chunk_target(&mut self, id: &str, part: Option<&PackURI>) -> Result<()> {
        let document_uri = PackURI::new("/word/document.xml")
            .map_err(|error| Error::InvalidUri(format!("document URI: {error}")))?;
        self.opc.get_part_mut(&document_uri)?.rels_mut().remove(id);
        if let Some(part) = part {
            self.opc.remove_part(part);
        }
        Ok(())
    }

    fn remove_alt_relationship(&mut self, chunk: &Chunk, target: Option<&PackURI>) -> Result<()> {
        if self.mutable_doc.as_ref().is_some_and(|document| {
            document
                .alts()
                .iter()
                .any(|remaining| remaining.relationship() == chunk.relationship())
        }) {
            return Ok(());
        }
        let document_uri = PackURI::new("/word/document.xml")
            .map_err(|error| Error::InvalidUri(format!("document URI: {error}")))?;
        self.opc
            .get_part_mut(&document_uri)?
            .rels_mut()
            .remove(chunk.relationship().as_str());
        let Some(target) = target else {
            return Ok(());
        };
        let package_reference = self.opc.rels().iter().any(|relationship| {
            !relationship.is_external()
                && relationship
                    .target_partname()
                    .is_ok_and(|part| &part == target)
        });
        let part_reference = self.opc.iter_parts().any(|part| {
            part.rels().iter().any(|relationship| {
                !relationship.is_external()
                    && relationship
                        .target_partname()
                        .is_ok_and(|part| &part == target)
            })
        });
        if !package_reference && !part_reference {
            self.opc.remove_part(target);
        }
        Ok(())
    }

    /// Discover every validated Custom XML Data Storage relationship occurrence.
    pub fn custom_xml(&self) -> Result<Vec<CustomXmlItem>> {
        Ok(custom_xml::discover(&self.opc)?)
    }

    /// Discover typed, inert bibliography source stores from Custom XML.
    ///
    /// Word stores its current bibliography source list in a document Custom
    /// XML data store. This method exposes stored source values and style
    /// metadata only. It never matches source tags to citations, resolves
    /// schemas or styles, runs transforms, refreshes fields, or changes data.
    pub fn bibliography_source_stores(&self) -> Result<Vec<SourceStore>> {
        let items = custom_xml::discover(&self.opc)?;
        discover_bibliography_source_stores(&items)
    }

    /// Discover typed, inert bibliography sources in package and XML order.
    ///
    /// This flattens [`Self::bibliography_source_stores`] without resolving
    /// `CITATION` fields or applying bibliography style rules.
    pub fn bibliography_sources(&self) -> Result<Vec<BibliographySource>> {
        let stores = self.bibliography_source_stores()?;
        Ok(stores
            .iter()
            .flat_map(|store| store.sources().iter().cloned())
            .collect())
    }

    /// Return the number of typed, inert bibliography sources.
    pub fn bibliography_source_count(&self) -> Result<usize> {
        Ok(self.bibliography_sources()?.len())
    }

    /// Find a Custom XML data store by its case-insensitive datastore item GUID.
    pub fn custom_xml_by_id(&self, id: &str) -> Result<Option<CustomXmlItem>> {
        Ok(custom_xml::discover(&self.opc)?.into_iter().find(|item| {
            item.props()
                .is_some_and(|props| props.id.eq_ignore_ascii_case(id))
        }))
    }

    /// Add a collision-safe `/customXml/itemN.xml` data store to the main document.
    pub fn add_custom_xml(&mut self, store: NewStore) -> Result<CustomXmlItem> {
        custom_xml::validate_content_type(&store.content_type)?;
        custom_xml::validate_payload(&store.xml)?;
        let props = CustomXmlProps {
            id: store.id,
            schemas: store.schemas,
        };
        custom_xml::validate_props(&props)?;
        let source_part = self.opc.main_document_part()?.partname().clone();
        let source = self.opc.get_part(&source_part)?;
        let rel_id = (1usize..=MAX_ITEMS + 1)
            .map(|number| format!("rIdCustomXml{number}"))
            .find(|id| source.rels().get(id).is_none())
            .ok_or_else(|| {
                Error::InvalidFormat("Custom XML relationship ID space is exhausted".into())
            })?;
        let mut part_names = None;
        for number in 1usize..=MAX_ITEMS + 1 {
            let data =
                PackURI::new(format!("/customXml/item{number}.xml")).map_err(Error::InvalidUri)?;
            let props = PackURI::new(format!("/customXml/itemProps{number}.xml"))
                .map_err(Error::InvalidUri)?;
            let conflict = self.opc.iter_parts().any(|part| {
                part.partname().as_str().eq_ignore_ascii_case(data.as_str())
                    || part
                        .partname()
                        .as_str()
                        .eq_ignore_ascii_case(props.as_str())
            });
            if !conflict {
                part_names = Some((data, props));
                break;
            }
        }
        let (data_part, props_part) = part_names.ok_or_else(|| {
            Error::InvalidFormat("Custom XML part-name space is exhausted".into())
        })?;
        custom_xml::add(
            &mut self.opc,
            NewCustomXmlItem {
                source: source_part,
                rel_id,
                part: data_part.clone(),
                content_type: store.content_type,
                xml: store.xml,
                props: Some(NewCustomXmlProps {
                    part: props_part,
                    rel_id: "rIdProps1".to_string(),
                    value: props,
                }),
                conformance: store.conformance,
            },
        )?;
        custom_xml::discover(&self.opc)?
            .into_iter()
            .find(|item| item.part() == &data_part)
            .ok_or_else(|| {
                Error::InvalidFormat("new Custom XML data store was not discoverable".into())
            })
    }

    /// Replace only the inert XML payload of a data store.
    pub fn set_custom_xml(&mut self, id: &str, xml: Vec<u8>) -> Result<()> {
        custom_xml::validate_payload(&xml)?;
        let item = self
            .custom_xml_by_id(id)?
            .ok_or_else(|| Error::PartNotFound(format!("Custom XML itemID '{id}'")))?;
        self.opc.get_part_mut(item.part())?.set_blob(xml);
        self.opc.unsign();
        Ok(())
    }

    /// Replace payload, content type, schema references, and canonical properties.
    pub fn replace_custom_xml(&mut self, id: &str, replacement: NewStore) -> Result<()> {
        custom_xml::validate_content_type(&replacement.content_type)?;
        custom_xml::validate_payload(&replacement.xml)?;
        if !replacement.id.eq_ignore_ascii_case(id) {
            return Err(Error::InvalidFormat(
                "replacement itemID must identify the existing data store".into(),
            ));
        }
        let props = CustomXmlProps {
            id: replacement.id,
            schemas: replacement.schemas,
        };
        let props_xml = custom_xml::write_props(&props, replacement.conformance)?;
        let item = self
            .custom_xml_by_id(id)?
            .ok_or_else(|| Error::PartNotFound(format!("Custom XML itemID '{id}'")))?;
        let props_part = item.props_part().cloned().ok_or_else(|| {
            Error::InvalidFormat("Custom XML data store has no properties part".into())
        })?;
        let existing_relationships = self
            .opc
            .get_part(item.part())?
            .rels()
            .iter()
            .map(|relationship| {
                (
                    relationship.reltype().to_string(),
                    relationship.target_ref().to_string(),
                    relationship.r_id().to_string(),
                    relationship.is_external(),
                )
            })
            .collect::<Vec<_>>();
        let mut data_part = BlobPart::new(
            item.part().clone(),
            replacement.content_type,
            replacement.xml,
        );
        for (reltype, target, id, external) in existing_relationships {
            data_part
                .rels_mut()
                .add_relationship(reltype, target, id, external);
        }
        self.opc.add_part(Box::new(data_part));
        self.opc.get_part_mut(&props_part)?.set_blob(props_xml);
        self.opc.unsign();
        Ok(())
    }

    /// Remove a data store unless an SDT still binds to its item GUID.
    pub fn remove_custom_xml(&mut self, id: &str) -> Result<bool> {
        let items = custom_xml::discover(&self.opc)?;
        let matching = items
            .iter()
            .filter(|item| {
                item.props()
                    .is_some_and(|props| props.id.eq_ignore_ascii_case(id))
            })
            .collect::<Vec<_>>();
        let Some(first) = matching.first() else {
            return Ok(false);
        };
        if self
            .custom_xml_bindings()?
            .iter()
            .any(|binding| binding.store_id.eq_ignore_ascii_case(id))
        {
            return Err(Error::InvalidFormat(format!(
                "Custom XML itemID '{id}' is still referenced by a content control"
            )));
        }
        for item in &matching {
            self.opc
                .get_part_mut(item.source())?
                .rels_mut()
                .remove(item.rel_id());
        }
        let data_part = first.part().clone();
        let props_part = first.props_part().cloned();
        if !self.part_is_referenced(&data_part) {
            self.opc.remove_part(&data_part);
            if let Some(props_part) = props_part
                && !self.part_is_referenced(&props_part)
            {
                self.opc.remove_part(&props_part);
            }
        }
        self.opc.unsign();
        Ok(true)
    }

    /// Locate the document's bibliography source store, if one exists.
    ///
    /// Returns the Custom XML item GUID and the store payload. Word keeps a
    /// single current source list; when several stores exist the first in
    /// package order is used and the rest are left untouched.
    fn bibliography_store_item(&self) -> Result<Option<(String, Vec<u8>)>> {
        let stores = self.bibliography_source_stores()?;
        let Some(store) = stores.first() else {
            return Ok(None);
        };
        let item_id = store
            .data_store_item_id()
            .ok_or_else(|| {
                Error::InvalidFormat("bibliography source store has no Custom XML item GUID".into())
            })?
            .to_owned();
        let item = self.custom_xml_by_id(&item_id)?.ok_or_else(|| {
            Error::PartNotFound(format!("bibliography source store item '{item_id}'"))
        })?;
        Ok(Some((item_id, item.xml().to_vec())))
    }

    /// Add a typed bibliography source to the document's source store.
    ///
    /// When no store exists, one is created as a Custom XML data store with
    /// the bibliography namespace registered. Otherwise the source is
    /// appended in place, preserving untouched entries, style metadata, and
    /// the store's relationship/content-type graph. Duplicate tags are
    /// rejected. Returns the Custom XML item GUID of the store.
    pub fn add_bibliography_source(
        &mut self,
        source: crate::bibliography::BibliographySourceBuilder,
    ) -> Result<String> {
        if let Some((item_id, xml)) = self.bibliography_store_item()? {
            let updated = crate::bibliography::add_source_xml(&xml, &source)?;
            self.set_custom_xml(&item_id, updated.into_bytes())?;
            // Re-validate the mutated store through the read side.
            self.bibliography_source_stores()?;
            Ok(item_id)
        } else {
            let xml = crate::bibliography::new_store_xml(&[source])?;
            let item = self.add_custom_xml(NewStore {
                xml: xml.into_bytes(),
                content_type: "application/xml".to_string(),
                id: crate::bibliography::DEFAULT_STORE_ITEM_ID.to_string(),
                schemas: vec![crate::OOXML_BIBLIOGRAPHY_NAMESPACE.to_string()],
                conformance: custom_xml::Conformance::Transitional,
            })?;
            item.props().map(|props| props.id.clone()).ok_or_else(|| {
                Error::InvalidFormat("new bibliography Custom XML store has no item GUID".into())
            })
        }
    }

    /// Remove the bibliography source with the given tag from the source
    /// store. Returns whether a source was removed.
    pub fn remove_bibliography_source(&mut self, tag: &str) -> Result<bool> {
        let Some((item_id, xml)) = self.bibliography_store_item()? else {
            return Ok(false);
        };
        let (updated, removed) = crate::bibliography::remove_source_xml(&xml, tag)?;
        if removed {
            self.set_custom_xml(&item_id, updated.into_bytes())?;
            // Re-validate the mutated store through the read side.
            self.bibliography_source_stores()?;
        }
        Ok(removed)
    }

    /// Replace the bibliography source with the given tag, preserving entry
    /// order and all untouched entries. Fails when the tag does not exist.
    pub fn replace_bibliography_source(
        &mut self,
        tag: &str,
        source: crate::bibliography::BibliographySourceBuilder,
    ) -> Result<()> {
        let Some((item_id, xml)) = self.bibliography_store_item()? else {
            return Err(Error::PartNotFound(
                "no bibliography source store exists".to_string(),
            ));
        };
        let updated = crate::bibliography::replace_source_xml(&xml, tag, &source)?;
        self.set_custom_xml(&item_id, updated.into_bytes())?;
        // Re-validate the mutated store through the read side.
        self.bibliography_source_stores()?;
        Ok(())
    }

    /// Reorder main-document data-store relationships by item GUID.
    pub fn order_custom_xml(&mut self, ordered_ids: &[String]) -> Result<()> {
        let source_part = self.opc.main_document_part()?.partname().clone();
        let items = custom_xml::discover(&self.opc)?
            .into_iter()
            .filter(|item| item.source() == &source_part)
            .collect::<Vec<_>>();
        if items.len() != ordered_ids.len() {
            return Err(Error::InvalidFormat(
                "reorder list must contain every main-document Custom XML item".into(),
            ));
        }
        let mut by_id = std::collections::HashMap::new();
        for item in &items {
            let id = item
                .props()
                .ok_or_else(|| {
                    Error::InvalidFormat("Custom XML item has no datastore itemID".into())
                })?
                .id
                .to_ascii_lowercase();
            if by_id.insert(id, item).is_some() {
                return Err(Error::InvalidFormat(
                    "main-document Custom XML items are not uniquely reorderable".into(),
                ));
            }
        }
        let mut ordered = Vec::with_capacity(items.len());
        let mut seen = std::collections::HashSet::new();
        for id in ordered_ids {
            let key = id.to_ascii_lowercase();
            if !seen.insert(key.clone()) {
                return Err(Error::InvalidFormat("duplicate reorder itemID".into()));
            }
            let item = *by_id
                .get(&key)
                .ok_or_else(|| Error::InvalidFormat(format!("unknown reorder itemID '{id}'")))?;
            let reltype = self
                .opc
                .get_part(&source_part)?
                .rels()
                .get(item.rel_id())
                .ok_or_else(|| {
                    Error::InvalidRelationship(format!(
                        "Custom XML relationship '{}' disappeared during reorder",
                        item.rel_id()
                    ))
                })?
                .reltype()
                .to_string();
            ordered.push((item, reltype));
        }
        let source = self.opc.get_part(&source_part)?;
        let reserved = source
            .rels()
            .iter()
            .filter(|relationship| {
                !items
                    .iter()
                    .any(|item| item.rel_id() == relationship.r_id())
            })
            .map(|relationship| relationship.r_id().to_string())
            .collect::<std::collections::HashSet<_>>();
        let ids = (1usize..=MAX_ITEMS + 1)
            .filter_map(|batch| {
                let candidates = (0..ordered.len())
                    .map(|index| format!("rIdCustomXmlOrder{batch:04}_{index:06}"))
                    .collect::<Vec<_>>();
                candidates
                    .iter()
                    .all(|id| !reserved.contains(id))
                    .then_some(candidates)
            })
            .next()
            .ok_or_else(|| {
                Error::InvalidFormat("Custom XML reorder relationship ID space is exhausted".into())
            })?;
        let source = self.opc.get_part_mut(&source_part)?;
        let source_base_uri = source.partname().base_uri().to_string();
        for item in &items {
            source.rels_mut().remove(item.rel_id());
        }
        for ((item, reltype), id) in ordered.into_iter().zip(ids) {
            source.rels_mut().add_relationship(
                reltype,
                item.part().relative_ref(&source_base_uri),
                id,
                false,
            );
        }
        self.opc.unsign();
        Ok(())
    }

    /// Collect and lexically validate SDT bindings from every permitted Word container.
    pub fn custom_xml_bindings(&self) -> Result<Vec<Binding>> {
        let permitted = [
            ct::WML_DOCUMENT_MAIN,
            ct::WML_DOCUMENT_GLOSSARY,
            ct::WML_HEADER,
            ct::WML_FOOTER,
            ct::WML_FOOTNOTES,
            ct::WML_ENDNOTES,
        ];
        let mut bindings = Vec::new();
        for part in self
            .opc
            .iter_parts()
            .filter(|part| permitted.contains(&part.content_type()))
        {
            for control in ContentControl::extract_from_document(part.blob())? {
                control.validate_data_binding()?;
                if let (Some(xpath), Some(store_item_id)) = (
                    control.data_binding_xpath(),
                    control.data_binding_store_item_id(),
                ) {
                    bindings.push(Binding {
                        source: part.partname().clone(),
                        control_id: control.id(),
                        xpath: xpath.to_string(),
                        store_id: store_item_id.to_string(),
                        prefixes: control.data_binding_prefix_mappings().map(str::to_string),
                    });
                }
            }
        }
        bindings.sort_unstable_by(|left, right| {
            left.source
                .as_str()
                .cmp(right.source.as_str())
                .then_with(|| left.control_id.cmp(&right.control_id))
        });
        Ok(bindings)
    }

    /// Validate that every permitted SDT binding resolves to a datastore item GUID.
    pub fn validate_custom_xml_bindings(&self) -> Result<()> {
        let item_ids = custom_xml::discover(&self.opc)?
            .into_iter()
            .filter_map(|item| item.props().map(|props| props.id.to_ascii_lowercase()))
            .collect::<std::collections::HashSet<_>>();
        for binding in self.custom_xml_bindings()? {
            if !item_ids.contains(&binding.store_id.to_ascii_lowercase()) {
                return Err(Error::InvalidFormat(format!(
                    "content control {} in '{}' references missing Custom XML itemID '{}'",
                    binding.control_id,
                    binding.source.as_str(),
                    binding.store_id
                )));
            }
        }
        Ok(())
    }

    fn part_is_referenced(&self, target: &PackURI) -> bool {
        self.opc.rels().iter().any(|relationship| {
            !relationship.is_external()
                && relationship
                    .target_partname()
                    .is_ok_and(|part| &part == target)
        }) || self.opc.iter_parts().any(|part| {
            part.rels().iter().any(|relationship| {
                !relationship.is_external()
                    && relationship
                        .target_partname()
                        .is_ok_and(|part| &part == target)
            })
        })
    }

    /// Return the validated inert mail-merge settings, if configured.
    pub fn mail_merge_settings(&self) -> Result<Option<MailMergeSettings>> {
        let snapshot = self.settings_part_snapshot()?;
        let part = settings_part_from_snapshot(&snapshot, snapshot.xml.clone(), None);
        Ok(DocumentSettings::extract_from_part(&part)?
            .mail_merge()
            .cloned())
    }

    /// Resolve a mail-merge relationship without opening or fetching its target.
    pub fn mail_merge_target(&self, relationship_id: &str) -> Result<Target> {
        let snapshot = self.settings_part_snapshot()?;
        let part = self.opc.get_part(&snapshot.target)?;
        let relationship = part.rels().get(relationship_id).ok_or_else(|| {
            Error::InvalidFormat(format!(
                "mail-merge relationship '{relationship_id}' is missing"
            ))
        })?;
        if !is_mail_merge_relationship_type(relationship.reltype()) {
            return Err(Error::InvalidFormat(format!(
                "relationship '{relationship_id}' is not a mail-merge source"
            )));
        }
        if relationship.is_external() {
            return Ok(Target::External(relationship.target_ref().to_string()));
        }
        let target = relationship.target_partname()?;
        let target_part = self.opc.get_part(&target)?;
        Ok(Target::Internal {
            part_name: target,
            bytes: target_part.blob().to_vec(),
            content_type: target_part.content_type().to_string(),
        })
    }

    /// Set or replace the complete mail-merge graph atomically.
    pub fn set_mail_merge(
        &mut self,
        mut settings: MailMergeSettings,
        data_source: Option<Source>,
        header_source: Option<Source>,
        recipients: Option<Recipients>,
        conformance: mail_merge::Conformance,
    ) -> Result<()> {
        let snapshot = self.settings_part_snapshot()?;
        let original = settings_part_from_snapshot(&snapshot, snapshot.xml.clone(), None);
        DocumentSettings::extract_from_part(&original)?;
        let old_targets = self.mail_merge_internal_targets(&snapshot)?;
        let mut used_ids = snapshot
            .relationships
            .iter()
            .filter(|relationship| !is_mail_merge_relationship_type(&relationship.reltype))
            .map(|relationship| relationship.id.clone())
            .collect::<std::collections::HashSet<_>>();
        let mut staged_parts = Vec::new();
        let mut staged_relationships = Vec::new();

        let data_id = if let Some(source) = data_source {
            let (relationship, part) = self.stage_mail_merge_source(
                source,
                "Data",
                "mailMergeSource",
                conformance,
                &mut used_ids,
            )?;
            let id = relationship.id.clone();
            staged_relationships.push(relationship);
            if let Some(part) = part {
                staged_parts.push(part);
            }
            Some(id)
        } else {
            None
        };
        let header_id = if let Some(source) = header_source {
            let (relationship, part) = self.stage_mail_merge_source(
                source,
                "Header",
                "mailMergeHeaderSource",
                conformance,
                &mut used_ids,
            )?;
            let id = relationship.id.clone();
            staged_relationships.push(relationship);
            if let Some(part) = part {
                staged_parts.push(part);
            }
            Some(id)
        } else {
            None
        };
        let recipient_id = if let Some(recipients) = recipients {
            let xml = recipients
                .to_xml(conformance)
                .map_err(map_docx_error)?
                .into_bytes();
            let id = allocate_mail_merge_relationship_id("Recipients", &mut used_ids)?;
            let uri = self.allocate_mail_merge_part_name("recipientData", "xml")?;
            let target = uri.relative_ref(snapshot.target.base_uri());
            staged_parts.push(BlobPart::new(
                uri,
                Recipients::content_type().to_string(),
                xml,
            ));
            staged_relationships.push(StoredRelationship {
                reltype: mail_merge_relationship_type(conformance, "recipientData"),
                target,
                id: id.clone(),
                external: false,
            });
            Some(id)
        } else {
            None
        };
        settings.assign_package_relationships(data_id, header_id, recipient_id);
        let patched = patch_mail_merge(&snapshot.xml, Some(&settings), conformance)?;
        let mut replacement = settings_part_from_snapshot(&snapshot, patched, None);
        let old_ids = replacement
            .rels()
            .iter()
            .filter(|relationship| is_mail_merge_relationship_type(relationship.reltype()))
            .map(|relationship| relationship.r_id().to_string())
            .collect::<Vec<_>>();
        for id in old_ids {
            replacement.rels_mut().remove(&id);
        }
        for relationship in staged_relationships {
            replacement.rels_mut().add_relationship(
                relationship.reltype,
                relationship.target,
                relationship.id,
                relationship.external,
            );
        }
        DocumentSettings::extract_from_part(&replacement)?;

        let mut installed = Vec::new();
        for part in staged_parts {
            let name = part.partname().clone();
            if let Err(error) = self.opc.try_add_part(Box::new(part)) {
                for installed_name in installed {
                    self.opc.remove_part(&installed_name);
                }
                return Err(error.into());
            }
            installed.push(name);
        }
        if let Err(error) = self.commit_settings_part(&snapshot, replacement) {
            for installed_name in installed {
                self.opc.remove_part(&installed_name);
            }
            return Err(error);
        }
        for old_target in old_targets {
            if !self.part_is_referenced(&old_target) {
                self.opc.remove_part(&old_target);
            }
        }
        self.opc.unsign();
        Ok(())
    }

    /// Update settings and sources using the same atomic replacement semantics.
    pub fn update_mail_merge(
        &mut self,
        settings: MailMergeSettings,
        data_source: Option<Source>,
        header_source: Option<Source>,
        recipients: Option<Recipients>,
        conformance: mail_merge::Conformance,
    ) -> Result<()> {
        self.set_mail_merge(
            settings,
            data_source,
            header_source,
            recipients,
            conformance,
        )
    }

    /// Replace recipient inclusion flags while retaining inert source targets and settings.
    pub fn update_mail_merge_recipients(
        &mut self,
        recipients: Recipients,
        conformance: mail_merge::Conformance,
    ) -> Result<()> {
        let settings = self
            .mail_merge_settings()?
            .ok_or_else(|| Error::InvalidFormat("document has no mail-merge settings".into()))?;
        let data_source = settings
            .data_source_relationship_id()
            .map(|id| self.mail_merge_target(id).map(mail_merge_target_as_source))
            .transpose()?;
        let header_source = settings
            .header_source_relationship_id()
            .map(|id| self.mail_merge_target(id).map(mail_merge_target_as_source))
            .transpose()?;
        self.set_mail_merge(
            settings,
            data_source,
            header_source,
            Some(recipients),
            conformance,
        )
    }

    /// Clear mail-merge XML, relationships, and unreachable owned targets.
    pub fn clear_mail_merge(&mut self) -> Result<bool> {
        let snapshot = self.settings_part_snapshot()?;
        let original = settings_part_from_snapshot(&snapshot, snapshot.xml.clone(), None);
        if DocumentSettings::extract_from_part(&original)?
            .mail_merge()
            .is_none()
        {
            return Ok(false);
        }
        let old_targets = self.mail_merge_internal_targets(&snapshot)?;
        let patched = patch_mail_merge(&snapshot.xml, None, mail_merge::Conformance::Transitional)?;
        let mut replacement = settings_part_from_snapshot(&snapshot, patched, None);
        let ids = replacement
            .rels()
            .iter()
            .filter(|relationship| is_mail_merge_relationship_type(relationship.reltype()))
            .map(|relationship| relationship.r_id().to_string())
            .collect::<Vec<_>>();
        for id in ids {
            replacement.rels_mut().remove(&id);
        }
        DocumentSettings::extract_from_part(&replacement)?;
        self.commit_settings_part(&snapshot, replacement)?;
        for old_target in old_targets {
            if !self.part_is_referenced(&old_target) {
                self.opc.remove_part(&old_target);
            }
        }
        self.opc.unsign();
        Ok(true)
    }

    fn stage_mail_merge_source(
        &self,
        source: Source,
        label: &str,
        relationship_suffix: &str,
        conformance: mail_merge::Conformance,
        used_ids: &mut std::collections::HashSet<String>,
    ) -> Result<(StoredRelationship, Option<BlobPart>)> {
        let id = allocate_mail_merge_relationship_id(label, used_ids)?;
        let settings_target = self.settings_part_snapshot()?.target;
        match source {
            Source::External(uri) => {
                validate_mail_merge_external_uri(&uri)?;
                Ok((
                    StoredRelationship {
                        reltype: mail_merge_relationship_type(conformance, relationship_suffix),
                        target: uri,
                        id,
                        external: true,
                    },
                    None,
                ))
            },
            Source::Internal {
                bytes,
                content_type,
                extension,
            } => {
                validate_mail_merge_internal_source(&bytes, &content_type, &extension)?;
                let uri = self.allocate_mail_merge_part_name(label, &extension)?;
                let target = uri.relative_ref(settings_target.base_uri());
                let part = BlobPart::new(uri, content_type, bytes);
                Ok((
                    StoredRelationship {
                        reltype: mail_merge_relationship_type(conformance, relationship_suffix),
                        target,
                        id,
                        external: false,
                    },
                    Some(part),
                ))
            },
        }
    }

    fn allocate_mail_merge_part_name(&self, stem: &str, extension: &str) -> Result<PackURI> {
        for number in 1usize.. {
            let candidate = PackURI::new(format!("/word/mailMerge/{stem}{number}.{extension}"))
                .map_err(Error::InvalidUri)?;
            if self.opc.iter_parts().all(|part| {
                !part
                    .partname()
                    .as_str()
                    .eq_ignore_ascii_case(candidate.as_str())
            }) {
                return Ok(candidate);
            }
        }
        unreachable!("the mail-merge part-name space is unbounded")
    }

    fn mail_merge_internal_targets(&self, snapshot: &SettingsPartSnapshot) -> Result<Vec<PackURI>> {
        let Ok(part) = self.opc.get_part(&snapshot.target) else {
            return Ok(Vec::new());
        };
        part.rels()
            .iter()
            .filter(|relationship| {
                is_mail_merge_relationship_type(relationship.reltype())
                    && !relationship.is_external()
            })
            .map(|relationship| relationship.target_partname().map_err(Into::into))
            .collect()
    }

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

    /// Update the footnotes.xml part with new content.
    #[allow(unused)] // Kept for future use
    fn update_footnotes_part(&mut self, xml: String) -> Result<()> {
        use litchi_opc::constants::content_type as ct;
        use litchi_opc::constants::relationship_type as rt;
        use litchi_opc::part::BlobPart;

        let footnotes_uri = PackURI::new("/word/footnotes.xml")
            .map_err(|e| Error::InvalidUri(format!("footnotes URI: {}", e)))?;

        let content_type = ct::WML_FOOTNOTES.to_string();
        let footnotes_part = BlobPart::new(footnotes_uri.clone(), content_type, xml.into_bytes());

        // Add the footnotes part
        self.opc.add_part(Box::new(footnotes_part));

        // Create relationship from document to footnotes (use relative path)
        let doc_uri = PackURI::new("/word/document.xml")
            .map_err(|e| Error::InvalidUri(format!("document URI: {}", e)))?;

        if let Ok(doc_part) = self.opc.get_part_mut(&doc_uri) {
            let _ = doc_part.relate_to("footnotes.xml", rt::FOOTNOTES);
        }

        Ok(())
    }

    /// Update the endnotes.xml part with new content.
    #[allow(unused)] // Kept for future use
    fn update_endnotes_part(&mut self, xml: String) -> Result<()> {
        use litchi_opc::constants::content_type as ct;
        use litchi_opc::constants::relationship_type as rt;
        use litchi_opc::part::BlobPart;

        let endnotes_uri = PackURI::new("/word/endnotes.xml")
            .map_err(|e| Error::InvalidUri(format!("endnotes URI: {}", e)))?;

        let content_type = ct::WML_ENDNOTES.to_string();
        let endnotes_part = BlobPart::new(endnotes_uri.clone(), content_type, xml.into_bytes());

        // Add the endnotes part
        self.opc.add_part(Box::new(endnotes_part));

        // Create relationship from document to endnotes (use relative path)
        let doc_uri = PackURI::new("/word/document.xml")
            .map_err(|e| Error::InvalidUri(format!("document URI: {}", e)))?;

        if let Ok(doc_part) = self.opc.get_part_mut(&doc_uri) {
            let _ = doc_part.relate_to("endnotes.xml", rt::ENDNOTES);
        }

        Ok(())
    }

    /// Update or create the comments part with the given XML content.
    pub(super) fn update_comments_part(&mut self, xml: String) -> Result<()> {
        use litchi_opc::constants::content_type as ct;
        use litchi_opc::constants::relationship_type as rt;
        use litchi_opc::part::BlobPart;

        let comments_uri = PackURI::new("/word/comments.xml")
            .map_err(|e| Error::InvalidUri(format!("comments URI: {}", e)))?;

        let content_type = ct::WML_COMMENTS.to_string();
        let comments_part = BlobPart::new(comments_uri.clone(), content_type, xml.into_bytes());

        // Add the comments part
        self.opc.add_part(Box::new(comments_part));

        // Create relationship from document to comments
        let doc_uri = PackURI::new("/word/document.xml")
            .map_err(|e| Error::InvalidUri(format!("document URI: {}", e)))?;

        if let Ok(doc_part) = self.opc.get_part_mut(&doc_uri) {
            let _ = doc_part.relate_to("/word/comments.xml", rt::COMMENTS);
        }

        Ok(())
    }

    /// Update the settings.xml part with new content.
    pub(super) fn update_settings_part(&mut self, xml: Vec<u8>) -> Result<()> {
        let snapshot = self.settings_part_snapshot()?;
        let part = settings_part_from_snapshot(&snapshot, xml, None);
        DocumentSettings::extract_from_part(&part)?;
        self.commit_settings_part(&snapshot, part)
    }

    fn settings_part_snapshot(&self) -> Result<SettingsPartSnapshot> {
        use litchi_opc::constants::relationship_type as rt;

        const STRICT_SETTINGS_RELATIONSHIP: &str =
            "http://purl.oclc.org/ooxml/officeDocument/relationships/settings";
        let document = self.opc.main_document_part()?;
        let document_uri = document.partname().clone();
        let mut matches = document.rels().iter().filter(|relationship| {
            matches!(
                relationship.reltype(),
                rt::SETTINGS | STRICT_SETTINGS_RELATIONSHIP
            )
        });
        let relationship = matches.next();
        if matches.next().is_some() {
            return Err(Error::InvalidFormat(
                "document has multiple settings relationships".into(),
            ));
        }
        let (target, relationship_exists) = match relationship {
            Some(relationship) if relationship.is_external() => {
                return Err(Error::InvalidFormat(
                    "settings relationship cannot be external".into(),
                ));
            },
            Some(relationship) => (relationship.target_partname()?, true),
            None => (
                PackURI::new("/word/settings.xml")
                    .map_err(|error| Error::InvalidUri(error.to_string()))?,
                false,
            ),
        };

        let (content_type, xml, relationships) = match self.opc.get_part(&target) {
            Ok(part) => {
                if part.content_type() != ct::WML_SETTINGS {
                    return Err(Error::InvalidFormat(format!(
                        "settings part has content type {:?}, expected {:?}",
                        part.content_type(),
                        ct::WML_SETTINGS
                    )));
                }
                (
                    part.content_type().to_owned(),
                    part.blob().to_vec(),
                    part.rels()
                        .iter()
                        .map(|relationship| StoredRelationship {
                            reltype: relationship.reltype().to_owned(),
                            target: relationship.target_ref().to_owned(),
                            id: relationship.r_id().to_owned(),
                            external: relationship.is_external(),
                        })
                        .collect(),
                )
            },
            Err(_) if relationship_exists => {
                return Err(Error::PartNotFound(format!("settings part {target}")));
            },
            Err(_) => (
                ct::WML_SETTINGS.to_owned(),
                crate::template::default_settings_xml().as_bytes().to_vec(),
                Vec::new(),
            ),
        };
        Ok(SettingsPartSnapshot {
            document_uri,
            target,
            relationship_exists,
            content_type,
            xml,
            relationships,
        })
    }

    fn commit_settings_part(
        &mut self,
        snapshot: &SettingsPartSnapshot,
        part: BlobPart,
    ) -> Result<()> {
        use litchi_opc::constants::relationship_type as rt;

        if !snapshot.relationship_exists {
            // Acquire the only fallible mutable reference before changing package state.
            self.opc
                .get_part_mut(&snapshot.document_uri)?
                .relate_to("settings.xml", rt::SETTINGS);
        }
        self.opc.add_part(Box::new(part));
        Ok(())
    }

    pub(super) fn update_theme_part(&mut self, xml: String) -> Result<()> {
        use litchi_opc::part::BlobPart;

        let theme_uri = PackURI::new("/word/theme/theme1.xml")
            .map_err(|e| Error::InvalidUri(format!("theme URI: {}", e)))?;

        let content_type = "application/vnd.openxmlformats-officedocument.theme+xml".to_string();
        let theme_part = BlobPart::new(theme_uri.clone(), content_type, xml.into_bytes());

        // Add/replace the theme part
        self.opc.add_part(Box::new(theme_part));

        // Add relationship from document to theme if not exists
        let doc_uri = PackURI::new("/word/document.xml")
            .map_err(|e| Error::InvalidUri(format!("document URI: {}", e)))?;

        if let Ok(doc_part) = self.opc.get_part_mut(&doc_uri) {
            // Check if theme relationship already exists
            let has_theme_rel = doc_part.rels().iter().any(|rel| {
                rel.reltype()
                    == "http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme"
            });

            if !has_theme_rel {
                doc_part.relate_to(
                    "theme/theme1.xml",
                    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme",
                );
            }
        }

        Ok(())
    }

    #[allow(unused)] // Kept for future use
    fn update_watermark_headers(&mut self, mutable_doc: &MutableDocument) -> Result<()> {
        use litchi_opc::constants::content_type as ct;
        use litchi_opc::constants::relationship_type as rt;
        use litchi_opc::part::BlobPart;

        // Get watermark if present
        // Access watermark through a temporary reference
        let has_watermark = mutable_doc.has_watermark();
        if !has_watermark {
            return Ok(());
        }

        // Get user header content if it exists
        let user_header_content = if mutable_doc.has_header() {
            mutable_doc.generate_header_xml()?
        } else {
            None
        };

        // Create three headers (default, first, even) with watermark
        let header_types = [
            ("/word/header1.xml", "default"),
            ("/word/header2.xml", "first"),
            ("/word/header3.xml", "even"),
        ];

        let doc_uri = PackURI::new("/word/document.xml")
            .map_err(|e| Error::InvalidUri(format!("document URI: {}", e)))?;

        for (idx, (header_path, _header_type)) in header_types.iter().enumerate() {
            // Generate watermark XML for this header - need to get watermark again each iteration
            let watermark_xml = if let Some(wm) = mutable_doc.watermark.as_ref() {
                wm.to_header_xml((idx + 1) as u32)?
            } else {
                continue;
            };

            // Merge user header content with watermark for the default header
            let header_xml = if idx == 0
                && let Some(ref user_content) = user_header_content
            {
                // Extract user paragraphs from the <w:hdr>...</w:hdr> wrapper
                let user_paragraphs = if let Some(start) = user_content.find("<w:p") {
                    if let Some(end) = user_content.rfind("</w:hdr>") {
                        &user_content[start..end]
                    } else {
                        ""
                    }
                } else {
                    ""
                };

                // Combine watermark and user content
                format!(
                    r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:hdr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:v="urn:schemas-microsoft-com:vml" xmlns:o="urn:schemas-microsoft-com:office:office">{}{}</w:hdr>"#,
                    watermark_xml, user_paragraphs
                )
            } else {
                // Just watermark for first and even headers
                format!(
                    r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:hdr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:v="urn:schemas-microsoft-com:vml" xmlns:o="urn:schemas-microsoft-com:office:office">{}</w:hdr>"#,
                    watermark_xml
                )
            };

            let header_uri = PackURI::new(*header_path)
                .map_err(|e| Error::InvalidUri(format!("header URI: {}", e)))?;

            let header_part = BlobPart::new(
                header_uri,
                ct::WML_HEADER.to_string(),
                header_xml.into_bytes(),
            );

            self.opc.add_part(Box::new(header_part));

            // Add relationship from document to header (use relative path)
            // Extract filename from the absolute path (e.g., "/word/header1.xml" -> "header1.xml")
            let header_filename = header_path.rsplit('/').next().unwrap_or(header_path);
            if let Ok(doc_part) = self.opc.get_part_mut(&doc_uri) {
                doc_part.relate_to(header_filename, rt::HEADER);
            }
        }

        Ok(())
    }
}

fn settings_part_from_snapshot(
    snapshot: &SettingsPartSnapshot,
    xml: Vec<u8>,
    omitted_relationship_id: Option<&str>,
) -> BlobPart {
    use litchi_opc::part::{BlobPart, Part};

    let mut part = BlobPart::new(snapshot.target.clone(), snapshot.content_type.clone(), xml);
    for relationship in &snapshot.relationships {
        if omitted_relationship_id == Some(relationship.id.as_str()) {
            continue;
        }
        part.rels_mut().add_relationship(
            relationship.reltype.clone(),
            relationship.target.clone(),
            relationship.id.clone(),
            relationship.external,
        );
    }
    part
}

fn mail_merge_relationship_type(conformance: mail_merge::Conformance, suffix: &str) -> String {
    let base = match conformance {
        mail_merge::Conformance::Transitional => {
            "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
        },
        mail_merge::Conformance::Strict => {
            "http://purl.oclc.org/ooxml/officeDocument/relationships"
        },
    };
    format!("{base}/{suffix}")
}

fn allocate_mail_merge_relationship_id(
    label: &str,
    used: &mut std::collections::HashSet<String>,
) -> Result<String> {
    (1usize..=MAX_MAIL_MERGE_RELATIONSHIPS)
        .map(|number| format!("rIdMailMerge{label}{number}"))
        .find(|id| used.insert(id.clone()))
        .ok_or_else(|| Error::InvalidFormat("mail-merge relationship ID space is exhausted".into()))
}

fn validate_mail_merge_external_uri(uri: &str) -> Result<()> {
    if uri.is_empty() || uri.len() > 32 * 1024 || uri.chars().any(char::is_control) {
        return Err(Error::InvalidFormat(
            "mail-merge external target is empty or exceeds URI limits".into(),
        ));
    }
    Ok(())
}

fn validate_mail_merge_internal_source(
    bytes: &[u8],
    content_type: &str,
    extension: &str,
) -> Result<()> {
    if bytes.len() > 128 * 1024 * 1024 {
        return Err(Error::InvalidFormat(
            "mail-merge source exceeds the 128 MiB authoring limit".into(),
        ));
    }
    if content_type.is_empty()
        || content_type.len() > 1024
        || content_type.chars().any(char::is_control)
    {
        return Err(Error::InvalidFormat(
            "mail-merge source content type is invalid".into(),
        ));
    }
    if extension.is_empty()
        || extension.len() > 16
        || !extension.bytes().all(|byte| byte.is_ascii_alphanumeric())
    {
        return Err(Error::InvalidFormat(
            "mail-merge source extension is invalid".into(),
        ));
    }
    Ok(())
}

fn mail_merge_target_as_source(target: Target) -> Source {
    match target {
        Target::External(uri) => Source::External(uri),
        Target::Internal {
            part_name,
            bytes,
            content_type,
        } => {
            let extension = part_name
                .as_str()
                .rsplit_once('.')
                .map(|(_, extension)| extension)
                .filter(|extension| {
                    !extension.is_empty()
                        && extension.len() <= 16
                        && extension.bytes().all(|byte| byte.is_ascii_alphanumeric())
                })
                .unwrap_or("bin")
                .to_string();
            Source::Internal {
                bytes,
                content_type,
                extension,
            }
        },
    }
}
