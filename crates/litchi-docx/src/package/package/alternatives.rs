//! Package-backed alternative-format anchors and relationship lifecycle.

use super::super::model::*;

impl Package {
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
}
