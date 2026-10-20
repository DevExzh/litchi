#![expect(
    clippy::shadow_reuse,
    reason = "parser bindings are intentionally refined after validation"
)]
#![expect(
    clippy::shadow_unrelated,
    reason = "local parser names mirror the OOXML role currently being decoded"
)]
//! Package-backed alternative-format anchors and relationship lifecycle.

use super::super::model::{
    BlobPart, Chunk, Conformance, DraftDocument, Error, Import, MAX_CHUNKS, PackURI, Package, Rel,
    Result, TargetMode, is_relationship,
};
use crate::alt::LimitKind;

/// The body whose anchors an alternative-format verb edits: a builder's
/// draft document, or an opened package's main-document XML
/// ([`crate::alt::XmlBody`], record 0355a).
pub(crate) trait AltBody {
    /// The active body anchors, in document order.
    fn alts(&self) -> Result<Vec<Chunk>>;
    /// Insert `chunk` before anchor `index`, or after the body content.
    fn insert_alt(&mut self, index: usize, chunk: Chunk, namespace: Conformance) -> Result<()>;
    /// Replace anchor `index` with `chunk`; returns the old anchor.
    fn replace_alt(&mut self, index: usize, chunk: Chunk, namespace: Conformance) -> Result<Chunk>;
    /// Remove anchor `index`; returns it.
    fn remove_alt(&mut self, index: usize) -> Result<Chunk>;
}

impl AltBody for DraftDocument {
    fn alts(&self) -> Result<Vec<Chunk>> {
        Ok(DraftDocument::alts(self))
    }
    fn insert_alt(&mut self, index: usize, chunk: Chunk, namespace: Conformance) -> Result<()> {
        DraftDocument::insert_alt(self, index, chunk, namespace)
    }
    fn replace_alt(&mut self, index: usize, chunk: Chunk, namespace: Conformance) -> Result<Chunk> {
        DraftDocument::replace_alt(self, index, chunk, namespace)
    }
    fn remove_alt(&mut self, index: usize) -> Result<Chunk> {
        DraftDocument::remove_alt(self, index)
    }
}

fn out_of_bounds(index: usize, len: usize) -> Error {
    Error::OutOfBounds {
        object: "altChunk",
        index,
        len,
    }
}

fn limit(kind: LimitKind) -> Error {
    Error::AltChunkLimit {
        kind,
        observed: MAX_CHUNKS.saturating_add(1),
        maximum: MAX_CHUNKS,
    }
}

impl Package {
    /// Insert a package-backed alternative-format import by anchor-relative index.
    ///
    /// Part, relationship, and body mutations are rolled back together on error.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation cannot be completed.
    pub(crate) fn insert_alt(
        &mut self,
        document: &mut impl AltBody,
        index: usize,
        import: Import,
        match_source: Option<bool>,
    ) -> Result<Chunk> {
        let count = document.alts()?.len();
        if index > count {
            return Err(out_of_bounds(index, count));
        }
        if count >= MAX_CHUNKS {
            return Err(Error::AltChunkLimit {
                kind: LimitKind::Anchors,
                observed: count.saturating_add(1),
                maximum: MAX_CHUNKS,
            });
        }
        let namespace = self.alt_chunk_namespace()?;
        let (chunk, installed_part) =
            self.install_alt_chunk_target(import, match_source, namespace)?;
        if let Err(error) = document.insert_alt(index, chunk.clone(), namespace) {
            self.discard_alt_chunk_target(chunk.relationship().as_str(), installed_part.as_ref())?;
            return Err(error);
        }
        Ok(chunk)
    }

    /// Replace an anchor and its relationship as one package mutation.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation cannot be completed.
    pub(crate) fn replace_alt(
        &mut self,
        document: &mut impl AltBody,
        index: usize,
        import: Import,
        match_source: Option<bool>,
    ) -> Result<Chunk> {
        let alts = document.alts()?;
        let old = alts
            .get(index)
            .cloned()
            .ok_or_else(|| out_of_bounds(index, alts.len()))?;
        let old_target = self.validate_alt_chunk_relationship(&old)?;
        let namespace = self.alt_chunk_namespace()?;
        let (new, installed_part) =
            self.install_alt_chunk_target(import, match_source, namespace)?;
        if let Err(error) = document.replace_alt(index, new.clone(), namespace) {
            self.discard_alt_chunk_target(new.relationship().as_str(), installed_part.as_ref())?;
            return Err(error);
        }
        self.remove_alt_relationship(document, &old, old_target.as_ref())?;
        Ok(old)
    }

    /// Remove an anchor, its relationship, and an unreachable internal payload.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation cannot be completed.
    pub(crate) fn remove_alt(
        &mut self,
        document: &mut impl AltBody,
        index: usize,
    ) -> Result<Chunk> {
        let alts = document.alts()?;
        let old = alts
            .get(index)
            .cloned()
            .ok_or_else(|| out_of_bounds(index, alts.len()))?;
        let old_target = self.validate_alt_chunk_relationship(&old)?;
        document.remove_alt(index)?;
        self.remove_alt_relationship(document, &old, old_target.as_ref())?;
        Ok(old)
    }

    /// The main document part, whose relationships anchors use.
    fn alt_document_uri(&self) -> Result<PackURI> {
        Ok(self.opc.main_document_part()?.partname().clone())
    }

    fn alt_chunk_namespace(&self) -> Result<Conformance> {
        let document_uri = self.alt_document_uri()?;
        // The namespace scan needs the payload, so a failure to read it is
        // reported rather than taken for a Transitional document (ADR 0030).
        let strict = self
            .opc
            .get_part(&document_uri)?
            .blob()
            .windows(b"http://purl.oclc.org/ooxml/wordprocessingml/main".len())
            .any(|window| window == b"http://purl.oclc.org/ooxml/wordprocessingml/main");
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
        let document_uri = self.alt_document_uri()?;
        let document = self.opc.get_part(&document_uri)?;
        let relationship_id = (1usize..=MAX_CHUNKS)
            .map(|number| format!("rIdAltChunk{number}"))
            .find(|id| document.rels().get(id).is_none())
            .ok_or_else(|| limit(LimitKind::RelationshipIds))?;
        let relationship = Rel::new(relationship_id.clone())?;
        let relationship_type = namespace.relationship();
        let (target_ref, target_mode, installed_part) = match import {
            Import::Link(uri) => (uri.into_string(), TargetMode::External, None),
            Import::Data(data) => {
                data.validate()?;
                let media_type = data.media_type();
                let mut selected = None;
                for number in 1usize..=MAX_CHUNKS {
                    let target_ref = format!("afchunk{number}.{}", data.extension());
                    let uri = PackURI::from_rel_ref(document_uri.base_uri(), &target_ref)
                        .map_err(Error::InvalidUri)?;
                    match self.opc.get_part(&uri) {
                        Ok(_) => {},
                        Err(litchi_opc::OpcError::PartNotFound(_)) => {
                            selected = Some((uri, target_ref));
                            break;
                        },
                        Err(error) => return Err(error.into()),
                    }
                }
                let (uri, target_ref) = selected.ok_or_else(|| limit(LimitKind::PartNames))?;
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
        let document_uri = self.alt_document_uri()?;
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
        self.opc.get_part(&target).map_err(|_source_error| {
            Error::InvalidFormat(format!(
                "altChunk relationship {:?} targets a missing part",
                chunk.relationship().as_str()
            ))
        })?;
        Ok(Some(target))
    }

    fn discard_alt_chunk_target(&mut self, id: &str, part: Option<&PackURI>) -> Result<()> {
        let document_uri = self.alt_document_uri()?;
        self.opc.get_part_mut(&document_uri)?.rels_mut().remove(id);
        if let Some(part) = part {
            self.opc.remove_part(part);
        }
        Ok(())
    }

    fn remove_alt_relationship(
        &mut self,
        document: &impl AltBody,
        chunk: &Chunk,
        target: Option<&PackURI>,
    ) -> Result<()> {
        if document
            .alts()?
            .iter()
            .any(|remaining| remaining.relationship() == chunk.relationship())
        {
            return Ok(());
        }
        let document_uri = self.alt_document_uri()?;
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

/// An opened package's main-document XML as an [`AltBody`] (record
/// 0355a): anchors are the active body-level `w:altChunk` elements, found
/// with the scan the reader uses, and every change is an exact splice of
/// the source bytes, so everything else in the part is kept byte for byte.
pub(crate) struct XmlBody {
    xml: Vec<u8>,
    changed: bool,
}

impl XmlBody {
    /// The active body anchors with their exact source ranges.
    fn anchors(&self) -> Result<Vec<(usize, usize, Chunk)>> {
        // Target index 2 of the scan is `altChunk`.
        const ALT_CHUNK: usize = 2;
        let (chunks, ranges) = crate::alt::scan_with_block_ranges(&self.xml)?;
        let mut anchors = Vec::new();
        for (target, start, length) in ranges {
            if target != ALT_CHUNK {
                continue;
            }
            let chunk = chunks.get(&start).cloned().ok_or_else(|| {
                Error::InvalidFormat("altChunk range has no parsed anchor".into())
            })?;
            let start = usize::try_from(start)
                .map_err(|_| Error::InvalidFormat("altChunk offset overflow".into()))?;
            let end = start
                .checked_add(
                    usize::try_from(length)
                        .map_err(|_| Error::InvalidFormat("altChunk length overflow".into()))?,
                )
                .ok_or_else(|| Error::InvalidFormat("altChunk range overflow".into()))?;
            anchors.push((start, end, chunk));
        }
        Ok(anchors)
    }

    fn splice(&mut self, start: usize, end: usize, bytes: &[u8]) -> Result<()> {
        if start > end || end > self.xml.len() {
            return Err(Error::InvalidFormat(
                "altChunk splice is out of range".into(),
            ));
        }
        let mut next = Vec::new();
        next.try_reserve_exact(self.xml.len() - (end - start) + bytes.len())
            .map_err(|source| Error::Allocation {
                resource: "alternative-format document XML",
                source,
            })?;
        next.extend_from_slice(&self.xml[..start]);
        next.extend_from_slice(bytes);
        next.extend_from_slice(&self.xml[end..]);
        self.xml = next;
        self.changed = true;
        Ok(())
    }

    /// The offset new trailing body content takes: before the body-final
    /// `w:sectPr`, or before `</w:body>`. An empty `<w:body/>` is first
    /// opened into a start and an end tag.
    fn content_end(&mut self) -> Result<usize> {
        use quick_xml::events::Event;
        let mut reader = quick_xml::reader::NsReader::from_reader(self.xml.as_slice());
        let origin = litchi_core::xml::ReaderOrigin::of(&self.xml);
        let mut depth = 0usize;
        let mut body_depth = None;
        let mut last_section = None::<usize>;
        let mut last_child_is_section = false;
        loop {
            let start = origin
                .offset(reader.buffer_position())
                .ok_or_else(|| Error::InvalidFormat("document XML offset overflow".into()))?;
            let event = reader
                .read_event()
                .map_err(|error| Error::Xml(error.to_string()))?;
            let end = origin
                .offset(reader.buffer_position())
                .ok_or_else(|| Error::InvalidFormat("document XML offset overflow".into()))?;
            let (namespace, event) = reader.resolver().resolve_event(event);
            let word = crate::namespace::is_wordprocessing_namespace(&namespace);
            match event {
                Event::Start(element) => {
                    depth += 1;
                    if body_depth.is_none() && word && element.local_name().as_ref() == b"body" {
                        body_depth = Some(depth);
                    } else if body_depth.is_some_and(|body| depth == body + 1) {
                        last_child_is_section = word && element.local_name().as_ref() == b"sectPr";
                        if last_child_is_section {
                            last_section = Some(start);
                        }
                    }
                },
                Event::Empty(element) => {
                    if body_depth.is_none() && word && element.local_name().as_ref() == b"body" {
                        // Open `<w:body .../>` into `<w:body ...></w:body>`.
                        let tag = self.xml[start..end].to_vec();
                        let name = element.name().as_ref().to_vec();
                        let trimmed = tag
                            .strip_suffix(b"/>")
                            .ok_or_else(|| Error::InvalidFormat("malformed empty body".into()))?;
                        let trimmed = trimmed.trim_ascii_end();
                        let mut opened = trimmed.to_vec();
                        opened.push(b'>');
                        let offset = start + opened.len();
                        opened.extend_from_slice(b"</");
                        opened.extend_from_slice(&name);
                        opened.push(b'>');
                        self.splice(start, end, &opened)?;
                        return Ok(offset);
                    }
                    if body_depth == Some(depth) {
                        last_child_is_section = word && element.local_name().as_ref() == b"sectPr";
                        if last_child_is_section {
                            last_section = Some(start);
                        }
                    }
                },
                Event::End(_) => {
                    if body_depth == Some(depth) {
                        return Ok(if last_child_is_section {
                            last_section.unwrap_or(start)
                        } else {
                            start
                        });
                    }
                    depth = depth.saturating_sub(1);
                },
                Event::Eof => {
                    return Err(Error::InvalidFormat("document XML has no w:body".into()));
                },
                _ => {},
            }
        }
    }

    /// Move anchor `from` to anchor position `to`, the builder's semantics.
    pub(crate) fn move_alt(&mut self, from: usize, to: usize) -> Result<()> {
        let anchors = self.anchors()?;
        let count = anchors.len();
        if from >= count || to >= count {
            return Err(out_of_bounds(from.max(to), count));
        }
        if from == to {
            return Ok(());
        }
        let (start, end, _) = anchors[from].clone();
        let bytes = self.xml[start..end].to_vec();
        self.splice(start, end, &[])?;
        let remaining = self.anchors()?;
        let destination = match remaining.get(to) {
            Some((start, _, _)) => *start,
            None => self.content_end()?,
        };
        self.splice(destination, destination, &bytes)
    }
}

impl AltBody for XmlBody {
    fn alts(&self) -> Result<Vec<Chunk>> {
        Ok(self
            .anchors()?
            .into_iter()
            .map(|(_, _, chunk)| chunk)
            .collect())
    }

    fn insert_alt(&mut self, index: usize, chunk: Chunk, namespace: Conformance) -> Result<()> {
        let anchors = self.anchors()?;
        if index > anchors.len() {
            return Err(out_of_bounds(index, anchors.len()));
        }
        let offset = match anchors.get(index) {
            Some((start, _, _)) => *start,
            None => self.content_end()?,
        };
        self.splice(offset, offset, chunk.xml(namespace).as_bytes())
    }

    fn replace_alt(&mut self, index: usize, chunk: Chunk, namespace: Conformance) -> Result<Chunk> {
        let anchors = self.anchors()?;
        let (start, end, old) = anchors
            .get(index)
            .cloned()
            .ok_or_else(|| out_of_bounds(index, anchors.len()))?;
        self.splice(start, end, chunk.xml(namespace).as_bytes())?;
        Ok(old)
    }

    fn remove_alt(&mut self, index: usize) -> Result<Chunk> {
        let anchors = self.anchors()?;
        let (start, end, old) = anchors
            .get(index)
            .cloned()
            .ok_or_else(|| out_of_bounds(index, anchors.len()))?;
        self.splice(start, end, &[])?;
        Ok(old)
    }
}

impl Package {
    /// Run an alternative-format verb on this package's main-document XML
    /// and store the XML back when it changed (record 0355a).
    pub(crate) fn with_alt_body<T>(
        &mut self,
        verb: impl FnOnce(&mut Self, &mut XmlBody) -> Result<T>,
    ) -> Result<T> {
        let uri = self.alt_document_uri()?;
        let xml = {
            let part = self.opc.get_part(&uri)?;
            part.ensure_payload()?;
            part.blob().to_vec()
        };
        let mut body = XmlBody {
            xml,
            changed: false,
        };
        let value = verb(self, &mut body)?;
        if body.changed {
            self.opc.get_part_mut(&uri)?.set_blob(body.xml);
        }
        Ok(value)
    }

    /// The active body anchors of the main document.
    pub(crate) fn alt_chunks(&self) -> Result<Vec<Chunk>> {
        let uri = self.alt_document_uri()?;
        let part = self.opc.get_part(&uri)?;
        part.ensure_payload()?;
        XmlBody {
            xml: part.blob().to_vec(),
            changed: false,
        }
        .alts()
    }
}
