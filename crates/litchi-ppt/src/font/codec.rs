//! Bounded, lossless-or-refuse `PowerPoint` font record I/O.

use super::model::{
    EmbeddedFont, Font, FontCollection, FontCollections, FontEmbeddingFlags, Limits, Scope,
};
use super::validation;
use crate::consts::RecordType;
use crate::package::{Error, Result};
use crate::records::Record;
use crate::records::RecordParseSession;

impl FontCollections {
    /// Validate every owned collection and embedding flag against `limits`.
    ///
    /// # Errors
    ///
    /// Returns an error if the limits are invalid or any collection violates
    /// them.
    pub fn validate_with_limits(&self, limits: Limits) -> Result<()> {
        validation::validate_collections(self, limits)
    }

    /// Parse font semantics from a live `DocumentContainer` with default limits.
    ///
    /// # Errors
    ///
    /// Returns an error if the record tree is not a valid live document font
    /// owner or contains malformed or unsupported font records.
    pub fn parse(root: &Record) -> Result<Self> {
        Self::parse_with_limits(root, Limits::default())
    }

    /// Parse only the live document's direct `Environment` and document-level
    /// `___PPT10` owners. Slide/master programmable tags are not searched.
    ///
    /// # Errors
    ///
    /// Returns an error if the input cannot be read or is malformed.
    pub fn parse_with_limits(root: &Record, limits: Limits) -> Result<Self> {
        validation::validate_limits(limits)?;
        let mut session = RecordParseSession::new(limits.records, root.data.len())?;
        session.account_existing(0)?;
        if root.record_type != RecordType::Document || root.version != 0x0f || root.instance != 0 {
            return Err(Error::Corrupted(
                "font owner is not a valid live DocumentContainer".into(),
            ));
        }
        let environments: Vec<_> = root
            .children
            .iter()
            .filter(|child| child.record_type == RecordType::Environment)
            .collect();
        if environments.len() > 1 {
            return Err(Error::Corrupted(
                "live document contains multiple Environment owners".into(),
            ));
        }
        if let Some(environment) = environments.first()
            && (environment.version != 0x0f || environment.instance != 0)
        {
            return Err(Error::Corrupted(
                "Environment has an invalid owner header".into(),
            ));
        }
        if !environments.is_empty() {
            session.account_existing(1)?;
        }
        let base_records: Vec<_> = environments
            .first()
            .into_iter()
            .flat_map(|environment| environment.children.iter())
            .filter(|child| child.record_type == RecordType::FontCollection)
            .collect();
        if base_records.len() > 1 {
            return Err(Error::Corrupted(
                "Environment contains multiple base font collections".into(),
            ));
        }
        let base = if let Some(record) = base_records.first() {
            Some(FontCollection::parse_with_session(
                record,
                limits,
                &mut session,
                2,
                false,
                false,
            )?)
        } else {
            None
        };

        let records = ppt10_records(root, limits, &mut session)?;
        let mut international = None;
        let mut embedding_flags = None;
        for record in records.unwrap_or_default() {
            #[allow(
                clippy::wildcard_enum_match_arm,
                reason = "`RecordType` has hundreds of variants; every record other than `FontCollection10` and `FontEmbedFlags10Atom` is intentionally ignored here"
            )]
            match record.record_type {
                RecordType::FontCollection10 if international.is_some() => {
                    return Err(Error::Corrupted(
                        "___PPT10 contains multiple international font collections".into(),
                    ));
                },
                RecordType::FontCollection10 => {
                    international = Some(FontCollection::parse_with_session(
                        &record,
                        limits,
                        &mut session,
                        5,
                        true,
                        false,
                    )?);
                },
                RecordType::FontEmbedFlags10Atom if embedding_flags.is_some() => {
                    return Err(Error::Corrupted(
                        "___PPT10 contains multiple font embedding flag atoms".into(),
                    ));
                },
                RecordType::FontEmbedFlags10Atom => {
                    embedding_flags = Some(FontEmbeddingFlags::parse(&record)?);
                },
                _ => {},
            }
        }
        let value = Self {
            base,
            international,
            embedding_flags,
        };
        validation::validate_collections(&value, limits)?;
        Ok(value)
    }

    /// Apply these font semantics to a live `DocumentContainer`, verifying the
    /// rewrite by reparsing the edited tree.
    ///
    /// # Errors
    ///
    /// Returns an error if the collections violate `limits`, the document
    /// cannot be edited losslessly, or the rewrite fails its reparse
    /// verification.
    pub fn apply_to_document(&self, root: &mut Record, limits: Limits) -> Result<()> {
        validation::validate_collections(self, limits)?;
        let current = Self::parse_with_limits(root, limits)?;
        self.apply_to_document_from(&current, root, limits)?;
        let reparsed = Self::parse_with_limits(root, limits)?;
        if &reparsed != self {
            return Err(Error::Corrupted(
                "font document rewrite did not preserve the requested semantics".into(),
            ));
        }
        Ok(())
    }

    /// Apply these font semantics to a live document whose current font state
    /// is already known as `current`.
    ///
    /// # Errors
    ///
    /// Returns an error if either collection violates `limits` or the edit
    /// cannot be applied losslessly.
    pub fn apply_to_document_from(
        &self,
        current: &Self,
        root: &mut Record,
        limits: Limits,
    ) -> Result<()> {
        validation::validate_collections(current, limits)?;
        validation::validate_collections(self, limits)?;
        apply_base(root, current.base.as_ref(), self.base.as_ref(), limits)?;
        apply_ppt10(root, current, self, limits)?;
        Ok(())
    }

    /// Restore every font-owned record into a normalized document tree whose
    /// large facet payloads were drained into [`SharedFontData`](super::SharedFontData).
    ///
    /// # Errors
    ///
    /// Returns an error if the collections violate `limits` or a font owner
    /// cannot be materialized losslessly.
    pub fn materialize_to_document(&self, root: &mut Record, limits: Limits) -> Result<()> {
        validation::validate_collections(self, limits)?;
        let mut environments: Vec<_> = root
            .children
            .iter_mut()
            .filter(|record| record.record_type == RecordType::Environment)
            .collect();
        let [environment] = environments.as_mut_slice() else {
            return Err(Error::Corrupted(
                "font materialization requires exactly one Environment".into(),
            ));
        };
        let positions: Vec<_> = environment
            .children
            .iter()
            .enumerate()
            .filter_map(|(index, record)| {
                (record.record_type == RecordType::FontCollection).then_some(index)
            })
            .collect();
        match (positions.as_slice(), self.base.as_ref()) {
            ([position], Some(collection)) => {
                environment.children[*position] =
                    collection.to_record_with_limits(limits, false)?;
            },
            ([], None) => {},
            _ => {
                return Err(Error::InvalidFormat(
                    "base font owner cannot be materialized losslessly".into(),
                ));
            },
        }
        if self.international.is_some() || self.embedding_flags.is_some() {
            edit_ppt10_records(root, limits, |records| {
                let collection_positions: Vec<_> = records
                    .iter()
                    .enumerate()
                    .filter_map(|(index, record)| {
                        (record.record_type == RecordType::FontCollection10).then_some(index)
                    })
                    .collect();
                match (collection_positions.as_slice(), self.international.as_ref()) {
                    ([position], Some(collection)) => {
                        records[*position] = collection.to_record_with_limits(limits, false)?;
                    },
                    ([], None) => {},
                    _ => {
                        return Err(Error::InvalidFormat(
                            "international font owner cannot be materialized losslessly".into(),
                        ));
                    },
                }
                let flag_positions: Vec<_> = records
                    .iter()
                    .enumerate()
                    .filter_map(|(index, record)| {
                        (record.record_type == RecordType::FontEmbedFlags10Atom).then_some(index)
                    })
                    .collect();
                let before = match flag_positions.as_slice() {
                    [position] => Some(FontEmbeddingFlags::parse(&records[*position])?),
                    [] => None,
                    _ => {
                        return Err(Error::InvalidFormat(
                            "font embedding flags cannot be materialized losslessly".into(),
                        ));
                    },
                };
                replace_flags(records, before, self.embedding_flags)?;
                Ok(())
            })?;
        }
        Ok(())
    }

    /// Drain large embedded payload vectors out of a parsed live Document tree
    /// into shared semantic owners, then discard redundant container payloads.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation fails.
    pub fn take_from_document(root: &mut Record, limits: Limits) -> Result<Self> {
        validation::validate_limits(limits)?;
        if root.record_type != RecordType::Document || root.version != 0x0f || root.instance != 0 {
            return Err(Error::Corrupted(
                "font owner is not a valid live DocumentContainer".into(),
            ));
        }
        let environment_positions: Vec<_> = root
            .children
            .iter()
            .enumerate()
            .filter_map(|(index, record)| {
                (record.record_type == RecordType::Environment).then_some(index)
            })
            .collect();
        if environment_positions.len() > 1 {
            return Err(Error::Corrupted(
                "live document has multiple Environment owners".into(),
            ));
        }
        let base = if let Some(position) = environment_positions.first().copied() {
            let environment = &mut root.children[position];
            if environment.version != 0x0f || environment.instance != 0 {
                return Err(Error::Corrupted(
                    "Environment has an invalid owner header".into(),
                ));
            }
            let positions: Vec<_> = environment
                .children
                .iter()
                .enumerate()
                .filter_map(|(index, record)| {
                    (record.record_type == RecordType::FontCollection).then_some(index)
                })
                .collect();
            match positions.as_slice() {
                [collection_position] => Some(FontCollection::take_with_limits(
                    &mut environment.children[*collection_position],
                    limits,
                )?),
                [] => None,
                _ => {
                    return Err(Error::Corrupted(
                        "Environment has multiple font collections".into(),
                    ));
                },
            }
        } else {
            None
        };

        let mut international = None;
        let mut embedding_flags = None;
        let had_ppt10 = edit_ppt10_records_optional(root, limits, |records| {
            let positions: Vec<_> = records
                .iter()
                .enumerate()
                .filter_map(|(index, record)| {
                    (record.record_type == RecordType::FontCollection10).then_some(index)
                })
                .collect();
            match positions.as_slice() {
                [position] => {
                    international = Some(FontCollection::take_with_limits(
                        &mut records[*position],
                        limits,
                    )?);
                },
                [] => {},
                _ => {
                    return Err(Error::Corrupted(
                        "___PPT10 has multiple font collections".into(),
                    ));
                },
            }
            let flags: Vec<_> = records
                .iter()
                .filter(|record| record.record_type == RecordType::FontEmbedFlags10Atom)
                .collect();
            match flags.as_slice() {
                [record] => embedding_flags = Some(FontEmbeddingFlags::parse(record)?),
                [] => {},
                _ => {
                    return Err(Error::Corrupted(
                        "___PPT10 has multiple font flag atoms".into(),
                    ));
                },
            }
            Ok(())
        })?;
        if !had_ppt10 {
            international = None;
            embedding_flags = None;
        }
        clear_redundant_container_data(root);
        let value = Self {
            base,
            international,
            embedding_flags,
        };
        validation::validate_collections(&value, limits)?;
        Ok(value)
    }

    /// Canonical full base collection record for fresh writers.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation fails.
    pub fn base_record_bytes(&self) -> Result<Option<Vec<u8>>> {
        self.base
            .as_ref()
            .map(FontCollection::to_record_bytes)
            .transpose()
    }

    /// Canonical full PP10 records in grammar order: collection, then flags.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation fails.
    pub fn powerpoint10_records(&self) -> Result<Vec<Vec<u8>>> {
        validation::validate_collections(self, Limits::default())?;
        let mut records = Vec::with_capacity(2);
        if let Some(collection) = &self.international {
            records.push(collection.to_record_bytes()?);
        }
        if let Some(flags) = self.embedding_flags {
            records.push(encode_record(
                &flags.to_record()?,
                Limits::default().records,
            )?);
        }
        Ok(records)
    }
}

impl FontEmbeddingFlags {
    /// Parse a `FontEmbedFlags10Atom` record.
    ///
    /// # Errors
    ///
    /// Returns an error if the record is not a `FontEmbedFlags10Atom` with a
    /// four-byte payload.
    pub fn parse(record: &Record) -> Result<Self> {
        if record.record_type != RecordType::FontEmbedFlags10Atom
            || record.version != 0
            || record.instance != 0
            || record.data.len() != 4
        {
            return Err(Error::Corrupted(
                "FontEmbedFlags10Atom has an invalid header or size".into(),
            ));
        }
        let raw = u32::from_le_bytes([
            record.data[0],
            record.data[1],
            record.data[2],
            record.data[3],
        ]);
        Ok(Self {
            raw,
            subset: raw & 1 != 0,
            subset_option_confirmed: raw & 2 != 0,
        })
    }

    /// Serialize back to a `FontEmbedFlags10Atom` record.
    ///
    /// # Errors
    ///
    /// Returns an error if the projected flag bits disagree with the retained
    /// raw bits.
    pub fn to_record(self) -> Result<Record> {
        if self.subset != (self.raw & 1 != 0) || self.subset_option_confirmed != (self.raw & 2 != 0)
        {
            return Err(Error::Corrupted(
                "font embedding flag projections disagree with raw bits".into(),
            ));
        }
        Ok(atom(
            RecordType::FontEmbedFlags10Atom,
            0,
            self.raw.to_le_bytes().to_vec(),
        ))
    }
}

impl FontCollection {
    /// Parse a font collection record with default limits.
    ///
    /// # Errors
    ///
    /// Returns an error if the record is malformed or violates the default
    /// limits.
    pub fn parse(record: &Record) -> Result<Self> {
        Self::parse_with_limits(record, Limits::default())
    }

    /// Parse a font collection record, enforcing `limits`.
    ///
    /// # Errors
    ///
    /// Returns an error if the limits are invalid, the record is malformed, or
    /// it violates `limits`.
    pub fn parse_with_limits(record: &Record, limits: Limits) -> Result<Self> {
        validation::validate_limits(limits)?;
        let mut session = RecordParseSession::new(limits.records, record.data.len())?;
        Self::parse_with_session(record, limits, &mut session, 0, false, false)
    }

    fn parse_with_session(
        record: &Record,
        limits: Limits,
        session: &mut RecordParseSession,
        logical_depth: usize,
        owner_already_accounted: bool,
        children_already_accounted: bool,
    ) -> Result<Self> {
        if !owner_already_accounted {
            session.account_existing(logical_depth)?;
        }
        #[allow(
            clippy::wildcard_enum_match_arm,
            reason = "`RecordType` has hundreds of variants; every record other than `FontCollection` and `FontCollection10` is rejected here"
        )]
        let scope = match record.record_type {
            RecordType::FontCollection => Scope::Base,
            RecordType::FontCollection10 => Scope::International,
            _ => return Err(Error::Corrupted("record is not a font collection".into())),
        };
        if record.version != 0x0f || record.instance != 0 {
            return Err(Error::Corrupted(
                "font collection has an invalid header".into(),
            ));
        }
        let child_depth = logical_depth
            .checked_add(1)
            .ok_or_else(|| Error::ResourceLimit("font record depth overflow".into()))?;
        let parsed;
        let children = if record.children.is_empty() {
            parsed = session.parse_sequence(&record.data, "font collection", child_depth)?;
            parsed.as_slice()
        } else if !children_already_accounted {
            for _ in &record.children {
                session.account_existing(child_depth)?;
            }
            record.children.as_slice()
        } else {
            record.children.as_slice()
        };
        if children.len() > limits.records.max_records
            || (!children.is_empty() && limits.records.max_depth == 0)
        {
            return Err(Error::ResourceLimit(
                "font collection exceeds its record-count or depth limit".into(),
            ));
        }
        let mut collection = FontCollection::new(scope);
        collection
            .fonts
            .try_reserve(children.len().min(limits.max_fonts_per_collection))
            .map_err(|_err| Error::AllocationFailed("font collection"))?;
        let mut current: Option<Font> = None;
        let mut facets = 0usize;
        let mut embedded_bytes = 0usize;
        for child in children {
            let encoded =
                child.data.len().checked_add(8).ok_or_else(|| {
                    Error::ResourceLimit("font child record size overflow".into())
                })?;
            if child.data.len() > limits.records.max_record_payload_bytes
                || encoded > limits.records.max_record_bytes
            {
                return Err(Error::ResourceLimit(
                    "font child exceeds its record byte limit".into(),
                ));
            }
            #[allow(
                clippy::wildcard_enum_match_arm,
                reason = "`RecordType` has hundreds of variants; every record other than `FontEntityAtom` and `FontEmbeddedData` is intentionally rejected here"
            )]
            match child.record_type {
                RecordType::FontEntityAtom => {
                    if let Some(font) = current.take() {
                        collection.fonts.push(font);
                    }
                    if collection.fonts.len() >= limits.max_fonts_per_collection {
                        return Err(Error::Corrupted("font collection exceeds its limit".into()));
                    }
                    current = Some(parse_font_entity(child, collection.fonts.len())?);
                },
                RecordType::FontEmbeddedData => {
                    let font = current.as_mut().ok_or_else(|| {
                        Error::Corrupted("embedded font precedes its FontEntityAtom".into())
                    })?;
                    if child.version != 0 || child.instance > 3 {
                        return Err(Error::Corrupted(
                            "embedded font has an invalid header".into(),
                        ));
                    }
                    #[allow(
                        clippy::cast_possible_truncation,
                        reason = "`child.instance` is checked to be at most 3 immediately above"
                    )]
                    let style = child.instance as u8;
                    if font
                        .embedded_fonts
                        .last()
                        .is_some_and(|old| old.style >= style)
                    {
                        return Err(Error::Corrupted(
                            "embedded font facets are duplicated or out of order".into(),
                        ));
                    }
                    facets = facets
                        .checked_add(1)
                        .ok_or_else(|| Error::Corrupted("facet count overflow".into()))?;
                    embedded_bytes =
                        embedded_bytes
                            .checked_add(child.data.len())
                            .ok_or_else(|| {
                                Error::Corrupted("embedded font byte count overflow".into())
                            })?;
                    if facets > limits.max_facets
                        || child.data.len() > limits.max_facet_bytes
                        || embedded_bytes > limits.max_embedded_bytes
                    {
                        return Err(Error::Corrupted(
                            "embedded fonts exceed configured limits".into(),
                        ));
                    }
                    font.embedded_fonts.push(EmbeddedFont::from_preserved(
                        style.try_into()?,
                        child.data.as_slice(),
                    ));
                },
                _ => {
                    return Err(Error::Corrupted(format!(
                        "unsupported record {} in font collection; lossless edit refused",
                        child.record_type_raw
                    )));
                },
            }
        }
        if let Some(font) = current {
            collection.fonts.push(font);
        }
        validation::validate_collection(&collection, limits)?;
        Ok(collection)
    }

    /// Drain embedded facet payloads out of an owned font collection record,
    /// returning the semantic collection.
    ///
    /// # Errors
    ///
    /// Returns an error if the record is malformed, violates `limits`, or the
    /// drained payloads lose alignment with the parsed facets.
    pub fn take_with_limits(record: &mut Record, limits: Limits) -> Result<Self> {
        let mut session = RecordParseSession::new(limits.records, record.data.len())?;
        session.account_existing(0)?;
        if record.children.is_empty() {
            record.children = session.parse_sequence(&record.data, "owned font collection", 1)?;
        } else {
            for _ in &record.children {
                session.account_existing(1)?;
            }
        }
        let mut payloads = Vec::new();
        payloads
            .try_reserve(record.children.len().min(limits.max_facets))
            .map_err(|_err| Error::AllocationFailed("owned embedded font facets"))?;
        for child in &mut record.children {
            if child.record_type == RecordType::FontEmbeddedData {
                payloads.push(std::mem::take(&mut child.data));
                child.data_length = 0;
            }
        }
        record.data.clear();
        record.data_length = 0;
        let mut collection = Self::parse_with_session(record, limits, &mut session, 0, true, true)?;
        let mut payload_queue = payloads.into_iter();
        for facet in collection
            .fonts
            .iter_mut()
            .flat_map(|font| font.embedded_fonts.iter_mut())
        {
            let payload = payload_queue.next().ok_or_else(|| {
                Error::Corrupted("owned font facet extraction lost semantic alignment".into())
            })?;
            facet.data = super::SharedFontData::from(payload);
        }
        if payload_queue.next().is_some() {
            return Err(Error::Corrupted(
                "owned font facet extraction lost semantic alignment".into(),
            ));
        }
        validation::validate_collection(&collection, limits)?;
        Ok(collection)
    }

    /// Lossless checked serialization, retaining ignored bits and source name padding.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation fails.
    pub fn to_record(&self) -> Result<Record> {
        self.to_record_with_limits(Limits::default(), false)
    }

    /// Canonical checked serialization for newly authored content.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation fails.
    pub fn to_record_canonical(&self) -> Result<Record> {
        self.to_record_with_limits(Limits::default(), true)
    }

    /// Lossless checked serialization to bytes with default limits.
    ///
    /// # Errors
    ///
    /// Returns an error if the collection fails validation or serialization
    /// exceeds the default record limits.
    pub fn to_record_bytes(&self) -> Result<Vec<u8>> {
        self.to_record_bytes_with_limits(Limits::default())
    }

    /// Canonical checked serialization to bytes, enforcing `limits`.
    ///
    /// # Errors
    ///
    /// Returns an error if the collection fails validation or serialization
    /// exceeds `limits`.
    pub fn to_record_bytes_with_limits(&self, limits: Limits) -> Result<Vec<u8>> {
        encode_record(&self.to_record_with_limits(limits, true)?, limits.records)
    }

    pub(crate) fn to_record_with_limits(&self, limits: Limits, canonical: bool) -> Result<Record> {
        if canonical {
            validation::validate_authored_collection(self, limits)?;
        } else {
            validation::validate_collection(self, limits)?;
        }
        let estimated = self
            .fonts
            .iter()
            .try_fold(0usize, |total, font| {
                font.embedded_fonts
                    .iter()
                    .try_fold(total + 76, |subtotal, facet| {
                        subtotal.checked_add(8 + facet.data.len())
                    })
            })
            .ok_or_else(|| Error::Corrupted("font collection size overflow".into()))?;
        let mut payload = Vec::new();
        payload
            .try_reserve(estimated)
            .map_err(|_err| Error::AllocationFailed("font collection serialization"))?;
        for font in &self.fonts {
            payload.extend_from_slice(&encode_record(
                &font_entity(font, canonical)?,
                limits.records,
            )?);
            for embedded in &font.embedded_fonts {
                let facet = embedded.facet()?;
                append_leaf_record(
                    &mut payload,
                    0,
                    facet as u16,
                    RecordType::FontEmbeddedData.as_u16(),
                    embedded.data.as_ref(),
                    limits.records,
                )?;
            }
        }
        let kind = if self.international {
            RecordType::FontCollection10
        } else {
            RecordType::FontCollection
        };
        Ok(Record {
            record_type: kind,
            record_type_raw: kind.as_u16(),
            version: 0x0f,
            instance: 0,
            data_length: u32::try_from(payload.len())
                .map_err(|_err| Error::Corrupted("font collection exceeds u32".into()))?,
            data: payload,
            children: Vec::new(),
        })
    }
}

fn parse_font_entity(record: &Record, ordinal: usize) -> Result<Font> {
    if record.version != 0 || record.instance > 128 || record.data.len() != 68 {
        return Err(Error::Corrupted(
            "FontEntityAtom has an invalid header or size".into(),
        ));
    }
    let mut source_name = [0u8; 64];
    source_name.copy_from_slice(&record.data[..64]);
    let units: Vec<u16> = source_name
        .chunks_exact(2)
        .map(|bytes| u16::from_le_bytes([bytes[0], bytes[1]]))
        .collect();
    let end = units
        .iter()
        .position(|unit| *unit == 0)
        .unwrap_or(units.len());
    let name = String::from_utf16(&units[..end])
        .map_err(|_err| Error::Corrupted("FontEntityAtom name is invalid UTF-16".into()))?;
    let font_flags = record.data[65];
    let font_type_flags = record.data[66];
    Ok(Font {
        index: u16::try_from(ordinal)
            .map_err(|_err| Error::Corrupted("font ordinal exceeds u16".into()))?,
        raw_instance: record.instance,
        name,
        charset: record.data[64],
        font_flags,
        embedded_subset: font_flags & 1 != 0,
        font_type_flags,
        raster: font_type_flags & 1 != 0,
        device: font_type_flags & 2 != 0,
        truetype: font_type_flags & 4 != 0,
        no_substitution: font_type_flags & 8 != 0,
        pitch_and_family: record.data[67],
        embedded_fonts: Vec::new(),
        source_name: Some(source_name),
    })
}

fn font_entity(font: &Font, canonical: bool) -> Result<Record> {
    validation::validate_font(font)?;
    let mut data = vec![0u8; 68];
    if !canonical
        && let Some(source) = &font.source_name
        && decode_name(source).ok().as_deref() == Some(font.name.as_str())
    {
        data[..64].copy_from_slice(source);
    } else {
        let units: Vec<u16> = font.name.encode_utf16().collect();
        for (index, unit) in units.into_iter().enumerate() {
            data[index * 2..index * 2 + 2].copy_from_slice(&unit.to_le_bytes());
        }
    }
    data[64] = font.charset;
    data[65] = if canonical {
        font.font_flags & 1
    } else {
        font.font_flags
    };
    data[66] = if canonical {
        font.font_type_flags & 0x0f
    } else {
        font.font_type_flags
    };
    data[67] = font.pitch_and_family;
    Ok(atom(RecordType::FontEntityAtom, font.raw_instance, data))
}

fn decode_name(bytes: &[u8; 64]) -> Result<String> {
    let units: Vec<u16> = bytes
        .chunks_exact(2)
        .map(|v| u16::from_le_bytes([v[0], v[1]]))
        .collect();
    let end = units
        .iter()
        .position(|unit| *unit == 0)
        .unwrap_or(units.len());
    String::from_utf16(&units[..end])
        .map_err(|_err| Error::Corrupted("font name is invalid UTF-16".into()))
}

fn apply_base(
    root: &mut Record,
    before: Option<&FontCollection>,
    after: Option<&FontCollection>,
    limits: Limits,
) -> Result<()> {
    if before == after {
        return Ok(());
    }
    let mut environments: Vec<_> = root
        .children
        .iter_mut()
        .filter(|r| r.record_type == RecordType::Environment)
        .collect();
    let [environment] = environments.as_mut_slice() else {
        return Err(Error::Corrupted(
            "font edit requires exactly one live Environment".into(),
        ));
    };
    let positions: Vec<_> = environment
        .children
        .iter()
        .enumerate()
        .filter_map(|(i, r)| (r.record_type == RecordType::FontCollection).then_some(i))
        .collect();
    match (positions.as_slice(), after) {
        ([position], Some(collection)) => {
            environment.children[*position] = collection.to_record_with_limits(limits, false)?;
        },
        ([], Some(_)) => {
            return Err(Error::InvalidFormat(
                "adding a missing base FontCollection owner is not losslessly provable".into(),
            ));
        },
        ([_], None) => {
            return Err(Error::InvalidFormat(
                "removing a font collection requires a complete reference remap".into(),
            ));
        },
        _ => {
            return Err(Error::Corrupted(
                "Environment font owner is ambiguous".into(),
            ));
        },
    }
    Ok(())
}

fn apply_ppt10(
    root: &mut Record,
    before: &FontCollections,
    after: &FontCollections,
    limits: Limits,
) -> Result<()> {
    if before.international == after.international
        && before.embedding_flags == after.embedding_flags
    {
        return Ok(());
    }
    edit_ppt10_records(root, limits, |records| {
        replace_optional_record(
            records,
            RecordType::FontCollection10,
            before.international.as_ref(),
            after
                .international
                .as_ref()
                .map(|v| v.to_record_with_limits(limits, false))
                .transpose()?,
        )?;
        replace_flags(records, before.embedding_flags, after.embedding_flags)?;
        Ok(())
    })
}

fn replace_optional_record(
    records: &mut Vec<Record>,
    kind: RecordType,
    before: Option<&FontCollection>,
    after: Option<Record>,
) -> Result<()> {
    let positions: Vec<_> = records
        .iter()
        .enumerate()
        .filter_map(|(i, r)| (r.record_type == kind).then_some(i))
        .collect();
    match (positions.as_slice(), before, after) {
        ([position], Some(_), Some(record)) => records[*position] = record,
        ([], None, Some(record)) => records.insert(0, record),
        ([_], Some(_), None) => return Err(Error::InvalidFormat(
            "removing an international font collection requires a complete FontIndexRef10 remap"
                .into(),
        )),
        ([], None, None) => {},
        _ => {
            return Err(Error::Corrupted(
                "PowerPoint 10 font owner is ambiguous".into(),
            ));
        },
    }
    Ok(())
}

fn replace_flags(
    records: &mut Vec<Record>,
    before: Option<FontEmbeddingFlags>,
    after: Option<FontEmbeddingFlags>,
) -> Result<()> {
    let positions: Vec<_> = records
        .iter()
        .enumerate()
        .filter_map(|(i, r)| (r.record_type == RecordType::FontEmbedFlags10Atom).then_some(i))
        .collect();
    match (positions.as_slice(), before, after) {
        ([position], Some(_), Some(flags)) => records[*position] = flags.to_record()?,
        ([position], Some(_), None) => {
            records.remove(*position);
        },
        ([], None, Some(flags)) => {
            // Every record through CommentIndex10 precedes the flags. All
            // later grammar members, including an opaque preserved tail,
            // follow it.
            let position = records
                .iter()
                .position(|record| !is_before_font_embed_flags(record))
                .unwrap_or(records.len());
            if records[position..].iter().any(is_before_font_embed_flags) {
                return Err(Error::Corrupted(
                    "PowerPoint 10 extension order is ambiguous around font flags".into(),
                ));
            }
            records.insert(position, flags.to_record()?);
        },
        ([], None, None) => {},
        _ => {
            return Err(Error::Corrupted(
                "PowerPoint 10 embedding flags are ambiguous".into(),
            ));
        },
    }
    Ok(())
}

fn is_before_font_embed_flags(record: &Record) -> bool {
    matches!(
        record.record_type,
        RecordType::FontCollection10
            | RecordType::TextMasterStyle10Atom
            | RecordType::TextDefaults10Atom
            | RecordType::GridSpacing10Atom
            | RecordType::CommentIndex10
    )
}

fn ppt10_records(
    root: &Record,
    _limits: Limits,
    session: &mut RecordParseSession,
) -> Result<Option<Vec<Record>>> {
    let doc_info_lists: Vec<_> = root
        .children
        .iter()
        .filter(|record| record.record_type == RecordType::DocInfoList)
        .collect();
    if doc_info_lists.len() > 1 {
        return Err(Error::Corrupted(
            "live document contains multiple DocInfoList owners".into(),
        ));
    }
    let Some(doc_info) = doc_info_lists.first().copied() else {
        return Ok(None);
    };
    session.account_existing(1)?;
    if doc_info.version != 0x0f || doc_info.instance != 0 {
        return Err(Error::Corrupted(
            "DocInfoList has an invalid owner header".into(),
        ));
    }
    let prog_tag_owners: Vec<_> = doc_info
        .children
        .iter()
        .filter(|record| record.record_type == RecordType::ProgTags)
        .collect();
    if prog_tag_owners.len() > 1 {
        return Err(Error::Corrupted(
            "DocInfoList contains multiple DocProgTags owners".into(),
        ));
    }
    let Some(prog_tags) = prog_tag_owners.first().copied() else {
        return Ok(None);
    };
    session.account_existing(2)?;
    if prog_tags.version != 0x0f || prog_tags.instance != 0 {
        return Err(Error::Corrupted(
            "DocProgTags has an invalid owner header".into(),
        ));
    }
    let mut found = None;
    let tags = session.parse_sequence(&prog_tags.data, "document ProgTags", 3)?;
    for tag in tags
        .into_iter()
        .filter(|r| r.record_type == RecordType::ProgBinaryTag)
    {
        if tag.version != 0x0f || tag.instance != 0 {
            return Err(Error::Corrupted(
                "ProgBinaryTag has an invalid owner header".into(),
            ));
        }
        let pair = session.parse_sequence(&tag.data, "document ProgBinaryTag", 4)?;
        if tag_version(&pair) != Some(10) {
            continue;
        }
        let blob = one_blob(&pair)?;
        if blob.version != 0 || blob.instance != 0 {
            return Err(Error::Corrupted(
                "BinaryTagData has an invalid owner header".into(),
            ));
        }
        if found.is_some() {
            return Err(Error::Corrupted(
                "live document contains multiple ___PPT10 tags".into(),
            ));
        }
        let records = session.parse_sequence(&blob.data, "___PPT10 BinaryTagData", 5)?;
        found = Some(records);
    }
    Ok(found)
}

fn edit_ppt10_records(
    root: &mut Record,
    limits: Limits,
    edit: impl FnOnce(&mut Vec<Record>) -> Result<()>,
) -> Result<()> {
    if !edit_ppt10_records_optional(root, limits, edit)? {
        return Err(Error::InvalidFormat(
            "font edit requires an existing document ___PPT10 owner".into(),
        ));
    }
    Ok(())
}

fn edit_ppt10_records_optional(
    root: &mut Record,
    limits: Limits,
    edit: impl FnOnce(&mut Vec<Record>) -> Result<()>,
) -> Result<bool> {
    let mut session = RecordParseSession::new(limits.records, root.data.len())?;
    session.account_existing(0)?;
    let doc_info_positions: Vec<_> = root
        .children
        .iter()
        .enumerate()
        .filter_map(|(index, record)| {
            (record.record_type == RecordType::DocInfoList).then_some(index)
        })
        .collect();
    if doc_info_positions.len() != 1 {
        return Err(Error::InvalidFormat(
            "font edit requires exactly one DocInfoList owner".into(),
        ));
    }
    let doc_info = &mut root.children[doc_info_positions[0]];
    session.account_existing(1)?;
    if doc_info.version != 0x0f || doc_info.instance != 0 {
        return Err(Error::Corrupted(
            "DocInfoList has an invalid owner header".into(),
        ));
    }
    let prog_positions: Vec<_> = doc_info
        .children
        .iter()
        .enumerate()
        .filter_map(|(index, record)| (record.record_type == RecordType::ProgTags).then_some(index))
        .collect();
    if prog_positions.len() != 1 {
        return Err(Error::InvalidFormat(
            "font edit requires exactly one DocProgTags owner".into(),
        ));
    }
    let prog_tags = &mut doc_info.children[prog_positions[0]];
    session.account_existing(2)?;
    if prog_tags.version != 0x0f || prog_tags.instance != 0 {
        return Err(Error::Corrupted(
            "DocProgTags has an invalid owner header".into(),
        ));
    }
    let mut pending_edit = Some(edit);
    let mut found = false;
    let mut tags = session.parse_sequence(&prog_tags.data, "document ProgTags", 3)?;
    let mut changed = false;
    for tag in tags
        .iter_mut()
        .filter(|r| r.record_type == RecordType::ProgBinaryTag)
    {
        if tag.version != 0x0f || tag.instance != 0 {
            return Err(Error::Corrupted(
                "ProgBinaryTag has an invalid owner header".into(),
            ));
        }
        let mut pair = session.parse_sequence(&tag.data, "document ProgBinaryTag", 4)?;
        if tag_version(&pair) != Some(10) {
            continue;
        }
        if found {
            return Err(Error::Corrupted(
                "live document contains multiple ___PPT10 tags".into(),
            ));
        }
        found = true;
        let blob = one_blob_mut(&mut pair)?;
        if blob.version != 0 || blob.instance != 0 {
            return Err(Error::Corrupted(
                "BinaryTagData has an invalid owner header".into(),
            ));
        }
        let mut records = session.parse_sequence(&blob.data, "___PPT10 BinaryTagData", 5)?;
        let apply_edit = pending_edit
            .take()
            .ok_or_else(|| Error::Corrupted("___PPT10 edit closure was already consumed".into()))?;
        apply_edit(&mut records)?;
        blob.data = encode_sequence(&records, limits.records)?;
        blob.data_length = u32::try_from(blob.data.len())
            .map_err(|_err| Error::Corrupted("PPT10 blob exceeds u32".into()))?;
        blob.children.clear();
        tag.data = encode_sequence(&pair, limits.records)?;
        tag.data_length = u32::try_from(tag.data.len())
            .map_err(|_err| Error::Corrupted("PPT10 tag exceeds u32".into()))?;
        tag.children.clear();
        changed = true;
    }
    if changed {
        prog_tags.data = encode_sequence(&tags, limits.records)?;
        prog_tags.data_length = u32::try_from(prog_tags.data.len())
            .map_err(|_err| Error::Corrupted("ProgTags exceeds u32".into()))?;
        prog_tags.children.clear();
    }
    Ok(found)
}

fn clear_redundant_container_data(record: &mut Record) {
    for child in &mut record.children {
        clear_redundant_container_data(child);
    }
    if !record.children.is_empty() {
        record.data.clear();
        record.data_length = 0;
    }
}

fn tag_version(pair: &[Record]) -> Option<u8> {
    if pair.len() != 2
        || pair[0].record_type != RecordType::CString
        || pair[1].record_type != RecordType::BinaryTagData
    {
        return None;
    }
    let expected: Vec<u8> = "___PPT10"
        .encode_utf16()
        .flat_map(u16::to_le_bytes)
        .collect();
    (pair[0].version == 0
        && pair[0].instance == 0
        && pair[0].data == expected
        && pair[1].version == 0
        && pair[1].instance == 0)
        .then_some(10)
}

fn one_blob(pair: &[Record]) -> Result<&Record> {
    if pair.len() != 2
        || pair[0].record_type != RecordType::CString
        || pair[1].record_type != RecordType::BinaryTagData
        || pair[1].version != 0
        || pair[1].instance != 0
    {
        return Err(Error::Corrupted(
            "___PPT10 tag does not contain its exact CString/BinaryTagData pair".into(),
        ));
    }
    Ok(&pair[1])
}

fn one_blob_mut(pair: &mut [Record]) -> Result<&mut Record> {
    if pair.len() != 2
        || pair[0].record_type != RecordType::CString
        || pair[1].record_type != RecordType::BinaryTagData
        || pair[1].version != 0
        || pair[1].instance != 0
    {
        return Err(Error::Corrupted(
            "___PPT10 tag does not contain its exact CString/BinaryTagData pair".into(),
        ));
    }
    Ok(&mut pair[1])
}

pub(crate) fn encode_sequence(
    records: &[Record],
    limits: crate::package::RecordLimits,
) -> Result<Vec<u8>> {
    let capacity = records
        .iter()
        .try_fold(0usize, |total, record| {
            total.checked_add(8)?.checked_add(record.data.len())
        })
        .ok_or_else(|| Error::Corrupted("record sequence size overflow".into()))?;
    let mut output = Vec::new();
    output
        .try_reserve(capacity)
        .map_err(|_err| Error::AllocationFailed("record sequence serialization"))?;
    for record in records {
        output.extend_from_slice(&encode_record(record, limits)?);
    }
    Ok(output)
}

pub(crate) fn encode_record(
    record: &Record,
    limits: crate::package::RecordLimits,
) -> Result<Vec<u8>> {
    let nested;
    let payload = if record.children.is_empty() {
        record.data.as_slice()
    } else {
        nested = encode_sequence(&record.children, limits)?;
        nested.as_slice()
    };
    let total = payload
        .len()
        .checked_add(8)
        .ok_or_else(|| Error::Corrupted("record size overflow".into()))?;
    if total > limits.max_record_bytes || payload.len() > limits.max_record_payload_bytes {
        return Err(Error::Corrupted(
            "serialized PPT record exceeds configured limits".into(),
        ));
    }
    let length = u32::try_from(payload.len())
        .map_err(|_err| Error::Corrupted("record payload exceeds u32".into()))?;
    let mut bytes = Vec::new();
    bytes
        .try_reserve(total)
        .map_err(|_err| Error::AllocationFailed("record serialization"))?;
    bytes.extend_from_slice(&((record.instance << 4) | (record.version & 0x0f)).to_le_bytes());
    bytes.extend_from_slice(&record.record_type_raw.to_le_bytes());
    bytes.extend_from_slice(&length.to_le_bytes());
    bytes.extend_from_slice(payload);
    Ok(bytes)
}

fn append_leaf_record(
    output: &mut Vec<u8>,
    version: u16,
    instance: u16,
    record_type: u16,
    payload: &[u8],
    limits: crate::package::RecordLimits,
) -> Result<()> {
    let total = payload
        .len()
        .checked_add(8)
        .ok_or_else(|| Error::Corrupted("record size overflow".into()))?;
    if total > limits.max_record_bytes || payload.len() > limits.max_record_payload_bytes {
        return Err(Error::ResourceLimit(
            "serialized font record exceeds configured limits".into(),
        ));
    }
    let length = u32::try_from(payload.len())
        .map_err(|_err| Error::Corrupted("record payload exceeds u32".into()))?;
    output
        .try_reserve(total)
        .map_err(|_err| Error::AllocationFailed("font record serialization"))?;
    output.extend_from_slice(&((instance << 4) | (version & 0x0f)).to_le_bytes());
    output.extend_from_slice(&record_type.to_le_bytes());
    output.extend_from_slice(&length.to_le_bytes());
    output.extend_from_slice(payload);
    Ok(())
}

#[allow(
    clippy::cast_possible_truncation,
    reason = "`atom` callers pass fixed tiny payloads (4 or 68 bytes), far below the u32 length field range"
)]
fn atom(kind: RecordType, instance: u16, data: Vec<u8>) -> Record {
    Record {
        record_type: kind,
        record_type_raw: kind.as_u16(),
        version: 0,
        instance,
        data_length: data.len() as u32,
        data,
        children: Vec::new(),
    }
}
