//! Hyperlink definitions and PowerPoint 9 hyperlink extensions.

use super::package::{PptError, Result};
use super::records::PptRecord;
use crate::consts::PptRecordType;

/// Additional hyperlink data introduced by PowerPoint 9.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct PowerPointHyperlinkExtension {
    /// Optional text displayed as a hover screen tip.
    pub screen_tip: Option<String>,
    /// Whether the hyperlink was created in the Insert Hyperlink dialog.
    pub inserted_with_dialog: bool,
    /// Whether the base hyperlink location names a custom slide show.
    pub location_is_named_show: bool,
    /// Whether a named show returns to the originating slide.
    pub named_show_returns_to_slide: bool,
}

impl PowerPointHyperlinkExtension {
    /// Parse an `ExHyperlink9Container` record and return its referenced ID.
    pub fn parse(record: &PptRecord) -> Result<(u32, Self)> {
        if record.record_type != PptRecordType::ExternalHyperlink9
            || record.version != 0x0f
            || record.instance != 0
        {
            return Err(PptError::Corrupted(
                "ExHyperlink9Container has an invalid record header".to_string(),
            ));
        }
        let children = PptRecord::parse_sequence_strict(&record.data, "PowerPoint 9 hyperlink")?;
        if !matches!(children.len(), 2 | 3) {
            return Err(PptError::Corrupted(
                "ExHyperlink9Container has an invalid child count".to_string(),
            ));
        }
        let reference = parse_hyperlink_atom(&children[0])?;
        if reference == 0 {
            return Err(PptError::Corrupted(
                "ExHyperlink9Container has a null hyperlink reference".to_string(),
            ));
        }
        let (screen_tip, flags_index) = if children.len() == 3 {
            let tip = &children[1];
            if tip.record_type != PptRecordType::CString || tip.version != 0 || tip.instance != 0 {
                return Err(PptError::Corrupted(
                    "ScreenTipAtom has an invalid record header".to_string(),
                ));
            }
            (Some(parse_unicode_string(&tip.data)?), 2)
        } else {
            (None, 1)
        };
        let flags = &children[flags_index];
        if flags.record_type != PptRecordType::ExternalHyperlinkFlagsAtom
            || flags.version != 0
            || flags.instance != 0
            || flags.data.len() != 4
        {
            return Err(PptError::Corrupted(
                "ExHyperlinkFlagsAtom has an invalid record header or size".to_string(),
            ));
        }
        let value =
            u32::from_le_bytes([flags.data[0], flags.data[1], flags.data[2], flags.data[3]]);
        if value & !0x07 != 0 {
            return Err(PptError::Corrupted(
                "ExHyperlinkFlagsAtom has nonzero reserved bits".to_string(),
            ));
        }
        Ok((
            reference,
            Self {
                screen_tip,
                inserted_with_dialog: value & 0x01 != 0,
                location_is_named_show: value & 0x02 != 0,
                named_show_returns_to_slide: value & 0x04 != 0,
            },
        ))
    }
}

/// One base PowerPoint hyperlink definition.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct PowerPointHyperlink {
    /// Positive identifier referenced by interactive information records.
    pub id: u32,
    /// Optional user-readable hyperlink name.
    pub friendly_name: Option<String>,
    /// Optional full destination-file path or URL.
    pub target: Option<String>,
    /// Optional location within the destination.
    pub location: Option<String>,
    /// Optional PowerPoint 9 metadata for this hyperlink.
    pub extension: Option<PowerPointHyperlinkExtension>,
}

impl PowerPointHyperlink {
    /// Parse an `ExHyperlinkContainer` record.
    pub fn parse(record: &PptRecord) -> Result<Self> {
        if record.record_type != PptRecordType::ExternalHyperlink
            || record.version != 0x0f
            || record.instance != 0
        {
            return Err(PptError::Corrupted(
                "ExHyperlinkContainer has an invalid record header".to_string(),
            ));
        }
        let children = PptRecord::parse_sequence_strict(&record.data, "external hyperlink")?;
        let Some(atom) = children.first() else {
            return Err(PptError::Corrupted(
                "ExHyperlinkContainer is missing ExHyperlinkAtom".to_string(),
            ));
        };
        let id = parse_hyperlink_atom(atom)?;
        if id == 0 {
            return Err(PptError::Corrupted(
                "ExHyperlinkAtom has a zero hyperlink ID".to_string(),
            ));
        }

        let mut friendly_name = None;
        let mut target = None;
        let mut location = None;
        let mut previous_instance = None;
        for child in &children[1..] {
            if child.record_type != PptRecordType::CString || child.version != 0 {
                return Err(PptError::Corrupted(
                    "ExHyperlinkContainer has an unexpected child record".to_string(),
                ));
            }
            if previous_instance.is_some_and(|previous| previous >= child.instance) {
                return Err(PptError::Corrupted(
                    "Hyperlink strings are duplicated or out of order".to_string(),
                ));
            }
            previous_instance = Some(child.instance);
            let value = Some(parse_unicode_string(&child.data)?);
            match child.instance {
                0 => friendly_name = value,
                1 => target = value,
                3 => location = value,
                _ => {
                    return Err(PptError::Corrupted(
                        "Hyperlink CString has an invalid record instance".to_string(),
                    ));
                },
            }
        }
        Ok(Self {
            id,
            friendly_name,
            target,
            location,
            extension: None,
        })
    }
}

/// Hyperlink definitions resolved with their PowerPoint 9 extensions.
#[derive(Debug, Clone, Default, PartialEq, Eq)]
pub struct PowerPointHyperlinks {
    /// Seed used when allocating new external-object or hyperlink identifiers.
    pub id_seed: Option<i32>,
    /// Hyperlinks in base `ExObjListContainer` order.
    pub hyperlinks: Vec<PowerPointHyperlink>,
}

impl PowerPointHyperlinks {
    /// Discover base hyperlinks and merge all `___PPT9` hyperlink extensions.
    pub fn parse(root: &PptRecord) -> Result<Self> {
        let mut lists = Vec::new();
        collect_records(root, PptRecordType::ExObjList, &mut lists);
        if lists.len() > 1 {
            return Err(PptError::Corrupted(
                "Record tree contains multiple external-object lists".to_string(),
            ));
        }
        let mut result = if let Some(list) = lists.first() {
            Self::parse_external_object_list(list)?
        } else {
            Self::default()
        };

        let mut extension_ids = Vec::new();
        for record in root.versioned_binary_tag_records(9)? {
            if record.record_type != PptRecordType::ExternalHyperlink9 {
                continue;
            }
            let (id, extension) = PowerPointHyperlinkExtension::parse(&record)?;
            if extension_ids.contains(&id) {
                return Err(PptError::Corrupted(
                    "PowerPoint 9 contains duplicate hyperlink extensions".to_string(),
                ));
            }
            extension_ids.push(id);
            let hyperlink = result.get_mut(id).ok_or_else(|| {
                PptError::Corrupted(
                    "PowerPoint 9 hyperlink extension references an unknown hyperlink".to_string(),
                )
            })?;
            hyperlink.extension = Some(extension);
        }
        Ok(result)
    }

    /// Resolve a hyperlink identifier.
    pub fn get(&self, id: u32) -> Option<&PowerPointHyperlink> {
        self.hyperlinks.iter().find(|hyperlink| hyperlink.id == id)
    }

    fn get_mut(&mut self, id: u32) -> Option<&mut PowerPointHyperlink> {
        self.hyperlinks
            .iter_mut()
            .find(|hyperlink| hyperlink.id == id)
    }

    fn parse_external_object_list(record: &PptRecord) -> Result<Self> {
        if record.record_type != PptRecordType::ExObjList
            || record.version != 0x0f
            || record.instance != 0
        {
            return Err(PptError::Corrupted(
                "ExObjListContainer has an invalid record header".to_string(),
            ));
        }
        let children = PptRecord::parse_sequence_strict(&record.data, "external-object list")?;
        let Some(atom) = children.first() else {
            return Err(PptError::Corrupted(
                "ExObjListContainer is missing ExObjListAtom".to_string(),
            ));
        };
        if atom.record_type != PptRecordType::ExObjListAtom
            || atom.version != 0
            || atom.instance != 0
            || atom.data.len() != 4
        {
            return Err(PptError::Corrupted(
                "ExObjListAtom has an invalid record header or size".to_string(),
            ));
        }
        let id_seed = i32::from_le_bytes([atom.data[0], atom.data[1], atom.data[2], atom.data[3]]);
        if id_seed < 1 {
            return Err(PptError::Corrupted(
                "ExObjListAtom has an invalid identifier seed".to_string(),
            ));
        }

        let mut hyperlinks = Vec::new();
        for child in &children[1..] {
            if child.record_type != PptRecordType::ExternalHyperlink {
                continue;
            }
            let hyperlink = PowerPointHyperlink::parse(child)?;
            if hyperlinks
                .iter()
                .any(|existing: &PowerPointHyperlink| existing.id == hyperlink.id)
            {
                return Err(PptError::Corrupted(
                    "External-object list has duplicate hyperlink IDs".to_string(),
                ));
            }
            hyperlinks.push(hyperlink);
        }
        if hyperlinks
            .iter()
            .any(|hyperlink| hyperlink.id > id_seed as u32)
        {
            return Err(PptError::Corrupted(
                "External-object identifier seed is below a hyperlink ID".to_string(),
            ));
        }
        Ok(Self {
            id_seed: Some(id_seed),
            hyperlinks,
        })
    }
}

fn parse_hyperlink_atom(record: &PptRecord) -> Result<u32> {
    if record.record_type != PptRecordType::ExternalHyperlinkAtom
        || record.version != 0
        || record.instance != 0
        || record.data.len() != 4
    {
        return Err(PptError::Corrupted(
            "ExHyperlinkAtom has an invalid record header or size".to_string(),
        ));
    }
    Ok(u32::from_le_bytes([
        record.data[0],
        record.data[1],
        record.data[2],
        record.data[3],
    ]))
}

fn parse_unicode_string(data: &[u8]) -> Result<String> {
    if data.len() & 1 != 0 {
        return Err(PptError::Corrupted(
            "Hyperlink string has an odd byte length".to_string(),
        ));
    }
    let mut units = Vec::with_capacity(data.len() / 2);
    for bytes in data.chunks_exact(2) {
        let unit = u16::from_le_bytes([bytes[0], bytes[1]]);
        if unit == 0 {
            break;
        }
        units.push(unit);
    }
    String::from_utf16(&units)
        .map_err(|_| PptError::Corrupted("Hyperlink string is invalid UTF-16".to_string()))
}

fn collect_records<'a>(
    record: &'a PptRecord,
    record_type: PptRecordType,
    records: &mut Vec<&'a PptRecord>,
) {
    if record.record_type == record_type {
        records.push(record);
    }
    for child in &record.children {
        collect_records(child, record_type, records);
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    fn record_bytes(version: u16, instance: u16, kind: u16, payload: &[u8]) -> Vec<u8> {
        let mut data = Vec::new();
        data.extend_from_slice(&((instance << 4) | version).to_le_bytes());
        data.extend_from_slice(&kind.to_le_bytes());
        data.extend_from_slice(&(payload.len() as u32).to_le_bytes());
        data.extend_from_slice(payload);
        data
    }

    fn unicode(value: &str) -> Vec<u8> {
        value.encode_utf16().flat_map(u16::to_le_bytes).collect()
    }

    fn hyperlink(id: u32) -> Vec<u8> {
        let mut payload = record_bytes(0, 0, 4051, &id.to_le_bytes());
        payload.extend_from_slice(&record_bytes(0, 0, 4026, &unicode("Example")));
        payload.extend_from_slice(&record_bytes(0, 1, 4026, &unicode("https://example.test")));
        payload.extend_from_slice(&record_bytes(0, 3, 4026, &unicode("section")));
        record_bytes(0x0f, 0, 4055, &payload)
    }

    fn external_object_list(seed: i32, hyperlinks: &[Vec<u8>]) -> PptRecord {
        let mut payload = record_bytes(0, 0, 1034, &seed.to_le_bytes());
        for hyperlink in hyperlinks {
            payload.extend_from_slice(hyperlink);
        }
        PptRecord {
            record_type: PptRecordType::ExObjList,
            record_type_raw: 1033,
            version: 0x0f,
            instance: 0,
            data_length: payload.len() as u32,
            data: payload,
            children: Vec::new(),
        }
    }

    fn hyperlink9(id: u32, screen_tip: Option<&str>, flags: u32) -> Vec<u8> {
        let mut payload = record_bytes(0, 0, 4051, &id.to_le_bytes());
        if let Some(screen_tip) = screen_tip {
            payload.extend_from_slice(&record_bytes(0, 0, 4026, &unicode(screen_tip)));
        }
        payload.extend_from_slice(&record_bytes(0, 0, 4120, &flags.to_le_bytes()));
        record_bytes(0x0f, 0, 4068, &payload)
    }

    fn prog_tags_record(blob_payload: &[u8]) -> PptRecord {
        let tag_name: Vec<u8> = "___PPT9"
            .encode_utf16()
            .flat_map(u16::to_le_bytes)
            .collect();
        let name = record_bytes(0, 0, 4026, &tag_name);
        let blob = record_bytes(0, 0, 0x138b, blob_payload);
        let mut tag_payload = name;
        tag_payload.extend_from_slice(&blob);
        let tag = record_bytes(0x0f, 0, 0x138a, &tag_payload);
        PptRecord {
            record_type: PptRecordType::ProgTags,
            record_type_raw: 0x1388,
            version: 0x0f,
            instance: 0,
            data_length: tag.len() as u32,
            data: tag,
            children: Vec::new(),
        }
    }

    fn root(list: Option<PptRecord>, extensions: &[Vec<u8>]) -> PptRecord {
        let mut children = Vec::new();
        if let Some(list) = list {
            children.push(list);
        }
        if !extensions.is_empty() {
            let blob: Vec<u8> = extensions.iter().flatten().copied().collect();
            children.push(prog_tags_record(&blob));
        }
        PptRecord {
            record_type: PptRecordType::Document,
            record_type_raw: 1000,
            version: 0x0f,
            instance: 0,
            data_length: 0,
            data: Vec::new(),
            children,
        }
    }

    #[test]
    fn parses_and_merges_powerpoint9_hyperlinks() {
        let root = root(
            Some(external_object_list(7, &[hyperlink(3)])),
            &[hyperlink9(3, Some("Open example"), 7)],
        );
        let hyperlinks = PowerPointHyperlinks::parse(&root).unwrap();
        assert_eq!(hyperlinks.id_seed, Some(7));
        let hyperlink = hyperlinks.get(3).unwrap();
        assert_eq!(hyperlink.friendly_name.as_deref(), Some("Example"));
        assert_eq!(hyperlink.target.as_deref(), Some("https://example.test"));
        assert_eq!(hyperlink.location.as_deref(), Some("section"));
        let extension = hyperlink.extension.as_ref().unwrap();
        assert_eq!(extension.screen_tip.as_deref(), Some("Open example"));
        assert!(extension.inserted_with_dialog);
        assert!(extension.location_is_named_show);
        assert!(extension.named_show_returns_to_slide);
    }

    #[test]
    fn accepts_optional_base_strings_and_absent_extensions() {
        let atom_only = record_bytes(
            0x0f,
            0,
            4055,
            &record_bytes(0, 0, 4051, &1u32.to_le_bytes()),
        );
        let hyperlinks =
            PowerPointHyperlinks::parse(&root(Some(external_object_list(1, &[atom_only])), &[]))
                .unwrap();
        assert_eq!(hyperlinks.get(1).unwrap().target, None);
    }

    #[test]
    fn rejects_invalid_hyperlink_ids_and_extensions() {
        assert!(
            PowerPointHyperlinks::parse(
                &root(Some(external_object_list(2, &[hyperlink(3)])), &[],)
            )
            .is_err()
        );
        assert!(
            PowerPointHyperlinks::parse(&root(
                Some(external_object_list(3, &[hyperlink(3), hyperlink(3)])),
                &[],
            ))
            .is_err()
        );
        assert!(
            PowerPointHyperlinks::parse(&root(
                Some(external_object_list(3, &[hyperlink(3)])),
                &[hyperlink9(4, None, 0)],
            ))
            .is_err()
        );
        assert!(
            PowerPointHyperlinks::parse(&root(
                Some(external_object_list(3, &[hyperlink(3)])),
                &[hyperlink9(3, None, 8)],
            ))
            .is_err()
        );
        assert!(
            PowerPointHyperlinks::parse(&root(
                Some(external_object_list(3, &[hyperlink(3)])),
                &[hyperlink9(3, None, 0), hyperlink9(3, None, 0)],
            ))
            .is_err()
        );
    }

    #[test]
    fn rejects_malformed_hyperlink_strings_and_child_order() {
        let mut invalid_utf16 = hyperlink(1);
        invalid_utf16[28..30].copy_from_slice(&0xd800u16.to_le_bytes());
        assert!(
            PowerPointHyperlinks::parse(&root(
                Some(external_object_list(1, &[invalid_utf16])),
                &[],
            ))
            .is_err()
        );

        let mut payload = record_bytes(0, 0, 4051, &1u32.to_le_bytes());
        payload.extend_from_slice(&record_bytes(0, 3, 4026, &unicode("late")));
        payload.extend_from_slice(&record_bytes(0, 1, 4026, &unicode("early")));
        let out_of_order = record_bytes(0x0f, 0, 4055, &payload);
        assert!(
            PowerPointHyperlinks::parse(
                &root(Some(external_object_list(1, &[out_of_order])), &[],)
            )
            .is_err()
        );
    }
}
