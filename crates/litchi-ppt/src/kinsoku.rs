//! East Asian line-breaking preferences in legacy `PowerPoint` files.

use super::package::{Error, Result};
use super::records::Record;
use crate::consts::RecordType;

/// Language whose East Asian line-breaking behavior is being queried.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
#[allow(
    clippy::module_name_repetitions,
    reason = "`KinsokuLanguage` is the established public API name; renaming it would break downstream crates"
)]
pub enum KinsokuLanguage {
    /// Korean.
    Korean,
    /// Simplified Chinese.
    SimplifiedChinese,
    /// Traditional Chinese.
    TraditionalChinese,
    /// Japanese.
    Japanese,
}

/// Line-breaking level stored by `PowerPoint`.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
#[allow(
    clippy::module_name_repetitions,
    reason = "`KinsokuLevel` is the established public API name; renaming it would break downstream crates"
)]
pub enum KinsokuLevel {
    /// Use standard line-breaking settings.
    Standard,
    /// Use strict Japanese line-breaking settings.
    Strict,
    /// Use the document's custom leading and following character lists.
    Custom,
}

impl KinsokuLevel {
    fn parse_base(value: u32) -> Result<Self> {
        match value {
            0 => Ok(Self::Standard),
            1 => Ok(Self::Strict),
            2 => Ok(Self::Custom),
            _ => Err(Error::Corrupted(
                "KinsokuAtom has an invalid line-breaking level".to_string(),
            )),
        }
    }

    fn parse_language(value: u8, japanese: bool) -> Result<Self> {
        match value {
            0 => Ok(Self::Standard),
            1 if japanese => Ok(Self::Strict),
            2 => Ok(Self::Custom),
            _ => Err(Error::Corrupted(
                "Kinsoku9Atom has an invalid language level".to_string(),
            )),
        }
    }
}

/// Base `KinsokuContainer` settings from `DocumentTextInfoContainer`.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct BaseKinsokuSettings {
    /// Base line-breaking level.
    pub level: KinsokuLevel,
    /// Characters that cannot immediately follow a line break.
    pub leading_characters: Option<String>,
    /// Characters that cannot immediately precede a line break.
    pub following_characters: Option<String>,
}

impl BaseKinsokuSettings {
    /// Parse a base `KinsokuContainer` record.
    ///
    /// # Errors
    ///
    /// Returns an error if the input cannot be read or is malformed.
    pub fn parse(record: &Record) -> Result<Self> {
        parse_base(record)
    }
}

/// Per-language settings introduced by `PowerPoint` 9.
#[derive(Debug, Clone, PartialEq, Eq)]
#[allow(
    clippy::module_name_repetitions,
    reason = "`KinsokuSettings9` is the established public API name; renaming it would break downstream crates"
)]
pub struct KinsokuSettings9 {
    /// Korean line-breaking level.
    pub korean: KinsokuLevel,
    /// Simplified Chinese line-breaking level.
    pub simplified_chinese: KinsokuLevel,
    /// Traditional Chinese line-breaking level.
    pub traditional_chinese: KinsokuLevel,
    /// Japanese line-breaking level.
    pub japanese: KinsokuLevel,
    /// `PowerPoint` 9 custom leading-character list, when not supplied by the base container.
    pub leading_characters: Option<String>,
    /// `PowerPoint` 9 custom following-character list, when not supplied by the base container.
    pub following_characters: Option<String>,
}

impl KinsokuSettings9 {
    /// Parse a `PowerPoint` 9 `Kinsoku9Container` record.
    ///
    /// # Errors
    ///
    /// Returns an error if the input cannot be read or is malformed.
    pub fn parse(record: &Record) -> Result<Self> {
        parse_powerpoint9(record)
    }

    /// Return the stored level for one language.
    #[must_use]
    pub const fn level(&self, language: KinsokuLanguage) -> KinsokuLevel {
        match language {
            KinsokuLanguage::Korean => self.korean,
            KinsokuLanguage::SimplifiedChinese => self.simplified_chinese,
            KinsokuLanguage::TraditionalChinese => self.traditional_chinese,
            KinsokuLanguage::Japanese => self.japanese,
        }
    }

    fn custom_count(&self) -> usize {
        [
            self.korean,
            self.simplified_chinese,
            self.traditional_chinese,
            self.japanese,
        ]
        .into_iter()
        .filter(|level| *level == KinsokuLevel::Custom)
        .count()
    }
}

/// Resolved base and `PowerPoint` 9 East Asian line-breaking preferences.
#[derive(Debug, Clone, Default, PartialEq, Eq)]
pub struct Kinsoku {
    /// Base document settings.
    pub base: Option<BaseKinsokuSettings>,
    /// `PowerPoint` 9 per-language settings, which take precedence over the base.
    pub powerpoint9: Option<KinsokuSettings9>,
}

impl Kinsoku {
    /// Discover, parse, and cross-validate line-breaking records below `root`.
    ///
    /// # Errors
    ///
    /// Returns an error if the input cannot be read or is malformed.
    pub fn parse(root: &Record) -> Result<Self> {
        let mut base_records = Vec::new();
        collect_records(root, RecordType::Kinsoku, &mut base_records);
        if base_records.len() > 1 {
            return Err(Error::Corrupted(
                "Record tree contains multiple base Kinsoku containers".to_string(),
            ));
        }
        let base = base_records
            .first()
            .map(|record| BaseKinsokuSettings::parse(record))
            .transpose()?;

        let mut powerpoint9 = None;
        for record in root.versioned_binary_tag_records(9)? {
            if record.record_type != RecordType::Kinsoku {
                continue;
            }
            if powerpoint9
                .replace(KinsokuSettings9::parse(&record)?)
                .is_some()
            {
                return Err(Error::Corrupted(
                    "Record tree contains multiple PowerPoint 9 Kinsoku containers".to_string(),
                ));
            }
        }

        validate_cross_version(base.as_ref(), powerpoint9.as_ref())?;
        Ok(Self { base, powerpoint9 })
    }

    /// Return the effective setting for one language.
    #[must_use]
    pub fn effective_level(&self, language: KinsokuLanguage) -> KinsokuLevel {
        if let Some(settings) = &self.powerpoint9 {
            return settings.level(language);
        }
        match self.base.as_ref().map(|settings| settings.level) {
            Some(KinsokuLevel::Strict) if language == KinsokuLanguage::Japanese => {
                KinsokuLevel::Strict
            },
            Some(KinsokuLevel::Custom) => KinsokuLevel::Custom,
            _ => KinsokuLevel::Standard,
        }
    }

    /// Return the effective custom leading and following lists for a language.
    #[must_use]
    pub fn effective_custom_characters(&self, language: KinsokuLanguage) -> Option<(&str, &str)> {
        if self.effective_level(language) != KinsokuLevel::Custom {
            return None;
        }
        if let Some(base) = &self.base
            && base.level == KinsokuLevel::Custom
        {
            return Some((
                base.leading_characters.as_deref()?,
                base.following_characters.as_deref()?,
            ));
        }
        let settings = self.powerpoint9.as_ref()?;
        Some((
            settings.leading_characters.as_deref()?,
            settings.following_characters.as_deref()?,
        ))
    }
}

fn parse_base(record: &Record) -> Result<BaseKinsokuSettings> {
    validate_container_header(record)?;
    let children = Record::parse_sequence_strict(&record.data, "base Kinsoku container")?;
    let (level_word, leading_characters, following_characters) = parse_children(&children, false)?;
    let level = KinsokuLevel::parse_base(level_word)?;
    let has_both_lists = leading_characters.is_some() && following_characters.is_some();
    if (level == KinsokuLevel::Custom) != has_both_lists {
        return Err(Error::Corrupted(
            "Base Kinsoku custom lists do not match its level".to_string(),
        ));
    }
    Ok(BaseKinsokuSettings {
        level,
        leading_characters,
        following_characters,
    })
}

fn parse_powerpoint9(record: &Record) -> Result<KinsokuSettings9> {
    validate_container_header(record)?;
    let children = Record::parse_sequence_strict(&record.data, "PowerPoint 9 Kinsoku container")?;
    let (levels, leading_characters, following_characters) = parse_children(&children, true)?;
    if levels & 0xffff_ff00 != 0 {
        return Err(Error::Corrupted(
            "Kinsoku9Atom has nonzero reserved bits".to_string(),
        ));
    }
    let settings = KinsokuSettings9 {
        korean: KinsokuLevel::parse_language((levels & 0x03) as u8, false)?,
        simplified_chinese: KinsokuLevel::parse_language(((levels >> 2) & 0x03) as u8, false)?,
        traditional_chinese: KinsokuLevel::parse_language(((levels >> 4) & 0x03) as u8, false)?,
        japanese: KinsokuLevel::parse_language(((levels >> 6) & 0x03) as u8, true)?,
        leading_characters,
        following_characters,
    };
    if settings.custom_count() > 1 {
        return Err(Error::Corrupted(
            "Kinsoku9Atom customizes more than one language".to_string(),
        ));
    }
    Ok(settings)
}

fn validate_container_header(record: &Record) -> Result<()> {
    if record.record_type != RecordType::Kinsoku || record.version != 0x0f || record.instance != 2 {
        return Err(Error::Corrupted(
            "KinsokuContainer has an invalid record header".to_string(),
        ));
    }
    Ok(())
}

fn parse_children(
    children: &[Record],
    powerpoint9: bool,
) -> Result<(u32, Option<String>, Option<String>)> {
    let Some(atom) = children.first() else {
        return Err(Error::Corrupted(
            "KinsokuContainer is missing its settings atom".to_string(),
        ));
    };
    if atom.record_type != RecordType::KinsokuAtom
        || atom.version != 0
        || atom.instance != 3
        || atom.data.len() != 4
    {
        return Err(Error::Corrupted(
            "Kinsoku settings atom has an invalid header or size".to_string(),
        ));
    }
    let level_word = u32::from_le_bytes([atom.data[0], atom.data[1], atom.data[2], atom.data[3]]);
    let mut leading = None;
    let mut following = None;
    for child in &children[1..] {
        if child.record_type != RecordType::CString || child.version != 0 {
            return Err(Error::Corrupted(
                "KinsokuContainer has an unexpected child record".to_string(),
            ));
        }
        let target = match child.instance {
            0 => &mut leading,
            1 if leading.is_some() => &mut following,
            _ => {
                return Err(Error::Corrupted(
                    "Kinsoku character lists are duplicated or out of order".to_string(),
                ));
            },
        };
        if target.is_some() {
            return Err(Error::Corrupted(
                "Kinsoku character list is duplicated".to_string(),
            ));
        }
        *target = Some(parse_utf16(&child.data)?);
    }
    if leading.is_some() != following.is_some() {
        return Err(Error::Corrupted(
            "KinsokuContainer contains only one custom character list".to_string(),
        ));
    }
    if !powerpoint9 && children.len() > 3 {
        return Err(Error::Corrupted(
            "Base KinsokuContainer has too many child records".to_string(),
        ));
    }
    Ok((level_word, leading, following))
}

fn parse_utf16(data: &[u8]) -> Result<String> {
    if data.len() & 1 != 0 {
        return Err(Error::Corrupted(
            "Kinsoku character list has an odd byte length".to_string(),
        ));
    }
    let units: Vec<u16> = data
        .chunks_exact(2)
        .map(|bytes| u16::from_le_bytes([bytes[0], bytes[1]]))
        .collect();
    String::from_utf16(&units)
        .map_err(|_err| Error::Corrupted("Kinsoku character list is invalid UTF-16".to_string()))
}

fn validate_cross_version(
    base: Option<&BaseKinsokuSettings>,
    powerpoint9: Option<&KinsokuSettings9>,
) -> Result<()> {
    let Some(extension) = powerpoint9 else {
        return Ok(());
    };
    let extension_has_lists =
        extension.leading_characters.is_some() && extension.following_characters.is_some();
    let base_supplies_lists = base.is_some_and(|settings| settings.level == KinsokuLevel::Custom);
    let extension_needs_lists = extension.custom_count() == 1 && !base_supplies_lists;
    if extension_has_lists != extension_needs_lists {
        return Err(Error::Corrupted(
            "PowerPoint 9 Kinsoku custom lists conflict with the base settings".to_string(),
        ));
    }
    Ok(())
}

fn collect_records<'a>(record: &'a Record, record_type: RecordType, records: &mut Vec<&'a Record>) {
    if record.record_type == record_type {
        records.push(record);
    }
    for child in &record.children {
        collect_records(child, record_type, records);
    }
}

#[cfg(test)]
#[allow(
    clippy::unwrap_used,
    clippy::expect_used,
    reason = "test assertions panic on failure by design"
)]
mod tests {
    use super::*;

    fn record_bytes(version: u16, instance: u16, kind: u16, payload: &[u8]) -> Vec<u8> {
        let mut data = Vec::new();
        data.extend_from_slice(&((instance << 4) | version).to_le_bytes());
        data.extend_from_slice(&kind.to_le_bytes());
        data.extend_from_slice(&u32::try_from(payload.len()).unwrap().to_le_bytes());
        data.extend_from_slice(payload);
        data
    }

    fn container(levels: u32, lists: Option<(&str, &str)>) -> Record {
        let mut payload = record_bytes(0, 3, 4050, &levels.to_le_bytes());
        if let Some((leading, following)) = lists {
            let leading_bytes: Vec<u8> =
                leading.encode_utf16().flat_map(u16::to_le_bytes).collect();
            let following_bytes: Vec<u8> = following
                .encode_utf16()
                .flat_map(u16::to_le_bytes)
                .collect();
            payload.extend_from_slice(&record_bytes(0, 0, 4026, &leading_bytes));
            payload.extend_from_slice(&record_bytes(0, 1, 4026, &following_bytes));
        }
        Record {
            record_type: RecordType::Kinsoku,
            record_type_raw: 4040,
            version: 0x0f,
            instance: 2,
            data_length: u32::try_from(payload.len()).unwrap(),
            data: payload,
            children: Vec::new(),
        }
    }

    fn prog_tags_record(blob_payload: &[u8]) -> Record {
        let tag_name: Vec<u8> = "___PPT9"
            .encode_utf16()
            .flat_map(u16::to_le_bytes)
            .collect();
        let name = record_bytes(0, 0, 4026, &tag_name);
        let blob = record_bytes(0, 0, 0x138b, blob_payload);
        let mut tag_payload = name;
        tag_payload.extend_from_slice(&blob);
        let tag = record_bytes(0x0f, 0, 0x138a, &tag_payload);
        Record {
            record_type: RecordType::ProgTags,
            record_type_raw: 0x1388,
            version: 0x0f,
            instance: 0,
            data_length: u32::try_from(tag.len()).unwrap(),
            data: tag,
            children: Vec::new(),
        }
    }

    fn root(base: Option<Record>, extension: Option<Record>) -> Record {
        let mut children = Vec::new();
        if let Some(base_record) = base {
            children.push(base_record);
        }
        if let Some(extension_record) = extension {
            let bytes = record_bytes(
                extension_record.version,
                extension_record.instance,
                extension_record.record_type_raw,
                &extension_record.data,
            );
            children.push(prog_tags_record(&bytes));
        }
        Record {
            record_type: RecordType::Document,
            record_type_raw: 1000,
            version: 0x0f,
            instance: 0,
            data_length: 0,
            data: Vec::new(),
            children,
        }
    }

    #[test]
    fn resolves_base_and_powerpoint9_precedence() {
        let settings = Kinsoku::parse(&root(Some(container(1, None)), None)).unwrap();
        assert_eq!(
            settings.effective_level(KinsokuLanguage::Japanese),
            KinsokuLevel::Strict
        );
        assert_eq!(
            settings.effective_level(KinsokuLanguage::Korean),
            KinsokuLevel::Standard
        );

        let custom_settings = Kinsoku::parse(&root(
            Some(container(2, Some(("（", "）")))),
            Some(container(2, None)),
        ))
        .unwrap();
        assert_eq!(
            custom_settings.effective_level(KinsokuLanguage::Korean),
            KinsokuLevel::Custom
        );
        assert_eq!(
            custom_settings.effective_custom_characters(KinsokuLanguage::Korean),
            Some(("（", "）"))
        );
        assert_eq!(
            custom_settings.effective_level(KinsokuLanguage::Japanese),
            KinsokuLevel::Standard
        );
    }

    #[test]
    fn resolves_powerpoint9_owned_custom_lists() {
        let japanese_custom = 2 << 6;
        let settings = Kinsoku::parse(&root(
            None,
            Some(container(japanese_custom, Some(("「", "」")))),
        ))
        .unwrap();
        assert_eq!(
            settings.effective_custom_characters(KinsokuLanguage::Japanese),
            Some(("「", "」"))
        );
    }

    #[test]
    fn rejects_malformed_kinsoku_settings() {
        assert!(Kinsoku::parse(&root(Some(container(3, None)), None)).is_err());
        assert!(Kinsoku::parse(&root(Some(container(2, None)), None)).is_err());
        assert!(Kinsoku::parse(&root(None, Some(container(1, None)))).is_err());
        assert!(Kinsoku::parse(&root(None, Some(container(2 | (2 << 2), None)))).is_err());
        assert!(Kinsoku::parse(&root(None, Some(container(1 << 8, None)))).is_err());
        assert!(Kinsoku::parse(&root(None, Some(container(0, Some(("A", "B")))))).is_err());
        assert!(
            Kinsoku::parse(&root(
                Some(container(2, Some(("A", "B")))),
                Some(container(2, Some(("C", "D")))),
            ))
            .is_err()
        );

        let mut invalid_utf16 = container(2, Some(("A", "B")));
        invalid_utf16.data[20..22].copy_from_slice(&0xd800u16.to_le_bytes());
        assert!(parse_base(&invalid_utf16).is_err());

        let mut missing_following = container(2, Some(("A", "B")));
        missing_following.data.truncate(22);
        missing_following.data_length = u32::try_from(missing_following.data.len()).unwrap();
        assert!(parse_base(&missing_following).is_err());

        let mut invalid_header = container(0, None);
        invalid_header.instance = 0;
        assert!(BaseKinsokuSettings::parse(&invalid_header).is_err());
    }
}
