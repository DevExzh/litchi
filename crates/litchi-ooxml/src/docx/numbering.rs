//! Immutable WordprocessingML numbering definitions.

use crate::docx::namespace::{is_wordprocessing_namespace, word_attribute_value};
use crate::error::{OoxmlError, Result};
use litchi_opc::part::Part;
use quick_xml::XmlVersion;
use quick_xml::encoding::Decoder;
use quick_xml::events::{BytesStart, Event};
use quick_xml::name::{Namespace, NamespaceResolver, ResolveResult};
use quick_xml::reader::NsReader;

const TRANSITIONAL_RELATIONSHIPS_NAMESPACE: &[u8] =
    b"http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const STRICT_RELATIONSHIPS_NAMESPACE: &[u8] =
    b"http://purl.oclc.org/ooxml/officeDocument/relationships";
const VML_NAMESPACE: &[u8] = b"urn:schemas-microsoft-com:vml";
const DRAWINGML_NAMESPACE: &[u8] = b"http://schemas.openxmlformats.org/drawingml/2006/main";
const STRICT_DRAWINGML_NAMESPACE: &[u8] = b"http://purl.oclc.org/ooxml/drawingml/main";

#[derive(Debug, Clone)]
pub struct Numbering {
    pub(crate) abstract_nums: Vec<AbstractNum>,
    pub(crate) nums: Vec<Num>,
    pub(crate) picture_bullets: Vec<PictureBullet>,
}

/// A picture bullet definition (`w:numPicBullet`) from `numbering.xml`.
///
/// The image itself lives in a package part referenced through a relationship;
/// only the inert relationship ID is captured here, never the image bytes.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct PictureBullet {
    id: u32,
    image_relationship_id: Option<String>,
}

impl PictureBullet {
    /// The `w:numPicBulletId` key referenced by `w:lvlPicBulletId` on a level.
    pub fn id(&self) -> u32 {
        self.id
    }

    /// Relationship ID of the bullet image, when the definition carries one.
    pub fn image_relationship_id(&self) -> Option<&str> {
        self.image_relationship_id.as_deref()
    }
}

#[derive(Debug, Clone)]
pub struct AbstractNum {
    pub(crate) id: u32,
    pub(crate) num_type: Option<String>,
    pub(crate) num_style_link: Option<String>,
    pub(crate) style_link: Option<String>,
    pub(crate) levels: Vec<NumberingLevel>,
}

#[derive(Debug, Clone)]
pub struct Num {
    pub(crate) id: u32,
    pub(crate) abstract_num_id: u32,
    pub(crate) overrides: Vec<LevelOverride>,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct ParagraphNumbering {
    pub num_id: u32,
    pub level: u8,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum LevelRestart {
    Default,
    Never,
    After(u8),
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum NumberingSuffix {
    Tab,
    Space,
    Nothing,
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub enum NumberFormat {
    Decimal,
    UpperRoman,
    LowerRoman,
    UpperLetter,
    LowerLetter,
    Ordinal,
    CardinalText,
    OrdinalText,
    Hex,
    Chicago,
    IdeographDigital,
    JapaneseCounting,
    Aiueo,
    Iroha,
    DecimalFullWidth,
    DecimalHalfWidth,
    JapaneseLegal,
    JapaneseDigitalTenThousand,
    DecimalEnclosedCircle,
    DecimalFullWidth2,
    AiueoFullWidth,
    IrohaFullWidth,
    DecimalZero,
    Bullet,
    Ganada,
    Chosung,
    DecimalEnclosedFullStop,
    DecimalEnclosedParen,
    DecimalEnclosedCircleChinese,
    IdeographEnclosedCircle,
    IdeographTraditional,
    IdeographZodiac,
    IdeographZodiacTraditional,
    TaiwaneseCounting,
    IdeographLegalTraditional,
    TaiwaneseCountingThousand,
    TaiwaneseDigital,
    ChineseCounting,
    ChineseLegalSimplified,
    ChineseCountingThousand,
    KoreanDigital,
    KoreanCounting,
    KoreanLegal,
    KoreanDigital2,
    VietnameseCounting,
    RussianLower,
    RussianUpper,
    None,
    NumberInDash,
    Hebrew1,
    Hebrew2,
    ArabicAlpha,
    ArabicAbjad,
    HindiVowels,
    HindiConsonants,
    HindiNumbers,
    HindiCounting,
    ThaiLetters,
    ThaiNumbers,
    ThaiCounting,
    Custom,
    /// A format token outside the standardized `ST_NumberFormat` value set.
    Other(String),
}

impl NumberFormat {
    pub(crate) fn parse(value: &str) -> Self {
        match value {
            "decimal" => Self::Decimal,
            "upperRoman" => Self::UpperRoman,
            "lowerRoman" => Self::LowerRoman,
            "upperLetter" => Self::UpperLetter,
            "lowerLetter" => Self::LowerLetter,
            "ordinal" => Self::Ordinal,
            "cardinalText" => Self::CardinalText,
            "ordinalText" => Self::OrdinalText,
            "hex" => Self::Hex,
            "chicago" => Self::Chicago,
            "ideographDigital" => Self::IdeographDigital,
            "japaneseCounting" => Self::JapaneseCounting,
            "aiueo" => Self::Aiueo,
            "iroha" => Self::Iroha,
            "decimalFullWidth" => Self::DecimalFullWidth,
            "decimalHalfWidth" => Self::DecimalHalfWidth,
            "japaneseLegal" => Self::JapaneseLegal,
            "japaneseDigitalTenThousand" => Self::JapaneseDigitalTenThousand,
            "decimalEnclosedCircle" => Self::DecimalEnclosedCircle,
            "decimalFullWidth2" => Self::DecimalFullWidth2,
            "aiueoFullWidth" => Self::AiueoFullWidth,
            "irohaFullWidth" => Self::IrohaFullWidth,
            "decimalZero" => Self::DecimalZero,
            "bullet" => Self::Bullet,
            "ganada" => Self::Ganada,
            "chosung" => Self::Chosung,
            "decimalEnclosedFullstop" => Self::DecimalEnclosedFullStop,
            "decimalEnclosedParen" => Self::DecimalEnclosedParen,
            "decimalEnclosedCircleChinese" => Self::DecimalEnclosedCircleChinese,
            "ideographEnclosedCircle" => Self::IdeographEnclosedCircle,
            "ideographTraditional" => Self::IdeographTraditional,
            "ideographZodiac" => Self::IdeographZodiac,
            "ideographZodiacTraditional" => Self::IdeographZodiacTraditional,
            "taiwaneseCounting" => Self::TaiwaneseCounting,
            "ideographLegalTraditional" => Self::IdeographLegalTraditional,
            "taiwaneseCountingThousand" => Self::TaiwaneseCountingThousand,
            "taiwaneseDigital" => Self::TaiwaneseDigital,
            "chineseCounting" => Self::ChineseCounting,
            "chineseLegalSimplified" => Self::ChineseLegalSimplified,
            "chineseCountingThousand" => Self::ChineseCountingThousand,
            "koreanDigital" => Self::KoreanDigital,
            "koreanCounting" => Self::KoreanCounting,
            "koreanLegal" => Self::KoreanLegal,
            "koreanDigital2" => Self::KoreanDigital2,
            "vietnameseCounting" => Self::VietnameseCounting,
            "russianLower" => Self::RussianLower,
            "russianUpper" => Self::RussianUpper,
            "none" => Self::None,
            "numberInDash" => Self::NumberInDash,
            "hebrew1" => Self::Hebrew1,
            "hebrew2" => Self::Hebrew2,
            "arabicAlpha" => Self::ArabicAlpha,
            "arabicAbjad" => Self::ArabicAbjad,
            "hindiVowels" => Self::HindiVowels,
            "hindiConsonants" => Self::HindiConsonants,
            "hindiNumbers" => Self::HindiNumbers,
            "hindiCounting" => Self::HindiCounting,
            "thaiLetters" => Self::ThaiLetters,
            "thaiNumbers" => Self::ThaiNumbers,
            "thaiCounting" => Self::ThaiCounting,
            "custom" => Self::Custom,
            other => Self::Other(other.to_owned()),
        }
    }

    pub fn as_str(&self) -> &str {
        match self {
            Self::Decimal => "decimal",
            Self::UpperRoman => "upperRoman",
            Self::LowerRoman => "lowerRoman",
            Self::UpperLetter => "upperLetter",
            Self::LowerLetter => "lowerLetter",
            Self::Ordinal => "ordinal",
            Self::CardinalText => "cardinalText",
            Self::OrdinalText => "ordinalText",
            Self::Hex => "hex",
            Self::Chicago => "chicago",
            Self::IdeographDigital => "ideographDigital",
            Self::JapaneseCounting => "japaneseCounting",
            Self::Aiueo => "aiueo",
            Self::Iroha => "iroha",
            Self::DecimalFullWidth => "decimalFullWidth",
            Self::DecimalHalfWidth => "decimalHalfWidth",
            Self::JapaneseLegal => "japaneseLegal",
            Self::JapaneseDigitalTenThousand => "japaneseDigitalTenThousand",
            Self::DecimalEnclosedCircle => "decimalEnclosedCircle",
            Self::DecimalFullWidth2 => "decimalFullWidth2",
            Self::AiueoFullWidth => "aiueoFullWidth",
            Self::IrohaFullWidth => "irohaFullWidth",
            Self::DecimalZero => "decimalZero",
            Self::Bullet => "bullet",
            Self::Ganada => "ganada",
            Self::Chosung => "chosung",
            Self::DecimalEnclosedFullStop => "decimalEnclosedFullstop",
            Self::DecimalEnclosedParen => "decimalEnclosedParen",
            Self::DecimalEnclosedCircleChinese => "decimalEnclosedCircleChinese",
            Self::IdeographEnclosedCircle => "ideographEnclosedCircle",
            Self::IdeographTraditional => "ideographTraditional",
            Self::IdeographZodiac => "ideographZodiac",
            Self::IdeographZodiacTraditional => "ideographZodiacTraditional",
            Self::TaiwaneseCounting => "taiwaneseCounting",
            Self::IdeographLegalTraditional => "ideographLegalTraditional",
            Self::TaiwaneseCountingThousand => "taiwaneseCountingThousand",
            Self::TaiwaneseDigital => "taiwaneseDigital",
            Self::ChineseCounting => "chineseCounting",
            Self::ChineseLegalSimplified => "chineseLegalSimplified",
            Self::ChineseCountingThousand => "chineseCountingThousand",
            Self::KoreanDigital => "koreanDigital",
            Self::KoreanCounting => "koreanCounting",
            Self::KoreanLegal => "koreanLegal",
            Self::KoreanDigital2 => "koreanDigital2",
            Self::VietnameseCounting => "vietnameseCounting",
            Self::RussianLower => "russianLower",
            Self::RussianUpper => "russianUpper",
            Self::None => "none",
            Self::NumberInDash => "numberInDash",
            Self::Hebrew1 => "hebrew1",
            Self::Hebrew2 => "hebrew2",
            Self::ArabicAlpha => "arabicAlpha",
            Self::ArabicAbjad => "arabicAbjad",
            Self::HindiVowels => "hindiVowels",
            Self::HindiConsonants => "hindiConsonants",
            Self::HindiNumbers => "hindiNumbers",
            Self::HindiCounting => "hindiCounting",
            Self::ThaiLetters => "thaiLetters",
            Self::ThaiNumbers => "thaiNumbers",
            Self::ThaiCounting => "thaiCounting",
            Self::Custom => "custom",
            Self::Other(value) => value,
        }
    }
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct NumberingLevel {
    pub level: u8,
    pub start: i64,
    pub format: NumberFormat,
    pub custom_format: Option<String>,
    pub level_text: Option<String>,
    pub suffix: NumberingSuffix,
    pub restart: LevelRestart,
    pub legal: bool,
    pub paragraph_style: Option<String>,
    pub picture_bullet_id: Option<u32>,
}

impl NumberingLevel {
    fn new(level: u8) -> Self {
        Self {
            level,
            start: 0,
            format: NumberFormat::Decimal,
            custom_format: None,
            level_text: None,
            suffix: NumberingSuffix::Tab,
            restart: LevelRestart::Default,
            legal: false,
            paragraph_style: None,
            picture_bullet_id: None,
        }
    }
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct LevelOverride {
    pub level: u8,
    pub start_override: Option<i64>,
    pub definition: Option<NumberingLevel>,
}

struct PendingAbstract {
    depth: usize,
    value: AbstractNum,
}

struct PendingNum {
    depth: usize,
    id: u32,
    abstract_num_id: Option<u32>,
    overrides: Vec<LevelOverride>,
}

struct PendingOverride {
    depth: usize,
    level: u8,
    start_override: Option<i64>,
    definition: Option<NumberingLevel>,
}

struct PendingLevel {
    depth: usize,
    value: NumberingLevel,
    in_override: bool,
}

struct PendingPictureBullet {
    depth: usize,
    id: u32,
    image_relationship_id: Option<String>,
}

impl Numbering {
    pub fn new() -> Self {
        Self {
            abstract_nums: Vec::new(),
            nums: Vec::new(),
            picture_bullets: Vec::new(),
        }
    }

    pub fn abstract_nums(&self) -> &[AbstractNum] {
        &self.abstract_nums
    }
    pub fn nums(&self) -> &[Num] {
        &self.nums
    }
    pub fn abstract_num_count(&self) -> usize {
        self.abstract_nums.len()
    }
    pub fn num_count(&self) -> usize {
        self.nums.len()
    }
    pub fn get_abstract_num(&self, id: u32) -> Option<&AbstractNum> {
        self.abstract_nums.iter().find(|value| value.id == id)
    }
    pub fn get_num(&self, id: u32) -> Option<&Num> {
        self.nums.iter().find(|value| value.id == id)
    }
    pub fn picture_bullets(&self) -> &[PictureBullet] {
        &self.picture_bullets
    }
    pub fn get_picture_bullet(&self, id: u32) -> Option<&PictureBullet> {
        self.picture_bullets.iter().find(|value| value.id == id)
    }

    pub(crate) fn extract_from_part(part: &dyn Part) -> Result<Self> {
        let xml = litchi_ooxml_common::mce::process_part(part)?;
        let mut reader = NsReader::from_reader(xml.as_ref());
        let mut result = Self::new();
        let mut abstract_num: Option<PendingAbstract> = None;
        let mut num: Option<PendingNum> = None;
        let mut level_override: Option<PendingOverride> = None;
        let mut level: Option<PendingLevel> = None;
        let mut picture_bullet: Option<PendingPictureBullet> = None;
        let mut depth = 0usize;

        loop {
            let decoder = reader.decoder();
            let event = reader
                .read_event()
                .map_err(|error| OoxmlError::Xml(error.to_string()))?
                .into_owned();
            let resolver = reader.resolver().clone();
            let (namespace, event) = resolver.resolve_event(event);
            match event {
                Event::Start(element) => {
                    depth = depth.checked_add(1).ok_or_else(too_deep)?;
                    begin_element(
                        &namespace,
                        &element,
                        decoder,
                        &resolver,
                        depth,
                        &mut abstract_num,
                        &mut num,
                        &mut level_override,
                        &mut level,
                        &mut picture_bullet,
                    )?;
                },
                Event::Empty(element) => {
                    let child_depth = depth.checked_add(1).ok_or_else(too_deep)?;
                    empty_element(
                        &namespace,
                        &element,
                        decoder,
                        &resolver,
                        child_depth,
                        &mut result,
                        &mut abstract_num,
                        &mut num,
                        &mut level_override,
                        &mut level,
                        &mut picture_bullet,
                    )?;
                },
                Event::End(element) => {
                    if is_wordprocessing_namespace(&namespace) {
                        match element.local_name().as_ref() {
                            b"numPicBullet"
                                if picture_bullet
                                    .as_ref()
                                    .is_some_and(|value| value.depth == depth) =>
                            {
                                let pending = picture_bullet.take().expect("bullet checked");
                                push_picture_bullet(
                                    &mut result.picture_bullets,
                                    PictureBullet {
                                        id: pending.id,
                                        image_relationship_id: pending.image_relationship_id,
                                    },
                                )?;
                            },
                            b"lvl" if level.as_ref().is_some_and(|value| value.depth == depth) => {
                                finish_level(
                                    &mut abstract_num,
                                    &mut level_override,
                                    level.take().expect("level checked"),
                                )?;
                            },
                            b"lvlOverride"
                                if level_override
                                    .as_ref()
                                    .is_some_and(|value| value.depth == depth) =>
                            {
                                if level.is_some() {
                                    return Err(invalid("unterminated level in lvlOverride"));
                                }
                                finish_override(
                                    &mut num,
                                    level_override.take().expect("override checked"),
                                )?;
                            },
                            b"abstractNum"
                                if abstract_num
                                    .as_ref()
                                    .is_some_and(|value| value.depth == depth) =>
                            {
                                if level.is_some() {
                                    return Err(invalid("unterminated abstract numbering level"));
                                }
                                push_abstract(
                                    &mut result.abstract_nums,
                                    abstract_num.take().expect("abstract checked").value,
                                )?;
                            },
                            b"num" if num.as_ref().is_some_and(|value| value.depth == depth) => {
                                if level.is_some() || level_override.is_some() {
                                    return Err(invalid("unterminated numbering override"));
                                }
                                let pending = num.take().expect("num checked");
                                let abstract_num_id = pending.abstract_num_id.ok_or_else(|| {
                                    invalid(&format!(
                                        "numbering instance {} is missing abstractNumId",
                                        pending.id
                                    ))
                                })?;
                                push_num(
                                    &mut result.nums,
                                    Num {
                                        id: pending.id,
                                        abstract_num_id,
                                        overrides: pending.overrides,
                                    },
                                )?;
                            },
                            _ => {},
                        }
                    }
                    depth = depth
                        .checked_sub(1)
                        .ok_or_else(|| invalid("invalid numbering XML nesting"))?;
                },
                Event::Eof
                    if depth != 0
                        || abstract_num.is_some()
                        || num.is_some()
                        || level_override.is_some()
                        || level.is_some()
                        || picture_bullet.is_some() =>
                {
                    return Err(invalid("unterminated numbering XML"));
                },
                Event::Eof => break,
                _ => {},
            }
        }

        for value in &result.nums {
            if result.get_abstract_num(value.abstract_num_id).is_none() {
                return Err(invalid(&format!(
                    "numbering instance {} references missing abstractNum {}",
                    value.id, value.abstract_num_id
                )));
            }
        }
        Ok(result)
    }
}

#[allow(clippy::too_many_arguments)]
fn begin_element(
    namespace: &ResolveResult<'_>,
    element: &BytesStart<'_>,
    decoder: Decoder,
    resolver: &NamespaceResolver,
    depth: usize,
    abstract_num: &mut Option<PendingAbstract>,
    num: &mut Option<PendingNum>,
    level_override: &mut Option<PendingOverride>,
    level: &mut Option<PendingLevel>,
    picture_bullet: &mut Option<PendingPictureBullet>,
) -> Result<()> {
    if let Some(pending) = picture_bullet.as_mut() {
        return capture_picture_bullet_image(namespace, element, decoder, resolver, pending);
    }
    if !is_wordprocessing_namespace(namespace) {
        return Ok(());
    }
    let name = element.local_name();
    match name.as_ref() {
        b"numPicBullet" if abstract_num.is_none() && num.is_none() => {
            *picture_bullet = Some(PendingPictureBullet {
                depth,
                id: required_u32(element, b"numPicBulletId", decoder, resolver)?,
                image_relationship_id: None,
            });
        },
        b"abstractNum" => {
            if abstract_num.is_some() || num.is_some() {
                return Err(invalid("nested numbering definitions are invalid"));
            }
            *abstract_num = Some(PendingAbstract {
                depth,
                value: AbstractNum {
                    id: required_u32(element, b"abstractNumId", decoder, resolver)?,
                    num_type: None,
                    num_style_link: None,
                    style_link: None,
                    levels: Vec::new(),
                },
            });
        },
        b"num" => {
            if abstract_num.is_some() || num.is_some() {
                return Err(invalid("nested numbering definitions are invalid"));
            }
            *num = Some(PendingNum {
                depth,
                id: required_u32(element, b"numId", decoder, resolver)?,
                abstract_num_id: None,
                overrides: Vec::new(),
            });
        },
        b"lvl"
            if level.is_none()
                && level_override
                    .as_ref()
                    .is_some_and(|value| depth == value.depth + 1) =>
        {
            *level = Some(PendingLevel {
                depth,
                value: NumberingLevel::new(required_level(element, decoder, resolver)?),
                in_override: true,
            });
        },
        b"lvl"
            if level.is_none()
                && abstract_num
                    .as_ref()
                    .is_some_and(|value| depth == value.depth + 1) =>
        {
            *level = Some(PendingLevel {
                depth,
                value: NumberingLevel::new(required_level(element, decoder, resolver)?),
                in_override: false,
            });
        },
        b"lvlOverride"
            if level_override.is_none()
                && num.as_ref().is_some_and(|value| depth == value.depth + 1) =>
        {
            *level_override = Some(PendingOverride {
                depth,
                level: required_level(element, decoder, resolver)?,
                start_override: None,
                definition: None,
            });
        },
        _ => apply_child(
            element,
            decoder,
            resolver,
            depth,
            abstract_num,
            num,
            level_override,
            level,
        )?,
    }
    Ok(())
}

#[allow(clippy::too_many_arguments)]
fn empty_element(
    namespace: &ResolveResult<'_>,
    element: &BytesStart<'_>,
    decoder: Decoder,
    resolver: &NamespaceResolver,
    depth: usize,
    result: &mut Numbering,
    abstract_num: &mut Option<PendingAbstract>,
    num: &mut Option<PendingNum>,
    level_override: &mut Option<PendingOverride>,
    level: &mut Option<PendingLevel>,
    picture_bullet: &mut Option<PendingPictureBullet>,
) -> Result<()> {
    if let Some(pending) = picture_bullet.as_mut() {
        return capture_picture_bullet_image(namespace, element, decoder, resolver, pending);
    }
    if !is_wordprocessing_namespace(namespace) {
        return Ok(());
    }
    match element.local_name().as_ref() {
        b"numPicBullet" if abstract_num.is_none() && num.is_none() => {
            push_picture_bullet(
                &mut result.picture_bullets,
                PictureBullet {
                    id: required_u32(element, b"numPicBulletId", decoder, resolver)?,
                    image_relationship_id: None,
                },
            )?;
        },
        b"abstractNum" => {
            if abstract_num.is_some() || num.is_some() {
                return Err(invalid("nested numbering definitions are invalid"));
            }
            push_abstract(
                &mut result.abstract_nums,
                AbstractNum {
                    id: required_u32(element, b"abstractNumId", decoder, resolver)?,
                    num_type: None,
                    num_style_link: None,
                    style_link: None,
                    levels: Vec::new(),
                },
            )?;
        },
        b"num" => return Err(invalid("numbering instance is missing abstractNumId")),
        b"lvl"
            if level_override
                .as_ref()
                .is_some_and(|value| depth == value.depth + 1) =>
        {
            finish_level(
                abstract_num,
                level_override,
                PendingLevel {
                    depth,
                    value: NumberingLevel::new(required_level(element, decoder, resolver)?),
                    in_override: true,
                },
            )?;
        },
        b"lvl"
            if abstract_num
                .as_ref()
                .is_some_and(|value| depth == value.depth + 1) =>
        {
            finish_level(
                abstract_num,
                level_override,
                PendingLevel {
                    depth,
                    value: NumberingLevel::new(required_level(element, decoder, resolver)?),
                    in_override: false,
                },
            )?;
        },
        b"lvlOverride" if num.as_ref().is_some_and(|value| depth == value.depth + 1) => {
            finish_override(
                num,
                PendingOverride {
                    depth,
                    level: required_level(element, decoder, resolver)?,
                    start_override: None,
                    definition: None,
                },
            )?;
        },
        _ => apply_child(
            element,
            decoder,
            resolver,
            depth,
            abstract_num,
            num,
            level_override,
            level,
        )?,
    }
    Ok(())
}

#[allow(clippy::too_many_arguments)]
fn apply_child(
    element: &BytesStart<'_>,
    decoder: Decoder,
    resolver: &NamespaceResolver,
    depth: usize,
    abstract_num: &mut Option<PendingAbstract>,
    num: &mut Option<PendingNum>,
    level_override: &mut Option<PendingOverride>,
    level: &mut Option<PendingLevel>,
) -> Result<()> {
    if let Some(value) = level.as_mut().filter(|value| depth == value.depth + 1) {
        match element.local_name().as_ref() {
            b"start" => value.value.start = required_i64(element, b"val", decoder, resolver)?,
            b"numFmt" => {
                let raw = required_string(element, b"val", decoder, resolver)?;
                value.value.format = NumberFormat::parse(&raw);
                value.value.custom_format =
                    word_attribute_value(element, b"format", decoder, resolver)?;
            },
            b"lvlText" => {
                value.value.level_text = Some(
                    word_attribute_value(element, b"val", decoder, resolver)?.unwrap_or_default(),
                )
            },
            b"suff" => {
                value.value.suffix = match required_string(element, b"val", decoder, resolver)?
                    .as_str()
                {
                    "tab" => NumberingSuffix::Tab,
                    "space" => NumberingSuffix::Space,
                    "nothing" => NumberingSuffix::Nothing,
                    other => return Err(invalid(&format!("invalid numbering suffix '{other}'"))),
                }
            },
            b"lvlRestart" => {
                let raw = required_u32(element, b"val", decoder, resolver)?;
                value.value.restart = match raw {
                    0 => LevelRestart::Never,
                    1..=9 => LevelRestart::After((raw - 1) as u8),
                    _ => return Err(invalid(&format!("invalid lvlRestart '{raw}'"))),
                };
            },
            b"isLgl" => value.value.legal = on_off(element, decoder, resolver)?,
            b"pStyle" => {
                value.value.paragraph_style =
                    Some(required_string(element, b"val", decoder, resolver)?)
            },
            b"lvlPicBulletId" => {
                value.value.picture_bullet_id =
                    Some(required_u32(element, b"val", decoder, resolver)?)
            },
            _ => {},
        }
        return Ok(());
    }
    if let Some(value) = level_override
        .as_mut()
        .filter(|value| depth == value.depth + 1)
    {
        if element.local_name().as_ref() == b"startOverride" {
            if value.start_override.is_some() {
                return Err(invalid("duplicate startOverride"));
            }
            value.start_override = Some(required_i64(element, b"val", decoder, resolver)?);
        }
        return Ok(());
    }
    if let Some(value) = abstract_num
        .as_mut()
        .filter(|value| depth == value.depth + 1)
    {
        match element.local_name().as_ref() {
            b"multiLevelType" => {
                let raw = required_string(element, b"val", decoder, resolver)?;
                if !matches!(
                    raw.as_str(),
                    "singleLevel" | "multilevel" | "hybridMultilevel"
                ) {
                    return Err(invalid(&format!("invalid multiLevelType '{raw}'")));
                }
                set_once(&mut value.value.num_type, raw, "multiLevelType")?;
            },
            b"numStyleLink" => {
                let raw = required_string(element, b"val", decoder, resolver)?;
                set_once(&mut value.value.num_style_link, raw, "numStyleLink")?;
            },
            b"styleLink" => {
                let raw = required_string(element, b"val", decoder, resolver)?;
                set_once(&mut value.value.style_link, raw, "styleLink")?;
            },
            _ => {},
        }
    }
    if let Some(value) = num.as_mut().filter(|value| depth == value.depth + 1)
        && element.local_name().as_ref() == b"abstractNumId"
    {
        if value.abstract_num_id.is_some() {
            return Err(invalid("duplicate abstractNumId"));
        }
        value.abstract_num_id = Some(required_u32(element, b"val", decoder, resolver)?);
    }
    Ok(())
}

fn finish_level(
    abstract_num: &mut Option<PendingAbstract>,
    level_override: &mut Option<PendingOverride>,
    pending: PendingLevel,
) -> Result<()> {
    if pending.in_override {
        let target = level_override
            .as_mut()
            .ok_or_else(|| invalid("level outside lvlOverride"))?;
        if target.definition.is_some() {
            return Err(invalid("duplicate override level definition"));
        }
        if pending.value.level != target.level {
            return Err(invalid("override level indices do not match"));
        }
        target.definition = Some(pending.value);
    } else {
        let target = abstract_num
            .as_mut()
            .ok_or_else(|| invalid("level outside abstractNum"))?;
        if target
            .value
            .levels
            .iter()
            .any(|value| value.level == pending.value.level)
        {
            return Err(invalid("duplicate abstract numbering level"));
        }
        target.value.levels.push(pending.value);
    }
    Ok(())
}

fn finish_override(num: &mut Option<PendingNum>, pending: PendingOverride) -> Result<()> {
    let target = num
        .as_mut()
        .ok_or_else(|| invalid("lvlOverride outside num"))?;
    if target
        .overrides
        .iter()
        .any(|value| value.level == pending.level)
    {
        return Err(invalid("duplicate lvlOverride"));
    }
    target.overrides.push(LevelOverride {
        level: pending.level,
        start_override: pending.start_override,
        definition: pending.definition,
    });
    Ok(())
}

fn set_once(slot: &mut Option<String>, value: String, name: &str) -> Result<()> {
    if slot.is_some() {
        return Err(invalid(&format!("duplicate {name}")));
    }
    *slot = Some(value);
    Ok(())
}

fn push_abstract(values: &mut Vec<AbstractNum>, value: AbstractNum) -> Result<()> {
    if values.iter().any(|item| item.id == value.id) {
        return Err(invalid(&format!(
            "duplicate abstract numbering ID {}",
            value.id
        )));
    }
    values.push(value);
    Ok(())
}

fn push_num(values: &mut Vec<Num>, value: Num) -> Result<()> {
    if values.iter().any(|item| item.id == value.id) {
        return Err(invalid(&format!(
            "duplicate numbering instance ID {}",
            value.id
        )));
    }
    values.push(value);
    Ok(())
}

fn push_picture_bullet(values: &mut Vec<PictureBullet>, value: PictureBullet) -> Result<()> {
    if values.iter().any(|item| item.id == value.id) {
        return Err(invalid(&format!(
            "duplicate picture bullet ID {}",
            value.id
        )));
    }
    values.push(value);
    Ok(())
}

/// Capture the first image relationship inside a `w:numPicBullet` definition.
///
/// Word writes the bullet picture either as VML (`v:imagedata r:id`) or as
/// DrawingML (`a:blip r:embed`/`a:link`); everything else inside `w:pict` is
/// inert shape geometry and is ignored.
fn capture_picture_bullet_image(
    namespace: &ResolveResult<'_>,
    element: &BytesStart<'_>,
    decoder: Decoder,
    resolver: &NamespaceResolver,
    pending: &mut PendingPictureBullet,
) -> Result<()> {
    if pending.image_relationship_id.is_some() {
        return Ok(());
    }
    let names: &[&[u8]] = match namespace {
        ResolveResult::Bound(Namespace(uri)) if *uri == VML_NAMESPACE => {
            match element.local_name().as_ref() {
                b"imagedata" => &[b"id"],
                _ => return Ok(()),
            }
        },
        ResolveResult::Bound(Namespace(uri))
            if *uri == DRAWINGML_NAMESPACE || *uri == STRICT_DRAWINGML_NAMESPACE =>
        {
            match element.local_name().as_ref() {
                b"blip" => &[b"embed", b"link"],
                _ => return Ok(()),
            }
        },
        _ => return Ok(()),
    };
    pending.image_relationship_id = relationship_attribute(element, names, decoder, resolver)?;
    Ok(())
}

fn relationship_attribute(
    element: &BytesStart<'_>,
    names: &[&[u8]],
    decoder: Decoder,
    resolver: &NamespaceResolver,
) -> Result<Option<String>> {
    let mut value = None;
    for attribute in element.attributes() {
        let attribute = attribute.map_err(|error| OoxmlError::Xml(error.to_string()))?;
        if !names.contains(&attribute.key.local_name().as_ref()) {
            continue;
        }
        let (namespace, _) = resolver.resolve_attribute(attribute.key);
        let is_relationship_attribute = matches!(
            namespace,
            ResolveResult::Bound(Namespace(uri))
                if uri == TRANSITIONAL_RELATIONSHIPS_NAMESPACE
                    || uri == STRICT_RELATIONSHIPS_NAMESPACE
        );
        if !is_relationship_attribute {
            continue;
        }
        if value.is_some() {
            return Err(invalid("duplicate picture bullet image relationship"));
        }
        value = Some(
            attribute
                .decoded_and_normalized_value(XmlVersion::Explicit1_0, decoder)
                .map_err(|error| OoxmlError::Xml(error.to_string()))?
                .into_owned(),
        );
    }
    Ok(value)
}

fn required_string(
    element: &BytesStart<'_>,
    name: &[u8],
    decoder: Decoder,
    resolver: &NamespaceResolver,
) -> Result<String> {
    word_attribute_value(element, name, decoder, resolver)?.ok_or_else(|| {
        invalid(&format!(
            "Word numbering element is missing required '{}' attribute",
            String::from_utf8_lossy(name)
        ))
    })
}

fn required_u32(
    element: &BytesStart<'_>,
    name: &[u8],
    decoder: Decoder,
    resolver: &NamespaceResolver,
) -> Result<u32> {
    let value = required_string(element, name, decoder, resolver)?;
    value
        .parse()
        .map_err(|_| invalid(&format!("invalid Word numbering integer '{value}'")))
}

fn required_i64(
    element: &BytesStart<'_>,
    name: &[u8],
    decoder: Decoder,
    resolver: &NamespaceResolver,
) -> Result<i64> {
    let value = required_string(element, name, decoder, resolver)?;
    value
        .parse()
        .map_err(|_| invalid(&format!("invalid Word numbering integer '{value}'")))
}

fn required_level(
    element: &BytesStart<'_>,
    decoder: Decoder,
    resolver: &NamespaceResolver,
) -> Result<u8> {
    let value = required_u32(element, b"ilvl", decoder, resolver)?;
    u8::try_from(value)
        .ok()
        .filter(|value| *value <= 8)
        .ok_or_else(|| invalid(&format!("invalid numbering level '{value}'")))
}

fn on_off(
    element: &BytesStart<'_>,
    decoder: Decoder,
    resolver: &NamespaceResolver,
) -> Result<bool> {
    Ok(
        match word_attribute_value(element, b"val", decoder, resolver)?.as_deref() {
            None | Some("1" | "true" | "on") => true,
            Some("0" | "false" | "off") => false,
            Some(value) => return Err(invalid(&format!("invalid on/off value '{value}'"))),
        },
    )
}

fn too_deep() -> OoxmlError {
    invalid("numbering XML nesting is too deep")
}
fn invalid(message: &str) -> OoxmlError {
    OoxmlError::InvalidFormat(message.to_owned())
}

impl Default for Numbering {
    fn default() -> Self {
        Self::new()
    }
}

impl AbstractNum {
    pub fn id(&self) -> u32 {
        self.id
    }
    pub fn num_type(&self) -> Option<&str> {
        self.num_type.as_deref()
    }
    pub fn num_style_link(&self) -> Option<&str> {
        self.num_style_link.as_deref()
    }
    pub fn style_link(&self) -> Option<&str> {
        self.style_link.as_deref()
    }
    pub fn levels(&self) -> &[NumberingLevel] {
        &self.levels
    }
    pub fn level(&self, level: u8) -> Option<&NumberingLevel> {
        self.levels.iter().find(|value| value.level == level)
    }
}

impl Num {
    pub fn id(&self) -> u32 {
        self.id
    }
    pub fn abstract_num_id(&self) -> u32 {
        self.abstract_num_id
    }
    pub fn overrides(&self) -> &[LevelOverride] {
        &self.overrides
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use litchi_opc::PackURI;
    use litchi_opc::part::BlobPart;

    fn parse(xml: &[u8]) -> Result<Numbering> {
        Numbering::extract_from_part(&BlobPart::new(
            PackURI::new("/word/numbering.xml").unwrap(),
            "application/xml".to_owned(),
            xml.to_vec(),
        ))
    }

    #[test]
    fn parses_complete_level_and_override() {
        let value = parse(br#"<w:numbering xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:abstractNum w:abstractNumId="1"><w:multiLevelType w:val="multilevel"/><w:styleLink w:val="List"/><w:lvl w:ilvl="0"><w:start w:val="3"/><w:numFmt w:val="decimal"/><w:lvlText w:val="%1."/><w:suff w:val="space"/><w:lvlRestart w:val="0"/><w:isLgl/><w:pStyle w:val="ListParagraph"/><w:lvlPicBulletId w:val="7"/></w:lvl></w:abstractNum><w:num w:numId="9"><w:abstractNumId w:val="1"/><w:lvlOverride w:ilvl="0"><w:startOverride w:val="5"/></w:lvlOverride></w:num></w:numbering>"#).unwrap();
        let level = &value.abstract_nums()[0].levels()[0];
        assert_eq!(level.start, 3);
        assert_eq!(level.level_text.as_deref(), Some("%1."));
        assert_eq!(level.restart, LevelRestart::Never);
        assert!(level.legal);
        assert_eq!(value.nums()[0].overrides()[0].start_override, Some(5));
    }

    #[test]
    fn parses_vml_picture_bullet_definition() {
        let value = parse(br##"<w:numbering xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:v="urn:schemas-microsoft-com:vml" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><w:numPicBullet w:numPicBulletId="3"><w:pict><v:shapetype id="_x0000_t75" coordsize="21600,21600" o:spt="75" xmlns:o="urn:schemas-microsoft-com:office:office"/><v:shape id="_x0000_i1025" type="#_x0000_t75" style="width:12pt;height:12pt" o:bullet="t"><v:imagedata r:id="rId4" o:title="bullet"/></v:shape></w:pict></w:numPicBullet><w:abstractNum w:abstractNumId="1"><w:lvl w:ilvl="0"><w:numFmt w:val="bullet"/><w:lvlPicBulletId w:val="3"/></w:lvl></w:abstractNum></w:numbering>"##).unwrap();
        assert_eq!(value.picture_bullets().len(), 1);
        let bullet = value.get_picture_bullet(3).expect("picture bullet 3");
        assert_eq!(bullet.id(), 3);
        assert_eq!(bullet.image_relationship_id(), Some("rId4"));
        assert!(value.get_picture_bullet(4).is_none());
        assert_eq!(
            value.abstract_nums()[0].levels()[0].picture_bullet_id,
            Some(3)
        );
    }

    #[test]
    fn parses_drawingml_picture_bullet_definition() {
        let value = parse(br#"<w:numbering xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><w:numPicBullet w:numPicBulletId="1"><w:pict><a:blip r:embed="rId9"/></w:pict></w:numPicBullet></w:numbering>"#).unwrap();
        let bullet = value.get_picture_bullet(1).expect("picture bullet 1");
        assert_eq!(bullet.image_relationship_id(), Some("rId9"));
    }

    #[test]
    fn parses_picture_bullet_without_image() {
        let value = parse(br#"<w:numbering xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:numPicBullet w:numPicBulletId="0"><w:pict/></w:numPicBullet><w:numPicBullet w:numPicBulletId="2"/></w:numbering>"#).unwrap();
        assert_eq!(value.picture_bullets().len(), 2);
        assert_eq!(
            value.get_picture_bullet(0).unwrap().image_relationship_id(),
            None
        );
        assert_eq!(
            value.get_picture_bullet(2).unwrap().image_relationship_id(),
            None
        );
    }

    #[test]
    fn rejects_duplicate_picture_bullet_ids() {
        let duplicate = br#"<w:numbering xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:numPicBullet w:numPicBulletId="1"/><w:numPicBullet w:numPicBulletId="1"/></w:numbering>"#;
        assert!(parse(duplicate).is_err());
    }

    #[test]
    fn parses_libreoffice_picture_bullet_fixture() {
        let root = std::path::Path::new(env!("CARGO_MANIFEST_DIR")).join("../..");
        let relative =
            "test-data/libreoffice-core/sw/qa/extras/ooxmlexport/data/lvlPicBulletId.docx";
        let package = litchi_opc::phys_pkg::OwnedPhysPkgReader::open(root.join(relative))
            .unwrap_or_else(|error| panic!("failed to open {relative}: {error}"));
        let uri = PackURI::new("/word/numbering.xml").expect("valid numbering URI");
        let bytes = package
            .blob_for(&uri)
            .unwrap_or_else(|error| panic!("failed to load numbering part: {error}"));
        let part = BlobPart::new(uri, "application/xml".to_owned(), bytes);
        let numbering = Numbering::extract_from_part(&part).unwrap();
        let bullet = numbering
            .get_picture_bullet(0)
            .expect("fixture defines picture bullet 0");
        // LibreOffice stripped the image payload from this fixture; only the
        // definition shell and the level linkage remain.
        assert_eq!(bullet.image_relationship_id(), None);
        let level = numbering
            .abstract_nums()
            .iter()
            .flat_map(|abstract_num| abstract_num.levels())
            .find(|level| level.picture_bullet_id.is_some())
            .expect("fixture level references a picture bullet");
        assert_eq!(level.picture_bullet_id, Some(0));
        assert!(
            numbering
                .get_picture_bullet(level.picture_bullet_id.unwrap())
                .is_some()
        );
    }

    #[test]
    fn rejects_duplicate_levels_and_bad_level_indices() {
        let duplicate = br#"<w:numbering xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:abstractNum w:abstractNumId="1"><w:lvl w:ilvl="0"/><w:lvl w:ilvl="0"/></w:abstractNum></w:numbering>"#;
        assert!(parse(duplicate).is_err());
        let bad = br#"<w:numbering xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:abstractNum w:abstractNumId="1"><w:lvl w:ilvl="9"/></w:abstractNum></w:numbering>"#;
        assert!(parse(bad).is_err());
    }

    #[test]
    fn recognizes_and_round_trips_every_standard_number_format() {
        for raw in [
            "decimal",
            "upperRoman",
            "lowerRoman",
            "upperLetter",
            "lowerLetter",
            "ordinal",
            "cardinalText",
            "ordinalText",
            "hex",
            "chicago",
            "ideographDigital",
            "japaneseCounting",
            "aiueo",
            "iroha",
            "decimalFullWidth",
            "decimalHalfWidth",
            "japaneseLegal",
            "japaneseDigitalTenThousand",
            "decimalEnclosedCircle",
            "decimalFullWidth2",
            "aiueoFullWidth",
            "irohaFullWidth",
            "decimalZero",
            "bullet",
            "ganada",
            "chosung",
            "decimalEnclosedFullstop",
            "decimalEnclosedParen",
            "decimalEnclosedCircleChinese",
            "ideographEnclosedCircle",
            "ideographTraditional",
            "ideographZodiac",
            "ideographZodiacTraditional",
            "taiwaneseCounting",
            "ideographLegalTraditional",
            "taiwaneseCountingThousand",
            "taiwaneseDigital",
            "chineseCounting",
            "chineseLegalSimplified",
            "chineseCountingThousand",
            "koreanDigital",
            "koreanCounting",
            "koreanLegal",
            "koreanDigital2",
            "vietnameseCounting",
            "russianLower",
            "russianUpper",
            "none",
            "numberInDash",
            "hebrew1",
            "hebrew2",
            "arabicAlpha",
            "arabicAbjad",
            "hindiVowels",
            "hindiConsonants",
            "hindiNumbers",
            "hindiCounting",
            "thaiLetters",
            "thaiNumbers",
            "thaiCounting",
            "custom",
        ] {
            let parsed = NumberFormat::parse(raw);
            assert!(
                !matches!(parsed, NumberFormat::Other(_)),
                "untyped standard token: {raw}"
            );
            assert_eq!(parsed.as_str(), raw);
        }

        let extension = NumberFormat::parse("vendorNumbering");
        assert_eq!(extension, NumberFormat::Other("vendorNumbering".to_owned()));
        assert_eq!(extension.as_str(), "vendorNumbering");
    }

    #[test]
    fn parses_poi_and_libreoffice_numbering_fixtures() {
        let root = std::path::Path::new(env!("CARGO_MANIFEST_DIR")).join("../..");
        for relative in [
            "test-data/poi/test-data/document/Numbering.docx",
            "test-data/poi/test-data/document/NumberingWOverrides.docx",
            "test-data/poi/test-data/document/ComplexNumberedLists.docx",
            "test-data/poi/test-data/document/NumberingWithOutOfOrderId.docx",
            "test-data/libreoffice-core/sw/qa/extras/ooxmlexport/data/listWithLgl.docx",
            "test-data/libreoffice-core/sw/qa/extras/ooxmlexport/data/decimal-numbering-no-leveltext.docx",
            "test-data/libreoffice-core/sw/qa/extras/ooxmlimport/data/numbering-circle.docx",
            "test-data/libreoffice-core/sw/qa/extras/ooxmlexport/data/NumberedList.docx",
            "test-data/libreoffice-core/sw/qa/extras/ooxmlexport/data/lvlPicBulletId.docx",
        ] {
            let package = litchi_opc::phys_pkg::OwnedPhysPkgReader::open(root.join(relative))
                .unwrap_or_else(|error| panic!("failed to open {relative}: {error}"));
            let uri = PackURI::new("/word/numbering.xml").expect("valid numbering URI");
            let bytes = package.blob_for(&uri).unwrap_or_else(|error| {
                panic!("failed to load numbering part in {relative}: {error}")
            });
            let part = BlobPart::new(uri, "application/xml".to_owned(), bytes);
            let numbering = Numbering::extract_from_part(&part)
                .unwrap_or_else(|error| panic!("failed to parse numbering in {relative}: {error}"));
            assert!(
                numbering.abstract_num_count() != 0,
                "fixture has no definitions: {relative}"
            );
        }
    }
}
