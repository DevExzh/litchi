//! Immutable static worksheet auto-filter and sort-state read model.

use crate::error::{Result, invalid};
use crate::sort::{SortBy, SortMethod};

pub(crate) const CORE: &[u8] = b"http://schemas.openxmlformats.org/spreadsheetml/2006/main";
pub(crate) const STRICT: &[u8] = b"http://purl.oclc.org/ooxml/spreadsheetml/main";
pub(crate) const MAX_COLUMNS: usize = 16_384;
pub(crate) const MAX_ITEMS: usize = 10_000;
pub(crate) const MAX_SORT_CONDITIONS: usize = 64;
pub(crate) const MAX_FRAGMENT_BYTES: usize = 8 * 1024 * 1024;
/// `SpreadsheetML` `ST_Xstring` values used by the filter family are limited to
/// fewer than 65,536 Unicode characters.  This is deliberately a character
/// limit; the codec still applies a separate byte budget to the XML owner.
pub(crate) const MAX_TEXT_CHARS: usize = 65_535;
pub(crate) const MAX_UNKNOWN_ATTRIBUTES: usize = 4096;
pub(crate) const MAX_UNKNOWN_ELEMENTS: usize = 4096;
pub(crate) const MAX_UNKNOWN_BYTES: usize = MAX_FRAGMENT_BYTES;

/// An inert attribute retained because this owner does not interpret it.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct UnknownAttribute {
    pub(crate) name: Box<str>,
    pub(crate) value: Box<str>,
}

impl UnknownAttribute {
    #[must_use]
    pub fn name(&self) -> &str {
        &self.name
    }

    #[must_use]
    pub fn value(&self) -> &str {
        &self.value
    }

    pub(crate) fn new(name: String, value: String) -> Result<Self> {
        if name.is_empty() || name.len() > MAX_TEXT_CHARS || value.chars().count() > MAX_TEXT_CHARS
        {
            return Err(invalid("unknown auto-filter attribute is out of bounds"));
        }
        Ok(Self {
            name: name.into_boxed_str(),
            value: value.into_boxed_str(),
        })
    }
}

/// A self-contained unknown element retained without interpretation.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct UnknownElement {
    pub(crate) xml: Box<[u8]>,
}

impl UnknownElement {
    pub fn new(xml: impl Into<Vec<u8>>) -> Result<Self> {
        let xml = xml.into();
        if xml.is_empty() || xml.len() > MAX_UNKNOWN_BYTES {
            return Err(invalid("unknown auto-filter element is out of bounds"));
        }
        Ok(Self {
            xml: xml.into_boxed_slice(),
        })
    }

    #[must_use]
    pub fn as_xml(&self) -> &[u8] {
        &self.xml
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub(crate) enum ChildOrder {
    Column(usize),
    SortState,
    Payload,
    Item(usize),
    Custom(usize),
    Condition(usize),
    Unknown(usize),
}

#[derive(Debug, Clone, PartialEq, Eq, Default)]
pub(crate) struct OpaqueFields {
    pub(crate) attributes: Vec<UnknownAttribute>,
    pub(crate) elements: Vec<UnknownElement>,
    pub(crate) order: Vec<ChildOrder>,
    pub(crate) retained_bytes: usize,
}

impl OpaqueFields {
    pub(crate) fn push_attribute(&mut self, value: UnknownAttribute) -> Result<()> {
        if self.attributes.len() >= MAX_UNKNOWN_ATTRIBUTES {
            return Err(invalid("too many unknown auto-filter attributes"));
        }
        self.retained_bytes = self
            .retained_bytes
            .checked_add(value.name().len())
            .and_then(|size| size.checked_add(value.value().len()))
            .ok_or_else(|| invalid("unknown auto-filter size overflow"))?;
        if self.retained_bytes > MAX_UNKNOWN_BYTES {
            return Err(invalid("unknown auto-filter data exceeds safety limit"));
        }
        self.attributes.push(value);
        Ok(())
    }

    pub(crate) fn push_element(&mut self, value: UnknownElement) -> Result<usize> {
        if self.elements.len() >= MAX_UNKNOWN_ELEMENTS {
            return Err(invalid("too many unknown auto-filter elements"));
        }
        self.retained_bytes = self
            .retained_bytes
            .checked_add(value.as_xml().len())
            .ok_or_else(|| invalid("unknown auto-filter size overflow"))?;
        if self.retained_bytes > MAX_UNKNOWN_BYTES {
            return Err(invalid("unknown auto-filter data exceeds safety limit"));
        }
        let index = self.elements.len();
        self.elements.push(value);
        Ok(index)
    }

    pub(crate) fn push_order(&mut self, value: ChildOrder) -> Result<()> {
        if self.order.len() >= MAX_COLUMNS + MAX_ITEMS + MAX_SORT_CONDITIONS {
            return Err(invalid("auto-filter child order exceeds safety limit"));
        }
        self.order.push(value);
        Ok(())
    }
}

pub(crate) fn opaque_mut(slot: &mut Option<Box<OpaqueFields>>) -> &mut OpaqueFields {
    slot.get_or_insert_with(|| Box::new(OpaqueFields::default()))
}

fn unknown_attributes(slot: Option<&OpaqueFields>) -> &[UnknownAttribute] {
    slot.map_or(&[], |value| value.attributes.as_slice())
}

fn unknown_elements(slot: Option<&OpaqueFields>) -> &[UnknownElement] {
    slot.map_or(&[], |value| value.elements.as_slice())
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Range(String);
impl Range {
    pub fn new(value: impl Into<String>) -> Result<Self> {
        let value = value.into();
        super::codec::parse_range(&value)?;
        Ok(Self(value))
    }
    #[must_use]
    pub fn as_str(&self) -> &str {
        &self.0
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum Calendar {
    None,
    Gregorian,
    GregorianUs,
    GregorianMeFrench,
    GregorianArabic,
    Hijri,
    Hebrew,
    Taiwan,
    Japan,
    Thai,
    Korea,
    Saka,
}
impl Calendar {
    pub(crate) fn parse(v: &str) -> Result<Self> {
        match v {
            "none" => Ok(Self::None),
            "gregorian" => Ok(Self::Gregorian),
            "gregorianUs" => Ok(Self::GregorianUs),
            "gregorianMeFrench" => Ok(Self::GregorianMeFrench),
            "gregorianArabic" => Ok(Self::GregorianArabic),
            "hijri" => Ok(Self::Hijri),
            "hebrew" => Ok(Self::Hebrew),
            "taiwan" => Ok(Self::Taiwan),
            "japan" => Ok(Self::Japan),
            "thai" => Ok(Self::Thai),
            "korea" => Ok(Self::Korea),
            "saka" => Ok(Self::Saka),
            _ => Err(invalid(format!("invalid calendarType '{v}'"))),
        }
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum Grouping {
    Year,
    Month,
    Day,
    Hour,
    Minute,
    Second,
}
impl Grouping {
    pub(crate) fn parse(v: &str) -> Result<Self> {
        match v {
            "year" => Ok(Self::Year),
            "month" => Ok(Self::Month),
            "day" => Ok(Self::Day),
            "hour" => Ok(Self::Hour),
            "minute" => Ok(Self::Minute),
            "second" => Ok(Self::Second),
            _ => Err(invalid(format!("invalid dateTimeGrouping '{v}'"))),
        }
    }
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct DateGroup {
    pub(crate) year: u16,
    pub(crate) month: Option<u8>,
    pub(crate) day: Option<u8>,
    pub(crate) hour: Option<u8>,
    pub(crate) minute: Option<u8>,
    pub(crate) second: Option<u8>,
    pub(crate) grouping: Grouping,
    pub(crate) opaque: Option<Box<OpaqueFields>>,
}
impl DateGroup {
    #[allow(
        clippy::too_many_arguments,
        reason = "arguments map one-to-one to the immutable auto-filter schema fields"
    )]
    pub fn new(
        year: u16,
        month: Option<u8>,
        day: Option<u8>,
        hour: Option<u8>,
        minute: Option<u8>,
        second: Option<u8>,
        grouping: Grouping,
    ) -> Result<Self> {
        if year > 9999
            || month.is_some_and(|value| !(1..=12).contains(&value))
            || day.is_some_and(|value| !(1..=31).contains(&value))
            || hour.is_some_and(|value| value > 23)
            || minute.is_some_and(|value| value > 59)
            || second.is_some_and(|value| value > 59)
        {
            return Err(invalid("date-group component is out of range"));
        }
        let required = match grouping {
            Grouping::Year => 0,
            Grouping::Month => 1,
            Grouping::Day => 2,
            Grouping::Hour => 3,
            Grouping::Minute => 4,
            Grouping::Second => 5,
        };
        if ![month, day, hour, minute, second]
            .iter()
            .take(required)
            .all(Option::is_some)
        {
            return Err(invalid("date-group components do not match grouping"));
        }
        Ok(Self {
            year,
            month,
            day,
            hour,
            minute,
            second,
            grouping,
            opaque: None,
        })
    }

    #[must_use]
    pub fn year(&self) -> u16 {
        self.year
    }
    #[must_use]
    pub fn month(&self) -> Option<u8> {
        self.month
    }
    #[must_use]
    pub fn day(&self) -> Option<u8> {
        self.day
    }
    #[must_use]
    pub fn hour(&self) -> Option<u8> {
        self.hour
    }
    #[must_use]
    pub fn minute(&self) -> Option<u8> {
        self.minute
    }
    #[must_use]
    pub fn second(&self) -> Option<u8> {
        self.second
    }
    #[must_use]
    pub fn grouping(&self) -> Grouping {
        self.grouping
    }

    #[must_use]
    pub fn unknown_attributes(&self) -> &[UnknownAttribute] {
        unknown_attributes(self.opaque.as_deref())
    }

    #[must_use]
    pub fn unknown_elements(&self) -> &[UnknownElement] {
        unknown_elements(self.opaque.as_deref())
    }
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub enum Item {
    Value(String),
    DateGroup(DateGroup),
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Values {
    pub(crate) blank: bool,
    pub(crate) calendar_type: Calendar,
    pub(crate) items: Vec<Item>,
    pub(crate) opaque: Option<Box<OpaqueFields>>,
}
impl Values {
    pub fn new(blank: bool, calendar_type: Calendar, items: Vec<Item>) -> Result<Self> {
        if items.len() > MAX_ITEMS {
            return Err(invalid("too many filter values"));
        }
        for item in &items {
            if let Item::Value(value) = item {
                bounded(value)?;
            }
        }
        Ok(Self {
            blank,
            calendar_type,
            items,
            opaque: None,
        })
    }

    #[must_use]
    pub fn blank(&self) -> bool {
        self.blank
    }
    #[must_use]
    pub fn calendar_type(&self) -> Calendar {
        self.calendar_type
    }
    #[must_use]
    pub fn items(&self) -> &[Item] {
        &self.items
    }

    #[must_use]
    pub fn unknown_attributes(&self) -> &[UnknownAttribute] {
        unknown_attributes(self.opaque.as_deref())
    }

    #[must_use]
    pub fn unknown_elements(&self) -> &[UnknownElement] {
        unknown_elements(self.opaque.as_deref())
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum Operator {
    LessThan,
    LessThanOrEqual,
    NotEqual,
    Equal,
    GreaterThanOrEqual,
    GreaterThan,
}
impl Operator {
    pub(crate) fn parse(v: &str) -> Result<Self> {
        match v {
            "lessThan" => Ok(Self::LessThan),
            "lessThanOrEqual" => Ok(Self::LessThanOrEqual),
            "notEqual" => Ok(Self::NotEqual),
            "equal" => Ok(Self::Equal),
            "greaterThanOrEqual" => Ok(Self::GreaterThanOrEqual),
            "greaterThan" => Ok(Self::GreaterThan),
            _ => Err(invalid(format!("invalid custom-filter operator '{v}'"))),
        }
    }
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Custom {
    pub(crate) operator: Operator,
    pub(crate) value: String,
    pub(crate) opaque: Option<Box<OpaqueFields>>,
}
impl Custom {
    pub fn new(operator: Operator, value: impl Into<String>) -> Result<Self> {
        let value = value.into();
        bounded(&value)?;
        Ok(Self {
            operator,
            value,
            opaque: None,
        })
    }

    #[must_use]
    pub fn operator(&self) -> Operator {
        self.operator
    }
    #[must_use]
    pub fn value(&self) -> &str {
        &self.value
    }

    #[must_use]
    pub fn unknown_attributes(&self) -> &[UnknownAttribute] {
        unknown_attributes(self.opaque.as_deref())
    }

    #[must_use]
    pub fn unknown_elements(&self) -> &[UnknownElement] {
        unknown_elements(self.opaque.as_deref())
    }
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Customs {
    pub(crate) and: bool,
    pub(crate) filters: Vec<Custom>,
    pub(crate) opaque: Option<Box<OpaqueFields>>,
}
impl Customs {
    pub fn new(and: bool, filters: Vec<Custom>) -> Result<Self> {
        if !(1..=2).contains(&filters.len()) {
            return Err(invalid(
                "customFilters requires one or two customFilter children",
            ));
        }
        Ok(Self {
            and,
            filters,
            opaque: None,
        })
    }

    #[must_use]
    pub fn and(&self) -> bool {
        self.and
    }
    #[must_use]
    pub fn filters(&self) -> &[Custom] {
        &self.filters
    }

    #[must_use]
    pub fn unknown_attributes(&self) -> &[UnknownAttribute] {
        unknown_attributes(self.opaque.as_deref())
    }

    #[must_use]
    pub fn unknown_elements(&self) -> &[UnknownElement] {
        unknown_elements(self.opaque.as_deref())
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum DynamicType {
    AboveAverage,
    BelowAverage,
    Tomorrow,
    Today,
    Yesterday,
    NextWeek,
    ThisWeek,
    LastWeek,
    NextMonth,
    ThisMonth,
    LastMonth,
    NextQuarter,
    ThisQuarter,
    LastQuarter,
    NextYear,
    ThisYear,
    LastYear,
    YearToDate,
    Q1,
    Q2,
    Q3,
    Q4,
    M1,
    M2,
    M3,
    M4,
    M5,
    M6,
    M7,
    M8,
    M9,
    M10,
    M11,
    M12,
    Null,
}
impl DynamicType {
    pub(crate) fn parse(v: &str) -> Result<Self> {
        use DynamicType::{
            AboveAverage, BelowAverage, LastMonth, LastQuarter, LastWeek, LastYear, M1, M2, M3, M4,
            M5, M6, M7, M8, M9, M10, M11, M12, NextMonth, NextQuarter, NextWeek, NextYear, Null,
            Q1, Q2, Q3, Q4, ThisMonth, ThisQuarter, ThisWeek, ThisYear, Today, Tomorrow,
            YearToDate, Yesterday,
        };
        match v {
            "aboveAverage" => Ok(AboveAverage),
            "belowAverage" => Ok(BelowAverage),
            "tomorrow" => Ok(Tomorrow),
            "today" => Ok(Today),
            "yesterday" => Ok(Yesterday),
            "nextWeek" => Ok(NextWeek),
            "thisWeek" => Ok(ThisWeek),
            "lastWeek" => Ok(LastWeek),
            "nextMonth" => Ok(NextMonth),
            "thisMonth" => Ok(ThisMonth),
            "lastMonth" => Ok(LastMonth),
            "nextQuarter" => Ok(NextQuarter),
            "thisQuarter" => Ok(ThisQuarter),
            "lastQuarter" => Ok(LastQuarter),
            "nextYear" => Ok(NextYear),
            "thisYear" => Ok(ThisYear),
            "lastYear" => Ok(LastYear),
            "yearToDate" => Ok(YearToDate),
            "Q1" => Ok(Q1),
            "Q2" => Ok(Q2),
            "Q3" => Ok(Q3),
            "Q4" => Ok(Q4),
            "M1" => Ok(M1),
            "M2" => Ok(M2),
            "M3" => Ok(M3),
            "M4" => Ok(M4),
            "M5" => Ok(M5),
            "M6" => Ok(M6),
            "M7" => Ok(M7),
            "M8" => Ok(M8),
            "M9" => Ok(M9),
            "M10" => Ok(M10),
            "M11" => Ok(M11),
            "M12" => Ok(M12),
            "null" => Ok(Null),
            _ => Err(invalid(format!("invalid dynamic-filter type '{v}'"))),
        }
    }
}

#[derive(Debug, Clone, PartialEq)]
pub struct Dynamic {
    pub(crate) filter_type: DynamicType,
    pub(crate) value: Option<f64>,
    pub(crate) max_value: Option<f64>,
    pub(crate) opaque: Option<Box<OpaqueFields>>,
}
impl Dynamic {
    pub fn new(
        filter_type: DynamicType,
        value: Option<f64>,
        max_value: Option<f64>,
    ) -> Result<Self> {
        if value.is_some_and(|value| !value.is_finite())
            || max_value.is_some_and(|value| !value.is_finite())
        {
            return Err(invalid("non-finite filter number"));
        }
        if max_value.is_some() && value.is_none() {
            return Err(invalid("dynamicFilter maxVal requires val"));
        }
        if let (Some(value), Some(max_value)) = (value, max_value)
            && value >= max_value
        {
            return Err(invalid("dynamicFilter val must be less than maxVal"));
        }
        Ok(Self {
            filter_type,
            value,
            max_value,
            opaque: None,
        })
    }

    #[must_use]
    pub fn filter_type(&self) -> DynamicType {
        self.filter_type
    }
    #[must_use]
    pub fn value(&self) -> Option<f64> {
        self.value
    }
    #[must_use]
    pub fn max_value(&self) -> Option<f64> {
        self.max_value
    }

    #[must_use]
    pub fn unknown_attributes(&self) -> &[UnknownAttribute] {
        unknown_attributes(self.opaque.as_deref())
    }

    #[must_use]
    pub fn unknown_elements(&self) -> &[UnknownElement] {
        unknown_elements(self.opaque.as_deref())
    }
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Color {
    pub(crate) differential_format_id: u32,
    pub(crate) cell_color: bool,
    pub(crate) opaque: Option<Box<OpaqueFields>>,
}
impl Color {
    #[must_use]
    pub fn new(differential_format_id: u32, cell_color: bool) -> Self {
        Self {
            differential_format_id,
            cell_color,
            opaque: None,
        }
    }

    #[must_use]
    pub fn differential_format_id(&self) -> u32 {
        self.differential_format_id
    }
    #[must_use]
    pub fn cell_color(&self) -> bool {
        self.cell_color
    }

    #[must_use]
    pub fn unknown_attributes(&self) -> &[UnknownAttribute] {
        unknown_attributes(self.opaque.as_deref())
    }

    #[must_use]
    pub fn unknown_elements(&self) -> &[UnknownElement] {
        unknown_elements(self.opaque.as_deref())
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum IconSet {
    ThreeArrows,
    ThreeArrowsGray,
    ThreeFlags,
    ThreeTrafficLights1,
    ThreeTrafficLights2,
    ThreeSigns,
    ThreeSymbols,
    ThreeSymbols2,
    FourArrows,
    FourArrowsGray,
    FourRedToBlack,
    FourRating,
    FourTrafficLights,
    FiveArrows,
    FiveArrowsGray,
    FiveRating,
    FiveQuarters,
    ThreeStars,
    ThreeTriangles,
    FiveBoxes,
    NoIcons,
}
impl IconSet {
    pub(crate) fn parse(v: &str) -> Result<Self> {
        use IconSet::{
            FiveArrows, FiveArrowsGray, FiveBoxes, FiveQuarters, FiveRating, FourArrows,
            FourArrowsGray, FourRating, FourRedToBlack, FourTrafficLights, NoIcons, ThreeArrows,
            ThreeArrowsGray, ThreeFlags, ThreeSigns, ThreeStars, ThreeSymbols, ThreeSymbols2,
            ThreeTrafficLights1, ThreeTrafficLights2, ThreeTriangles,
        };
        match v {
            "3Arrows" => Ok(ThreeArrows),
            "3ArrowsGray" => Ok(ThreeArrowsGray),
            "3Flags" => Ok(ThreeFlags),
            "3TrafficLights1" => Ok(ThreeTrafficLights1),
            "3TrafficLights2" => Ok(ThreeTrafficLights2),
            "3Signs" => Ok(ThreeSigns),
            "3Symbols" => Ok(ThreeSymbols),
            "3Symbols2" => Ok(ThreeSymbols2),
            "4Arrows" => Ok(FourArrows),
            "4ArrowsGray" => Ok(FourArrowsGray),
            "4RedToBlack" => Ok(FourRedToBlack),
            "4Rating" => Ok(FourRating),
            "4TrafficLights" => Ok(FourTrafficLights),
            "5Arrows" => Ok(FiveArrows),
            "5ArrowsGray" => Ok(FiveArrowsGray),
            "5Rating" => Ok(FiveRating),
            "5Quarters" => Ok(FiveQuarters),
            "3Stars" => Ok(ThreeStars),
            "3Triangles" => Ok(ThreeTriangles),
            "5Boxes" => Ok(FiveBoxes),
            "NoIcons" => Ok(NoIcons),
            _ => Err(invalid(format!("invalid icon set '{v}'"))),
        }
    }
    pub(crate) fn cardinality(self) -> u32 {
        match self {
            Self::ThreeArrows
            | Self::ThreeArrowsGray
            | Self::ThreeFlags
            | Self::ThreeTrafficLights1
            | Self::ThreeTrafficLights2
            | Self::ThreeSigns
            | Self::ThreeSymbols
            | Self::ThreeSymbols2 => 3,
            Self::FourArrows
            | Self::FourArrowsGray
            | Self::FourRedToBlack
            | Self::FourRating
            | Self::FourTrafficLights => 4,
            Self::NoIcons => 1,
            Self::FiveArrows
            | Self::FiveArrowsGray
            | Self::FiveRating
            | Self::FiveQuarters
            | Self::ThreeStars
            | Self::ThreeTriangles
            | Self::FiveBoxes => 5,
        }
    }
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Icon {
    pub(crate) icon_set: IconSet,
    pub(crate) icon_id: u32,
    pub(crate) opaque: Option<Box<OpaqueFields>>,
}
impl Icon {
    pub fn new(icon_set: IconSet, icon_id: u32) -> Result<Self> {
        if icon_set == IconSet::NoIcons {
            if icon_id != 0 {
                return Err(invalid("NoIcons iconFilter requires iconId 0"));
            }
        } else if icon_id >= icon_set.cardinality() {
            return Err(invalid("iconFilter iconId exceeds icon-set cardinality"));
        }
        Ok(Self {
            icon_set,
            icon_id,
            opaque: None,
        })
    }

    #[must_use]
    pub fn icon_set(&self) -> IconSet {
        self.icon_set
    }
    #[must_use]
    pub fn icon_id(&self) -> u32 {
        self.icon_id
    }

    #[must_use]
    pub fn unknown_attributes(&self) -> &[UnknownAttribute] {
        unknown_attributes(self.opaque.as_deref())
    }

    #[must_use]
    pub fn unknown_elements(&self) -> &[UnknownElement] {
        unknown_elements(self.opaque.as_deref())
    }
}

#[derive(Debug, Clone, PartialEq)]
pub struct Top10 {
    pub(crate) top: bool,
    pub(crate) percent: bool,
    pub(crate) value: f64,
    pub(crate) filter_value: Option<f64>,
    pub(crate) opaque: Option<Box<OpaqueFields>>,
}
impl Top10 {
    pub fn new(top: bool, percent: bool, value: f64, filter_value: Option<f64>) -> Result<Self> {
        if !value.is_finite()
            || filter_value.is_some_and(|value| !value.is_finite())
            || (!percent && !(1.0..=500.0).contains(&value))
            || (percent && !(0.0..=100.0).contains(&value))
        {
            return Err(invalid("top10 val is out of range"));
        }
        Ok(Self {
            top,
            percent,
            value,
            filter_value,
            opaque: None,
        })
    }

    #[must_use]
    pub fn top(&self) -> bool {
        self.top
    }
    #[must_use]
    pub fn percent(&self) -> bool {
        self.percent
    }
    #[must_use]
    pub fn value(&self) -> f64 {
        self.value
    }
    #[must_use]
    pub fn filter_value(&self) -> Option<f64> {
        self.filter_value
    }

    #[must_use]
    pub fn unknown_attributes(&self) -> &[UnknownAttribute] {
        unknown_attributes(self.opaque.as_deref())
    }

    #[must_use]
    pub fn unknown_elements(&self) -> &[UnknownElement] {
        unknown_elements(self.opaque.as_deref())
    }
}

pub(crate) fn bounded(value: &str) -> Result<()> {
    if value.chars().count() > MAX_TEXT_CHARS {
        Err(invalid("filter value is too large"))
    } else {
        Ok(())
    }
}

#[derive(Debug, Clone, PartialEq)]
pub enum Payload {
    Values(Values),
    Custom(Customs),
    Dynamic(Dynamic),
    Color(Color),
    Icon(Icon),
    Top10(Top10),
}

#[derive(Debug, Clone, PartialEq)]
pub struct Column {
    pub column_id: u32,
    pub hidden_button: bool,
    pub show_button: bool,
    pub payload: Option<Payload>,
    pub(crate) opaque: Option<Box<OpaqueFields>>,
}
impl Column {
    pub fn new(column_id: u32) -> Result<Self> {
        let max_columns = u32::try_from(MAX_COLUMNS).map_err(|_source| {
            invalid("worksheet column limit exceeds the filterColumn wire type")
        })?;
        if column_id >= max_columns {
            return Err(invalid("filterColumn colId is outside worksheet range"));
        }
        Ok(Self {
            column_id,
            hidden_button: false,
            show_button: true,
            payload: None,
            opaque: None,
        })
    }

    pub fn set_hidden_button(&mut self, value: bool) -> &mut Self {
        self.hidden_button = value;
        self
    }

    pub fn set_show_button(&mut self, value: bool) -> &mut Self {
        self.show_button = value;
        self
    }

    pub fn set_payload(&mut self, value: Option<Payload>) -> &mut Self {
        self.payload = value;
        self
    }

    #[must_use]
    pub fn column_id(&self) -> u32 {
        self.column_id
    }
    #[must_use]
    pub fn hidden_button(&self) -> bool {
        self.hidden_button
    }
    #[must_use]
    pub fn show_button(&self) -> bool {
        self.show_button
    }
    #[must_use]
    pub fn payload(&self) -> Option<&Payload> {
        self.payload.as_ref()
    }

    #[must_use]
    pub fn unknown_attributes(&self) -> &[UnknownAttribute] {
        unknown_attributes(self.opaque.as_deref())
    }

    #[must_use]
    pub fn unknown_elements(&self) -> &[UnknownElement] {
        unknown_elements(self.opaque.as_deref())
    }
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Condition {
    pub(crate) reference: Range,
    pub(crate) descending: bool,
    pub(crate) sort_by: SortBy,
    pub(crate) custom_list: Option<String>,
    pub(crate) differential_format_id: Option<u32>,
    pub(crate) icon_set: Option<IconSet>,
    pub(crate) icon_id: Option<u32>,
    pub(crate) opaque: Option<Box<OpaqueFields>>,
}
impl Condition {
    #[must_use]
    pub fn reference(&self) -> &Range {
        &self.reference
    }
    #[must_use]
    pub fn descending(&self) -> bool {
        self.descending
    }
    #[must_use]
    pub fn sort_by(&self) -> SortBy {
        self.sort_by
    }
    #[must_use]
    pub fn custom_list(&self) -> Option<&str> {
        self.custom_list.as_deref()
    }
    #[must_use]
    pub fn differential_format_id(&self) -> Option<u32> {
        self.differential_format_id
    }
    #[must_use]
    pub fn icon_set(&self) -> Option<IconSet> {
        self.icon_set
    }
    #[must_use]
    pub fn icon_id(&self) -> Option<u32> {
        self.icon_id
    }

    #[must_use]
    pub fn unknown_attributes(&self) -> &[UnknownAttribute] {
        unknown_attributes(self.opaque.as_deref())
    }

    #[must_use]
    pub fn unknown_elements(&self) -> &[UnknownElement] {
        unknown_elements(self.opaque.as_deref())
    }
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct State {
    pub(crate) reference: Range,
    pub(crate) column_sort: bool,
    pub(crate) case_sensitive: bool,
    pub(crate) sort_method: Option<SortMethod>,
    pub(crate) conditions: Vec<Condition>,
    pub(crate) opaque: Option<Box<OpaqueFields>>,
}
impl State {
    #[must_use]
    pub fn reference(&self) -> &Range {
        &self.reference
    }
    #[must_use]
    pub fn column_sort(&self) -> bool {
        self.column_sort
    }
    #[must_use]
    pub fn case_sensitive(&self) -> bool {
        self.case_sensitive
    }
    #[must_use]
    pub fn sort_method(&self) -> Option<SortMethod> {
        self.sort_method
    }
    #[must_use]
    pub fn conditions(&self) -> &[Condition] {
        &self.conditions
    }

    #[must_use]
    pub fn unknown_attributes(&self) -> &[UnknownAttribute] {
        unknown_attributes(self.opaque.as_deref())
    }

    #[must_use]
    pub fn unknown_elements(&self) -> &[UnknownElement] {
        unknown_elements(self.opaque.as_deref())
    }
}

#[derive(Debug, Clone, PartialEq)]
pub struct Definition {
    pub reference: Option<Range>,
    pub columns: Vec<Column>,
    pub sort_state: Option<State>,
    pub(crate) opaque: Option<Box<OpaqueFields>>,
}
impl Definition {
    #[must_use]
    pub fn new(reference: Option<Range>) -> Self {
        Self {
            reference,
            columns: Vec::new(),
            sort_state: None,
            opaque: None,
        }
    }
    #[must_use]
    pub fn reference(&self) -> Option<&Range> {
        self.reference.as_ref()
    }
    #[must_use]
    pub fn columns(&self) -> &[Column] {
        &self.columns
    }
    #[must_use]
    pub fn sort_state(&self) -> Option<&State> {
        self.sort_state.as_ref()
    }

    #[must_use]
    pub fn unknown_attributes(&self) -> &[UnknownAttribute] {
        unknown_attributes(self.opaque.as_deref())
    }

    #[must_use]
    pub fn unknown_elements(&self) -> &[UnknownElement] {
        unknown_elements(self.opaque.as_deref())
    }
}

impl Range {
    pub(crate) fn from_parsed(value: String) -> Self {
        Self(value)
    }
}
