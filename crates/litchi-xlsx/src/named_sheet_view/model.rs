//! Semantic SpreadsheetML named-sheet-view model.
//!
//! The owner module supplies the contextual namespace; historical
//! prefixed names remain aliases in mod.rs.

use crate::auto_filter::FilterColumnDefinition;
use crate::error::Result;
use crate::sort::{SortBy, SortMethod};

use super::codec::{
    add_extension, filter_payload_markup, parse_authored_differential_format,
    parse_authored_extension, parse_guid, parse_range, remove_extension, validate_column_id,
    validate_name, view_has_filter_payload, write_named_sheet_views,
};
use super::{CORE, MAX_COLUMNS, MAX_FILTERS, MAX_VIEWS, NSV, invalid};

#[derive(Debug, Clone, PartialEq, Eq, Hash)]
pub struct Guid(pub(crate) String);
impl Guid {
    /// Parse a braced Named Sheet View GUID.
    pub fn new(value: impl Into<String>) -> Result<Self> {
        let value = value.into();
        parse_guid(&value)
    }

    /// Generate a fresh RFC 4122 v4 GUID in the braced OOXML representation.
    pub fn generate() -> Self {
        Self(litchi_core::id::generate_guid_braced())
    }

    pub fn as_str(&self) -> &str {
        &self.0
    }
}
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Range(pub(crate) String);
impl Range {
    /// Parse a non-reversed A1 cell or rectangular range.
    pub fn new(value: impl Into<String>) -> Result<Self> {
        let value = value.into();
        parse_range(&value)
    }

    pub fn as_str(&self) -> &str {
        &self.0
    }
}
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Markup(pub(crate) Vec<u8>);
impl Markup {
    pub fn xml(&self) -> &[u8] {
        &self.0
    }
}
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct DifferentialFormat {
    pub(crate) markup: Markup,
}
impl DifferentialFormat {
    /// Validate and retain one self-contained Named Sheet Views `dxf` element.
    ///
    /// The fragment remains inert: fonts, fills, borders, and extension
    /// payloads are never interpreted or applied by this API.
    pub fn from_xml(xml: impl AsRef<[u8]>) -> Result<Self> {
        parse_authored_differential_format(xml.as_ref())
    }

    /// Construct an empty differential-format element.
    pub fn empty() -> Self {
        Self {
            markup: Markup(
                format!(
                    r#"<dxf xmlns="{}"/>"#,
                    std::str::from_utf8(NSV).expect("constant namespace is UTF-8")
                )
                .into_bytes(),
            ),
        }
    }

    pub fn xml(&self) -> &[u8] {
        self.markup.xml()
    }
}
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Extension {
    pub(crate) uri: String,
    pub(crate) markup: Markup,
}
impl Extension {
    /// Construct one extension from its URI and bounded XML content.
    ///
    /// `content_xml` is inserted below a SpreadsheetML `ext` element and the
    /// complete result is validated before it can enter the model.
    pub fn new(uri: impl Into<String>, content_xml: impl AsRef<[u8]>) -> Result<Self> {
        parse_authored_extension(uri.into(), content_xml.as_ref())
    }

    pub fn uri(&self) -> &str {
        &self.uri
    }
    pub fn markup(&self) -> &Markup {
        &self.markup
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum SortConditionKind {
    Standard,
    RichValue,
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
        use IconSet::*;
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
            _ => Err(invalid(format!("invalid named-sheet-view icon set '{v}'"))),
        }
    }
    pub(crate) fn cardinality(self) -> Option<u32> {
        use IconSet::*;
        Some(match self {
            ThreeArrows | ThreeArrowsGray | ThreeFlags | ThreeTrafficLights1
            | ThreeTrafficLights2 | ThreeSigns | ThreeSymbols | ThreeSymbols2 | ThreeStars
            | ThreeTriangles => 3,
            FourArrows | FourArrowsGray | FourRedToBlack | FourRating | FourTrafficLights => 4,
            FiveArrows | FiveArrowsGray | FiveRating | FiveQuarters | FiveBoxes => 5,
            NoIcons => return None,
        })
    }

    pub(crate) fn as_str(self) -> &'static str {
        use IconSet::*;
        match self {
            ThreeArrows => "3Arrows",
            ThreeArrowsGray => "3ArrowsGray",
            ThreeFlags => "3Flags",
            ThreeTrafficLights1 => "3TrafficLights1",
            ThreeTrafficLights2 => "3TrafficLights2",
            ThreeSigns => "3Signs",
            ThreeSymbols => "3Symbols",
            ThreeSymbols2 => "3Symbols2",
            FourArrows => "4Arrows",
            FourArrowsGray => "4ArrowsGray",
            FourRedToBlack => "4RedToBlack",
            FourRating => "4Rating",
            FourTrafficLights => "4TrafficLights",
            FiveArrows => "5Arrows",
            FiveArrowsGray => "5ArrowsGray",
            FiveRating => "5Rating",
            FiveQuarters => "5Quarters",
            ThreeStars => "3Stars",
            ThreeTriangles => "3Triangles",
            FiveBoxes => "5Boxes",
            NoIcons => "NoIcons",
        }
    }
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct SortCondition {
    pub(crate) kind: SortConditionKind,
    pub(crate) reference: Range,
    pub(crate) descending: bool,
    pub(crate) sort_by: SortBy,
    pub(crate) custom_list: Option<String>,
    pub(crate) differential_format_id: Option<u32>,
    pub(crate) icon_set: Option<IconSet>,
    pub(crate) icon_id: Option<u32>,
    pub(crate) rich_sort_key: Option<String>,
}
impl SortCondition {
    /// Create a value-based sort condition for a worksheet range.
    pub fn new(kind: SortConditionKind, reference: Range) -> Self {
        Self {
            kind,
            reference,
            descending: false,
            sort_by: SortBy::Value,
            custom_list: None,
            differential_format_id: None,
            icon_set: None,
            icon_id: None,
            rich_sort_key: None,
        }
    }

    pub fn set_descending(&mut self, value: bool) -> &mut Self {
        self.descending = value;
        self
    }

    /// Set an optional custom value order.
    pub fn set_custom_list(&mut self, value: Option<String>) -> Result<&mut Self> {
        if value
            .as_ref()
            .is_some_and(|value| value.chars().count() > 32_767)
        {
            return Err(invalid("customList exceeds 32767 characters"));
        }
        self.custom_list = value;
        Ok(self)
    }

    /// Configure icon sorting, validating the icon index against its set.
    pub fn set_icon_sort(&mut self, icon_set: IconSet, icon_id: Option<u32>) -> Result<&mut Self> {
        if icon_id.is_some_and(|id| icon_set.cardinality().is_some_and(|count| id >= count)) {
            return Err(invalid("iconId is outside icon set"));
        }
        self.sort_by = SortBy::Icon;
        self.differential_format_id = None;
        self.icon_set = Some(icon_set);
        self.icon_id = icon_id;
        Ok(self)
    }

    /// Configure sorting by a cell fill or font color.
    pub fn set_color_sort(
        &mut self,
        sort_by: SortBy,
        differential_format_id: u32,
    ) -> Result<&mut Self> {
        if !matches!(sort_by, SortBy::CellColor | SortBy::FontColor) {
            return Err(invalid(
                "color sort requires CellColor or FontColor sort kind",
            ));
        }
        self.sort_by = sort_by;
        self.differential_format_id = Some(differential_format_id);
        self.icon_set = None;
        self.icon_id = None;
        Ok(self)
    }

    /// Set the rich-value sort key. Standard sort conditions reject this metadata.
    pub fn set_rich_sort_key(&mut self, value: Option<String>) -> Result<&mut Self> {
        if self.kind != SortConditionKind::RichValue && value.is_some() {
            return Err(invalid("standard sort condition cannot have richSortKey"));
        }
        if value
            .as_ref()
            .is_some_and(|value| value.chars().count() > 255)
        {
            return Err(invalid("richSortKey exceeds 255 characters"));
        }
        self.rich_sort_key = value;
        Ok(self)
    }

    pub fn kind(&self) -> SortConditionKind {
        self.kind
    }
    pub fn reference(&self) -> &Range {
        &self.reference
    }
    pub fn descending(&self) -> bool {
        self.descending
    }
    pub fn sort_by(&self) -> SortBy {
        self.sort_by
    }
    pub fn custom_list(&self) -> Option<&str> {
        self.custom_list.as_deref()
    }
    pub fn differential_format_id(&self) -> Option<u32> {
        self.differential_format_id
    }
    pub fn icon_set(&self) -> Option<IconSet> {
        self.icon_set
    }
    pub fn icon_id(&self) -> Option<u32> {
        self.icon_id
    }
    pub fn rich_sort_key(&self) -> Option<&str> {
        self.rich_sort_key.as_deref()
    }
}
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct SortRule {
    pub(crate) column_id: u32,
    pub(crate) id: Option<Guid>,
    pub(crate) differential_format: Option<Markup>,
    pub(crate) condition: Option<SortCondition>,
}
impl SortRule {
    pub fn new(column_id: u32) -> Result<Self> {
        validate_column_id(column_id)?;
        Ok(Self {
            column_id,
            id: None,
            differential_format: None,
            condition: None,
        })
    }

    pub fn set_id(&mut self, id: Option<Guid>) -> &mut Self {
        self.id = id;
        self
    }

    pub fn set_condition(&mut self, condition: Option<SortCondition>) -> Result<&mut Self> {
        if condition
            .as_ref()
            .is_some_and(|condition| condition.differential_format_id.is_some())
            != self.differential_format.is_some()
        {
            return Err(invalid(
                "sortRule dxf presence does not match sortCondition dxfId",
            ));
        }
        self.condition = condition;
        Ok(self)
    }

    pub fn set_differential_format(
        &mut self,
        value: Option<DifferentialFormat>,
    ) -> Result<&mut Self> {
        if let Some(condition) = self.condition.as_ref()
            && condition.differential_format_id.is_some() != value.is_some()
        {
            return Err(invalid(
                "sortRule dxf presence does not match sortCondition dxfId",
            ));
        }
        self.differential_format = value.map(|value| value.markup);
        Ok(self)
    }

    pub fn column_id(&self) -> u32 {
        self.column_id
    }
    pub fn id(&self) -> Option<&Guid> {
        self.id.as_ref()
    }
    pub fn differential_format(&self) -> Option<&Markup> {
        self.differential_format.as_ref()
    }
    pub fn condition(&self) -> Option<&SortCondition> {
        self.condition.as_ref()
    }
}
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct SortRules {
    pub(crate) sort_method: SortMethod,
    pub(crate) case_sensitive: bool,
    pub(crate) rules: Vec<SortRule>,
    pub(crate) extensions: Vec<Extension>,
}
impl SortRules {
    pub fn new() -> Self {
        Self {
            sort_method: SortMethod::None,
            case_sensitive: false,
            rules: Vec::new(),
            extensions: Vec::new(),
        }
    }

    pub fn set_sort_method(&mut self, value: SortMethod) -> &mut Self {
        self.sort_method = value;
        self
    }

    pub fn set_case_sensitive(&mut self, value: bool) -> &mut Self {
        self.case_sensitive = value;
        self
    }

    pub fn add_rule(&mut self, rule: SortRule) -> Result<&mut Self> {
        if self.rules.len() >= 64 {
            return Err(invalid("sortRules exceeds 64 rules"));
        }
        self.rules.push(rule);
        Ok(self)
    }

    pub fn remove_rule(&mut self, column_id: u32) -> Option<SortRule> {
        self.rules
            .iter()
            .position(|rule| rule.column_id == column_id)
            .map(|index| self.rules.remove(index))
    }

    pub fn add_extension(&mut self, value: Extension) -> Result<&mut Self> {
        add_extension(&mut self.extensions, value)?;
        Ok(self)
    }

    pub fn remove_extension(&mut self, uri: &str) -> Option<Extension> {
        remove_extension(&mut self.extensions, uri)
    }

    pub fn sort_method(&self) -> SortMethod {
        self.sort_method
    }
    pub fn case_sensitive(&self) -> bool {
        self.case_sensitive
    }
    pub fn rules(&self) -> &[SortRule] {
        &self.rules
    }
    pub fn extensions(&self) -> &[Extension] {
        &self.extensions
    }
}
impl Default for SortRules {
    fn default() -> Self {
        Self::new()
    }
}
#[derive(Debug, Clone, PartialEq)]
pub struct ColumnFilter {
    pub(crate) column_id: u32,
    pub(crate) id: Option<Guid>,
    pub(crate) differential_format: Option<Markup>,
    pub(crate) filters: Vec<FilterColumnDefinition>,
    pub(crate) extensions: Vec<Extension>,
}
impl ColumnFilter {
    pub fn new(column_id: u32) -> Result<Self> {
        validate_column_id(column_id)?;
        Ok(Self {
            column_id,
            id: None,
            differential_format: None,
            filters: Vec::new(),
            extensions: Vec::new(),
        })
    }

    pub fn set_id(&mut self, id: Option<Guid>) -> &mut Self {
        self.id = id;
        self
    }

    /// Add a filter payload for this column.
    ///
    /// The shared SpreadsheetML auto-filter serializer validates the payload
    /// before this model is mutated.
    pub fn add_filter(&mut self, filter: FilterColumnDefinition) -> Result<&mut Self> {
        if self.filters.len() >= MAX_FILTERS {
            return Err(invalid("too many filter payloads"));
        }
        if filter.column_id != self.column_id {
            return Err(invalid(
                "named-sheet-view filter colId does not match columnFilter colId",
            ));
        }
        filter_payload_markup(&filter)?;
        self.filters.push(filter);
        Ok(self)
    }

    pub fn clear_filters(&mut self) {
        self.filters.clear();
    }

    pub fn set_differential_format(&mut self, value: Option<DifferentialFormat>) -> &mut Self {
        self.differential_format = value.map(|value| value.markup);
        self
    }

    pub fn add_extension(&mut self, value: Extension) -> Result<&mut Self> {
        add_extension(&mut self.extensions, value)?;
        Ok(self)
    }

    pub fn remove_extension(&mut self, uri: &str) -> Option<Extension> {
        remove_extension(&mut self.extensions, uri)
    }

    pub fn column_id(&self) -> u32 {
        self.column_id
    }
    pub fn id(&self) -> Option<&Guid> {
        self.id.as_ref()
    }
    pub fn differential_format(&self) -> Option<&Markup> {
        self.differential_format.as_ref()
    }
    pub fn filters(&self) -> &[FilterColumnDefinition] {
        &self.filters
    }
    pub fn extensions(&self) -> &[Extension] {
        &self.extensions
    }
}
#[derive(Debug, Clone, PartialEq)]
pub struct Filter {
    pub(crate) filter_id: Guid,
    pub(crate) reference: Option<Range>,
    pub(crate) table_id: Option<u32>,
    pub(crate) column_filters: Vec<ColumnFilter>,
    pub(crate) sort_rules: Option<SortRules>,
    pub(crate) extensions: Vec<Extension>,
}
impl Filter {
    pub fn new(filter_id: Guid) -> Self {
        Self {
            filter_id,
            reference: None,
            table_id: None,
            column_filters: Vec::new(),
            sort_rules: None,
            extensions: Vec::new(),
        }
    }

    pub fn set_reference(&mut self, value: Option<Range>) -> &mut Self {
        self.reference = value;
        self
    }

    pub fn set_table_id(&mut self, value: Option<u32>) -> &mut Self {
        self.table_id = value;
        self
    }

    pub fn add_column_filter(&mut self, filter: ColumnFilter) -> Result<&mut Self> {
        if self.column_filters.len() >= MAX_COLUMNS {
            return Err(invalid("too many named-sheet-view column filters"));
        }
        self.column_filters.push(filter);
        Ok(self)
    }

    pub fn remove_column_filter(&mut self, column_id: u32) -> Option<ColumnFilter> {
        self.column_filters
            .iter()
            .position(|filter| filter.column_id == column_id)
            .map(|index| self.column_filters.remove(index))
    }

    pub fn set_sort_rules(&mut self, value: Option<SortRules>) -> &mut Self {
        self.sort_rules = value;
        self
    }

    pub fn add_extension(&mut self, value: Extension) -> Result<&mut Self> {
        add_extension(&mut self.extensions, value)?;
        Ok(self)
    }

    pub fn remove_extension(&mut self, uri: &str) -> Option<Extension> {
        remove_extension(&mut self.extensions, uri)
    }

    pub fn filter_id(&self) -> &Guid {
        &self.filter_id
    }
    pub fn reference(&self) -> Option<&Range> {
        self.reference.as_ref()
    }
    pub fn table_id(&self) -> Option<u32> {
        self.table_id
    }
    pub fn column_filters(&self) -> &[ColumnFilter] {
        &self.column_filters
    }
    pub fn sort_rules(&self) -> Option<&SortRules> {
        self.sort_rules.as_ref()
    }
    pub fn extensions(&self) -> &[Extension] {
        &self.extensions
    }
}
#[derive(Debug, Clone, PartialEq)]
pub struct View {
    pub(crate) name: String,
    pub(crate) id: Guid,
    pub(crate) filters: Vec<Filter>,
    pub(crate) extensions: Vec<Extension>,
}
impl View {
    /// Create an empty Named Sheet View with a fresh identifier.
    ///
    /// The resulting view is metadata only: it does not apply filters or sort
    /// worksheet data. Use [`Views::add_view`] to add it to the
    /// worksheet-scoped collection before storing that collection.
    pub fn new(name: impl Into<String>) -> Result<Self> {
        Self::with_id(name, Guid::generate())
    }

    /// Create an empty Named Sheet View with a caller-supplied identifier.
    pub fn with_id(name: impl Into<String>, id: Guid) -> Result<Self> {
        let name = name.into();
        validate_name(&name)?;
        Ok(Self {
            name,
            id,
            filters: Vec::new(),
            extensions: Vec::new(),
        })
    }

    pub fn name(&self) -> &str {
        &self.name
    }
    pub fn id(&self) -> &Guid {
        &self.id
    }
    pub fn filters(&self) -> &[Filter] {
        &self.filters
    }
    pub fn extensions(&self) -> &[Extension] {
        &self.extensions
    }

    pub fn add_filter(&mut self, filter: Filter) -> Result<&mut Self> {
        if self.filters.len() >= MAX_FILTERS {
            return Err(invalid("too many named-sheet-view filters"));
        }
        self.filters.push(filter);
        Ok(self)
    }

    pub fn remove_filter(&mut self, filter_id: &Guid) -> Option<Filter> {
        self.filters
            .iter()
            .position(|filter| &filter.filter_id == filter_id)
            .map(|index| self.filters.remove(index))
    }

    pub fn add_extension(&mut self, value: Extension) -> Result<&mut Self> {
        add_extension(&mut self.extensions, value)?;
        Ok(self)
    }

    pub fn remove_extension(&mut self, uri: &str) -> Option<Extension> {
        remove_extension(&mut self.extensions, uri)
    }
}
#[derive(Debug, Clone, PartialEq)]
pub struct Views {
    pub(crate) views: Vec<View>,
    pub(crate) extensions: Vec<Extension>,
    pub(crate) namespace_declarations: Vec<(String, String)>,
}
impl Views {
    /// Start a valid Named Sheet Views collection with one view.
    ///
    /// A Named Sheet Views part requires at least one `namedSheetView`, so an
    /// empty collection is deliberately not constructible through this API.
    pub fn new(view: View) -> Self {
        let mut namespace_declarations =
            vec![("xmlns".into(), std::str::from_utf8(NSV).unwrap().into())];
        if view_has_filter_payload(&view) {
            namespace_declarations
                .push(("xmlns:x".into(), std::str::from_utf8(CORE).unwrap().into()));
        }
        Self {
            views: vec![view],
            extensions: Vec::new(),
            // Match the parser's retained root declarations so a freshly
            // constructed value has parse/write round-trip equality too.
            namespace_declarations,
        }
    }

    /// Add a uniquely named and identified view to this worksheet collection.
    pub fn add_view(&mut self, view: View) -> Result<&mut Self> {
        if self.views.len() >= MAX_VIEWS {
            return Err(invalid("too many named sheet views"));
        }
        if self.views.iter().any(|existing| existing.name == view.name) {
            return Err(invalid("duplicate named sheet view name"));
        }
        if self.views.iter().any(|existing| existing.id == view.id) {
            return Err(invalid("duplicate named sheet view GUID"));
        }
        if view_has_filter_payload(&view)
            && !self
                .namespace_declarations
                .iter()
                .any(|(name, _)| name == "xmlns:x")
        {
            self.namespace_declarations
                .push(("xmlns:x".into(), std::str::from_utf8(CORE).unwrap().into()));
        }
        self.views.push(view);
        Ok(self)
    }

    /// Remove a view by its worksheet-scoped name.
    ///
    /// Removing the final view is rejected because an empty Named Sheet Views
    /// part is not valid OOXML. Remove the worksheet part instead with
    /// [`remove_worksheet_named_sheet_views`] or
    /// [`crate::Workbook::remove_named_sheet_views`].
    pub fn remove_view(&mut self, name: &str) -> Result<Option<View>> {
        let Some(index) = self.views.iter().position(|view| view.name == name) else {
            return Ok(None);
        };
        if self.views.len() == 1 {
            return Err(invalid(
                "Named Sheet Views part must contain a namedSheetView",
            ));
        }
        Ok(Some(self.views.remove(index)))
    }

    pub fn views(&self) -> &[View] {
        &self.views
    }
    pub fn extensions(&self) -> &[Extension] {
        &self.extensions
    }

    pub fn add_extension(&mut self, value: Extension) -> Result<&mut Self> {
        add_extension(&mut self.extensions, value)?;
        Ok(self)
    }

    pub fn remove_extension(&mut self, uri: &str) -> Option<Extension> {
        remove_extension(&mut self.extensions, uri)
    }

    /// Serialize this parsed Named Sheet Views value without evaluating filters
    /// or changing sort semantics.
    pub fn to_xml(&self) -> Result<Vec<u8>> {
        write_named_sheet_views(self)
    }
}
