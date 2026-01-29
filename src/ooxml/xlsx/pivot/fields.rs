use super::{AxisType, ItemType, SortType};
use crate::ooxml::pivot::PivotValueFunction;

#[derive(Debug, Clone)]
pub struct FieldItem {
    pub name: Option<String>,
    pub item_type: ItemType,
    pub hidden: Option<bool>,
    pub selected: Option<bool>,
    pub show_detail: bool,
    pub formula: Option<bool>,
    pub missing: Option<bool>,
    pub child: Option<bool>,
    pub index: Option<u32>,
    pub expanded: Option<bool>,
    pub drill_across_attributes: Option<bool>,
}

impl Default for FieldItem {
    fn default() -> Self {
        Self {
            name: None,
            item_type: ItemType::Data,
            hidden: None,
            selected: None,
            show_detail: true,
            formula: None,
            missing: None,
            child: None,
            index: None,
            expanded: None,
            drill_across_attributes: None,
        }
    }
}

#[derive(Debug, Clone)]
pub struct PivotField {
    pub name: Option<String>,
    pub axis: Option<AxisType>,
    pub data_field: Option<bool>,
    pub subtotal_caption: Option<String>,
    pub show_drop_downs: bool,
    pub hidden_level: Option<bool>,
    pub unique_member_property: Option<String>,
    pub compact: bool,
    pub all_drilled: Option<bool>,
    pub num_fmt_id: Option<u32>,
    pub outline: bool,
    pub subtotal_top: bool,
    pub drag_to_row: bool,
    pub drag_to_col: bool,
    pub multiple_item_selection_allowed: Option<bool>,
    pub drag_to_page: bool,
    pub drag_to_data: bool,
    pub drag_off: bool,
    pub show_all: bool,
    pub insert_blank_row: Option<bool>,
    pub server_field: Option<bool>,
    pub insert_page_break: Option<bool>,
    pub auto_show: Option<bool>,
    pub top_auto_show: bool,
    pub hide_new_items: Option<bool>,
    pub measure_filter: Option<bool>,
    pub include_new_items_in_filter: Option<bool>,
    pub item_page_count: u32,
    pub sort_type: SortType,
    pub data_source_sort: Option<bool>,
    pub non_auto_sort_default: Option<bool>,
    pub rank_by: Option<u32>,
    pub default_subtotal: bool,
    pub sum_subtotal: Option<bool>,
    pub count_a_subtotal: Option<bool>,
    pub avg_subtotal: Option<bool>,
    pub max_subtotal: Option<bool>,
    pub min_subtotal: Option<bool>,
    pub product_subtotal: Option<bool>,
    pub count_subtotal: Option<bool>,
    pub std_dev_subtotal: Option<bool>,
    pub std_dev_p_subtotal: Option<bool>,
    pub var_subtotal: Option<bool>,
    pub var_p_subtotal: Option<bool>,
    pub show_prop_cell: Option<bool>,
    pub show_prop_tip: Option<bool>,
    pub show_prop_as_caption: Option<bool>,
    pub default_attribute_drill_state: Option<bool>,
    pub items: Vec<FieldItem>,
}

impl Default for PivotField {
    fn default() -> Self {
        Self {
            name: None,
            axis: None,
            data_field: None,
            subtotal_caption: None,
            show_drop_downs: true,
            hidden_level: None,
            unique_member_property: None,
            compact: true,
            all_drilled: None,
            num_fmt_id: None,
            outline: true,
            subtotal_top: true,
            drag_to_row: true,
            drag_to_col: true,
            multiple_item_selection_allowed: None,
            drag_to_page: true,
            drag_to_data: true,
            drag_off: true,
            show_all: true,
            insert_blank_row: None,
            server_field: None,
            insert_page_break: None,
            auto_show: None,
            top_auto_show: true,
            hide_new_items: None,
            measure_filter: None,
            include_new_items_in_filter: None,
            item_page_count: 10,
            sort_type: SortType::Manual,
            data_source_sort: None,
            non_auto_sort_default: None,
            rank_by: None,
            default_subtotal: true,
            sum_subtotal: None,
            count_a_subtotal: None,
            avg_subtotal: None,
            max_subtotal: None,
            min_subtotal: None,
            product_subtotal: None,
            count_subtotal: None,
            std_dev_subtotal: None,
            std_dev_p_subtotal: None,
            var_subtotal: None,
            var_p_subtotal: None,
            show_prop_cell: None,
            show_prop_tip: None,
            show_prop_as_caption: None,
            default_attribute_drill_state: None,
            items: Vec::new(),
        }
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum Subtotal {
    Average,
    Count,
    CountNums,
    Max,
    Min,
    Product,
    StdDev,
    StdDevP,
    Sum,
    Var,
    VarP,
}

impl Subtotal {
    pub fn as_str(&self) -> &'static str {
        match self {
            Self::Average => "average",
            Self::Count => "count",
            Self::CountNums => "countNums",
            Self::Max => "max",
            Self::Min => "min",
            Self::Product => "product",
            Self::StdDev => "stdDev",
            Self::StdDevP => "stdDevp",
            Self::Sum => "sum",
            Self::Var => "var",
            Self::VarP => "varp",
        }
    }

    pub fn parse_str(s: &str) -> Option<Self> {
        match s {
            "average" => Some(Self::Average),
            "count" => Some(Self::Count),
            "countNums" => Some(Self::CountNums),
            "max" => Some(Self::Max),
            "min" => Some(Self::Min),
            "product" => Some(Self::Product),
            "stdDev" => Some(Self::StdDev),
            "stdDevp" => Some(Self::StdDevP),
            "sum" => Some(Self::Sum),
            "var" => Some(Self::Var),
            "varp" => Some(Self::VarP),
            _ => None,
        }
    }

    pub fn to_pivot_value_function(self) -> PivotValueFunction {
        match self {
            Self::Average => PivotValueFunction::Average,
            Self::Count | Self::CountNums => PivotValueFunction::Count,
            Self::Max => PivotValueFunction::Max,
            Self::Min => PivotValueFunction::Min,
            Self::Sum => PivotValueFunction::Sum,
            _ => PivotValueFunction::Custom,
        }
    }
}

#[derive(Debug, Clone)]
pub struct DataField {
    pub name: Option<String>,
    pub fld: u32,
    pub subtotal: Subtotal,
    pub show_data_as: String,
    pub base_field: i32,
    pub base_item: u32,
    pub num_fmt_id: Option<u32>,
}

impl Default for DataField {
    fn default() -> Self {
        Self {
            name: None,
            fld: 0,
            subtotal: Subtotal::Sum,
            show_data_as: "normal".to_string(),
            base_field: -1,
            base_item: 1048832,
            num_fmt_id: None,
        }
    }
}

#[derive(Debug, Clone)]
pub struct PageField {
    pub fld: u32,
    pub item: Option<u32>,
    pub hier: Option<u32>,
    pub name: Option<String>,
    pub cap: Option<String>,
}

#[derive(Debug, Clone)]
pub struct RowColItem {
    pub item_type: ItemType,
    pub r: u32,
    pub i: u32,
    pub x: Vec<u32>,
}

impl Default for RowColItem {
    fn default() -> Self {
        Self {
            item_type: ItemType::Data,
            r: 0,
            i: 0,
            x: Vec::new(),
        }
    }
}

#[derive(Debug, Clone)]
pub struct RowColField {
    pub x: u32,
}
