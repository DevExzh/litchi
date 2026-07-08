use super::AxisType;

#[derive(Debug, Clone)]
pub struct Index {
    pub v: u32,
}

#[derive(Debug, Clone)]
pub struct Reference {
    pub field: Option<u32>,
    pub selected: Option<bool>,
    pub by_position: Option<bool>,
    pub relative: Option<bool>,
    pub default_subtotal: Option<bool>,
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
    pub x: Vec<Index>,
}

#[derive(Debug, Clone)]
pub struct PivotArea {
    pub references: Vec<Reference>,
    pub field: Option<u32>,
    pub area_type: String,
    pub data_only: bool,
    pub label_only: Option<bool>,
    pub grand_row: Option<bool>,
    pub grand_col: Option<bool>,
    pub cache_index: Option<bool>,
    pub outline: bool,
    pub offset: Option<String>,
    pub collapsed_levels_are_subtotals: Option<bool>,
    pub axis: Option<AxisType>,
    pub field_position: Option<u32>,
}

impl Default for PivotArea {
    fn default() -> Self {
        Self {
            references: Vec::new(),
            field: None,
            area_type: "normal".to_string(),
            data_only: true,
            label_only: None,
            grand_row: None,
            grand_col: None,
            cache_index: None,
            outline: true,
            offset: None,
            collapsed_levels_are_subtotals: None,
            axis: None,
            field_position: None,
        }
    }
}

#[derive(Debug, Clone)]
pub struct PivotFilter {
    pub fld: u32,
    pub mp_fld: Option<u32>,
    pub filter_type: String,
    pub eval_order: Option<u32>,
    pub id: u32,
    pub i_measure_hier: Option<u32>,
    pub i_measure_fld: Option<u32>,
    pub name: Option<String>,
    pub description: Option<String>,
    pub string_value1: Option<String>,
    pub string_value2: Option<String>,
}

impl Default for PivotFilter {
    fn default() -> Self {
        Self {
            fld: 0,
            mp_fld: None,
            filter_type: "unknown".to_string(),
            eval_order: None,
            id: 0,
            i_measure_hier: None,
            i_measure_fld: None,
            name: None,
            description: None,
            string_value1: None,
            string_value2: None,
        }
    }
}
