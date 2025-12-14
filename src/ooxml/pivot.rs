use std::collections::HashMap;

#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub enum PivotAxis {
    Row,
    Column,
    Filter,
    Data,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub enum PivotValueFunction {
    Sum,
    Count,
    Average,
    Min,
    Max,
    Custom,
}

#[derive(Debug, Clone)]
pub struct PivotFieldRole {
    pub field_name: String,
    pub axis: PivotAxis,
    pub position: u32,
}

#[derive(Debug, Clone)]
pub struct PivotDataField {
    pub field_name: String,
    pub function: PivotValueFunction,
    pub display_name: Option<String>,
}

#[derive(Debug, Clone)]
pub struct PivotCacheField {
    pub name: String,
    pub shared_items: Vec<String>,
}

#[derive(Debug, Clone)]
pub struct PivotCacheDefinition {
    pub id: u32,
    pub source_ref: Option<String>,
    pub fields: Vec<PivotCacheField>,
}

#[derive(Debug, Clone)]
pub struct PivotTable {
    pub name: String,
    pub source_sheet: Option<String>,
    pub source_ref: Option<String>,
    pub field_names: Vec<String>,
    pub sheet_name: String,
    pub cache_id: u32,
    pub location_ref: String,
    pub row_fields: Vec<PivotFieldRole>,
    pub column_fields: Vec<PivotFieldRole>,
    pub filter_fields: Vec<PivotFieldRole>,
    pub data_fields: Vec<PivotDataField>,
}

impl PivotTable {
    pub fn fields_by_axis(&self, axis: PivotAxis) -> &[PivotFieldRole] {
        match axis {
            PivotAxis::Row => &self.row_fields,
            PivotAxis::Column => &self.column_fields,
            PivotAxis::Filter => &self.filter_fields,
            PivotAxis::Data => &[],
        }
    }

    pub fn data_fields_map(&self) -> HashMap<&str, &PivotDataField> {
        let mut map = HashMap::with_capacity(self.data_fields.len());
        for df in &self.data_fields {
            map.insert(df.field_name.as_str(), df);
        }
        map
    }
}
