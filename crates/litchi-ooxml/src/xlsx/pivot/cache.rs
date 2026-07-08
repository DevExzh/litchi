#[derive(Debug, Clone)]
pub enum SharedItem {
    Missing,
    Number(f64),
    Boolean(bool),
    Error(String),
    String(String),
    DateTime(String),
}

impl SharedItem {
    pub fn as_string(&self) -> String {
        match self {
            Self::Missing => String::new(),
            Self::Number(n) => n.to_string(),
            Self::Boolean(b) => (if *b { "TRUE" } else { "FALSE" }).to_string(),
            Self::Error(e) => e.clone(),
            Self::String(s) => s.clone(),
            Self::DateTime(d) => d.clone(),
        }
    }
}

#[derive(Debug, Clone)]
pub struct PivotCacheField {
    pub name: String,
    pub num_fmt_id: Option<u32>,
    pub database_field: bool,
    pub caption: Option<String>,
    pub property_name: Option<String>,
    pub server_field: Option<bool>,
    pub unique_list: bool,
    pub level: Option<u32>,
    pub formula: Option<String>,
    pub sql_type: Option<i32>,
    pub hierarchy: Option<i32>,
    pub member_property_field: Option<u32>,
    pub shared_items: Vec<SharedItem>,
}

impl Default for PivotCacheField {
    fn default() -> Self {
        Self {
            name: String::new(),
            num_fmt_id: None,
            database_field: true,
            caption: None,
            property_name: None,
            server_field: None,
            unique_list: true,
            level: None,
            formula: None,
            sql_type: None,
            hierarchy: None,
            member_property_field: None,
            shared_items: Vec::new(),
        }
    }
}

#[derive(Debug, Clone)]
pub struct PivotCacheDefinition {
    pub id: Option<String>,
    pub invalid: bool,
    pub save_data: bool,
    pub refresh_on_load: bool,
    pub optimize_memory: Option<bool>,
    pub enable_refresh: bool,
    pub refreshed_by: Option<String>,
    pub refreshed_date: Option<f64>,
    pub refreshed_date_iso: Option<String>,
    pub background_query: bool,
    pub missing_items_limit: Option<u32>,
    pub created_version: u8,
    pub refreshed_version: u8,
    pub min_refreshable_version: u8,
    pub record_count: Option<u32>,
    pub upgrade_on_refresh: Option<bool>,
    pub tuples_cache: Option<bool>,
    pub supports_subquery: Option<bool>,
    pub supports_advanced_drill: Option<bool>,
    pub source_type: String,
    pub source_worksheet: Option<String>,
    pub source_ref: Option<String>,
    pub source_name: Option<String>,
    pub cache_fields: Vec<PivotCacheField>,
}

impl Default for PivotCacheDefinition {
    fn default() -> Self {
        Self {
            id: None,
            invalid: false,
            save_data: true,
            refresh_on_load: false,
            optimize_memory: None,
            enable_refresh: true,
            refreshed_by: None,
            refreshed_date: None,
            refreshed_date_iso: None,
            background_query: true,
            missing_items_limit: None,
            created_version: 3,
            refreshed_version: 3,
            min_refreshable_version: 3,
            record_count: None,
            upgrade_on_refresh: None,
            tuples_cache: None,
            supports_subquery: None,
            supports_advanced_drill: None,
            source_type: "worksheet".to_string(),
            source_worksheet: None,
            source_ref: None,
            source_name: None,
            cache_fields: Vec::new(),
        }
    }
}

#[derive(Debug, Clone, Default)]
pub struct CacheRecord {
    pub values: Vec<SharedItem>,
}

#[derive(Debug, Clone, Default)]
pub struct PivotCacheRecords {
    pub records: Vec<CacheRecord>,
}
