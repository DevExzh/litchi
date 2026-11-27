//! Typed semantic values for the BIFF8 `PivotTable` OLAP extension records.

/// Typed `SXViewEx` record content (MS-XLS 2.4.314): the header of the
/// `PivotTable` view OLAP extension sequence.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct PivotViewOlapHeader {
    /// Number of `SXTH` records that follow (`csxth`). MUST be at least 1.
    pub hierarchy_count: u32,
    /// Number of `SXPIEx` records that follow the `SXTH` records (`csxpi`).
    pub page_extension_count: u32,
    /// Number of `SXVDTEx` records that follow the `SXPIEx` records
    /// (`csxvdtex`).
    pub field_extension_count: u32,
    /// Information from future versions (`rgbFuture`), at most 1024 bytes.
    pub future_bytes: Vec<u8>,
}

/// The `PivotTable` axis or axes a pivot hierarchy is present on (`SXAxis`,
/// MS-XLS 2.5.254).
#[derive(Debug, Clone, Copy, PartialEq, Eq, Default)]
pub struct PivotHierarchyAxis {
    /// Whether the hierarchy is on the row axis (`sxaxisRw`).
    pub row: bool,
    /// Whether the hierarchy is on the column axis (`sxaxisCol`).
    pub column: bool,
    /// Whether the hierarchy is on the page axis (`sxaxisPage`).
    pub page: bool,
    /// Whether the hierarchy is on the data axis (`sxaxisData`).
    pub data: bool,
}

/// A `HiddenMemberSet` structure (MS-XLS 2.5.157): the OLAP members hidden
/// from the `PivotTable` view at one level of a pivot hierarchy.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct HiddenMemberSet {
    /// Names of the hidden OLAP members (`rgMemberName`), each at most 255
    /// characters.
    pub member_names: Vec<String>,
}

/// Typed `SXTH` record content (MS-XLS 2.4.308): properties of one OLAP
/// pivot hierarchy.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct PivotHierarchy {
    /// Whether this hierarchy is an OLAP measure (`fMeasure`).
    pub is_measure: bool,
    /// Whether level fields are created with `SXVDEx.fOutline` set
    /// (`fOutlineMode`).
    pub outline_mode: bool,
    /// Whether multiple OLAP members can be selected on the page axis
    /// (`fEnableMultiplePageItems`).
    pub multiple_page_items: bool,
    /// Whether level fields are created with `SXVDEx.fSubtotalAtTop` set
    /// (`fSubtotalAtTop`).
    pub subtotal_at_top: bool,
    /// Whether this hierarchy is an OLAP named set (`fSet`).
    pub is_named_set: bool,
    /// Whether this hierarchy is hidden in the field list (`fDontShowFList`).
    pub hidden_from_field_list: bool,
    /// Whether this hierarchy is an attribute hierarchy
    /// (`fAttributeHierarchy`).
    pub is_attribute_hierarchy: bool,
    /// Whether this hierarchy is a time hierarchy (`fTimeHierarchy`).
    pub is_time_hierarchy: bool,
    /// Whether manual filters are inclusive rather than exclusive
    /// (`fFilterInclusive`).
    pub filter_inclusive: bool,
    /// Whether this is the key attribute hierarchy of its dimension
    /// (`fKeyAttributeHierarchy`).
    pub is_key_attribute_hierarchy: bool,
    /// Whether this hierarchy is a KPI hierarchy (`fKPI`).
    pub is_kpi: bool,
    /// The axis or axes this hierarchy is present on (`sxaxis`).
    pub axis: PivotHierarchyAxis,
    /// The associated pivot field index (`isxvd`).
    pub pivot_field_index: i32,
    /// Number of pivot fields on `PivotTable` axes for this hierarchy
    /// (`csxvdXl`). Related to `level_fields` by the `stAll` rule.
    pub axis_field_count: i32,
    /// Whether this hierarchy can be placed on the row axis (`fDragToRow`).
    pub drag_to_row: bool,
    /// Whether this hierarchy can be placed on the column axis
    /// (`fDragToColumn`).
    pub drag_to_column: bool,
    /// Whether this hierarchy can be placed on the page axis (`fDragToPage`).
    pub drag_to_page: bool,
    /// Whether this hierarchy can be placed on the data axis (`fDragToData`).
    pub drag_to_data: bool,
    /// Whether this hierarchy can be removed from the view (`fDragToHide`).
    pub drag_to_hide: bool,
    /// MDX unique name of this hierarchy (`stUnique`), 1..=255 characters.
    pub unique_name: String,
    /// Display name of this hierarchy (`stDisplay`), 1..=255 characters.
    pub display_name: String,
    /// MDX unique name of the default member (`stDefault`), at most 255
    /// characters.
    pub default_member: String,
    /// Unique name of the ALL member (`stAll`); empty when there is no ALL
    /// member.
    pub all_member: String,
    /// Unique name of the OLAP dimension this hierarchy belongs to
    /// (`stDimension`); empty for measures.
    pub dimension: String,
    /// Pivot fields associated with this hierarchy (`rgisxvd`); each element
    /// is a pivot field index or -1 for none.
    pub level_fields: Vec<i32>,
    /// Hidden OLAP members per level (`rgHiddenMemberSets`).
    pub hidden_member_sets: Vec<HiddenMemberSet>,
}

/// Typed `SXPIEx` record content (MS-XLS 2.4.299): the OLAP extension of one
/// page-axis entry of a `PivotTable` view.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct PivotPageItemOlapExt {
    /// Pivot hierarchy index of the hierarchy on the page axis (`isxth`).
    pub hierarchy_index: u32,
    /// Unique name of the OLAP member used for filtering (`stUnique`), at
    /// most 255 characters.
    pub unique_name: String,
    /// Caption of the OLAP member (`stDisplay`), at most 255 characters.
    pub display_name: String,
}

/// An `SXVIFlags` structure (MS-XLS 2.5.263): additional OLAP properties of
/// one pivot item.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Default)]
pub struct PivotItemOlapFlags {
    /// Whether child elements of this item are collapsed (`fDrilledMember`).
    pub drilled_member: bool,
    /// Whether the item has child OLAP members (`fHasChildren`).
    pub has_children: bool,
    /// Whether the subnodes of this item are collapsed (`fCollapsedMember`).
    pub collapsed_member: bool,
    /// Whether `has_children` is considered correct (`fHasChildrenEst`).
    pub has_children_estimated: bool,
    /// Whether the item is selected for OLAP manual filtering
    /// (`fOlapFilterSelected`).
    pub olap_filter_selected: bool,
}

/// Typed `SXVDTEx` record content (MS-XLS 2.4.311): the OLAP extension of
/// one pivot field.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct PivotFieldOlapExt {
    /// Whether the sort order is determined by the OLAP source data
    /// (`fTensorSort`).
    pub tensor_sort: bool,
    /// Whether all pivot items of this field are expanded (`fDrilledLevel`).
    pub drilled_level: bool,
    /// Whether this attribute hierarchy is expanded by default
    /// (`fItemsDrilledByDefault`).
    pub items_drilled_by_default: bool,
    /// Whether this member property field is displayed in the report
    /// (`fMemPropDisplayInReport`).
    pub member_property_in_report: bool,
    /// Whether this member property field is displayed in a `ToolTip`
    /// (`fMemPropDisplayInTip`).
    pub member_property_in_tip: bool,
    /// Whether member property captions replace pivot item captions
    /// (`fMemPropDisplayInCaption`).
    pub member_property_in_caption: bool,
    /// The pivot hierarchy this field is associated with (`isxth`): a pivot
    /// hierarchy index, or -1 when the field is not part of a hierarchy.
    pub hierarchy_index: i16,
    /// Zero-based index of the associated OLAP level (`isxtl`).
    pub olap_level_index: i32,
    /// Additional properties of the pivot items (`rgsxvi`); one element per
    /// pivot item of this field.
    pub item_flags: Vec<PivotItemOlapFlags>,
}
