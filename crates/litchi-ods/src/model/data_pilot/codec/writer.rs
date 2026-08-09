//! ODF data-pilot XML writer.

use crate::model::database_range::{write_database_source, write_filter};
use litchi_core::{Result, xml::escape_xml};

use super::super::{
    MAX_DATA_PILOT_TABLES, invalid_message,
    model::{Field, GrandTotal, GroupBoundary, Groups, Level, Source, Table},
};

pub(crate) fn write_data_pilot_tables(output: &mut String, tables: &[Table]) -> Result<()> {
    if tables.is_empty() {
        return Ok(());
    }
    if tables.len() > MAX_DATA_PILOT_TABLES {
        return Err(invalid_message(
            "data-pilot table count exceeds safety limit",
        ));
    }
    output.push_str("<table:data-pilot-tables>");
    for table in tables {
        write_table(output, table)?;
    }
    output.push_str("</table:data-pilot-tables>");
    Ok(())
}

pub(crate) fn write_data_pilot_table_fragment(table: &Table) -> Result<String> {
    let mut output = String::new();
    write_table(&mut output, table)?;
    Ok(output)
}

fn write_table(out: &mut String, table: &Table) -> Result<()> {
    table.validate()?;
    out.push_str(
        "<table:data-pilot-table xmlns:table=\"urn:oasis:names:tc:opendocument:xmlns:table:1.0\"",
    );
    attr(out, "table:name", Some(&table.name));
    attr(
        out,
        "table:application-data",
        table.application_data.as_deref(),
    );
    attr(
        out,
        "table:grand-total",
        table.grand_total.map(GrandTotal::as_str),
    );
    bool_attr(out, "table:ignore-empty-rows", table.ignore_empty_rows);
    bool_attr(out, "table:identify-categories", table.identify_categories);
    attr(
        out,
        "table:target-range-address",
        Some(&table.target_range_address),
    );
    attr(out, "table:buttons", table.buttons.as_deref());
    bool_attr(out, "table:show-filter-button", table.show_filter_button);
    bool_attr(
        out,
        "table:drill-down-on-double-click",
        table.drill_down_on_double_click,
    );
    out.push('>');
    for total in &table.grand_totals {
        out.push_str("<table-ext:data-pilot-grand-total xmlns:table-ext=\"urn:org:documentfoundation:names:experimental:office:xmlns:table:1.0\"");
        bool_attr(out, "table:display", Some(total.display));
        attr(out, "table:orientation", Some(total.orientation.as_str()));
        attr(out, "table-ext:display-name", total.display_name.as_deref());
        out.push_str("/>");
    }
    if let Some(source) = &table.source {
        write_source(out, source);
    }
    for field in &table.fields {
        write_field(out, field)?;
    }
    out.push_str("</table:data-pilot-table>");
    Ok(())
}

fn write_source(out: &mut String, source: &Source) {
    match source {
        Source::Database(source) => write_database_source(out, source),
        Source::Service {
            name,
            source_name,
            object_name,
            user_name,
            password,
        } => {
            out.push_str("<table:source-service");
            attr(out, "table:name", Some(name));
            attr(out, "table:source-name", Some(source_name));
            attr(out, "table:object-name", Some(object_name));
            attr(out, "table:user-name", user_name.as_deref());
            attr(out, "table:password", password.as_deref());
            out.push_str("/>");
        },
        Source::CellRange {
            name,
            cell_range_address,
            filter,
        } => {
            out.push_str("<table:source-cell-range");
            attr(out, "table:name", name.as_deref());
            attr(out, "table:cell-range-address", Some(cell_range_address));
            if let Some(filter) = filter {
                out.push('>');
                write_filter(out, filter);
                out.push_str("</table:source-cell-range>");
            } else {
                out.push_str("/>");
            }
        },
    }
}

fn write_field(out: &mut String, field: &Field) -> Result<()> {
    field.validate()?;
    out.push_str("<table:data-pilot-field");
    attr(
        out,
        "table:source-field-name",
        Some(&field.source_field_name),
    );
    attr(out, "table:orientation", Some(field.orientation.as_str()));
    attr(out, "table:selected-page", field.selected_page.as_deref());
    attr(
        out,
        "table:is-data-layout-field",
        field.is_data_layout_field.as_deref(),
    );
    attr(out, "table:function", field.function.as_deref());
    i64_attr(out, "table:used-hierarchy", field.used_hierarchy);
    if field.level.is_none() && field.reference.is_none() && field.groups.is_none() {
        out.push_str("/>");
        return Ok(());
    }
    out.push('>');
    if let Some(level) = &field.level {
        write_level(out, level);
    }
    if let Some(reference) = &field.reference {
        out.push_str("<table:data-pilot-field-reference");
        attr(out, "table:field-name", Some(&reference.field_name));
        attr(
            out,
            "table:member-type",
            Some(reference.member_type.as_str()),
        );
        attr(out, "table:member-name", reference.member_name.as_deref());
        attr(out, "table:type", Some(reference.reference_type.as_str()));
        out.push_str("/>");
    }
    if let Some(groups) = &field.groups {
        write_groups(out, groups);
    }
    out.push_str("</table:data-pilot-field>");
    Ok(())
}

fn write_level(out: &mut String, level: &Level) {
    out.push_str("<table:data-pilot-level");
    bool_attr(out, "table:show-empty", level.show_empty);
    if level.repeat_item_labels.is_some() {
        out.push_str(" xmlns:calcext=\"urn:org:documentfoundation:names:experimental:calc:xmlns:calcext:1.0\"");
        bool_attr(out, "calcext:repeat-item-labels", level.repeat_item_labels);
    }
    if level.subtotals.is_empty()
        && level.members.is_empty()
        && level.display.is_none()
        && level.sort.is_none()
        && level.layout.is_none()
    {
        out.push_str("/>");
        return;
    }
    out.push('>');
    if !level.subtotals.is_empty() {
        out.push_str("<table:data-pilot-subtotals>");
        for function in &level.subtotals {
            out.push_str("<table:data-pilot-subtotal");
            attr(out, "table:function", Some(function));
            out.push_str("/>");
        }
        out.push_str("</table:data-pilot-subtotals>");
    }
    if !level.members.is_empty() {
        out.push_str("<table:data-pilot-members>");
        for member in &level.members {
            out.push_str("<table:data-pilot-member");
            attr(out, "table:name", Some(&member.name));
            bool_attr(out, "table:display", member.display);
            bool_attr(out, "table:show-details", member.show_details);
            out.push_str("/>");
        }
        out.push_str("</table:data-pilot-members>");
    }
    if let Some(info) = &level.display {
        out.push_str("<table:data-pilot-display-info");
        bool_attr(out, "table:enabled", Some(info.enabled));
        attr(out, "table:data-field", Some(&info.data_field));
        u64_attr(out, "table:member-count", info.member_count);
        attr(out, "table:display-member-mode", Some(info.mode.as_str()));
        out.push_str("/>");
    }
    if let Some(info) = &level.sort {
        out.push_str("<table:data-pilot-sort-info");
        attr(out, "table:sort-mode", Some(info.mode.as_str()));
        attr(out, "table:data-field", info.data_field.as_deref());
        attr(out, "table:order", Some(info.order.as_str()));
        out.push_str("/>");
    }
    if let Some(info) = &level.layout {
        out.push_str("<table:data-pilot-layout-info");
        attr(out, "table:layout-mode", Some(info.mode.as_str()));
        bool_attr(out, "table:add-empty-lines", Some(info.add_empty_lines));
        out.push_str("/>");
    }
    out.push_str("</table:data-pilot-level>");
}

fn write_groups(out: &mut String, groups: &Groups) {
    out.push_str("<table:data-pilot-groups");
    attr(
        out,
        "table:source-field-name",
        Some(&groups.source_field_name),
    );
    write_boundary(out, "start", &groups.start);
    write_boundary(out, "end", &groups.end);
    attr(out, "table:step", Some(&groups.step.to_string()));
    attr(out, "table:grouped-by", Some(groups.grouped_by.as_str()));
    out.push('>');
    for group in &groups.groups {
        out.push_str("<table:data-pilot-group");
        attr(out, "table:name", Some(&group.name));
        out.push('>');
        for member in &group.members {
            out.push_str("<table:data-pilot-group-member");
            attr(out, "table:name", Some(member));
            out.push_str("/>");
        }
        out.push_str("</table:data-pilot-group>");
    }
    out.push_str("</table:data-pilot-groups>");
}

fn write_boundary(out: &mut String, suffix: &str, boundary: &GroupBoundary) {
    match boundary {
        GroupBoundary::AutomaticNumber => {
            attr(out, &format!("table:{suffix}"), Some("auto"));
        },
        GroupBoundary::AutomaticDate => {
            attr(out, &format!("table:date-{suffix}"), Some("auto"));
        },
        GroupBoundary::Number(value) => {
            attr(out, &format!("table:{suffix}"), Some(&value.to_string()));
        },
        GroupBoundary::Date(value) => attr(out, &format!("table:date-{suffix}"), Some(value)),
    }
}

fn attr(out: &mut String, name: &str, value: Option<&str>) {
    if let Some(value) = value {
        out.push(' ');
        out.push_str(name);
        out.push_str("=\"");
        out.push_str(&escape_xml(value));
        out.push('"');
    }
}

fn bool_attr(out: &mut String, name: &str, value: Option<bool>) {
    attr(
        out,
        name,
        value.map(|value| if value { "true" } else { "false" }),
    );
}

fn i64_attr(out: &mut String, name: &str, value: Option<i64>) {
    if let Some(value) = value {
        attr(out, name, Some(&value.to_string()));
    }
}

fn u64_attr(out: &mut String, name: &str, value: u64) {
    attr(out, name, Some(&value.to_string()));
}
