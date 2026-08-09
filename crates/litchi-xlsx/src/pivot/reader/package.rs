//! OPC package traversal and relationship validation for pivot parts.

use std::collections::HashMap;

use super::super::PivotTable;
use super::super::cache::Definition;
use super::codec::{
    parse_pivot_table_definition_with_cache, read_pivot_cache_definition,
    validate_pivot_cache_records,
};
use crate::raw::parse_catalog;
use litchi_core::sheet::Result as SheetResult;
use litchi_opc::constants::{content_type as ct, relationship_type as rt};
use litchi_opc::{OpcPackage, PackURI};

pub fn read_pivot_tables(package: &OpcPackage) -> SheetResult<Vec<PivotTable>> {
    let workbook_part = package.main_document_part()?;
    let workbook_xml = workbook_part.blob();

    let workbook = parse_catalog(workbook_xml)?;

    let workbook_rels = workbook_part.rels();
    let mut worksheet_uris = HashMap::with_capacity(workbook.sheets.len());
    let mut worksheet_names_by_uri = HashMap::with_capacity(workbook.sheets.len());
    for worksheet in &workbook.sheets {
        let rel = workbook_rels
            .get(&worksheet.relationship_id)
            .ok_or_else(|| {
                format!(
                    "worksheet '{}' references missing relationship '{}'",
                    worksheet.name, worksheet.relationship_id
                )
            })?;
        // Non-worksheet sheets (chartsheets, dialog sheets, macro sheets)
        // cannot host pivot tables and are skipped rather than rejected.
        if !matches!(rel.reltype(), rt::WORKSHEET | rt::STRICT_WORKSHEET) {
            continue;
        }
        if rel.is_external() {
            return Err(format!(
                "worksheet '{}' relationship cannot be external",
                worksheet.name
            )
            .into());
        }
        let worksheet_uri = rel.target_partname()?;
        let worksheet_part = package.get_part(&worksheet_uri)?;
        require_content_type(
            &worksheet_uri,
            worksheet_part.content_type(),
            ct::SML_WORKSHEET,
        )?;
        if worksheet_names_by_uri
            .insert(worksheet_uri.clone(), worksheet.name.clone())
            .is_some()
        {
            return Err(format!("multiple workbook sheets target part '{worksheet_uri}'").into());
        }
        worksheet_uris.insert(worksheet.relationship_id.clone(), worksheet_uri);
    }
    let (pivot_caches, pivot_cache_ids_by_uri) = resolve_workbook_pivot_caches(
        package,
        workbook_rels,
        &workbook.pivot_caches,
        &worksheet_names_by_uri,
    )?;
    if workbook.sheets.is_empty() {
        return Ok(Vec::new());
    }
    let mut tables = Vec::new();

    for ws_info in workbook.sheets {
        // Sheets skipped above (chartsheets and other non-worksheet kinds)
        // have no resolved worksheet URI and are ignored here as well.
        let Some(sheet_uri) = worksheet_uris.get(&ws_info.relationship_id) else {
            continue;
        };
        let sheet_uri = sheet_uri.clone();
        let sheet_part = package.get_part(&sheet_uri)?;
        let sheet_rels = sheet_part.rels();

        for rel in sheet_rels.iter() {
            if !matches!(rel.reltype(), rt::PIVOT_TABLE | rt::STRICT_PIVOT_TABLE) {
                continue;
            }
            if rel.is_external() {
                return Err(format!(
                    "worksheet '{}' pivot-table relationship cannot be external",
                    ws_info.name
                )
                .into());
            }

            let table_uri = rel.target_partname()?;
            let table_part = package.get_part(&table_uri)?;
            require_content_type(&table_uri, table_part.content_type(), ct::SML_PIVOT_TABLE)?;
            let cache_uri = resolve_pivot_table_cache_uri(table_part)?;
            let expected_cache_id = pivot_cache_ids_by_uri.get(&cache_uri).ok_or_else(|| {
                format!(
                    "pivot-table part '{table_uri}' references cache definition '{cache_uri}' that is not listed by the workbook"
                )
            })?;
            let cache = pivot_caches
                .get(expected_cache_id)
                .ok_or("resolved pivot cache is missing")?;
            let bytes = litchi_ooxml_common::mce::process_part(table_part)?;
            let xml = std::str::from_utf8(bytes.as_ref())?;

            let mut table = parse_pivot_table_definition_with_cache(
                xml,
                &ws_info.name,
                Some(&cache.definition.cache_fields),
            )?
            .ok_or_else(|| {
                format!("pivot-table part '{table_uri}' has no pivotTableDefinition root")
            })?;
            if table.cache_id != *expected_cache_id {
                return Err(format!(
                    "pivot-table part '{table_uri}' declares cache ID {}, but its relationship targets workbook cache ID {expected_cache_id}",
                    table.cache_id
                )
                .into());
            }
            table
                .source_sheet
                .clone_from(&cache.definition.source_worksheet);
            table.source_ref.clone_from(&cache.definition.source_ref);
            tables.push(table);
        }
    }

    Ok(tables)
}

struct ResolvedPivotCache {
    definition: Definition,
}

fn resolve_workbook_pivot_caches(
    package: &OpcPackage,
    workbook_rels: &litchi_opc::Relationships,
    cache_references: &[crate::raw::PivotCache],
    worksheet_names_by_uri: &HashMap<PackURI, String>,
) -> SheetResult<(HashMap<u32, ResolvedPivotCache>, HashMap<PackURI, u32>)> {
    let mut caches = HashMap::with_capacity(cache_references.len());
    let mut ids_by_uri = HashMap::with_capacity(cache_references.len());
    for cache_reference in cache_references {
        let rel = workbook_rels
            .get(&cache_reference.relationship_id)
            .ok_or_else(|| {
                format!(
                    "workbook pivot cache {} references missing relationship '{}'",
                    cache_reference.cache_id, cache_reference.relationship_id
                )
            })?;
        if !matches!(
            rel.reltype(),
            rt::PIVOT_CACHE_DEFINITION | rt::STRICT_PIVOT_CACHE_DEFINITION
        ) {
            return Err(format!(
                "workbook pivot cache {} relationship has invalid type '{}'",
                cache_reference.cache_id,
                rel.reltype()
            )
            .into());
        }
        if rel.is_external() {
            return Err(format!(
                "workbook pivot cache {} relationship cannot be external",
                cache_reference.cache_id
            )
            .into());
        }
        let cache_uri = rel.target_partname()?;
        let cache_part = package.get_part(&cache_uri)?;
        require_content_type(
            &cache_uri,
            cache_part.content_type(),
            ct::SML_PIVOT_CACHE_DEFINITION,
        )?;
        let bytes = litchi_ooxml_common::mce::process_part(cache_part)?;
        let xml = std::str::from_utf8(bytes.as_ref())?;
        let mut definition = read_pivot_cache_definition(xml)?.ok_or_else(|| {
            format!("pivot-cache part '{cache_uri}' has no pivotCacheDefinition root")
        })?;
        validate_pivot_cache_relationships(
            package,
            cache_part,
            &cache_uri,
            &mut definition,
            worksheet_names_by_uri,
        )?;
        if ids_by_uri
            .insert(cache_uri.clone(), cache_reference.cache_id)
            .is_some()
        {
            return Err(
                format!("multiple workbook pivot cache IDs target part '{cache_uri}'").into(),
            );
        }
        caches.insert(cache_reference.cache_id, ResolvedPivotCache { definition });
    }
    Ok((caches, ids_by_uri))
}

fn validate_pivot_cache_relationships(
    package: &OpcPackage,
    cache_part: &dyn litchi_opc::Part,
    cache_uri: &PackURI,
    definition: &mut Definition,
    worksheet_names_by_uri: &HashMap<PackURI, String>,
) -> SheetResult<()> {
    if let Some(relationship_id) = definition.id.as_deref() {
        let rel = cache_part.rels().get(relationship_id).ok_or_else(|| {
            format!(
                "pivot-cache part '{cache_uri}' references missing records relationship '{relationship_id}'"
            )
        })?;
        if !matches!(
            rel.reltype(),
            rt::PIVOT_CACHE_RECORDS | rt::STRICT_PIVOT_CACHE_RECORDS
        ) {
            return Err(format!(
                "pivot-cache part '{cache_uri}' records relationship has invalid type '{}'",
                rel.reltype()
            )
            .into());
        }
        if rel.is_external() {
            return Err(format!(
                "pivot-cache part '{cache_uri}' records relationship cannot be external"
            )
            .into());
        }
        let records_uri = rel.target_partname()?;
        let records_part = package.get_part(&records_uri)?;
        require_content_type(
            &records_uri,
            records_part.content_type(),
            ct::SML_PIVOT_CACHE_RECORDS,
        )?;
        let bytes = litchi_ooxml_common::mce::process_part(records_part)?;
        let records_xml = std::str::from_utf8(bytes.as_ref())?;
        validate_pivot_cache_records(
            records_xml,
            &definition.cache_fields,
            definition.record_count,
        )?;
    }

    if let Some(relationship_id) = definition.source_relationship_id.as_deref() {
        let rel = cache_part.rels().get(relationship_id).ok_or_else(|| {
            format!(
                "pivot-cache part '{cache_uri}' references missing source worksheet relationship '{relationship_id}'"
            )
        })?;
        if !matches!(rel.reltype(), rt::WORKSHEET | rt::STRICT_WORKSHEET) {
            return Err(format!(
                "pivot-cache part '{cache_uri}' source worksheet relationship has invalid type '{}'",
                rel.reltype()
            )
            .into());
        }
        if rel.is_external() {
            return Err(format!(
                "pivot-cache part '{cache_uri}' source worksheet relationship cannot be external"
            )
            .into());
        }
        let worksheet_uri = rel.target_partname()?;
        let worksheet_part = package.get_part(&worksheet_uri)?;
        require_content_type(
            &worksheet_uri,
            worksheet_part.content_type(),
            ct::SML_WORKSHEET,
        )?;
        let workbook_name = worksheet_names_by_uri.get(&worksheet_uri).ok_or_else(|| {
            format!(
                "pivot-cache part '{cache_uri}' source worksheet '{worksheet_uri}' is not listed by the workbook"
            )
        })?;
        if let Some(source_name) = definition.source_worksheet.as_deref() {
            if source_name != workbook_name {
                return Err(format!(
                    "pivot-cache part '{cache_uri}' names source worksheet '{source_name}', but its relationship targets '{workbook_name}'"
                )
                .into());
            }
        } else {
            definition.source_worksheet = Some(workbook_name.clone());
        }
    }
    Ok(())
}

fn resolve_pivot_table_cache_uri(table_part: &dyn litchi_opc::Part) -> SheetResult<PackURI> {
    let mut matching = table_part.rels().iter().filter(|rel| {
        matches!(
            rel.reltype(),
            rt::PIVOT_CACHE_DEFINITION | rt::STRICT_PIVOT_CACHE_DEFINITION
        )
    });
    let rel = matching
        .next()
        .ok_or("pivot-table part is missing its cache-definition relationship")?;
    if matching.next().is_some() {
        return Err("pivot-table part has multiple cache-definition relationships".into());
    }
    if rel.is_external() {
        return Err("pivot-table cache-definition relationship cannot be external".into());
    }
    rel.target_partname().map_err(Into::into)
}

fn require_content_type(uri: &PackURI, actual: &str, expected: &str) -> SheetResult<()> {
    if actual != expected {
        return Err(
            format!("part '{uri}' has content type '{actual}', expected '{expected}'").into(),
        );
    }
    Ok(())
}
