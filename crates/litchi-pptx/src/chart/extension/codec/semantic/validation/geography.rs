//! Geography and map data validation concerns for the ChartEx graph.

use super::*;
use std::collections::HashSet;

pub(super) fn parse_geography(node: &MiniNode) -> Result<Geography> {
    let allowed = &[
        ("", "projectionType"),
        ("", "viewedRegionType"),
        ("", "cultureLanguage"),
        ("", "cultureRegion"),
        ("", "attribution"),
    ];
    reject_unknown(&node.attributes, allowed, "geography")?;
    if !node.text.trim().is_empty() {
        return invalid("unexpected text in  geography");
    }
    let projection = optional(&node.attributes, "", "projectionType")
        .map(|value| match value {
            "mercator" => Ok(GeoProjection::Mercator),
            "miller" => Ok(GeoProjection::Miller),
            "robinson" => Ok(GeoProjection::Robinson),
            "albers" => Ok(GeoProjection::Albers),
            _ => invalid("invalid  geography projectionType"),
        })
        .transpose()?;
    let viewed_region = optional(&node.attributes, "", "viewedRegionType")
        .map(|value| match value {
            "dataOnly" => Ok(GeoMappingLevel::DataOnly),
            "postalCode" => Ok(GeoMappingLevel::PostalCode),
            "county" => Ok(GeoMappingLevel::County),
            "state" => Ok(GeoMappingLevel::State),
            "countryRegion" => Ok(GeoMappingLevel::CountryRegion),
            "countryRegionList" => Ok(GeoMappingLevel::CountryRegionList),
            "world" => Ok(GeoMappingLevel::World),
            _ => invalid("invalid  geography viewedRegionType"),
        })
        .transpose()?;
    let culture_language = bounded_required(node, "cultureLanguage", MAX_CULTURE_NAME_LEN)?;
    let culture_region = bounded_required(node, "cultureRegion", MAX_CULTURE_NAME_LEN)?;
    let attribution = bounded_required(node, "attribution", MAX_ATTRIBUTION_LEN)?;
    let mut cache = None;
    for child in &node.children {
        if child.namespace != CX || child.name != "geoCache" || cache.is_some() {
            return invalid("invalid or duplicate  geography child");
        }
        cache = Some(parse_geo_cache(child)?);
    }
    let has_cache = cache.is_some();
    Ok(Geography {
        projection,
        viewed_region,
        culture_language,
        culture_region,
        attribution,
        has_cache,
        cache,
    })
}

pub(super) fn parse_geo_cache(node: &MiniNode) -> Result<GeoCache> {
    reject_unknown(&node.attributes, &[("", "provider")], "geoCache")?;
    if !node.text.trim().is_empty() {
        return invalid("unexpected text in  geoCache");
    }
    let provider = geo_required_string(node, "provider", MAX_GEO_STRING_LEN)?;
    if node.children.is_empty() {
        return invalid(" geoCache requires binary or clear content");
    }
    if node.children.len() > MAX_GEO_CACHE_ENTRIES {
        return limit(" geography cache entries");
    }
    let mut entries = Vec::with_capacity(node.children.len());
    for child in &node.children {
        if child.namespace != CX {
            return invalid("foreign direct child in  geoCache");
        }
        entries.push(match child.name.as_str() {
            "binary" => {
                reject_unknown(&child.attributes, &[], "geography binary")?;
                if !child.children.is_empty() {
                    return invalid(" geography binary contains elements");
                }
                let (encoded_characters, decoded_bytes) = validate_geo_base64(&child.text)?;
                GeoCacheEntry::Binary {
                    encoded_characters,
                    decoded_bytes,
                }
            },
            "clear" => GeoCacheEntry::Clear(parse_geo_clear(child)?),
            _ => return invalid("invalid direct child in  geoCache"),
        });
    }
    Ok(GeoCache { provider, entries })
}

pub(super) fn parse_geo_clear(node: &MiniNode) -> Result<GeoClear> {
    reject_geo_container(node, "geography clear cache")?;
    let mut result = GeoClear::default();
    let mut rank = 0u8;
    let mut seen = HashSet::new();
    for child in &node.children {
        let current = geo_ordered_child(
            child,
            &[
                "geoLocationQueryResults",
                "geoDataEntityQueryResults",
                "geoDataPointToEntityQueryResults",
                "geoChildEntitiesQueryResults",
                "geoParentEntitiesQueryResults",
            ],
        )?;
        if current < rank || !seen.insert(child.name.as_str()) {
            return invalid(" clear geography cache children are out of order or duplicated");
        }
        rank = current;
        match child.name.as_str() {
            "geoLocationQueryResults" => {
                result.location_query_results = Some(parse_geo_collection(
                    child,
                    "geoLocationQueryResult",
                    parse_geo_location_query_result,
                )?)
            },
            "geoDataEntityQueryResults" => {
                result.data_entity_query_results = Some(parse_geo_collection(
                    child,
                    "geoDataEntityQueryResult",
                    parse_geo_data_entity_query_result,
                )?)
            },
            "geoDataPointToEntityQueryResults" => {
                result.data_point_to_entity_query_results = Some(parse_geo_collection(
                    child,
                    "geoDataPointToEntityQueryResult",
                    parse_geo_data_point_to_entity_query_result,
                )?)
            },
            "geoChildEntitiesQueryResults" => {
                result.child_entities_query_results = Some(parse_geo_collection(
                    child,
                    "geoChildEntitiesQueryResult",
                    parse_geo_child_entities_query_result,
                )?)
            },
            "geoParentEntitiesQueryResults" => {
                result.parent_entities_query_results = Some(parse_geo_collection(
                    child,
                    "geoParentEntitiesQueryResult",
                    parse_geo_parent_entities_query_result,
                )?)
            },
            _ => unreachable!(),
        }
    }
    Ok(result)
}

pub(super) fn parse_geo_collection<T>(
    node: &MiniNode,
    item_name: &str,
    parser: fn(&MiniNode) -> Result<T>,
) -> Result<Vec<T>> {
    reject_geo_container(node, &node.name)?;
    if node.children.len() > MAX_GEO_RESULTS {
        return limit(" geography query results");
    }
    node.children
        .iter()
        .map(|child| {
            if child.namespace != CX || child.name != item_name {
                return invalid(format!("invalid direct child in  {}", node.name));
            }
            parser(child)
        })
        .collect()
}

pub(super) fn parse_geo_location_query_result(node: &MiniNode) -> Result<GeoLocationQueryResult> {
    reject_geo_container(node, "geoLocationQueryResult")?;
    let mut result = GeoLocationQueryResult::default();
    for child in geo_unique_ordered(node, &["geoLocationQuery", "geoLocations"])? {
        if child.name == "geoLocationQuery" {
            result.query = Some(parse_geo_location_query(child)?);
        } else {
            reject_geo_container(child, "geoLocations")?;
            if child.children.len() > 1 {
                return invalid("geoLocations permits at most one geoLocation");
            }
            result.location = child
                .children
                .first()
                .map(|value| {
                    if value.namespace != CX || value.name != "geoLocation" {
                        return invalid("invalid direct child in geoLocations");
                    }
                    parse_geo_location(value)
                })
                .transpose()?;
        }
    }
    Ok(result)
}

pub(super) fn parse_geo_location_query(node: &MiniNode) -> Result<GeoLocationQuery> {
    let allowed = &[
        ("", "countryRegion"),
        ("", "adminDistrict1"),
        ("", "adminDistrict2"),
        ("", "postalCode"),
        ("", "entityType"),
    ];
    reject_unknown(&node.attributes, allowed, "geoLocationQuery")?;
    require_empty_content(node, "geoLocationQuery")?;
    Ok(GeoLocationQuery {
        country_region: geo_optional_string(node, "countryRegion", MAX_GEO_STRING_LEN)?,
        admin_district1: geo_optional_string(node, "adminDistrict1", MAX_GEO_STRING_LEN)?,
        admin_district2: geo_optional_string(node, "adminDistrict2", MAX_GEO_STRING_LEN)?,
        postal_code: geo_optional_string(node, "postalCode", MAX_GEO_STRING_LEN)?,
        entity_type: parse_geo_entity_type(required(&node.attributes, "", "entityType")?)?,
    })
}

pub(super) fn parse_geo_location(node: &MiniNode) -> Result<GeoLocation> {
    reject_unknown(
        &node.attributes,
        &[
            ("", "latitude"),
            ("", "longitude"),
            ("", "entityName"),
            ("", "entityType"),
        ],
        "geoLocation",
    )?;
    if !node.text.trim().is_empty() {
        return invalid("unexpected text in  geoLocation");
    }
    let address = match node.children.as_slice() {
        [] => None,
        [child] if child.namespace == CX && child.name == "address" => {
            Some(parse_geo_address(child)?)
        },
        _ => return invalid("geoLocation permits at most one ordered address"),
    };
    Ok(GeoLocation {
        latitude: geo_optional_double(node, "latitude")?,
        longitude: geo_optional_double(node, "longitude")?,
        entity_name: geo_required_string(node, "entityName", MAX_GEO_STRING_LEN)?,
        entity_type: parse_geo_entity_type(required(&node.attributes, "", "entityType")?)?,
        address,
    })
}

pub(super) fn parse_geo_address(node: &MiniNode) -> Result<GeoAddress> {
    let allowed = &[
        ("", "address1"),
        ("", "countryRegion"),
        ("", "adminDistrict1"),
        ("", "adminDistrict2"),
        ("", "postalCode"),
        ("", "locality"),
        ("", "isoCountryCode"),
    ];
    reject_unknown(&node.attributes, allowed, "geography address")?;
    require_empty_content(node, "geography address")?;
    Ok(GeoAddress {
        address1: geo_optional_string(node, "address1", MAX_GEO_STRING_LEN)?,
        country_region: geo_optional_string(node, "countryRegion", MAX_GEO_STRING_LEN)?,
        admin_district1: geo_optional_string(node, "adminDistrict1", MAX_GEO_STRING_LEN)?,
        admin_district2: geo_optional_string(node, "adminDistrict2", MAX_GEO_STRING_LEN)?,
        postal_code: geo_optional_string(node, "postalCode", MAX_GEO_STRING_LEN)?,
        locality: geo_optional_string(node, "locality", MAX_GEO_STRING_LEN)?,
        iso_country_code: geo_optional_string(node, "isoCountryCode", MAX_GEO_STRING_LEN)?,
    })
}

pub(super) fn parse_geo_data_entity_query_result(
    node: &MiniNode,
) -> Result<GeoDataEntityQueryResult> {
    reject_geo_container(node, "geoDataEntityQueryResult")?;
    let mut result = GeoDataEntityQueryResult::default();
    for child in geo_unique_ordered(node, &["geoDataEntityQuery", "geoData"])? {
        if child.name == "geoDataEntityQuery" {
            result.query = Some(parse_geo_data_entity_query(child)?);
        } else {
            result.data = Some(parse_geo_data(child)?);
        }
    }
    Ok(result)
}

pub(super) fn parse_geo_data_entity_query(node: &MiniNode) -> Result<GeoDataEntityQuery> {
    reject_unknown(
        &node.attributes,
        &[("", "entityType"), ("", "entityId")],
        "geoDataEntityQuery",
    )?;
    require_empty_content(node, "geoDataEntityQuery")?;
    Ok(GeoDataEntityQuery {
        entity_type: parse_geo_entity_type(required(&node.attributes, "", "entityType")?)?,
        entity_id: geo_required_string(node, "entityId", MAX_GEO_STRING_LEN)?,
    })
}

pub(super) fn parse_geo_data(node: &MiniNode) -> Result<GeoData> {
    reject_unknown(
        &node.attributes,
        &[
            ("", "entityName"),
            ("", "entityId"),
            ("", "east"),
            ("", "west"),
            ("", "north"),
            ("", "south"),
        ],
        "geoData",
    )?;
    if !node.text.trim().is_empty() {
        return invalid("unexpected text in  geoData");
    }
    let mut polygons = None;
    let mut copyrights = None;
    for child in geo_unique_ordered(node, &["geoPolygons", "copyrights"])? {
        if child.name == "geoPolygons" {
            polygons = Some(parse_geo_collection(
                child,
                "geoPolygon",
                parse_geo_polygon,
            )?);
        } else {
            copyrights = Some(parse_geo_copyrights(child)?);
        }
    }
    Ok(GeoData {
        entity_name: geo_required_string(node, "entityName", MAX_GEO_STRING_LEN)?,
        entity_id: geo_required_string(node, "entityId", MAX_GEO_STRING_LEN)?,
        east: geo_required_double(node, "east")?,
        west: geo_required_double(node, "west")?,
        north: geo_required_double(node, "north")?,
        south: geo_required_double(node, "south")?,
        polygons,
        copyrights,
    })
}

pub(super) fn parse_geo_polygon(node: &MiniNode) -> Result<GeoPolygon> {
    reject_unknown(
        &node.attributes,
        &[("", "polygonId"), ("", "numPoints"), ("", "pcaRings")],
        "geoPolygon",
    )?;
    require_empty_content(node, "geoPolygon")?;
    let num_points = geo_required_string(node, "numPoints", 128)?;
    validate_xsd_integer(&num_points, "geoPolygon numPoints")?;
    Ok(GeoPolygon {
        polygon_id: geo_required_string(node, "polygonId", MAX_GEO_STRING_LEN)?,
        num_points,
        pca_rings: geo_required_string(node, "pcaRings", MAX_GEO_POLYGON_DATA_LEN)?,
    })
}

pub(super) fn parse_geo_copyrights(node: &MiniNode) -> Result<Vec<String>> {
    reject_geo_container(node, "copyrights")?;
    if node.children.len() > MAX_GEO_RESULTS {
        return limit(" geography copyrights");
    }
    node.children
        .iter()
        .map(|child| {
            if child.namespace != CX
                || child.name != "copyright"
                || !child.attributes.is_empty()
                || !child.children.is_empty()
            {
                return invalid("invalid direct child in  copyrights");
            }
            if child.text.len() > MAX_GEO_STRING_LEN {
                return limit(" geography copyright");
            }
            Ok(child.text.clone())
        })
        .collect()
}

pub(super) fn parse_geo_data_point_to_entity_query_result(
    node: &MiniNode,
) -> Result<GeoDataPointToEntityQueryResult> {
    reject_geo_container(node, "geoDataPointToEntityQueryResult")?;
    let mut result = GeoDataPointToEntityQueryResult::default();
    for child in geo_unique_ordered(node, &["geoDataPointQuery", "geoDataPointToEntityQuery"])? {
        if child.name == "geoDataPointQuery" {
            result.point_query = Some(parse_geo_data_point_query(child)?);
        } else {
            result.entity_query = Some(parse_geo_data_point_to_entity_query(child)?);
        }
    }
    Ok(result)
}

pub(super) fn parse_geo_data_point_query(node: &MiniNode) -> Result<GeoDataPointQuery> {
    reject_unknown(
        &node.attributes,
        &[("", "entityType"), ("", "latitude"), ("", "longitude")],
        "geoDataPointQuery",
    )?;
    require_empty_content(node, "geoDataPointQuery")?;
    Ok(GeoDataPointQuery {
        entity_type: parse_geo_entity_type(required(&node.attributes, "", "entityType")?)?,
        latitude: geo_required_double(node, "latitude")?,
        longitude: geo_required_double(node, "longitude")?,
    })
}

pub(super) fn parse_geo_data_point_to_entity_query(
    node: &MiniNode,
) -> Result<GeoDataPointToEntityQuery> {
    reject_unknown(
        &node.attributes,
        &[("", "entityType"), ("", "entityId")],
        "geoDataPointToEntityQuery",
    )?;
    require_empty_content(node, "geoDataPointToEntityQuery")?;
    Ok(GeoDataPointToEntityQuery {
        entity_type: parse_geo_entity_type(required(&node.attributes, "", "entityType")?)?,
        entity_id: geo_required_string(node, "entityId", MAX_GEO_STRING_LEN)?,
    })
}

pub(super) fn parse_geo_child_entities_query_result(
    node: &MiniNode,
) -> Result<GeoChildEntitiesQueryResult> {
    reject_geo_container(node, "geoChildEntitiesQueryResult")?;
    let mut result = GeoChildEntitiesQueryResult::default();
    for child in geo_unique_ordered(node, &["geoChildEntitiesQuery", "geoChildEntities"])? {
        if child.name == "geoChildEntitiesQuery" {
            result.query = Some(parse_geo_child_entities_query(child)?);
        } else {
            result.children = Some(parse_geo_collection(
                child,
                "geoHierarchyEntity",
                parse_geo_hierarchy_entity,
            )?);
        }
    }
    Ok(result)
}

pub(super) fn parse_geo_child_entities_query(node: &MiniNode) -> Result<GeoChildEntitiesQuery> {
    reject_unknown(
        &node.attributes,
        &[("", "entityId")],
        "geoChildEntitiesQuery",
    )?;
    if !node.text.trim().is_empty() {
        return invalid("unexpected text in geoChildEntitiesQuery");
    }
    let child_types = match node.children.as_slice() {
        [] => None,
        [child] if child.namespace == CX && child.name == "geoChildTypes" => {
            reject_geo_container(child, "geoChildTypes")?;
            if child.children.len() > MAX_GEO_RESULTS {
                return limit(" geography child types");
            }
            Some(
                child
                    .children
                    .iter()
                    .map(|value| {
                        if value.namespace != CX
                            || value.name != "entityType"
                            || !value.attributes.is_empty()
                            || !value.children.is_empty()
                        {
                            return invalid("invalid direct child in geoChildTypes");
                        }
                        parse_geo_entity_type(value.text.trim())
                    })
                    .collect::<Result<Vec<_>>>()?,
            )
        },
        _ => return invalid("geoChildEntitiesQuery permits at most one geoChildTypes"),
    };
    Ok(GeoChildEntitiesQuery {
        entity_id: geo_required_string(node, "entityId", MAX_GEO_STRING_LEN)?,
        child_types,
    })
}

pub(super) fn parse_geo_hierarchy_entity(node: &MiniNode) -> Result<GeoHierarchyEntity> {
    reject_unknown(
        &node.attributes,
        &[("", "entityName"), ("", "entityId"), ("", "entityType")],
        "geoHierarchyEntity",
    )?;
    require_empty_content(node, "geoHierarchyEntity")?;
    Ok(GeoHierarchyEntity {
        entity_name: geo_required_string(node, "entityName", MAX_GEO_STRING_LEN)?,
        entity_id: geo_required_string(node, "entityId", MAX_GEO_STRING_LEN)?,
        entity_type: parse_geo_entity_type(required(&node.attributes, "", "entityType")?)?,
    })
}

pub(super) fn parse_geo_parent_entities_query_result(
    node: &MiniNode,
) -> Result<GeoParentEntitiesQueryResult> {
    reject_geo_container(node, "geoParentEntitiesQueryResult")?;
    let children = &node.children;
    if children.is_empty()
        || children[0].namespace != CX
        || children[0].name != "geoParentEntitiesQuery"
    {
        return invalid("geoParentEntitiesQueryResult requires geoParentEntitiesQuery first");
    }
    reject_unknown(
        &children[0].attributes,
        &[("", "entityId")],
        "geoParentEntitiesQuery",
    )?;
    require_empty_content(&children[0], "geoParentEntitiesQuery")?;
    let entity_id = geo_required_string(&children[0], "entityId", MAX_GEO_STRING_LEN)?;
    let mut entity = None;
    let mut parent_entity_id = None;
    let mut rank = 0u8;
    for child in children.iter().skip(1) {
        let current = geo_ordered_child(child, &["geoEntity", "geoParentEntity"])?;
        if current < rank {
            return invalid("invalid geoParentEntitiesQueryResult order");
        }
        rank = current;
        if child.name == "geoEntity" {
            if entity.is_some() {
                return invalid("duplicate geoEntity");
            }
            reject_unknown(
                &child.attributes,
                &[("", "entityName"), ("", "entityType")],
                "geoEntity",
            )?;
            require_empty_content(child, "geoEntity")?;
            entity = Some(GeoEntity {
                entity_name: geo_required_string(child, "entityName", MAX_GEO_STRING_LEN)?,
                entity_type: parse_geo_entity_type(required(&child.attributes, "", "entityType")?)?,
            });
        } else {
            if parent_entity_id.is_some() {
                return invalid("duplicate geoParentEntity");
            }
            reject_unknown(&child.attributes, &[("", "entityId")], "geoParentEntity")?;
            require_empty_content(child, "geoParentEntity")?;
            parent_entity_id = Some(geo_required_string(child, "entityId", MAX_GEO_STRING_LEN)?);
        }
    }
    Ok(GeoParentEntitiesQueryResult {
        entity_id,
        entity,
        parent_entity_id,
    })
}

pub(super) fn reject_geo_container(node: &MiniNode, label: &str) -> Result<()> {
    reject_unknown(&node.attributes, &[], label)?;
    if !node.text.trim().is_empty() {
        return invalid(format!("unexpected text in  {label}"));
    }
    Ok(())
}

pub(super) fn geo_ordered_child(child: &MiniNode, names: &[&str]) -> Result<u8> {
    if child.namespace != CX {
        return invalid("foreign child in  geography cache");
    }
    names
        .iter()
        .position(|name| *name == child.name)
        .map(|value| value as u8)
        .ok_or_else(|| invalid_error(format!("invalid geography cache child '{}'", child.name)))
}

pub(super) fn geo_unique_ordered<'a>(
    node: &'a MiniNode,
    names: &[&str],
) -> Result<Vec<&'a MiniNode>> {
    let mut rank = 0u8;
    let mut seen = HashSet::new();
    for child in &node.children {
        let current = geo_ordered_child(child, names)?;
        if current < rank || !seen.insert(child.name.as_str()) {
            return invalid(format!("invalid {} order or cardinality", node.name));
        }
        rank = current;
    }
    Ok(node.children.iter().collect())
}

pub(super) fn parse_geo_entity_type(value: &str) -> Result<GeoEntityType> {
    match value {
        "Address" => Ok(GeoEntityType::Address),
        "AdminDistrict" => Ok(GeoEntityType::AdminDistrict),
        "AdminDistrict2" => Ok(GeoEntityType::AdminDistrict2),
        "AdminDistrict3" => Ok(GeoEntityType::AdminDistrict3),
        "Continent" => Ok(GeoEntityType::Continent),
        "CountryRegion" => Ok(GeoEntityType::CountryRegion),
        "Locality" => Ok(GeoEntityType::Locality),
        "Ocean" => Ok(GeoEntityType::Ocean),
        "Planet" => Ok(GeoEntityType::Planet),
        "PostalCode" => Ok(GeoEntityType::PostalCode),
        "Region" => Ok(GeoEntityType::Region),
        "Unsupported" => Ok(GeoEntityType::Unsupported),
        _ => invalid("invalid  geography entity type"),
    }
}

pub(super) fn geo_required_string(node: &MiniNode, name: &str, maximum: usize) -> Result<String> {
    let value = required(&node.attributes, "", name)?;
    if value.len() > maximum {
        return limit(" geography string");
    }
    Ok(value.to_owned())
}

pub(super) fn geo_optional_string(
    node: &MiniNode,
    name: &str,
    maximum: usize,
) -> Result<Option<String>> {
    optional(&node.attributes, "", name)
        .map(|value| {
            if value.len() > maximum {
                return limit(" geography string");
            }
            Ok(value.to_owned())
        })
        .transpose()
}

pub(super) fn geo_required_double(node: &MiniNode, name: &str) -> Result<String> {
    let value = required(&node.attributes, "", name)?;
    if !valid_xml_double(value) {
        return invalid(format!("invalid  geography {name}"));
    }
    Ok(value.to_owned())
}

pub(super) fn geo_optional_double(node: &MiniNode, name: &str) -> Result<Option<String>> {
    optional(&node.attributes, "", name)
        .map(|value| {
            if !valid_xml_double(value) {
                return invalid(format!("invalid  geography {name}"));
            }
            Ok(value.to_owned())
        })
        .transpose()
}

pub(super) fn validate_xsd_integer(value: &str, label: &str) -> Result<()> {
    let digits = value.strip_prefix(['+', '-']).unwrap_or(value);
    if digits.is_empty() || !digits.bytes().all(|byte| byte.is_ascii_digit()) {
        return invalid(format!("invalid  {label}"));
    }
    Ok(())
}

pub(super) fn validate_geo_base64(value: &str) -> Result<(usize, usize)> {
    let mut encoded = 0usize;
    let mut padding = 0usize;
    let mut saw_padding = false;
    for byte in value.bytes() {
        if matches!(byte, b' ' | b'\t' | b'\r' | b'\n') {
            continue;
        }
        encoded += 1;
        if byte == b'=' {
            saw_padding = true;
            padding += 1;
            if padding > 2 {
                return invalid("invalid  geography base64 padding");
            }
        } else if saw_padding || !(byte.is_ascii_alphanumeric() || matches!(byte, b'+' | b'/')) {
            return invalid("invalid  geography base64 data");
        }
    }
    if !encoded.is_multiple_of(4) {
        return invalid("invalid  geography base64 length");
    }
    let decoded = encoded
        .checked_div(4)
        .and_then(|value| value.checked_mul(3))
        .and_then(|value| value.checked_sub(padding))
        .ok_or_else(|| invalid_error(" geography base64 size overflow"))?;
    if decoded > MAX_GEO_BINARY_BYTES {
        return limit(" geography binary data");
    }
    Ok((encoded, decoded))
}
