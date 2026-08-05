//! PPT-specific OfficeArt wire assembly.
//!
//! The record grammar here is deliberately thin: shared headers, properties,
//! flags, and record kinds come from the litchi-odraw-backed wire substrate.

use zerocopy::IntoBytes;

use litchi_core::unit::emu_i32_to_ppt_master_i16_round;

use super::{
    BG_SHAPE_PROPERTIES, DGG_DEFAULT_PROPERTIES, Error, EscherDgData, EscherDggHeader,
    EscherHeader, EscherProperty, EscherRecordHeader, EscherSpData, EscherSpgrData, FileIdCluster,
    PROPERTY_FLAG_COMPLEX, ShapeFlags, SplitMenuColors, UserShapeData, header_version,
    ppt_prop_value, ppt_record_type, prop_id, record_type, shape_type,
};

/// Escher record builder
pub(crate) struct EscherBuilder {
    header: EscherHeader,
    data: Vec<u8>,
}

impl EscherBuilder {
    /// Create a new Escher record builder
    pub(crate) fn new(version: u8, instance: u16, record_type: u16) -> Self {
        Self {
            header: EscherHeader::new(version, instance, record_type, 0),
            data: Vec::new(),
        }
    }

    /// Add data to the record
    pub(crate) fn add_data(&mut self, data: &[u8]) {
        self.data.extend_from_slice(data);
        self.header.length = self.data.len() as u32;
    }

    /// Build the complete record
    pub(crate) fn build(&self) -> Result<Vec<u8>, Error> {
        let mut record = Vec::new();
        self.header.write(&mut record)?;
        record.extend_from_slice(&self.data);
        Ok(record)
    }
}

/// Create a DggContainer (Drawing Group Container) per MS-ODRAW
///
/// # Arguments
/// * `master_shapes` - Number of shapes in the master (6 for POI template)
/// * `slide_shape_counts` - Shape count for each slide (including group+background, so user_shapes+2)
pub(crate) fn create_dgg_container(
    master_shapes: u32,
    slide_shape_counts: &[u32],
) -> Result<Vec<u8>, Error> {
    let mut container =
        EscherBuilder::new(header_version::CONTAINER, 0, record_type::DGG_CONTAINER);

    // Total drawings = 1 (master) + number of slides
    let drawing_count = (slide_shape_counts.len() as u32) + 1;

    // Calculate total shapes: master + sum of all slide shapes
    let total_slide_shapes: u32 = slide_shape_counts.iter().sum();
    let csp_saved = master_shapes + total_slide_shapes;

    // POI uses cidcl=4 (3 clusters) even for 1 drawing
    let num_clusters = std::cmp::max(3, drawing_count as usize);
    let cidcl = (num_clusters + 1) as u32;

    // spidMax: Calculate based on highest drawing ID * 1024 + shapes in that drawing
    let max_slide_shapes = slide_shape_counts.iter().max().copied().unwrap_or(2);
    let spid_max = if drawing_count == 1 && master_shapes == ppt_prop_value::POI_MASTER_SHAPE_COUNT
    {
        ppt_prop_value::POI_SPID_MAX
    } else {
        drawing_count * 1024 + max_slide_shapes
    };

    // Build EscherDgg record (OfficeArtFDGGBlock)
    let mut dgg = EscherBuilder::new(header_version::SIMPLE, 0, record_type::DGG);
    let mut dgg_data = Vec::with_capacity(16 + num_clusters * 8);

    // Write header using zerocopy struct
    let header = EscherDggHeader {
        spid_max,
        cidcl,
        csp_saved,
        cdg_saved: drawing_count,
    };
    dgg_data.extend_from_slice(header.as_bytes());

    // FileIdClusters: each drawing gets its own cluster
    // dg_id 1 = master, dg_id 2+ = slides
    for dg_id in 1..=drawing_count {
        let cspid_cur = if dg_id == 1 {
            master_shapes + 1
        } else {
            let slide_idx = (dg_id - 2) as usize;
            slide_shape_counts.get(slide_idx).copied().unwrap_or(2) + 1
        };
        let cluster = FileIdCluster::new(dg_id, cspid_cur);
        dgg_data.extend_from_slice(cluster.as_bytes());
    }

    // Add reserved cluster slots to match POI
    for _ in drawing_count..num_clusters as u32 {
        dgg_data.extend_from_slice(FileIdCluster::reserved().as_bytes());
    }
    dgg.add_data(&dgg_data);
    container.add_data(&dgg.build()?);

    // NOTE: BStore container goes here if pictures are present
    // Call create_dgg_container_with_blips() instead if you have pictures

    // Add EscherOpt with default properties using const array
    let mut opt = EscherBuilder::new(
        header_version::OPT,
        DGG_DEFAULT_PROPERTIES.len() as u16,
        record_type::OPT,
    );
    for prop in &DGG_DEFAULT_PROPERTIES {
        opt.add_data(prop.as_bytes());
    }
    container.add_data(&opt.build()?);

    // Add SplitMenuColors using zerocopy struct
    let mut colors = EscherBuilder::new(header_version::SIMPLE, 4, record_type::SPLIT_MENU_COLORS);
    colors.add_data(SplitMenuColors::DEFAULT.as_bytes());
    container.add_data(&colors.build()?);

    container.build()
}

/// Create a DggContainer with BStore for pictures
///
/// Same as `create_dgg_container` but includes a BStoreContainer for pictures.
/// `bstore_blob` is the raw output of `Pictures::store`.
pub(crate) fn create_dgg_container_with_blips(
    master_shapes: u32,
    slide_shape_counts: &[u32],
    bstore_blob: &[u8],
) -> Result<Vec<u8>, Error> {
    let mut container =
        EscherBuilder::new(header_version::CONTAINER, 0, record_type::DGG_CONTAINER);

    // Total drawings = 1 (master) + number of slides
    let drawing_count = (slide_shape_counts.len() as u32) + 1;

    // Calculate total shapes: master + sum of all slide shapes
    let total_slide_shapes: u32 = slide_shape_counts.iter().sum();
    let csp_saved = master_shapes + total_slide_shapes;

    // POI uses cidcl=4 (3 clusters) even for 1 drawing
    let num_clusters = std::cmp::max(3, drawing_count as usize);
    let cidcl = (num_clusters + 1) as u32;

    // spidMax: Calculate based on highest drawing ID * 1024 + shapes in that drawing
    let max_slide_shapes = slide_shape_counts.iter().max().copied().unwrap_or(2);
    let spid_max = if drawing_count == 1 && master_shapes == ppt_prop_value::POI_MASTER_SHAPE_COUNT
    {
        ppt_prop_value::POI_SPID_MAX
    } else {
        drawing_count * 1024 + max_slide_shapes
    };

    // Build EscherDgg record (OfficeArtFDGGBlock)
    let mut dgg = EscherBuilder::new(header_version::SIMPLE, 0, record_type::DGG);
    let mut dgg_data = Vec::with_capacity(16 + num_clusters * 8);

    let header = EscherDggHeader {
        spid_max,
        cidcl,
        csp_saved,
        cdg_saved: drawing_count,
    };
    dgg_data.extend_from_slice(header.as_bytes());

    for dg_id in 1..=drawing_count {
        let cspid_cur = if dg_id == 1 {
            master_shapes + 1
        } else {
            let slide_idx = (dg_id - 2) as usize;
            slide_shape_counts.get(slide_idx).copied().unwrap_or(2) + 1
        };
        let cluster = FileIdCluster::new(dg_id, cspid_cur);
        dgg_data.extend_from_slice(cluster.as_bytes());
    }

    for _ in drawing_count..num_clusters as u32 {
        dgg_data.extend_from_slice(FileIdCluster::reserved().as_bytes());
    }
    dgg.add_data(&dgg_data);
    container.add_data(&dgg.build()?);

    // BStore container (if not empty)
    if !bstore_blob.is_empty() {
        container.add_data(bstore_blob);
    }

    // Add EscherOpt with default properties
    let mut opt = EscherBuilder::new(
        header_version::OPT,
        DGG_DEFAULT_PROPERTIES.len() as u16,
        record_type::OPT,
    );
    for prop in &DGG_DEFAULT_PROPERTIES {
        opt.add_data(prop.as_bytes());
    }
    container.add_data(&opt.build()?);

    // Add SplitMenuColors
    let mut colors = EscherBuilder::new(header_version::SIMPLE, 4, record_type::SPLIT_MENU_COLORS);
    colors.add_data(SplitMenuColors::DEFAULT.as_bytes());
    container.add_data(&colors.build()?);

    container.build()
}

/// Create a DgContainer with user shapes
pub(crate) fn create_dg_container_with_shapes(
    drawing_id: u32,
    shapes: &[UserShapeData],
) -> Result<Vec<u8>, Error> {
    create_dg_container_with_charts(drawing_id, shapes, &[], &[])
}

/// Create a DgContainer with user shapes and table groups.
///
/// Tables are emitted after the plain shapes inside the slide's
/// SpgrContainer; each table occupies one group shape id plus one id per
/// cell. See [`crate::writer::table::build_table_spgr_container`] for the record
/// layout.
#[cfg(test)]
pub(crate) fn create_dg_container_with_tables(
    drawing_id: u32,
    shapes: &[UserShapeData],
    tables: &[crate::writer::table::PositionedTable],
) -> Result<Vec<u8>, Error> {
    create_dg_container_with_charts(drawing_id, shapes, tables, &[])
}

/// Create a DgContainer with user shapes, table groups, and chart frames.
///
/// Chart frames (OLE object shapes referencing an embedded chart object)
/// follow the tables, one shape id each. See
/// [`crate::writer::chart::build_chart_sp_container`] for the record layout.
pub(crate) fn create_dg_container_with_charts(
    drawing_id: u32,
    shapes: &[UserShapeData],
    tables: &[crate::writer::table::PositionedTable],
    charts: &[crate::writer::chart::ChartFrame],
) -> Result<Vec<u8>, Error> {
    let table_shape_count: u32 = tables.iter().map(|table| table.table.shape_count()).sum();
    let mut container = EscherBuilder::new(header_version::CONTAINER, 0, record_type::DG_CONTAINER);

    // Total shapes = group + background + user shapes + table groups/cells + chart frames
    let total_shapes = (shapes.len() as u32)
        .saturating_add(table_shape_count)
        .saturating_add(charts.len() as u32)
        .saturating_add(2);

    // Add DG record
    let mut dg = EscherBuilder::new(header_version::DG, drawing_id as u16, record_type::DG);
    let dg_data = EscherDgData::new(total_shapes, drawing_id);
    dg.add_data(dg_data.as_bytes());
    container.add_data(&dg.build()?);

    // SpgrContainer
    let mut spgr_container =
        EscherBuilder::new(header_version::CONTAINER, 0, record_type::SPGR_CONTAINER);

    // Group patriarch SpContainer
    let mut group_sp_container =
        EscherBuilder::new(header_version::CONTAINER, 0, record_type::SP_CONTAINER);

    let mut spgr = EscherBuilder::new(header_version::SPGR, 0, record_type::SPGR);
    spgr.add_data(EscherSpgrData::ZERO.as_bytes());
    group_sp_container.add_data(&spgr.build()?);

    let group_spid = drawing_id << 10;
    let mut sp = EscherBuilder::new(
        header_version::SP,
        shape_type::NOT_PRIMITIVE,
        record_type::SP,
    );
    sp.add_data(EscherSpData::group_patriarch(group_spid).as_bytes());
    group_sp_container.add_data(&sp.build()?);

    spgr_container.add_data(&group_sp_container.build()?);

    // User shapes go INSIDE SpgrContainer (after group patriarch)
    let bg_spid = group_spid + 1;
    for (i, shape) in shapes.iter().enumerate() {
        let shape_spid = bg_spid + 1 + (i as u32);
        let sp_container = create_user_shape_container(shape_spid, shape)?;
        spgr_container.add_data(&sp_container);
    }

    // Table groups follow the plain shapes; each table's cells take the
    // consecutive shape ids after its group shape id.
    let mut table_group_spid = bg_spid + 1 + (shapes.len() as u32);
    for table in tables {
        let table_container =
            crate::writer::table::build_table_spgr_container(table, table_group_spid)?;
        spgr_container.add_data(&table_container);
        table_group_spid += table.table.shape_count();
    }

    // Chart frames follow the tables, one shape id each.
    for (index, frame) in charts.iter().enumerate() {
        let chart_container =
            crate::writer::chart::build_chart_sp_container(frame, table_group_spid + index as u32)?;
        spgr_container.add_data(&chart_container);
    }

    // Add SpgrContainer to DgContainer
    container.add_data(&spgr_container.build()?);

    // Background shape container - added to DgContainer OUTSIDE SpgrContainer (per POI)
    let mut bg_sp_container =
        EscherBuilder::new(header_version::CONTAINER, 0, record_type::SP_CONTAINER);

    let mut bg_sp = EscherBuilder::new(header_version::SP, shape_type::RECTANGLE, record_type::SP);
    bg_sp.add_data(EscherSpData::background(bg_spid).as_bytes());
    bg_sp_container.add_data(&bg_sp.build()?);

    let mut opt = EscherBuilder::new(
        header_version::OPT,
        BG_SHAPE_PROPERTIES.len() as u16,
        record_type::OPT,
    );
    for prop in &BG_SHAPE_PROPERTIES {
        opt.add_data(prop.as_bytes());
    }
    bg_sp_container.add_data(&opt.build()?);

    // NOTE: Per POI's PPDrawing.create(), background SpContainer has NO ClientAnchor or ClientData
    // Only Sp + Opt records are present

    // Add background SpContainer to DgContainer (NOT to SpgrContainer)
    container.add_data(&bg_sp_container.build()?);

    container.build()
}

/// Create a user shape SpContainer
pub(crate) fn create_user_shape_container(
    shape_id: u32,
    shape: &UserShapeData,
) -> Result<Vec<u8>, Error> {
    if shape.adjust_values.len() > 10 {
        return Err(std::io::Error::new(
            std::io::ErrorKind::InvalidInput,
            "OfficeArt shapes support at most 10 adjustment values",
        ));
    }
    let mut container = EscherBuilder::new(header_version::CONTAINER, 0, record_type::SP_CONTAINER);

    // Shape flags
    let mut flags = ShapeFlags::HAVE_ANCHOR | ShapeFlags::HAVE_SPT;
    if shape.flip_h {
        flags |= ShapeFlags::FLIP_H;
    }
    if shape.flip_v {
        flags |= ShapeFlags::FLIP_V;
    }

    // SP record
    let mut sp = EscherBuilder::new(header_version::SP, shape.shape_type, record_type::SP);
    sp.add_data(EscherSpData::with_flags(shape_id, flags).as_bytes());
    container.add_data(&sp.build()?);

    // OPT record with shape properties (sorted by property number, not full ID)
    // Per POI: sort by getPropertyNumber() which masks out flags (id & 0x3FFF)
    let mut properties: Vec<(EscherProperty, Option<Vec<u8>>)> = build_shape_properties(shape)
        .into_iter()
        .map(|property| (property, None))
        .collect();
    if let Some(geometry) = &shape.freeform_geometry {
        let rect = geometry.coordinate_space();
        let (vertices, segments) = geometry.encode_arrays()?;
        properties.extend([
            (
                EscherProperty::new(prop_id::GEOM_LEFT, rect.left as u32),
                None,
            ),
            (
                EscherProperty::new(prop_id::GEOM_TOP, rect.top as u32),
                None,
            ),
            (
                EscherProperty::new(prop_id::GEOM_RIGHT, rect.right as u32),
                None,
            ),
            (
                EscherProperty::new(prop_id::GEOM_BOTTOM, rect.bottom as u32),
                None,
            ),
            (
                EscherProperty::new(prop_id::SHAPE_PATH, geometry.path_type() as u32),
                None,
            ),
            (
                EscherProperty::new(
                    prop_id::VERTICES | PROPERTY_FLAG_COMPLEX,
                    vertices.len() as u32,
                ),
                Some(vertices),
            ),
            (
                EscherProperty::new(
                    prop_id::SEGMENT_INFO | PROPERTY_FLAG_COMPLEX,
                    segments.len() as u32,
                ),
                Some(segments),
            ),
        ]);
    }
    properties.sort_by_key(|(property, _)| ({ property.prop_id }) & 0x3FFF);
    let mut opt = EscherBuilder::new(
        header_version::OPT,
        properties.len() as u16,
        record_type::OPT,
    );
    for (property, _) in &properties {
        opt.add_data(property.as_bytes());
    }
    for (_, complex_data) in &properties {
        if let Some(data) = complex_data {
            opt.add_data(data);
        }
    }
    container.add_data(&opt.build()?);

    // ClientAnchor with position/size (8-byte short format for PPT top-level shapes)
    // POI uses: flag(y1), col1(x1), dx1(x2), row1(y2) - all shorts in master units
    let mut anchor = EscherBuilder::new(header_version::SIMPLE, 0, record_type::CLIENT_ANCHOR);
    let x1 = emu_i32_to_ppt_master_i16_round(shape.x);
    let y1 = emu_i32_to_ppt_master_i16_round(shape.y);
    let x2 = emu_i32_to_ppt_master_i16_round(shape.x + shape.width);
    let y2 = emu_i32_to_ppt_master_i16_round(shape.y + shape.height);
    // Short record format: 8 bytes (4 shorts)
    anchor.add_data(&y1.to_le_bytes()); // flag/top
    anchor.add_data(&x1.to_le_bytes()); // col1/left
    anchor.add_data(&x2.to_le_bytes()); // dx1/right
    anchor.add_data(&y2.to_le_bytes()); // row1/bottom
    container.add_data(&anchor.build()?);

    // ClientData grammar order: animation, click, mouse-over, placeholder,
    // then round-trip records such as programmable tags.
    // MUST come BEFORE ClientTextbox per POI (addChildBefore(clientData, EscherTextboxRecord.RECORD_ID))
    let mut client_data = if let Some(ref animation_info) = shape.animation_info {
        Some(build_client_data_with_animation(animation_info)?)
    } else {
        None
    };
    let legacy_click = shape
        .hyperlink_id
        .map(|hyperlink_id| {
            legacy_hyperlink_interaction(
                hyperlink_id,
                shape.hyperlink_action,
                shape.hyperlink_jump,
                shape.hyperlink_type,
            )
        })
        .transpose()?;
    for trigger in [
        crate::InteractionTrigger::Click,
        crate::InteractionTrigger::MouseOver,
    ] {
        let mut matching = shape
            .interactions
            .iter()
            .filter(|interaction| interaction.trigger == trigger);
        let interaction = matching.next();
        if matching.next().is_some() {
            return Err(std::io::Error::other(
                "shape contains duplicate interactive triggers",
            ));
        }
        let interaction = interaction.or_else(|| {
            (trigger == crate::InteractionTrigger::Click)
                .then_some(legacy_click.as_ref())
                .flatten()
        });
        if let Some(interaction) = interaction {
            let bytes = interaction
                .to_bytes_with_limits(crate::InteractionLimits::default())
                .map_err(|error| std::io::Error::other(error.to_string()))?;
            append_client_data_payload(&mut client_data, &bytes)?;
        }
    }
    if let Some(placeholder_type) = shape.placeholder_type {
        let placeholder = build_client_data_with_placeholder(placeholder_type)?;
        append_client_data_record_payload(&mut client_data, &placeholder)?;
    }
    if let Some(runs) = &shape.smart_tag_runs
        && let Some(programmable_tags) =
            crate::writer::smart_tags::build_shape_programmable_tags(runs)
                .map_err(|error| std::io::Error::other(error.to_string()))?
    {
        append_client_data_payload(&mut client_data, &programmable_tags)?;
    }
    if let Some(client_data) = client_data {
        crate::ClientData::parse(&client_data)
            .map_err(|error| std::io::Error::other(error.to_string()))?;
        container.add_data(&client_data);
    }

    // ClientTextBox if text present (prefer paragraphs with formatting over plain text)
    if let Some(paragraphs) = &shape.paragraphs {
        if !paragraphs.is_empty() {
            let textbox = build_client_textbox_formatted_with_interactions(
                paragraphs,
                shape.text_type,
                &shape.text_interactions,
            )?;
            container.add_data(&textbox);
        } else if !shape.text_interactions.is_empty() {
            return Err(std::io::Error::other(
                "shape has text interactions but no corresponding text",
            ));
        }
    } else if let Some(text) = &shape.text {
        let textbox = build_client_textbox_with_interactions(
            text,
            shape.text_type,
            &shape.text_interactions,
        )?;
        container.add_data(&textbox);
    } else if !shape.text_interactions.is_empty() {
        return Err(std::io::Error::other(
            "shape has text interactions but no corresponding text",
        ));
    }

    container.build()
}

fn legacy_hyperlink_interaction(
    hyperlink_id: u32,
    action: u8,
    jump: u8,
    hyperlink_type: u8,
) -> Result<crate::Interaction, Error> {
    let mut atom_data = [0u8; 16];
    atom_data[4..8].copy_from_slice(&hyperlink_id.to_le_bytes());
    atom_data[8] = action;
    atom_data[10] = jump;
    atom_data[12] = hyperlink_type;
    let atom = crate::InteractiveInfoAtom::parse_payload(&atom_data)
        .map_err(|error| std::io::Error::other(error.to_string()))?;
    Ok(crate::Interaction {
        trigger: crate::InteractionTrigger::Click,
        sound_id: atom.sound_id,
        hyperlink_id: atom.hyperlink_id,
        action: atom.action,
        ole_verb: atom.ole_verb,
        jump: atom.jump,
        animated: atom.animated,
        stop_sound: atom.stop_sound,
        custom_show_return: atom.custom_show_return,
        visited: atom.visited,
        link_target: atom.link_target,
        macro_name: None,
        unused: atom.unused,
        macro_name_data: None,
    })
}

fn append_client_data_record_payload(
    client_data: &mut Option<Vec<u8>>,
    record: &[u8],
) -> Result<(), Error> {
    let payload_len = record
        .get(4..8)
        .and_then(|bytes| bytes.try_into().ok())
        .map(u32::from_le_bytes)
        .and_then(|length| usize::try_from(length).ok())
        .ok_or_else(|| std::io::Error::other("invalid ClientData record header"))?;
    let expected_options = 0x000fu16.to_le_bytes();
    let expected_type = record_type::CLIENT_DATA.to_le_bytes();
    if record.len() != payload_len.saturating_add(8)
        || record.get(0..2) != Some(expected_options.as_slice())
        || record.get(2..4) != Some(expected_type.as_slice())
    {
        return Err(std::io::Error::other("invalid ClientData record"));
    }
    append_client_data_payload(client_data, &record[8..])
}

fn append_client_data_payload(
    client_data: &mut Option<Vec<u8>>,
    payload: &[u8],
) -> Result<(), Error> {
    let data = client_data.get_or_insert_with(|| {
        let mut bytes = Vec::with_capacity(8);
        bytes.extend_from_slice(&0x000fu16.to_le_bytes());
        bytes.extend_from_slice(&record_type::CLIENT_DATA.to_le_bytes());
        bytes.extend_from_slice(&0u32.to_le_bytes());
        bytes
    });
    let declared = data
        .get(4..8)
        .and_then(|bytes| bytes.try_into().ok())
        .map(u32::from_le_bytes)
        .ok_or_else(|| std::io::Error::other("invalid ClientData record header"))?;
    if data.len()
        != usize::try_from(declared)
            .unwrap_or(usize::MAX)
            .saturating_add(8)
    {
        return Err(std::io::Error::other(
            "ClientData record length does not match its payload",
        ));
    }
    let new_length = u32::try_from(data.len().saturating_sub(8))
        .ok()
        .and_then(|length| {
            u32::try_from(payload.len())
                .ok()
                .and_then(|addition| length.checked_add(addition))
        })
        .ok_or_else(|| std::io::Error::other("ClientData payload exceeds u32"))?;
    data.extend_from_slice(payload);
    data[4..8].copy_from_slice(&new_length.to_le_bytes());
    Ok(())
}

/// Build ClientData record with InteractiveInfo for a legacy hyperlink.
#[cfg(test)]
pub(crate) fn build_client_data_with_hyperlink(
    hyperlink_id: u32,
    action: u8,
    jump: u8,
    hyperlink_type: u8,
) -> Result<Vec<u8>, Error> {
    let interaction = legacy_hyperlink_interaction(hyperlink_id, action, jump, hyperlink_type)?;
    let mut client_data = None;
    append_client_data_payload(
        &mut client_data,
        &interaction
            .to_bytes()
            .map_err(|error| std::io::Error::other(error.to_string()))?,
    )?;
    client_data.ok_or_else(|| std::io::Error::other("missing ClientData record"))
}

/// Build ClientData record with AnimationInfo.
///
/// Per LibreOffice reference files, animation sounds only need AnimationInfo in ClientData.
/// InteractiveInfo with action=6 (MEDIA) is for movie/media objects, NOT animation sounds.
/// The reference `sound.ppt` has AnimationInfo WITHOUT InteractiveInfo in its ClientData.
fn build_client_data_with_animation(
    animation_info: &crate::animation::AnimationInfo,
) -> Result<Vec<u8>, Error> {
    use crate::animation::writer::write_animation_info;

    // Write AnimationInfo container (contains AnimationInfoAtom with soundRef)
    let (animation_bytes, _sound_ref) = write_animation_info(animation_info).map_err(|error| {
        std::io::Error::new(std::io::ErrorKind::InvalidInput, error.to_string())
    })?;

    // ClientData Escher record (0xF011) wrapping AnimationInfo only
    let mut client_data =
        EscherBuilder::new(header_version::CONTAINER, 0, record_type::CLIENT_DATA);
    client_data.add_data(&animation_bytes);
    client_data.build()
}

/// Build ClientData record with OEPlaceholderAtom for placeholder shapes
/// Per POI HSLFSimpleShape - placeholders have OEPlaceholderAtom in ClientData
pub(crate) fn build_client_data_with_placeholder(placeholder_type: u8) -> Result<Vec<u8>, Error> {
    use crate::writer::records::RecordBuilder;

    // OEPlaceholderAtom (type 0x0BC3 = 3011)
    // Structure: position (4 bytes), placeholderType (1 byte), size (1 byte), unused (2 bytes)
    let mut oe_atom = RecordBuilder::new(0x00, 0, ppt_record_type::OE_PLACEHOLDER_ATOM);
    oe_atom.write_data(&0u32.to_le_bytes()); // position = 0
    oe_atom.write_data(&[placeholder_type]); // placeholder type (12 = NotesBody per MS-PPT)
    oe_atom.write_data(&[0x00]); // size = full
    oe_atom.write_data(&[0x00, 0x00]); // unused
    let oe_bytes = oe_atom.build()?;

    // ClientData Escher record (0xF011) wrapping OEPlaceholderAtom
    let mut client_data =
        EscherBuilder::new(header_version::CONTAINER, 0, record_type::CLIENT_DATA);
    client_data.add_data(&oe_bytes);

    client_data.build()
}

/// Build shape properties for OPT record
/// Based on Apache POI HSLFTextBox.createSpContainer() defaults
pub(crate) fn build_shape_properties(shape: &UserShapeData) -> Vec<EscherProperty> {
    let mut props = Vec::with_capacity(16);

    if let Some(rotation) = shape.rotation {
        props.push(EscherProperty::new(prop_id::ROTATION, rotation as u32));
    }
    for (index, &value) in shape.adjust_values.iter().enumerate() {
        props.push(EscherProperty::new(
            prop_id::ADJUST_VALUE + index as u16,
            value as u32,
        ));
    }

    // Picture shapes have special handling - BLIP reference only, no fill/line
    if let Some(picture_index) = shape.picture_index {
        // PROTECTION__LOCKAGAINSTGROUPING (0x007F) = 0x800080 per POI
        props.push(EscherProperty::new(0x007F, 0x0080_0080));
        // BLIP__BLIPTODISPLAY (0x4104) - with isBlipId flag (0x4000 + 0x0104)
        props.push(EscherProperty::new(0x4104, picture_index));
        // No fill for pictures (picture IS the fill)
        props.push(EscherProperty::new(
            prop_id::NO_FILL_HIT_TEST,
            ppt_prop_value::FILL_STYLE_DISABLED,
        ));
        // No line for pictures
        props.push(EscherProperty::new(prop_id::LINE_STYLE_BOOL, 0x0008_0000));
        return props;
    }

    // Fill properties
    if let Some(fill_color) = shape.fill_color {
        // Fill type (0=solid, 4=shade/gradient) - MUST be first
        if let Some(fill_type) = shape.fill_type {
            props.push(EscherProperty::new(prop_id::FILL_TYPE, fill_type));
        }

        // Fill color
        props.push(EscherProperty::new(prop_id::FILL_COLOR, fill_color));

        // Back color (for gradients) - before angle
        if let Some(back_color) = shape.fill_back_color {
            props.push(EscherProperty::new(prop_id::FILL_BACK_COLOR, back_color));
        }

        // Gradient angle (for gradient fills) - MUST be before opacity
        if let Some(angle) = shape.fill_angle {
            props.push(EscherProperty::new(prop_id::FILL_ANGLE, angle as u32));
        }

        // Fill opacity (after angle)
        if let Some(opacity) = shape.fill_opacity {
            props.push(EscherProperty::new(prop_id::FILL_OPACITY, opacity));
        }

        // Match POI setForegroundColor: filled=true, fillShape=false,
        // and noFillHitTest=true, with all three use bits set.
        props.push(EscherProperty::new(
            prop_id::NO_FILL_HIT_TEST,
            ppt_prop_value::FILL_STYLE_ENABLED,
        ));
    } else {
        // Default: scheme fill colors with no-fill flag
        props.push(EscherProperty::new(prop_id::FILL_COLOR, 0x0800_0004)); // scheme fill
        props.push(EscherProperty::new(prop_id::FILL_BACK_COLOR, 0x0800_0000));
        props.push(EscherProperty::new(
            prop_id::NO_FILL_HIT_TEST,
            ppt_prop_value::FILL_STYLE_DISABLED,
        ));
    }

    if let Some(blip_index) = shape.fill_blip_index {
        props.push(EscherProperty::new(prop_id::FILL_BLIP, blip_index));
    }

    // Line properties (based on POI HSLFSimpleShape)
    if let Some(line_color) = shape.line_color {
        props.push(EscherProperty::new(prop_id::LINE_COLOR, line_color));
        if let Some(opacity) = shape.line_opacity {
            props.push(EscherProperty::new(prop_id::LINE_OPACITY, opacity));
        }
        if let Some(width) = shape.line_width {
            props.push(EscherProperty::new(prop_id::LINE_WIDTH, width as u32));
        }
        if let Some(style) = shape.line_style {
            props.push(EscherProperty::new(prop_id::LINE_STYLE, style));
        }
        // Line dash style
        if let Some(dash) = shape.line_dash_style {
            props.push(EscherProperty::new(prop_id::LINE_DASH_STYLE, dash));
        }
        // Line start arrow
        if let Some(arrow) = shape.line_start_arrow {
            props.push(EscherProperty::new(prop_id::LINE_START_ARROW, arrow));
            props.push(EscherProperty::new(
                prop_id::LINE_START_ARROW_WIDTH,
                shape.line_start_arrow_width.unwrap_or(1),
            ));
            props.push(EscherProperty::new(
                prop_id::LINE_START_ARROW_LENGTH,
                shape.line_start_arrow_length.unwrap_or(1),
            ));
        }
        // Line end arrow
        if let Some(arrow) = shape.line_end_arrow {
            props.push(EscherProperty::new(prop_id::LINE_END_ARROW, arrow));
            props.push(EscherProperty::new(
                prop_id::LINE_END_ARROW_WIDTH,
                shape.line_end_arrow_width.unwrap_or(1),
            ));
            props.push(EscherProperty::new(
                prop_id::LINE_END_ARROW_LENGTH,
                shape.line_end_arrow_length.unwrap_or(1),
            ));
        }
        if let Some(join_style) = shape.line_join_style {
            props.push(EscherProperty::new(prop_id::LINE_JOIN_STYLE, join_style));
        }
        if let Some(end_cap_style) = shape.line_end_cap_style {
            props.push(EscherProperty::new(
                prop_id::LINE_END_CAP_STYLE,
                end_cap_style,
            ));
        }
        // Enable line: 0x180018 = line visible
        props.push(EscherProperty::new(prop_id::LINE_STYLE_BOOL, 0x0018_0018));
    } else {
        // No line: POI uses 0x80000 for no line
        props.push(EscherProperty::new(prop_id::LINE_COLOR, 0x0800_0001)); // scheme line
        props.push(EscherProperty::new(prop_id::LINE_STYLE_BOOL, 0x0008_0000));
    }

    // Shadow properties
    if shape.has_shadow {
        // Offset shadow is the specified default and is required for the offsets below.
        props.push(EscherProperty::new(
            prop_id::SHADOW_TYPE,
            shape.shadow_type.unwrap_or(0),
        ));

        // Shadow color
        let shadow_color = shape.shadow_color.unwrap_or(0x0800_0002); // default: scheme shadow
        props.push(EscherProperty::new(prop_id::SHADOW_COLOR, shadow_color));

        // Shadow offsets
        let offset_x = shape.shadow_offset_x.unwrap_or(25400) as u32; // default: 2pt
        let offset_y = shape.shadow_offset_y.unwrap_or(25400) as u32; // default: 2pt
        props.push(EscherProperty::new(prop_id::SHADOW_OFFSET_X, offset_x));
        props.push(EscherProperty::new(prop_id::SHADOW_OFFSET_Y, offset_y));

        // Shadow opacity
        if let Some(opacity) = shape.shadow_opacity {
            props.push(EscherProperty::new(prop_id::SHADOW_OPACITY, opacity));
        }

        // Enable shadow boolean
        props.push(EscherProperty::new(
            prop_id::SHADOW_BOOL,
            ppt_prop_value::SHADOW_STYLE_ENABLED,
        ));
    } else {
        // No shadow - still set scheme color for consistency
        props.push(EscherProperty::new(prop_id::SHADOW_COLOR, 0x0800_0002));
        props.push(EscherProperty::new(
            prop_id::SHADOW_BOOL,
            ppt_prop_value::SHADOW_STYLE_DISABLED,
        ));
    }

    props
}

/// Build ClientTextBox record with plain text content (no formatting)
/// Based on Apache POI EscherTextboxWrapper and HSLFTextShape
/// text_type: 0=Title, 1=Body, 2=Notes, 4=Other
pub(crate) fn build_client_textbox(text: &str, text_type: u32) -> Result<Vec<u8>, Error> {
    build_client_textbox_with_interactions(text, text_type, &[])
}

pub(crate) fn build_client_textbox_with_interactions(
    text: &str,
    text_type: u32,
    interactions: &[crate::TextInteraction],
) -> Result<Vec<u8>, Error> {
    use crate::writer::records::{RecordBuilder, record_type as ppt_rt};

    let mut result = Vec::new();
    let mut ppt_content = Vec::new();

    // TextHeaderAtom (type=3999): textType from parameter
    let mut text_header = RecordBuilder::new(0, 0, ppt_rt::TEXT_HEADER_ATOM);
    text_header.write_data(&text_type.to_le_bytes());
    ppt_content.extend_from_slice(&text_header.build()?);

    // TextBytesAtom (type=4008) for ASCII or TextCharsAtom (type=4000) for Unicode
    let is_ascii = text.is_ascii();
    if is_ascii {
        let mut text_atom = RecordBuilder::new(0, 0, ppt_rt::TEXT_BYTES_ATOM);
        text_atom.write_data(text.as_bytes());
        ppt_content.extend_from_slice(&text_atom.build()?);
    } else {
        let mut text_atom = RecordBuilder::new(0, 0, ppt_rt::TEXT_CHARS_ATOM);
        for ch in text.encode_utf16() {
            text_atom.write_data(&ch.to_le_bytes());
        }
        ppt_content.extend_from_slice(&text_atom.build()?);
    }

    // StyleTextPropAtom with no formatting
    let too_large = || {
        Error::new(
            std::io::ErrorKind::InvalidInput,
            "ClientTextbox text exceeds the PPT size limit",
        )
    };
    let text_units = u32::try_from(text.encode_utf16().count()).map_err(|_| too_large())?;
    let char_count = text_units.checked_add(1).ok_or_else(too_large)?;
    let mut style_atom = RecordBuilder::new(0, 0, ppt_rt::STYLE_TEXT_PROP_ATOM);
    style_atom.write_data(&char_count.to_le_bytes()); // para char count
    style_atom.write_data(&0u16.to_le_bytes()); // indent
    style_atom.write_data(&0u32.to_le_bytes()); // para mask
    style_atom.write_data(&char_count.to_le_bytes()); // char count
    style_atom.write_data(&0u32.to_le_bytes()); // char mask
    ppt_content.extend_from_slice(&style_atom.build()?);
    append_text_interactions(
        &mut ppt_content,
        text_units,
        interactions,
        crate::TextInteractionLimits::default(),
    )?;

    let header = EscherRecordHeader::new(0x0F, 0, 0xF00D, ppt_content.len() as u32);
    result.extend_from_slice(header.as_bytes());
    result.extend_from_slice(&ppt_content);

    Ok(result)
}

/// Build ClientTextBox record with rich text formatting (paragraphs with runs)
/// text_type: 0=Title, 1=Body, 2=Notes, 4=Other
#[cfg(test)]
pub(crate) fn build_client_textbox_formatted(
    paragraphs: &[crate::writer::text_format::Paragraph],
    text_type: u32,
) -> Result<Vec<u8>, Error> {
    build_client_textbox_formatted_with_interactions(paragraphs, text_type, &[])
}

fn build_client_textbox_formatted_with_interactions(
    paragraphs: &[crate::writer::text_format::Paragraph],
    text_type: u32,
    interactions: &[crate::TextInteraction],
) -> Result<Vec<u8>, Error> {
    use crate::writer::records::{RecordBuilder, record_type as ppt_rt};
    use crate::writer::text_format::TextPropsBuilder;

    let mut result = Vec::new();
    let mut ppt_content = Vec::new();

    // TextHeaderAtom (type=3999): textType from parameter
    let mut text_header = RecordBuilder::new(0, 0, ppt_rt::TEXT_HEADER_ATOM);
    text_header.write_data(&text_type.to_le_bytes());
    ppt_content.extend_from_slice(&text_header.build()?);

    // Build text content from paragraphs
    let mut builder = TextPropsBuilder::new();
    for para in paragraphs {
        builder.add_paragraph(para.clone());
    }

    // Use TextCharsAtom (UTF-16) since we might have unicode
    let text_chars = builder.build_text_chars();
    let text_units = u32::try_from(text_chars.len() / 2).map_err(|_| {
        std::io::Error::new(
            std::io::ErrorKind::InvalidInput,
            "ClientTextbox text exceeds the PPT size limit",
        )
    })?;
    let mut text_atom = RecordBuilder::new(0, 0, ppt_rt::TEXT_CHARS_ATOM);
    text_atom.write_data(&text_chars);
    ppt_content.extend_from_slice(&text_atom.build()?);

    // StyleTextPropAtom with full formatting
    let style_data = builder.build_style_text_prop()?;
    let mut style_atom = RecordBuilder::new(0, 0, ppt_rt::STYLE_TEXT_PROP_ATOM);
    style_atom.write_data(&style_data);
    ppt_content.extend_from_slice(&style_atom.build()?);
    append_text_interactions(
        &mut ppt_content,
        text_units,
        interactions,
        crate::TextInteractionLimits::default(),
    )?;

    let header = EscherRecordHeader::new(0x0F, 0, 0xF00D, ppt_content.len() as u32);
    result.extend_from_slice(header.as_bytes());
    result.extend_from_slice(&ppt_content);

    Ok(result)
}

fn append_text_interactions(
    output: &mut Vec<u8>,
    text_units: u32,
    interactions: &[crate::TextInteraction],
    limits: crate::TextInteractionLimits,
) -> Result<(), Error> {
    if interactions.len() > limits.max_interactions {
        return Err(std::io::Error::new(
            std::io::ErrorKind::InvalidInput,
            "ClientTextbox exceeds the text interaction count limit",
        ));
    }
    for interaction in interactions {
        output.extend_from_slice(
            &interaction
                .to_bytes_for_text(text_units, limits)
                .map_err(|error| std::io::Error::other(error.to_string()))?,
        );
    }
    Ok(())
}
