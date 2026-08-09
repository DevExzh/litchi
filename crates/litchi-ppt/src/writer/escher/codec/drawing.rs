//! Drawing-group and drawing-container record families.

use zerocopy::IntoBytes;

use litchi_odraw::write::{Sp, record_type, shape_type};

use super::super::{
    BG_SHAPE_PROPERTIES, DGG_DEFAULT_PROPERTIES, Error, EscherDgData, EscherDggHeader,
    EscherSpgrData, FileIdCluster, SplitMenuColors, UserShapeData, header_version, ppt_prop_value,
};
use super::shapes::create_user_shape_container;
use super::wire::EscherBuilder;

/// Creates the file-level `DggContainer` without a BLIP store.
pub(crate) fn create_dgg_container(
    master_shapes: u32,
    slide_shape_counts: &[u32],
) -> Result<Vec<u8>, Error> {
    build_dgg_container(master_shapes, slide_shape_counts, None)
}

/// Creates the file-level `DggContainer` with an already-encoded BLIP store.
pub(crate) fn create_dgg_container_with_blips(
    master_shapes: u32,
    slide_shape_counts: &[u32],
    bstore_blob: &[u8],
) -> Result<Vec<u8>, Error> {
    build_dgg_container(master_shapes, slide_shape_counts, Some(bstore_blob))
}

fn build_dgg_container(
    master_shapes: u32,
    slide_shape_counts: &[u32],
    bstore_blob: Option<&[u8]>,
) -> Result<Vec<u8>, Error> {
    let mut container =
        EscherBuilder::new(header_version::CONTAINER, 0, record_type::DGG_CONTAINER);
    let slide_count = u32::try_from(slide_shape_counts.len()).map_err(|_err| {
        Error::new(
            std::io::ErrorKind::InvalidInput,
            "slide count exceeds the Escher drawing-group limit",
        )
    })?;
    let drawing_count = slide_count + 1;
    let total_slide_shapes: u32 = slide_shape_counts.iter().sum();
    let csp_saved = master_shapes + total_slide_shapes;

    let num_clusters = std::cmp::max(3, drawing_count);
    let cidcl = num_clusters + 1;
    let max_slide_shapes = slide_shape_counts.iter().max().copied().unwrap_or(2);
    let spid_max = if drawing_count == 1 && master_shapes == ppt_prop_value::POI_MASTER_SHAPE_COUNT
    {
        ppt_prop_value::POI_SPID_MAX
    } else {
        drawing_count * 1024 + max_slide_shapes
    };

    let mut dgg = EscherBuilder::new(header_version::SIMPLE, 0, record_type::DGG);
    let mut dgg_data = Vec::with_capacity(16 + num_clusters as usize * 8);
    dgg_data.extend_from_slice(
        EscherDggHeader {
            spid_max,
            cidcl,
            csp_saved,
            cdg_saved: drawing_count,
        }
        .as_bytes(),
    );
    for dg_id in 1..=drawing_count {
        let cspid_cur = if dg_id == 1 {
            master_shapes + 1
        } else {
            let slide_idx = (dg_id - 2) as usize;
            slide_shape_counts.get(slide_idx).copied().unwrap_or(2) + 1
        };
        dgg_data.extend_from_slice(FileIdCluster::new(dg_id, cspid_cur).as_bytes());
    }
    for _ in drawing_count..num_clusters {
        dgg_data.extend_from_slice(FileIdCluster::reserved().as_bytes());
    }
    dgg.add_data(&dgg_data);
    container.add_data(&dgg.build()?);

    if let Some(nonempty_blob) = bstore_blob.filter(|blob| !blob.is_empty()) {
        container.add_data(nonempty_blob);
    }

    #[allow(
        clippy::cast_possible_truncation,
        reason = "`DGG_DEFAULT_PROPERTIES` is a fixed eight-element array"
    )]
    let mut opt = EscherBuilder::new(
        header_version::OPT,
        DGG_DEFAULT_PROPERTIES.len() as u16,
        record_type::OPT,
    );
    for property in &DGG_DEFAULT_PROPERTIES {
        opt.add_data(property.as_bytes());
    }
    container.add_data(&opt.build()?);

    let mut colors = EscherBuilder::new(header_version::SIMPLE, 4, record_type::SPLIT_MENU_COLORS);
    colors.add_data(SplitMenuColors::DEFAULT.as_bytes());
    container.add_data(&colors.build()?);
    container.build()
}

/// Creates a drawing container with ordinary top-level shapes.
pub(crate) fn create_dg_container_with_shapes(
    drawing_id: u32,
    shapes: &[UserShapeData],
) -> Result<Vec<u8>, Error> {
    create_dg_container_with_charts(drawing_id, shapes, &[], &[])
}

/// Creates a drawing container with shapes and test-only table groups.
#[cfg(test)]
pub(crate) fn create_dg_container_with_tables(
    drawing_id: u32,
    shapes: &[UserShapeData],
    tables: &[crate::writer::table::PositionedTable],
) -> Result<Vec<u8>, Error> {
    create_dg_container_with_charts(drawing_id, shapes, tables, &[])
}

/// Creates a drawing container with shapes, table groups, and chart frames.
pub(crate) fn create_dg_container_with_charts(
    drawing_id: u32,
    shapes: &[UserShapeData],
    tables: &[crate::writer::table::PositionedTable],
    charts: &[crate::writer::chart::ChartFrame],
) -> Result<Vec<u8>, Error> {
    let table_shape_count: u32 = tables.iter().map(|table| table.table.shape_count()).sum();
    let mut container = EscherBuilder::new(header_version::CONTAINER, 0, record_type::DG_CONTAINER);
    let shape_count = u32::try_from(shapes.len()).unwrap_or(u32::MAX);
    let total_shapes = shape_count
        .saturating_add(table_shape_count)
        .saturating_add(u32::try_from(charts.len()).unwrap_or(u32::MAX))
        .saturating_add(2);

    #[allow(
        clippy::cast_possible_truncation,
        reason = "the `DG` instance field carries the low 12 bits of the drawing id; `EscherHeader::new` masks it"
    )]
    let mut dg = EscherBuilder::new(header_version::DG, drawing_id as u16, record_type::DG);
    dg.add_data(EscherDgData::new(total_shapes, drawing_id).as_bytes());
    container.add_data(&dg.build()?);

    let mut spgr_container =
        EscherBuilder::new(header_version::CONTAINER, 0, record_type::SPGR_CONTAINER);
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
    sp.add_data(Sp::group_patriarch(group_spid).as_bytes());
    group_sp_container.add_data(&sp.build()?);
    spgr_container.add_data(&group_sp_container.build()?);

    let bg_spid = group_spid + 1;
    for (shape_spid, shape) in (bg_spid + 1..).zip(shapes) {
        spgr_container.add_data(&create_user_shape_container(shape_spid, shape)?);
    }

    let mut table_group_spid = bg_spid + 1 + shape_count;
    for table in tables {
        let table_container =
            crate::writer::table::build_table_spgr_container(table, table_group_spid)?;
        spgr_container.add_data(&table_container);
        table_group_spid += table.table.shape_count();
    }

    for (chart_spid, frame) in (table_group_spid..).zip(charts) {
        let chart_container = crate::writer::chart::build_chart_sp_container(frame, chart_spid)?;
        spgr_container.add_data(&chart_container);
    }
    container.add_data(&spgr_container.build()?);

    // PowerPoint keeps the background shape outside the root SpgrContainer.
    let mut background_container =
        EscherBuilder::new(header_version::CONTAINER, 0, record_type::SP_CONTAINER);
    let mut background_shape =
        EscherBuilder::new(header_version::SP, shape_type::RECTANGLE, record_type::SP);
    background_shape.add_data(Sp::background(bg_spid).as_bytes());
    background_container.add_data(&background_shape.build()?);
    #[allow(
        clippy::cast_possible_truncation,
        reason = "`BG_SHAPE_PROPERTIES` is a fixed eight-element array"
    )]
    let mut opt = EscherBuilder::new(
        header_version::OPT,
        BG_SHAPE_PROPERTIES.len() as u16,
        record_type::OPT,
    );
    for property in &BG_SHAPE_PROPERTIES {
        opt.add_data(property.as_bytes());
    }
    background_container.add_data(&opt.build()?);
    container.add_data(&background_container.build()?);

    container.build()
}
