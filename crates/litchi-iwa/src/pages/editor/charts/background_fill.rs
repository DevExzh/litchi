//! Native chart-background fill CRUD for Pages body charts.

use super::*;
use crate::charts::background_fill::{
    chart_background_fill as read_native_chart_background_fill,
    set_chart_background_fill as set_native_chart_background_fill,
    set_chart_background_image_fill_data,
};
use crate::shapes::{RgbaColor, ShapeFill, ShapeImageFill, ShapeImageFillTechnique};

impl PagesEditor {
    /// Read the effective color, gradient, or image background of one body chart.
    pub fn body_chart_background_fill(&self, drawable_object_id: u64) -> Result<ShapeFill> {
        body_chart_background_fill(self, drawable_object_id)
    }

    /// Replace one body chart's background fill transactionally.
    pub fn set_body_chart_background_fill(
        &mut self,
        drawable_object_id: u64,
        fill: &ShapeFill,
    ) -> Result<()> {
        set_body_chart_background_fill(self, drawable_object_id, fill)
    }

    /// Embed image bytes and use them as a simple or tinted chart background.
    pub fn set_body_chart_background_image_fill(
        &mut self,
        drawable_object_id: u64,
        preferred_filename: &str,
        data: &[u8],
        technique: ShapeImageFillTechnique,
        tint: Option<RgbaColor>,
    ) -> Result<ShapeImageFill> {
        set_body_chart_background_image_fill(
            self,
            drawable_object_id,
            preferred_filename,
            data,
            technique,
            tint,
        )
    }
}

fn body_chart_background_fill(editor: &PagesEditor, drawable_object_id: u64) -> Result<ShapeFill> {
    let graph = body_chart_graph(editor, drawable_object_id)?;
    read_native_chart_background_fill(
        editor.package(),
        &graph.archive_name,
        drawable_object_id,
        "Pages",
    )
}

fn set_body_chart_background_fill(
    editor: &mut PagesEditor,
    drawable_object_id: u64,
    fill: &ShapeFill,
) -> Result<()> {
    if &body_chart_background_fill(editor, drawable_object_id)? == fill {
        return Ok(());
    }
    let graph = body_chart_graph(editor, drawable_object_id)?;
    let mut staged = editor.package().clone();
    set_native_chart_background_fill(
        &mut staged,
        &graph.archive_name,
        drawable_object_id,
        "Pages",
        fill,
    )?;
    let verified = PagesEditor::from_bytes(&staged.to_bytes()?)?;
    if &verified.body_chart_background_fill(drawable_object_id)? != fill {
        return Err(Error::InvalidFormat(
            "Pages chart background fill update failed validation".to_owned(),
        ));
    }
    *editor = verified;
    Ok(())
}

fn set_body_chart_background_image_fill(
    editor: &mut PagesEditor,
    drawable_object_id: u64,
    preferred_filename: &str,
    data: &[u8],
    technique: ShapeImageFillTechnique,
    tint: Option<RgbaColor>,
) -> Result<ShapeImageFill> {
    let graph = body_chart_graph(editor, drawable_object_id)?;
    let fill_size = graph.info.geometry.size.ok_or_else(|| {
        Error::InvalidFormat(format!(
            "Pages chart {drawable_object_id} has no background image-fill dimensions"
        ))
    })?;
    let mut staged = editor.package().clone();
    let image = set_chart_background_image_fill_data(
        &mut staged,
        &graph.archive_name,
        drawable_object_id,
        "Pages",
        preferred_filename,
        data,
        technique,
        fill_size,
        tint,
    )?;
    let verified = PagesEditor::from_bytes(&staged.to_bytes()?)?;
    if verified.body_chart_background_fill(drawable_object_id)? != ShapeFill::Image(image.clone()) {
        return Err(Error::InvalidFormat(
            "Pages chart background image-fill update failed validation".to_owned(),
        ));
    }
    *editor = verified;
    Ok(image)
}
