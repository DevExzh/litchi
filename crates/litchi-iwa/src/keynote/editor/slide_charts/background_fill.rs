//! Native chart-background fill CRUD for Keynote slide charts.

use super::*;
use crate::charts::background_fill::{
    chart_background_fill as read_native_chart_background_fill,
    set_chart_background_fill as set_native_chart_background_fill,
    set_chart_background_image_fill_data,
};
use crate::shapes::{RgbaColor, ShapeFill, ShapeImageFill, ShapeImageFillTechnique};

impl KeynoteEditor {
    /// Read the effective color, gradient, or image background of one slide chart.
    pub fn slide_chart_background_fill(
        &self,
        slide_index: usize,
        drawable_object_id: u64,
    ) -> Result<ShapeFill> {
        slide_chart_background_fill(self, slide_index, drawable_object_id)
    }

    /// Replace one slide chart's background fill transactionally.
    pub fn set_slide_chart_background_fill(
        &mut self,
        slide_index: usize,
        drawable_object_id: u64,
        fill: &ShapeFill,
    ) -> Result<()> {
        set_slide_chart_background_fill(self, slide_index, drawable_object_id, fill)
    }

    /// Embed image bytes and use them as a simple or tinted chart background.
    pub fn set_slide_chart_background_image_fill(
        &mut self,
        slide_index: usize,
        drawable_object_id: u64,
        preferred_filename: &str,
        data: &[u8],
        technique: ShapeImageFillTechnique,
        tint: Option<RgbaColor>,
    ) -> Result<ShapeImageFill> {
        set_slide_chart_background_image_fill(
            self,
            slide_index,
            drawable_object_id,
            preferred_filename,
            data,
            technique,
            tint,
        )
    }
}

fn slide_chart_background_fill(
    editor: &KeynoteEditor,
    slide_index: usize,
    drawable_object_id: u64,
) -> Result<ShapeFill> {
    let graph = chart_graph(editor, slide_index, drawable_object_id)?;
    read_native_chart_background_fill(
        editor.package(),
        &graph.archive_name,
        drawable_object_id,
        "Keynote",
    )
}

fn set_slide_chart_background_fill(
    editor: &mut KeynoteEditor,
    slide_index: usize,
    drawable_object_id: u64,
    fill: &ShapeFill,
) -> Result<()> {
    if &slide_chart_background_fill(editor, slide_index, drawable_object_id)? == fill {
        return Ok(());
    }
    let graph = chart_graph(editor, slide_index, drawable_object_id)?;
    let mut staged = editor.package().clone();
    set_native_chart_background_fill(
        &mut staged,
        &graph.archive_name,
        drawable_object_id,
        "Keynote",
        fill,
    )?;
    let verified = KeynoteEditor::from_bytes(&staged.to_bytes()?)?;
    if &verified.slide_chart_background_fill(slide_index, drawable_object_id)? != fill {
        return Err(Error::InvalidFormat(
            "Keynote chart background fill update failed validation".to_owned(),
        ));
    }
    *editor = verified;
    Ok(())
}

fn set_slide_chart_background_image_fill(
    editor: &mut KeynoteEditor,
    slide_index: usize,
    drawable_object_id: u64,
    preferred_filename: &str,
    data: &[u8],
    technique: ShapeImageFillTechnique,
    tint: Option<RgbaColor>,
) -> Result<ShapeImageFill> {
    let graph = chart_graph(editor, slide_index, drawable_object_id)?;
    let fill_size = graph.info.geometry.size.ok_or_else(|| {
        Error::InvalidFormat(format!(
            "Keynote chart {drawable_object_id} has no background image-fill dimensions"
        ))
    })?;
    let mut staged = editor.package().clone();
    let image = set_chart_background_image_fill_data(
        &mut staged,
        &graph.archive_name,
        drawable_object_id,
        "Keynote",
        preferred_filename,
        data,
        technique,
        fill_size,
        tint,
    )?;
    let verified = KeynoteEditor::from_bytes(&staged.to_bytes()?)?;
    if verified.slide_chart_background_fill(slide_index, drawable_object_id)?
        != ShapeFill::Image(image.clone())
    {
        return Err(Error::InvalidFormat(
            "Keynote chart background image-fill update failed validation".to_owned(),
        ));
    }
    *editor = verified;
    Ok(image)
}
