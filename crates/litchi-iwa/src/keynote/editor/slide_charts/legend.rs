//! Native legend visibility CRUD for Keynote slide charts.

use super::*;
use crate::charts::legend_fill::{
    chart_legend_fill as read_native_chart_legend_fill,
    set_chart_legend_fill as set_native_chart_legend_fill,
};
use crate::charts::legend_font_size::{
    chart_legend_font_size as read_native_chart_legend_font_size,
    set_chart_legend_font_size as set_native_chart_legend_font_size,
};
use crate::charts::legend_shadow::{
    chart_legend_shadow as read_native_chart_legend_shadow,
    set_chart_legend_shadow as set_native_chart_legend_shadow,
};
use crate::charts::legend_stroke::{
    chart_legend_stroke as read_native_chart_legend_stroke,
    set_chart_legend_stroke as set_native_chart_legend_stroke,
};
use crate::charts::options::{
    chart_legend_visible as read_native_chart_legend_visible,
    set_chart_legend_visible as set_native_chart_legend_visible,
};
use crate::charts::{ChartLegendFill, ChartLegendFontSize, ChartLegendShadow, ChartLegendStroke};

impl KeynoteEditor {
    /// Read whether Keynote shows the native legend for one slide chart.
    pub fn slide_chart_legend_visible(
        &self,
        slide_index: usize,
        drawable_object_id: u64,
    ) -> Result<bool> {
        slide_chart_legend_visible(self, slide_index, drawable_object_id)
    }

    /// Set whether Keynote shows the native legend for one slide chart.
    pub fn set_slide_chart_legend_visible(
        &mut self,
        slide_index: usize,
        drawable_object_id: u64,
        visible: bool,
    ) -> Result<()> {
        set_slide_chart_legend_visible(self, slide_index, drawable_object_id, visible)
    }

    /// Read the exact inherited or direct native legend fill.
    pub fn slide_chart_legend_fill(
        &self,
        slide_index: usize,
        drawable_object_id: u64,
    ) -> Result<ChartLegendFill> {
        let graph = chart_graph(self, slide_index, drawable_object_id)?;
        read_native_chart_legend_fill(
            self.package(),
            &graph.archive_name,
            drawable_object_id,
            "Keynote",
        )
    }

    /// Set or remove the direct native legend-fill override.
    pub fn set_slide_chart_legend_fill(
        &mut self,
        slide_index: usize,
        drawable_object_id: u64,
        fill: &ChartLegendFill,
    ) -> Result<()> {
        if &self.slide_chart_legend_fill(slide_index, drawable_object_id)? == fill {
            return Ok(());
        }
        let graph = chart_graph(self, slide_index, drawable_object_id)?;
        let mut staged = self.package().clone();
        set_native_chart_legend_fill(
            &mut staged,
            &graph.archive_name,
            drawable_object_id,
            "Keynote",
            fill,
        )?;
        let verified = KeynoteEditor::from_bytes(&staged.to_bytes()?)?;
        if verified.slide_chart_legend_fill(slide_index, drawable_object_id)? != *fill {
            return Err(Error::InvalidFormat(
                "Keynote chart legend-fill update failed validation".to_owned(),
            ));
        }
        *self = verified;
        Ok(())
    }

    /// Read the exact inherited or direct native legend font size.
    pub fn slide_chart_legend_font_size(
        &self,
        slide_index: usize,
        drawable_object_id: u64,
    ) -> Result<ChartLegendFontSize> {
        let graph = chart_graph(self, slide_index, drawable_object_id)?;
        read_native_chart_legend_font_size(
            self.package(),
            &graph.archive_name,
            drawable_object_id,
            "Keynote",
        )
    }

    /// Set or remove the direct native legend font-size override.
    pub fn set_slide_chart_legend_font_size(
        &mut self,
        slide_index: usize,
        drawable_object_id: u64,
        size: ChartLegendFontSize,
    ) -> Result<()> {
        if self.slide_chart_legend_font_size(slide_index, drawable_object_id)? == size {
            return Ok(());
        }
        let graph = chart_graph(self, slide_index, drawable_object_id)?;
        let mut staged = self.package().clone();
        set_native_chart_legend_font_size(
            &mut staged,
            &graph.archive_name,
            drawable_object_id,
            "Keynote",
            size,
        )?;
        let verified = KeynoteEditor::from_bytes(&staged.to_bytes()?)?;
        if verified.slide_chart_legend_font_size(slide_index, drawable_object_id)? != size {
            return Err(Error::InvalidFormat(
                "Keynote chart legend font-size update failed validation".to_owned(),
            ));
        }
        *self = verified;
        Ok(())
    }

    /// Read the exact inherited, empty, or direct native legend stroke.
    pub fn slide_chart_legend_stroke(
        &self,
        slide_index: usize,
        drawable_object_id: u64,
    ) -> Result<ChartLegendStroke> {
        let graph = chart_graph(self, slide_index, drawable_object_id)?;
        read_native_chart_legend_stroke(
            self.package(),
            &graph.archive_name,
            drawable_object_id,
            "Keynote",
        )
    }

    /// Set or remove the direct native legend-stroke override.
    pub fn set_slide_chart_legend_stroke(
        &mut self,
        slide_index: usize,
        drawable_object_id: u64,
        stroke: ChartLegendStroke,
    ) -> Result<()> {
        if self.slide_chart_legend_stroke(slide_index, drawable_object_id)? == stroke {
            return Ok(());
        }
        let graph = chart_graph(self, slide_index, drawable_object_id)?;
        let mut staged = self.package().clone();
        set_native_chart_legend_stroke(
            &mut staged,
            &graph.archive_name,
            drawable_object_id,
            "Keynote",
            stroke,
        )?;
        let verified = KeynoteEditor::from_bytes(&staged.to_bytes()?)?;
        if verified.slide_chart_legend_stroke(slide_index, drawable_object_id)? != stroke {
            return Err(Error::InvalidFormat(
                "Keynote chart legend-stroke update failed validation".to_owned(),
            ));
        }
        *self = verified;
        Ok(())
    }

    /// Read the exact inherited, disabled, or direct native legend shadow.
    pub fn slide_chart_legend_shadow(
        &self,
        slide_index: usize,
        drawable_object_id: u64,
    ) -> Result<ChartLegendShadow> {
        let graph = chart_graph(self, slide_index, drawable_object_id)?;
        read_native_chart_legend_shadow(
            self.package(),
            &graph.archive_name,
            drawable_object_id,
            "Keynote",
        )
    }

    /// Set or remove the direct native legend-shadow override.
    pub fn set_slide_chart_legend_shadow(
        &mut self,
        slide_index: usize,
        drawable_object_id: u64,
        shadow: ChartLegendShadow,
    ) -> Result<()> {
        if self.slide_chart_legend_shadow(slide_index, drawable_object_id)? == shadow {
            return Ok(());
        }
        let graph = chart_graph(self, slide_index, drawable_object_id)?;
        let mut staged = self.package().clone();
        set_native_chart_legend_shadow(
            &mut staged,
            &graph.archive_name,
            drawable_object_id,
            "Keynote",
            shadow,
        )?;
        let verified = KeynoteEditor::from_bytes(&staged.to_bytes()?)?;
        if verified.slide_chart_legend_shadow(slide_index, drawable_object_id)? != shadow {
            return Err(Error::InvalidFormat(
                "Keynote chart legend-shadow update failed validation".to_owned(),
            ));
        }
        *self = verified;
        Ok(())
    }
}

fn slide_chart_legend_visible(
    editor: &KeynoteEditor,
    slide_index: usize,
    drawable_object_id: u64,
) -> Result<bool> {
    let graph = chart_graph(editor, slide_index, drawable_object_id)?;
    read_native_chart_legend_visible(
        editor.package(),
        &graph.archive_name,
        drawable_object_id,
        "Keynote",
    )
}

fn set_slide_chart_legend_visible(
    editor: &mut KeynoteEditor,
    slide_index: usize,
    drawable_object_id: u64,
    visible: bool,
) -> Result<()> {
    if slide_chart_legend_visible(editor, slide_index, drawable_object_id)? == visible {
        return Ok(());
    }
    let graph = chart_graph(editor, slide_index, drawable_object_id)?;
    let mut staged = editor.package().clone();
    set_native_chart_legend_visible(
        &mut staged,
        &graph.archive_name,
        drawable_object_id,
        "Keynote",
        visible,
    )?;
    let verified = KeynoteEditor::from_bytes(&staged.to_bytes()?)?;
    if verified.slide_chart_legend_visible(slide_index, drawable_object_id)? != visible {
        return Err(Error::InvalidFormat(
            "Keynote chart legend update failed validation".to_owned(),
        ));
    }
    *editor = verified;
    Ok(())
}
