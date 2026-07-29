//! Native legend visibility CRUD for Pages body charts.

use super::*;
use crate::charts::legend_fill::{
    chart_legend_fill as read_native_chart_legend_fill,
    set_chart_legend_fill as set_native_chart_legend_fill,
};
use crate::charts::legend_stroke::{
    chart_legend_stroke as read_native_chart_legend_stroke,
    set_chart_legend_stroke as set_native_chart_legend_stroke,
};
use crate::charts::options::{
    chart_legend_visible as read_native_chart_legend_visible,
    set_chart_legend_visible as set_native_chart_legend_visible,
};
use crate::charts::{ChartLegendFill, ChartLegendStroke};

impl PagesEditor {
    /// Read whether Pages shows the native legend for one body chart.
    pub fn body_chart_legend_visible(&self, drawable_object_id: u64) -> Result<bool> {
        body_chart_legend_visible(self, drawable_object_id)
    }

    /// Set whether Pages shows the native legend for one body chart.
    pub fn set_body_chart_legend_visible(
        &mut self,
        drawable_object_id: u64,
        visible: bool,
    ) -> Result<()> {
        set_body_chart_legend_visible(self, drawable_object_id, visible)
    }

    /// Read the exact inherited or direct native legend fill.
    pub fn body_chart_legend_fill(&self, drawable_object_id: u64) -> Result<ChartLegendFill> {
        let graph = body_chart_graph(self, drawable_object_id)?;
        read_native_chart_legend_fill(
            self.package(),
            &graph.archive_name,
            drawable_object_id,
            "Pages",
        )
    }

    /// Set or remove the direct native legend-fill override.
    pub fn set_body_chart_legend_fill(
        &mut self,
        drawable_object_id: u64,
        fill: &ChartLegendFill,
    ) -> Result<()> {
        if &self.body_chart_legend_fill(drawable_object_id)? == fill {
            return Ok(());
        }
        let graph = body_chart_graph(self, drawable_object_id)?;
        let mut staged = self.package().clone();
        set_native_chart_legend_fill(
            &mut staged,
            &graph.archive_name,
            drawable_object_id,
            "Pages",
            fill,
        )?;
        let verified = PagesEditor::from_bytes(&staged.to_bytes()?)?;
        if verified.body_chart_legend_fill(drawable_object_id)? != *fill {
            return Err(Error::InvalidFormat(
                "Pages chart legend-fill update failed validation".to_owned(),
            ));
        }
        *self = verified;
        Ok(())
    }

    /// Read the exact inherited, empty, or direct native legend stroke.
    pub fn body_chart_legend_stroke(&self, drawable_object_id: u64) -> Result<ChartLegendStroke> {
        let graph = body_chart_graph(self, drawable_object_id)?;
        read_native_chart_legend_stroke(
            self.package(),
            &graph.archive_name,
            drawable_object_id,
            "Pages",
        )
    }

    /// Set or remove the direct native legend-stroke override.
    pub fn set_body_chart_legend_stroke(
        &mut self,
        drawable_object_id: u64,
        stroke: ChartLegendStroke,
    ) -> Result<()> {
        if self.body_chart_legend_stroke(drawable_object_id)? == stroke {
            return Ok(());
        }
        let graph = body_chart_graph(self, drawable_object_id)?;
        let mut staged = self.package().clone();
        set_native_chart_legend_stroke(
            &mut staged,
            &graph.archive_name,
            drawable_object_id,
            "Pages",
            stroke,
        )?;
        let verified = PagesEditor::from_bytes(&staged.to_bytes()?)?;
        if verified.body_chart_legend_stroke(drawable_object_id)? != stroke {
            return Err(Error::InvalidFormat(
                "Pages chart legend-stroke update failed validation".to_owned(),
            ));
        }
        *self = verified;
        Ok(())
    }
}

fn body_chart_legend_visible(editor: &PagesEditor, drawable_object_id: u64) -> Result<bool> {
    let graph = body_chart_graph(editor, drawable_object_id)?;
    read_native_chart_legend_visible(
        editor.package(),
        &graph.archive_name,
        drawable_object_id,
        "Pages",
    )
}

fn set_body_chart_legend_visible(
    editor: &mut PagesEditor,
    drawable_object_id: u64,
    visible: bool,
) -> Result<()> {
    if body_chart_legend_visible(editor, drawable_object_id)? == visible {
        return Ok(());
    }
    let graph = body_chart_graph(editor, drawable_object_id)?;
    let mut staged = editor.package().clone();
    set_native_chart_legend_visible(
        &mut staged,
        &graph.archive_name,
        drawable_object_id,
        "Pages",
        visible,
    )?;
    let verified = PagesEditor::from_bytes(&staged.to_bytes()?)?;
    if verified.body_chart_legend_visible(drawable_object_id)? != visible {
        return Err(Error::InvalidFormat(
            "Pages chart legend update failed validation".to_owned(),
        ));
    }
    *editor = verified;
    Ok(())
}
