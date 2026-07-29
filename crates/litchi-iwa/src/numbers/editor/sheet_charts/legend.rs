//! Native legend visibility CRUD for Numbers sheet charts.

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

impl NumbersEditor {
    /// Read whether Numbers shows the native legend for one sheet chart.
    pub fn sheet_chart_legend_visible(
        &self,
        sheet_id: u64,
        drawable_object_id: u64,
    ) -> Result<bool> {
        sheet_chart_legend_visible(self, sheet_id, drawable_object_id)
    }

    /// Set whether Numbers shows the native legend for one sheet chart.
    pub fn set_sheet_chart_legend_visible(
        &mut self,
        sheet_id: u64,
        drawable_object_id: u64,
        visible: bool,
    ) -> Result<()> {
        set_sheet_chart_legend_visible(self, sheet_id, drawable_object_id, visible)
    }

    /// Read the exact inherited or direct native legend fill.
    pub fn sheet_chart_legend_fill(
        &self,
        sheet_id: u64,
        drawable_object_id: u64,
    ) -> Result<ChartLegendFill> {
        let graph = chart_graph(self, sheet_id, drawable_object_id)?;
        read_native_chart_legend_fill(
            &self.package,
            &graph.archive_name,
            drawable_object_id,
            "Numbers",
        )
    }

    /// Set or remove the direct native legend-fill override.
    pub fn set_sheet_chart_legend_fill(
        &mut self,
        sheet_id: u64,
        drawable_object_id: u64,
        fill: &ChartLegendFill,
    ) -> Result<()> {
        if &self.sheet_chart_legend_fill(sheet_id, drawable_object_id)? == fill {
            return Ok(());
        }
        let graph = chart_graph(self, sheet_id, drawable_object_id)?;
        let mut staged = self.package.clone();
        set_native_chart_legend_fill(
            &mut staged,
            &graph.archive_name,
            drawable_object_id,
            "Numbers",
            fill,
        )?;
        let verified = NumbersEditor::from_bytes(&staged.to_bytes()?)?;
        if verified.sheet_chart_legend_fill(sheet_id, drawable_object_id)? != *fill {
            return Err(Error::InvalidFormat(
                "Numbers chart legend-fill update failed validation".to_owned(),
            ));
        }
        *self = verified;
        Ok(())
    }

    /// Read the exact inherited, empty, or direct native legend stroke.
    pub fn sheet_chart_legend_stroke(
        &self,
        sheet_id: u64,
        drawable_object_id: u64,
    ) -> Result<ChartLegendStroke> {
        let graph = chart_graph(self, sheet_id, drawable_object_id)?;
        read_native_chart_legend_stroke(
            &self.package,
            &graph.archive_name,
            drawable_object_id,
            "Numbers",
        )
    }

    /// Set or remove the direct native legend-stroke override.
    pub fn set_sheet_chart_legend_stroke(
        &mut self,
        sheet_id: u64,
        drawable_object_id: u64,
        stroke: ChartLegendStroke,
    ) -> Result<()> {
        if self.sheet_chart_legend_stroke(sheet_id, drawable_object_id)? == stroke {
            return Ok(());
        }
        let graph = chart_graph(self, sheet_id, drawable_object_id)?;
        let mut staged = self.package.clone();
        set_native_chart_legend_stroke(
            &mut staged,
            &graph.archive_name,
            drawable_object_id,
            "Numbers",
            stroke,
        )?;
        let verified = NumbersEditor::from_bytes(&staged.to_bytes()?)?;
        if verified.sheet_chart_legend_stroke(sheet_id, drawable_object_id)? != stroke {
            return Err(Error::InvalidFormat(
                "Numbers chart legend-stroke update failed validation".to_owned(),
            ));
        }
        *self = verified;
        Ok(())
    }
}

fn sheet_chart_legend_visible(
    editor: &NumbersEditor,
    sheet_id: u64,
    drawable_object_id: u64,
) -> Result<bool> {
    let graph = chart_graph(editor, sheet_id, drawable_object_id)?;
    read_native_chart_legend_visible(
        &editor.package,
        &graph.archive_name,
        drawable_object_id,
        "Numbers",
    )
}

fn set_sheet_chart_legend_visible(
    editor: &mut NumbersEditor,
    sheet_id: u64,
    drawable_object_id: u64,
    visible: bool,
) -> Result<()> {
    if sheet_chart_legend_visible(editor, sheet_id, drawable_object_id)? == visible {
        return Ok(());
    }
    let graph = chart_graph(editor, sheet_id, drawable_object_id)?;
    let mut staged = editor.package.clone();
    set_native_chart_legend_visible(
        &mut staged,
        &graph.archive_name,
        drawable_object_id,
        "Numbers",
        visible,
    )?;
    let verified = NumbersEditor::from_bytes(&staged.to_bytes()?)?;
    if verified.sheet_chart_legend_visible(sheet_id, drawable_object_id)? != visible {
        return Err(Error::InvalidFormat(
            "Numbers chart legend update failed validation".to_owned(),
        ));
    }
    *editor = verified;
    Ok(())
}
