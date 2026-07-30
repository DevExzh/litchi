//! Native 3D lighting-style CRUD for Numbers sheet charts.

use super::*;
use crate::charts::Chart3dLightingStyle;
use crate::charts::lighting_3d::{
    chart_3d_lighting_style as read_native_chart_3d_lighting_style,
    set_chart_3d_lighting_style as set_native_chart_3d_lighting_style,
};

impl NumbersEditor {
    /// Read one sheet chart's native Lighting Style choice.
    pub fn sheet_chart_3d_lighting_style(
        &self,
        sheet_id: u64,
        drawable_object_id: u64,
    ) -> Result<Chart3dLightingStyle> {
        let graph = chart_graph(self, sheet_id, drawable_object_id)?;
        read_native_chart_3d_lighting_style(
            self.package(),
            &graph.archive_name,
            drawable_object_id,
            "Numbers",
            graph.info.kind,
        )
    }

    /// Set one sheet chart's native Lighting Style choice.
    pub fn set_sheet_chart_3d_lighting_style(
        &mut self,
        sheet_id: u64,
        drawable_object_id: u64,
        style: Chart3dLightingStyle,
    ) -> Result<()> {
        let graph = chart_graph(self, sheet_id, drawable_object_id)?;
        if read_native_chart_3d_lighting_style(
            self.package(),
            &graph.archive_name,
            drawable_object_id,
            "Numbers",
            graph.info.kind,
        )? == style
        {
            return Ok(());
        }
        let mut staged = self.package().clone();
        set_native_chart_3d_lighting_style(
            &mut staged,
            &graph.archive_name,
            drawable_object_id,
            "Numbers",
            graph.info.kind,
            style,
        )?;
        let verified = NumbersEditor::from_bytes(&staged.to_bytes()?)?;
        if verified.sheet_chart_3d_lighting_style(sheet_id, drawable_object_id)? != style {
            return Err(Error::InvalidFormat(
                "Numbers chart 3D lighting-style update failed validation".to_owned(),
            ));
        }
        *self = verified;
        Ok(())
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::numbers::NumbersDocumentBuilder;
    use crate::shapes::{DrawablePoint, DrawableSize};

    #[test]
    fn duplicated_sheet_charts_have_copy_on_write_lighting_styles() {
        let mut editor = NumbersDocumentBuilder::new().build().unwrap();
        let sheet_id = editor.sheets().unwrap()[0].object_id;
        let chart = editor
            .add_sheet_chart(
                sheet_id,
                ChartKind::Bar3d,
                data(),
                DrawablePoint { x: 20.0, y: 20.0 },
                DrawableSize {
                    width: 400.0,
                    height: 300.0,
                },
            )
            .unwrap();
        editor
            .set_sheet_chart_3d_lighting_style(
                sheet_id,
                chart.drawable_object_id,
                Chart3dLightingStyle::Glossy,
            )
            .unwrap();
        let duplicate = editor
            .duplicate_sheet_chart(sheet_id, chart.drawable_object_id)
            .unwrap();
        editor
            .set_sheet_chart_3d_lighting_style(
                sheet_id,
                duplicate.drawable_object_id,
                Chart3dLightingStyle::SoftFill,
            )
            .unwrap();
        assert_eq!(
            editor
                .sheet_chart_3d_lighting_style(sheet_id, chart.drawable_object_id)
                .unwrap(),
            Chart3dLightingStyle::Glossy
        );
        assert_eq!(
            editor
                .sheet_chart_3d_lighting_style(sheet_id, duplicate.drawable_object_id)
                .unwrap(),
            Chart3dLightingStyle::SoftFill
        );
    }

    fn data() -> ChartData {
        ChartData::new(
            vec!["Series".to_owned()],
            vec!["A".to_owned(), "B".to_owned()],
            vec![vec![Some(1.0), Some(2.0)]],
        )
        .unwrap()
    }
}
