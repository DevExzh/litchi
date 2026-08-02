//! Contiguous private-object identifier allocation for source-built charts.

use super::*;

#[derive(Debug, Clone, Copy)]
pub(crate) struct SourceChartObjectIds {
    pub(crate) drawable: u64,
    pub(crate) caption: u64,
    pub(crate) title: u64,
    pub(crate) mediator: Option<u64>,
    pub(crate) preset: u64,
    pub(crate) chart_style: u64,
    pub(crate) chart_non_style: u64,
    pub(crate) legend_style: u64,
    pub(crate) legend_non_style: u64,
    pub(crate) value_axis_styles: [u64; VALUE_AXIS_COUNT],
    pub(crate) value_axis_non_styles: [u64; VALUE_AXIS_COUNT],
    pub(crate) category_axis_style: u64,
    pub(crate) category_axis_non_style: u64,
    pub(crate) series_styles: [u64; SERIES_STYLE_COUNT],
}

impl SourceChartObjectIds {
    pub(crate) fn allocate(first: u64, profile: ChartApplicationProfile) -> Result<Self> {
        let mut next = first;
        Ok(Self {
            drawable: take_identifier(&mut next)?,
            caption: take_identifier(&mut next)?,
            title: take_identifier(&mut next)?,
            mediator: profile
                .uses_mediator()
                .then(|| take_identifier(&mut next))
                .transpose()?,
            preset: take_identifier(&mut next)?,
            chart_style: take_identifier(&mut next)?,
            chart_non_style: take_identifier(&mut next)?,
            legend_style: take_identifier(&mut next)?,
            legend_non_style: take_identifier(&mut next)?,
            value_axis_styles: [take_identifier(&mut next)?, take_identifier(&mut next)?],
            value_axis_non_styles: [take_identifier(&mut next)?, take_identifier(&mut next)?],
            category_axis_style: take_identifier(&mut next)?,
            category_axis_non_style: take_identifier(&mut next)?,
            series_styles: [
                take_identifier(&mut next)?,
                take_identifier(&mut next)?,
                take_identifier(&mut next)?,
                take_identifier(&mut next)?,
                take_identifier(&mut next)?,
                take_identifier(&mut next)?,
            ],
        })
    }

    pub(crate) fn all(self) -> Vec<u64> {
        let mut identifiers = Vec::with_capacity(if self.mediator.is_some() { 21 } else { 20 });
        identifiers.extend([self.drawable, self.caption, self.title]);
        identifiers.extend(self.mediator);
        identifiers.extend([
            self.preset,
            self.chart_style,
            self.chart_non_style,
            self.legend_style,
            self.legend_non_style,
        ]);
        identifiers.extend(self.value_axis_styles);
        identifiers.extend(self.value_axis_non_styles);
        identifiers.extend([self.category_axis_style, self.category_axis_non_style]);
        identifiers.extend(self.series_styles);
        identifiers
    }

    pub(crate) fn style_ids(self) -> Vec<u64> {
        let mut identifiers = Vec::with_capacity(16);
        identifiers.extend([
            self.chart_style,
            self.chart_non_style,
            self.legend_style,
            self.legend_non_style,
        ]);
        identifiers.extend(self.value_axis_styles);
        identifiers.extend(self.value_axis_non_styles);
        identifiers.extend([self.category_axis_style, self.category_axis_non_style]);
        identifiers.extend(self.series_styles);
        identifiers
    }

    pub(crate) const fn last(self) -> u64 {
        self.series_styles[SERIES_STYLE_COUNT - 1]
    }

    pub(super) fn chart_references(self) -> Vec<u64> {
        let mut references = Vec::with_capacity(self.all().len() - 1);
        references.extend([self.caption, self.title]);
        references.extend(self.mediator);
        references.extend([
            self.preset,
            self.chart_style,
            self.chart_non_style,
            self.legend_style,
            self.legend_non_style,
        ]);
        references.extend(self.value_axis_styles);
        references.extend(self.value_axis_non_styles);
        references.extend([self.category_axis_style, self.category_axis_non_style]);
        references.extend(self.series_styles);
        references
    }
}

fn take_identifier(next: &mut u64) -> Result<u64> {
    let identifier = *next;
    *next = next
        .checked_add(1)
        .ok_or_else(|| Error::ParseError("iWork object identifier overflow".to_owned()))?;
    Ok(identifier)
}
