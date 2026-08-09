//! Worksheet-level BIFF8 HLink/HLinkTooltip record linkage.
//!
//! A tooltip is associated only with the immediately preceding hyperlink
//! and only when its checked `Ref8U` range matches. Targets remain inert.

use super::codec::{invalid, parse_hlink_record, parse_tooltip};
use super::model::{Hyperlink, RECORD_TYPE, TOOLTIP_RECORD_TYPE};
use crate::error::{Error, Result};

#[derive(Debug, Default)]
pub(crate) struct HyperlinkCollector {
    hyperlinks: Vec<Hyperlink>,
    pending_tooltip_index: Option<usize>,
}
impl HyperlinkCollector {
    pub(crate) fn new() -> Self {
        Self::default()
    }
    pub(crate) fn feed_record(&mut self, record_type: u16, data: &[u8]) -> Result<()> {
        if record_type == TOOLTIP_RECORD_TYPE {
            let index = self.pending_tooltip_index.take().ok_or_else(|| {
                Error::InvalidData("HLinkTooltip must immediately follow an HLink".to_string())
            })?;
            let (range, tooltip) = parse_tooltip(data)?;
            if self.hyperlinks[index].range != range {
                return invalid(
                    "HLinkTooltip range does not match its preceding HLink".to_string(),
                );
            }
            self.hyperlinks[index].tooltip = Some(tooltip);
            return Ok(());
        }
        self.pending_tooltip_index = None;
        if record_type == RECORD_TYPE {
            self.hyperlinks.push(parse_hlink_record(data)?);
            self.pending_tooltip_index = Some(self.hyperlinks.len() - 1);
        }
        Ok(())
    }
    pub(crate) fn finish(self) -> Vec<Hyperlink> {
        self.hyperlinks
    }
}
