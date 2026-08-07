//! Native explicit paragraph tab-stop conversion and canonical-wire checks.

use crate::protobuf::tswp;
use crate::text::paragraph_tabs::{
    ParagraphTabAlignment, ParagraphTabLeader, ParagraphTabPosition, ParagraphTabStop,
    ParagraphTabStops, alignment_from_native, alignment_to_native,
};
use crate::wire::repeated_length_delimited_payloads;
use crate::{Error, Result};

use super::has_exact_fields;

const TABS_TAB_FIELD: u32 = 1;
const TAB_POSITION_FIELD: u32 = 1;
const TAB_ALIGNMENT_FIELD: u32 = 2;
const TAB_LEADER_FIELD: u32 = 3;

pub(super) fn from_archive(archive: &tswp::TabsArchive) -> Result<ParagraphTabStops> {
    let mut stops = Vec::with_capacity(archive.tabs.len());
    for tab in &archive.tabs {
        let position = tab.position.ok_or_else(|| {
            Error::InvalidFormat("native iWork tab stop has no position".to_owned())
        })?;
        let alignment = alignment_from_native(
            tab.alignment
                .unwrap_or(alignment_to_native(ParagraphTabAlignment::Left)),
        )?;
        let leader = tab
            .leader
            .as_deref()
            .map(ParagraphTabLeader::new)
            .transpose()?;
        stops.push(ParagraphTabStop {
            position: ParagraphTabPosition::from_points(position)?,
            alignment,
            leader,
        });
    }
    Ok(ParagraphTabStops::new(stops)?)
}

pub(super) fn archive(stops: &ParagraphTabStops) -> tswp::TabsArchive {
    tswp::TabsArchive {
        tabs: stops
            .as_slice()
            .iter()
            .map(|stop| tswp::TabArchive {
                position: Some(stop.position.points()),
                alignment: (stop.alignment != ParagraphTabAlignment::Left)
                    .then(|| alignment_to_native(stop.alignment)),
                leader: stop
                    .leader
                    .as_ref()
                    .map(|leader| leader.as_str().to_owned()),
            })
            .collect(),
    }
}

pub(super) fn has_canonical_wire(data: &[u8], stops: &ParagraphTabStops) -> Result<bool> {
    let payloads = repeated_length_delimited_payloads(data, TABS_TAB_FIELD)?;
    if payloads.len() != stops.as_slice().len()
        || !has_exact_fields(data, &vec![TABS_TAB_FIELD; payloads.len()])?
    {
        return Ok(false);
    }
    for (payload, stop) in payloads.into_iter().zip(stops.as_slice()) {
        let mut expected = vec![TAB_POSITION_FIELD];
        if stop.alignment != ParagraphTabAlignment::Left {
            expected.push(TAB_ALIGNMENT_FIELD);
        }
        if stop.leader.is_some() {
            expected.push(TAB_LEADER_FIELD);
        }
        if !has_exact_fields(payload, &expected)? {
            return Ok(false);
        }
    }
    Ok(true)
}

#[cfg(test)]
mod tests {
    use prost::Message;

    use super::*;

    #[test]
    fn every_native_alignment_and_leader_round_trips_canonically() {
        let stops = ParagraphTabStops::new(
            [
                ParagraphTabAlignment::Left,
                ParagraphTabAlignment::Center,
                ParagraphTabAlignment::Right,
                ParagraphTabAlignment::Decimal,
            ]
            .into_iter()
            .enumerate()
            .map(|(index, alignment)| {
                let stop = ParagraphTabStop::new(
                    ParagraphTabPosition::from_points((index as f32 + 1.0) * 12.0).unwrap(),
                    alignment,
                );
                if alignment == ParagraphTabAlignment::Decimal {
                    stop.with_leader(ParagraphTabLeader::new(".").unwrap())
                } else {
                    stop
                }
            })
            .collect(),
        )
        .unwrap();
        let archive = archive(&stops);
        let encoded = archive.encode_to_vec();
        assert_eq!(from_archive(&archive).unwrap(), stops);
        assert!(has_canonical_wire(&encoded, &stops).unwrap());
    }

    #[test]
    fn malformed_and_noncanonical_native_tabs_are_rejected() {
        let missing_position = tswp::TabsArchive {
            tabs: vec![tswp::TabArchive {
                alignment: Some(alignment_to_native(ParagraphTabAlignment::Center)),
                ..Default::default()
            }],
        };
        assert!(from_archive(&missing_position).is_err());

        let explicit_left = tswp::TabsArchive {
            tabs: vec![tswp::TabArchive {
                position: Some(12.0),
                alignment: Some(alignment_to_native(ParagraphTabAlignment::Left)),
                ..Default::default()
            }],
        };
        let stops = from_archive(&explicit_left).unwrap();
        assert!(!has_canonical_wire(&explicit_left.encode_to_vec(), &stops).unwrap());

        let descending = tswp::TabsArchive {
            tabs: vec![
                tswp::TabArchive {
                    position: Some(24.0),
                    ..Default::default()
                },
                tswp::TabArchive {
                    position: Some(12.0),
                    ..Default::default()
                },
            ],
        };
        assert!(from_archive(&descending).is_err());
    }
}
