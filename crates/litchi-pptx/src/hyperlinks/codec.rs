//! Bounded parsing for inert `PresentationML` hyperlink target strings.

use super::model::Hyperlink;
use crate::{Error, Result};

const MAX_TARGET_BYTES: usize = 1 << 20;
const MAX_TOOLTIP_BYTES: usize = 1 << 20;

pub(super) fn parse(target: &str, tooltip: Option<String>) -> Result<Hyperlink> {
    if target.len() > MAX_TARGET_BYTES {
        return Err(Error::Limit {
            resource: "hyperlink target bytes",
            limit: MAX_TARGET_BYTES,
        });
    }
    if tooltip
        .as_deref()
        .is_some_and(|value| value.len() > MAX_TOOLTIP_BYTES)
    {
        return Err(Error::Limit {
            resource: "hyperlink tooltip bytes",
            limit: MAX_TOOLTIP_BYTES,
        });
    }

    if target.starts_with("http://") || target.starts_with("https://") {
        return Ok(Hyperlink::External {
            url: target.to_owned(),
            tooltip,
        });
    }
    if let Some(value) = target.strip_prefix("ppaction://hlinksldjump") {
        let slide_number = value
            .split_once("sldNum=")
            .and_then(|(_, value)| value.split('&').next())
            .and_then(|value| value.parse().ok())
            .unwrap_or(1);
        return Ok(Hyperlink::Slide {
            slide_number,
            tooltip,
        });
    }
    if let Some(value) = target.strip_prefix("mailto:") {
        let (email, query) = value
            .split_once('?')
            .map_or((value, None), |(email, query)| (email, Some(query)));
        let subject = query.and_then(|query| {
            query
                .split('&')
                .find_map(|part| part.strip_prefix("subject=").map(str::to_owned))
        });
        return Ok(Hyperlink::Email {
            email: email.to_owned(),
            subject,
            tooltip,
        });
    }

    Ok(Hyperlink::External {
        url: target.to_owned(),
        tooltip,
    })
}
