use super::model::{MAX_STOPS, Position};
use crate::prop::Array;
use crate::{Error, Result};

pub(super) fn validate(array: &Array<'_>) -> Result<()> {
    if usize::from(array.element_count()) > MAX_STOPS {
        return Err(Error::MalformedProperties {
            reason: "gradient stop count exceeds the safe bound",
        });
    }

    let mut previous = None;
    for element in array.elements() {
        let raw = i32::from_le_bytes(
            element
                .get(4..8)
                .ok_or(Error::MalformedProperties {
                    reason: "gradient stop element is truncated",
                })?
                .try_into()
                .map_err(|_| Error::MalformedProperties {
                    reason: "gradient stop element is not eight bytes",
                })?,
        );
        let current = Position::new(raw)
            .ok_or(Error::MalformedProperties {
                reason: "gradient stop position is outside the inclusive 0..1 range",
            })?
            .raw();
        if let Some(previous) = previous {
            if current <= previous {
                return Err(Error::MalformedProperties {
                    reason: "gradient stop positions are not strictly ascending",
                });
            }
        }
        previous = Some(current);
    }
    Ok(())
}
