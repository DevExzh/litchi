//! Custom slide show (named show) support for PPT files.
//!
//! Implements NamedShows/NamedShow/NamedShowSlides records per [MS-PPT].
//! Custom shows are stored inside the DocumentContainer.
//!
use super::records::Error;
use crate::{NamedShow, NamedShows};
use std::io::ErrorKind;

/// A custom (named) slide show definition.
///
/// Custom shows allow defining named subsets of slides that can be
/// presented independently from the main slide deck.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct CustomShow {
    /// Show name. Validation occurs during serialization.
    pub name: String,
    /// Slide indices (0-based) included in this show, in presentation order.
    pub slide_indices: Vec<usize>,
}

impl CustomShow {
    /// Create a new custom show with the given name and slide indices.
    ///
    /// # Arguments
    ///
    /// * `name` - Show name
    /// * `slide_indices` - 0-based slide indices in presentation order
    ///
    /// # Example
    ///
    /// ```
    /// use litchi_ppt::writer::custom_shows::CustomShow;
    /// // Create a custom show with slides 0, 2, and 4
    /// let show = CustomShow::new("Executive Summary", &[0, 2, 4]);
    /// ```
    #[must_use]
    pub fn new(name: &str, slide_indices: &[usize]) -> Self {
        Self {
            name: name.to_string(),
            slide_indices: slide_indices.to_vec(),
        }
    }
}

/// Build the `NamedShows` container for the Document.
///
/// Returns the serialized `NamedShows` container bytes, or an empty Vec if
/// there are no custom shows.
///
/// # Arguments
///
/// * `shows` - Slice of custom show definitions
///
/// # Errors
///
/// Returns an error if serialization fails or the underlying writer reports an error.
pub fn build_named_shows(shows: &[CustomShow]) -> Result<Vec<u8>, Error> {
    if shows.is_empty() {
        return Ok(Vec::new());
    }
    let checked_shows = shows
        .iter()
        .map(checked_named_show)
        .collect::<Result<Vec<_>, _>>()?;
    build_named_shows_typed(&NamedShows {
        shows: checked_shows,
    })
}

/// Serialize the strict typed named-show model, including an empty container.
///
/// # Errors
///
/// Returns an error if serialization fails or the underlying writer reports an error.
pub fn build_named_shows_typed(shows: &NamedShows) -> Result<Vec<u8>, Error> {
    shows
        .to_record_bytes()
        .map_err(|error| Error::new(ErrorKind::InvalidData, error.to_string()))
}

fn checked_named_show(show: &CustomShow) -> Result<NamedShow, Error> {
    let mut slide_ids = Vec::with_capacity(show.slide_indices.len());
    for &index in &show.slide_indices {
        let index_u32 = u32::try_from(index).map_err(|_err| {
            Error::new(
                ErrorKind::InvalidInput,
                "custom-show slide index exceeds u32",
            )
        })?;
        let slide_id = index_u32
            .checked_add(0x100)
            .filter(|id| *id <= 0x7fff_ffff)
            .ok_or_else(|| {
                Error::new(
                    ErrorKind::InvalidInput,
                    "custom-show slide index exceeds the SlideId range",
                )
            })?;
        slide_ids.push(slide_id);
    }
    Ok(NamedShow {
        name: show.name.clone(),
        slide_ids: Some(slide_ids),
    })
}

#[cfg(test)]
#[allow(
    clippy::unwrap_used,
    clippy::expect_used,
    reason = "test assertions panic on failure by design"
)]
mod tests {
    use super::*;
    use crate::consts::RecordType;
    use crate::records::Record;

    fn root(child: Record) -> Record {
        Record {
            version: 0x0f,
            instance: 0,
            record_type: RecordType::Document,
            record_type_raw: RecordType::Document.as_u16(),
            data_length: 0,
            data: Vec::new(),
            children: vec![child],
        }
    }

    fn parse(data: &[u8]) -> NamedShows {
        let record = Record::parse(data, 0).unwrap().0;
        NamedShows::parse(&root(record)).unwrap().unwrap()
    }

    #[test]
    fn test_build_named_shows_empty() {
        let data = build_named_shows(&[]).unwrap();
        assert!(data.is_empty());
    }

    #[test]
    fn test_build_named_shows_single() {
        let shows = vec![CustomShow::new("Test Show", &[0, 1, 2])];
        let data = build_named_shows(&shows).unwrap();
        assert!(!data.is_empty());

        // Verify outer container type = 1040 (NamedShows)
        let rtype = u16::from_le_bytes([data[2], data[3]]);
        assert_eq!(rtype, 1040);
        assert_eq!(
            parse(&data).shows[0].slide_ids.as_deref(),
            Some(&[0x100, 0x101, 0x102][..])
        );
    }

    #[test]
    fn test_build_named_shows_multiple() {
        let shows = vec![
            CustomShow::new("Short Version", &[0, 3]),
            CustomShow::new("Full Version", &[0, 1, 2, 3, 4]),
        ];
        let data = build_named_shows(&shows).unwrap();
        assert!(!data.is_empty());

        // Verify outer container type = 1040 (NamedShows)
        let rtype = u16::from_le_bytes([data[2], data[3]]);
        assert_eq!(rtype, 1040);
        let parsed = parse(&data);
        assert_eq!(parsed.shows.len(), 2);
        assert_eq!(parsed.shows[0].name, "Short Version");
        assert!(parsed.shows.iter().all(|show| !show.name.is_empty()));
    }

    #[test]
    fn test_slide_id_encoding() {
        let show = CustomShow::new("Test", &[0, 5, 10]);
        let data = build_named_shows(&[show]).unwrap();
        assert!(!data.is_empty());
        assert_eq!(
            parse(&data).shows[0].slide_ids.as_deref(),
            Some(&[0x100, 0x105, 0x10a][..])
        );
    }

    #[test]
    fn strict_writer_preserves_long_names_and_zero_instances() {
        let name = "A named show longer than thirty-one Unicode characters";
        let data =
            build_named_shows(&[CustomShow::new(name, &[0]), CustomShow::new("Second", &[1])])
                .unwrap();
        let outer = Record::parse(&data, 0).unwrap().0;
        assert!(outer.children.iter().all(|show| show.instance == 0));
        assert_eq!(parse(&data).shows[0].name, name);
    }

    #[test]
    fn rejects_non_printable_names_and_overflowing_indices() {
        assert!(build_named_shows(&[CustomShow::new("bad\nname", &[0])]).is_err());
        assert!(build_named_shows(&[CustomShow::new("overflow", &[usize::MAX])]).is_err());
        let empty = NamedShows::default();
        assert!(!build_named_shows_typed(&empty).unwrap().is_empty());
        assert!(
            parse(&build_named_shows_typed(&empty).unwrap())
                .shows
                .is_empty()
        );
    }
}
