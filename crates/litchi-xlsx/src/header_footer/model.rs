//! Semantic worksheet header/footer values.

/// A logical left, center, or right header/footer section.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum SectionKind {
    Left,
    Center,
    Right,
}

/// Header/footer text with its unambiguous alignment sections extracted.
///
/// Formatting and field control codes inside each section are intentionally
/// preserved. They can be localized, and OOXML requires their text to remain
/// available even when an application does not interpret the formatting.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Text {
    pub(crate) raw: String,
    pub(crate) left: Option<String>,
    pub(crate) center: Option<String>,
    pub(crate) right: Option<String>,
}

impl Text {
    pub(crate) fn new(raw: String) -> Self {
        let (left, center, right) = split_sections(&raw);
        Self {
            raw,
            left,
            center,
            right,
        }
    }

    /// Complete decoded text, including alignment and formatting controls.
    pub fn raw(&self) -> &str {
        &self.raw
    }

    /// Content of a logical alignment section, excluding its alignment marker.
    pub fn section(&self, kind: SectionKind) -> Option<&str> {
        match kind {
            SectionKind::Left => self.left.as_deref(),
            SectionKind::Center => self.center.as_deref(),
            SectionKind::Right => self.right.as_deref(),
        }
    }

    pub fn left(&self) -> Option<&str> {
        self.left.as_deref()
    }

    pub fn center(&self) -> Option<&str> {
        self.center.as_deref()
    }

    pub fn right(&self) -> Option<&str> {
        self.right.as_deref()
    }
}

/// Complete core headerFooter settings for one worksheet.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Settings {
    pub(crate) different_odd_even: bool,
    pub(crate) different_first: bool,
    pub(crate) scale_with_document: bool,
    pub(crate) align_with_margins: bool,
    pub(crate) odd_header: Option<Text>,
    pub(crate) odd_footer: Option<Text>,
    pub(crate) even_header: Option<Text>,
    pub(crate) even_footer: Option<Text>,
    pub(crate) first_header: Option<Text>,
    pub(crate) first_footer: Option<Text>,
}

impl Settings {
    pub fn different_odd_even(&self) -> bool {
        self.different_odd_even
    }

    pub fn different_first(&self) -> bool {
        self.different_first
    }

    pub fn scale_with_document(&self) -> bool {
        self.scale_with_document
    }

    pub fn align_with_margins(&self) -> bool {
        self.align_with_margins
    }

    pub fn odd_header(&self) -> Option<&Text> {
        self.odd_header.as_ref()
    }

    pub fn odd_footer(&self) -> Option<&Text> {
        self.odd_footer.as_ref()
    }

    pub fn even_header(&self) -> Option<&Text> {
        self.even_header.as_ref()
    }

    pub fn even_footer(&self) -> Option<&Text> {
        self.even_footer.as_ref()
    }

    pub fn first_header(&self) -> Option<&Text> {
        self.first_header.as_ref()
    }

    pub fn first_footer(&self) -> Option<&Text> {
        self.first_footer.as_ref()
    }
}

impl Default for Settings {
    fn default() -> Self {
        Self {
            different_odd_even: false,
            different_first: false,
            scale_with_document: true,
            align_with_margins: true,
            odd_header: None,
            odd_footer: None,
            even_header: None,
            even_footer: None,
            first_header: None,
            first_footer: None,
        }
    }
}

fn split_sections(raw: &str) -> (Option<String>, Option<String>, Option<String>) {
    let mut sections = [None::<String>, None::<String>, None::<String>];
    let mut current = 1usize;
    let mut index = 0usize;
    while index < raw.len() {
        let tail = &raw[index..];
        if let Some(marker) = tail
            .as_bytes()
            .get(1)
            .copied()
            .filter(|_| tail.as_bytes()[0] == b'&')
        {
            if marker == b'&' {
                sections[current]
                    .get_or_insert_with(String::new)
                    .push_str("&&");
                index += 2;
                continue;
            }
            if let Some(next) = match marker {
                b'L' => Some(0),
                b'C' => Some(1),
                b'R' => Some(2),
                _ => None,
            } {
                current = next;
                sections[current].get_or_insert_with(String::new);
                index += 2;
                continue;
            }
        }
        let character = tail.chars().next().expect("non-empty tail");
        sections[current]
            .get_or_insert_with(String::new)
            .push(character);
        index += character.len_utf8();
    }
    if raw.is_empty() {
        sections[1] = Some(String::new());
    }
    (sections[0].take(), sections[1].take(), sections[2].take())
}
