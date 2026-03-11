//! Slide comment support for PPT files.
//!
//! Implements Comment2000 (EPP_Comment10) containers per [MS-PPT] Section 2.4.
//! Comments are stored inside `ProgTags/ProgBinaryTag/BinaryTagData` within each slide.
//!
//! # Binary Structure
//!
//! ```text
//! Comment2000 (container, type=12000)
//! ├── CString (instance=0): author name (UTF-16LE)
//! ├── CString (instance=1): comment text (UTF-16LE)
//! ├── CString (instance=2): author initials (UTF-16LE)
//! └── Comment2000Atom (type=12001, 28 bytes)
//!     ├── index (i32): 1-based comment index
//!     ├── year (i16)
//!     ├── month (u16)
//!     ├── day_of_week (u16): day of week (unused, set to day)
//!     ├── day (u16)
//!     ├── hour (u16)
//!     ├── minute (u16)
//!     ├── second (u16)
//!     ├── millisecond (i16)
//!     ├── x (i32): position in master units (576/inch)
//!     └── y (i32): position in master units
//! ```

use super::records::{PptError, RecordBuilder, record_type};

/// A single comment on a slide.
#[derive(Debug, Clone)]
pub struct SlideComment {
    /// Comment author name.
    pub author: String,
    /// Comment text content.
    pub text: String,
    /// Author initials (e.g. "JD" for "John Doe").
    pub initials: String,
    /// X position in points (72 points = 1 inch).
    pub x: i32,
    /// Y position in points.
    pub y: i32,
    /// Comment date/time.
    pub date: CommentDateTime,
}

/// Date and time for a comment.
#[derive(Debug, Clone, Default)]
pub struct CommentDateTime {
    /// Year (e.g. 2025).
    pub year: i16,
    /// Month (1-12).
    pub month: u16,
    /// Day of month (1-31).
    pub day: u16,
    /// Hour (0-23).
    pub hour: u16,
    /// Minute (0-59).
    pub minute: u16,
    /// Second (0-59).
    pub second: u16,
    /// Millisecond (0-999).
    pub millisecond: i16,
}

impl SlideComment {
    /// Create a new comment with author, text, and position.
    ///
    /// # Arguments
    ///
    /// * `author` - Comment author name
    /// * `text` - Comment text content
    /// * `x` - X position in points
    /// * `y` - Y position in points
    ///
    /// # Example
    ///
    /// ```
    /// use litchi::ole::ppt::writer::comments::SlideComment;
    /// let comment = SlideComment::new("John Doe", "Great slide!", 100, 50);
    /// ```
    pub fn new(author: &str, text: &str, x: i32, y: i32) -> Self {
        // Derive initials from author name (first letter of each word)
        let initials: String = author
            .split_whitespace()
            .filter_map(|w| w.chars().next())
            .collect();

        Self {
            author: author.to_string(),
            text: text.to_string(),
            initials,
            x,
            y,
            date: CommentDateTime::default(),
        }
    }

    /// Set the comment date/time.
    pub fn with_date(mut self, date: CommentDateTime) -> Self {
        self.date = date;
        self
    }

    /// Set the author initials explicitly.
    pub fn with_initials(mut self, initials: &str) -> Self {
        self.initials = initials.to_string();
        self
    }
}

/// Convert points to PPT master units (576 units per inch, 72 points per inch).
fn pt_to_master(pt: i32) -> i32 {
    (pt as i64 * 576 / 72) as i32
}

/// Write a CString record with the given instance and UTF-16LE text.
fn write_cstring(instance: u16, text: &str) -> Result<Vec<u8>, PptError> {
    let utf16: Vec<u16> = text.encode_utf16().collect();
    let mut data = Vec::with_capacity(utf16.len() * 2);
    for ch in &utf16 {
        data.extend_from_slice(&ch.to_le_bytes());
    }
    let mut builder = RecordBuilder::new(0x00, instance, record_type::CSTRING);
    builder.write_data(&data);
    builder.build()
}

/// Build a single Comment2000 container for one comment.
///
/// # Arguments
///
/// * `comment` - The comment to serialize
/// * `index` - 1-based comment index within the slide
fn build_comment_container(comment: &SlideComment, index: i32) -> Result<Vec<u8>, PptError> {
    let mut children = Vec::new();

    // CString instance 0: author name
    if !comment.author.is_empty() {
        children.extend(write_cstring(0, &comment.author)?);
    }

    // CString instance 1: comment text
    if !comment.text.is_empty() {
        children.extend(write_cstring(1, &comment.text)?);
    }

    // CString instance 2: author initials
    if !comment.initials.is_empty() {
        children.extend(write_cstring(2, &comment.initials)?);
    }

    // Comment2000Atom (28 bytes)
    let mut atom_data = Vec::with_capacity(28);
    atom_data.extend_from_slice(&index.to_le_bytes()); // index (i32)
    atom_data.extend_from_slice(&comment.date.year.to_le_bytes()); // year (i16)
    atom_data.extend_from_slice(&comment.date.month.to_le_bytes()); // month (u16)
    atom_data.extend_from_slice(&comment.date.day.to_le_bytes()); // day of week (u16, set to day)
    atom_data.extend_from_slice(&comment.date.day.to_le_bytes()); // day (u16)
    atom_data.extend_from_slice(&comment.date.hour.to_le_bytes()); // hour (u16)
    atom_data.extend_from_slice(&comment.date.minute.to_le_bytes()); // minute (u16)
    atom_data.extend_from_slice(&comment.date.second.to_le_bytes()); // second (u16)
    atom_data.extend_from_slice(&comment.date.millisecond.to_le_bytes()); // millisecond (i16)
    atom_data.extend_from_slice(&pt_to_master(comment.x).to_le_bytes()); // x (i32)
    atom_data.extend_from_slice(&pt_to_master(comment.y).to_le_bytes()); // y (i32)

    let mut atom = RecordBuilder::new(0x00, 0, record_type::COMMENT2000_ATOM);
    atom.write_data(&atom_data);
    children.extend(atom.build()?);

    // Comment2000 container
    let mut container = RecordBuilder::new(0x0F, 0, record_type::COMMENT2000);
    container.write_data(&children);
    container.build()
}

/// Build all Comment2000 records for a slide's comments.
///
/// Returns the raw bytes to be embedded inside the slide's BinaryTagData.
/// If there are no comments, returns an empty Vec.
pub fn build_slide_comments(comments: &[SlideComment]) -> Result<Vec<u8>, PptError> {
    if comments.is_empty() {
        return Ok(Vec::new());
    }

    let mut data = Vec::new();
    for (i, comment) in comments.iter().enumerate() {
        data.extend(build_comment_container(comment, (i + 1) as i32)?);
    }
    Ok(data)
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_slide_comment_new() {
        let comment = SlideComment::new("John Doe", "Hello!", 100, 50);
        assert_eq!(comment.author, "John Doe");
        assert_eq!(comment.text, "Hello!");
        assert_eq!(comment.x, 100);
        assert_eq!(comment.y, 50);
        assert_eq!(comment.initials, "JD");
    }

    #[test]
    fn test_slide_comment_initials_single_word() {
        let comment = SlideComment::new("Alice", "Test", 0, 0);
        assert_eq!(comment.initials, "A");
    }

    #[test]
    fn test_slide_comment_initials_multiple_words() {
        let comment = SlideComment::new("John Jacob Jingleheimer Schmidt", "Test", 0, 0);
        assert_eq!(comment.initials, "JJJS");
    }

    #[test]
    fn test_slide_comment_with_initials() {
        let comment = SlideComment::new("John Doe", "Test", 0, 0).with_initials("JDOE");
        assert_eq!(comment.initials, "JDOE");
    }

    #[test]
    fn test_slide_comment_with_date() {
        let date = CommentDateTime {
            year: 2025,
            month: 3,
            day: 13,
            hour: 14,
            minute: 30,
            second: 45,
            millisecond: 500,
        };
        let comment = SlideComment::new("Test", "Test", 0, 0).with_date(date.clone());
        assert_eq!(comment.date.year, 2025);
        assert_eq!(comment.date.month, 3);
        assert_eq!(comment.date.day, 13);
    }

    #[test]
    fn test_slide_comment_clone() {
        let comment = SlideComment::new("Author", "Text", 100, 200)
            .with_initials("A")
            .with_date(CommentDateTime {
                year: 2025,
                month: 1,
                day: 1,
                hour: 0,
                minute: 0,
                second: 0,
                millisecond: 0,
            });
        let cloned = comment.clone();
        assert_eq!(cloned.author, comment.author);
        assert_eq!(cloned.text, comment.text);
        assert_eq!(cloned.x, comment.x);
        assert_eq!(cloned.y, comment.y);
        assert_eq!(cloned.initials, comment.initials);
    }

    #[test]
    fn test_slide_comment_debug() {
        let comment = SlideComment::new("Test", "Text", 100, 200);
        let debug = format!("{:?}", comment);
        assert!(debug.contains("SlideComment"));
        assert!(debug.contains("Test"));
    }

    #[test]
    fn test_comment_date_time_default() {
        let date = CommentDateTime::default();
        assert_eq!(date.year, 0);
        assert_eq!(date.month, 0);
        assert_eq!(date.day, 0);
        assert_eq!(date.hour, 0);
        assert_eq!(date.minute, 0);
        assert_eq!(date.second, 0);
        assert_eq!(date.millisecond, 0);
    }

    #[test]
    fn test_comment_date_time_clone() {
        let date = CommentDateTime {
            year: 2025,
            month: 6,
            day: 15,
            hour: 12,
            minute: 30,
            second: 0,
            millisecond: 0,
        };
        let cloned = date.clone();
        assert_eq!(cloned.year, 2025);
        assert_eq!(cloned.month, 6);
        assert_eq!(cloned.day, 15);
    }

    #[test]
    fn test_pt_to_master() {
        // 72 points = 1 inch = 576 master units
        assert_eq!(pt_to_master(72), 576);
        // 0 points = 0 master units
        assert_eq!(pt_to_master(0), 0);
        // 144 points = 2 inches = 1152 master units
        assert_eq!(pt_to_master(144), 1152);
    }

    #[test]
    fn test_build_comment_container() {
        let comment = SlideComment::new("John Doe", "Hello!", 100, 50);
        let data = build_comment_container(&comment, 1).unwrap();
        // Should contain Comment2000 container header (8 bytes) + children
        assert!(data.len() > 8);
        // Verify record type = 12000
        let rtype = u16::from_le_bytes([data[2], data[3]]);
        assert_eq!(rtype, 12000);
    }

    #[test]
    fn test_build_comment_container_with_date() {
        let comment = SlideComment::new("Author", "Text", 100, 50).with_date(CommentDateTime {
            year: 2025,
            month: 3,
            day: 13,
            hour: 14,
            minute: 30,
            second: 0,
            millisecond: 0,
        });
        let data = build_comment_container(&comment, 1).unwrap();
        assert!(!data.is_empty());
    }

    #[test]
    fn test_build_slide_comments_empty() {
        let data = build_slide_comments(&[]).unwrap();
        assert!(data.is_empty());
    }

    #[test]
    fn test_build_slide_comments_single() {
        let comments = vec![SlideComment::new("Alice", "First comment", 10, 20)];
        let data = build_slide_comments(&comments).unwrap();
        assert!(!data.is_empty());

        // Count Comment2000 containers (type=12000)
        let mut count = 0;
        let mut offset = 0;
        while offset + 8 <= data.len() {
            let rtype = u16::from_le_bytes([data[offset + 2], data[offset + 3]]);
            let rlen = u32::from_le_bytes([
                data[offset + 4],
                data[offset + 5],
                data[offset + 6],
                data[offset + 7],
            ]);
            let ver = data[offset] & 0x0F;
            if rtype == 12000 {
                count += 1;
            }
            if ver == 0x0F {
                offset += 8; // container: descend
            } else {
                offset += 8 + rlen as usize; // atom: skip data
            }
        }
        assert_eq!(count, 1);
    }

    #[test]
    fn test_build_slide_comments_multiple() {
        let comments = vec![
            SlideComment::new("Alice", "First comment", 10, 20),
            SlideComment::new("Bob", "Second comment", 30, 40),
        ];
        let data = build_slide_comments(&comments).unwrap();
        assert!(!data.is_empty());

        // Count Comment2000 containers (type=12000)
        let mut count = 0;
        let mut offset = 0;
        while offset + 8 <= data.len() {
            let rtype = u16::from_le_bytes([data[offset + 2], data[offset + 3]]);
            let rlen = u32::from_le_bytes([
                data[offset + 4],
                data[offset + 5],
                data[offset + 6],
                data[offset + 7],
            ]);
            let ver = data[offset] & 0x0F;
            if rtype == 12000 {
                count += 1;
            }
            if ver == 0x0F {
                offset += 8; // container: descend
            } else {
                offset += 8 + rlen as usize; // atom: skip data
            }
        }
        assert_eq!(count, 2);
    }

    #[test]
    fn test_build_slide_comments_unicode() {
        let comments = vec![SlideComment::new(
            "\u{4f60}\u{597d}",
            "\u{4e16}\u{754c}",
            100,
            100,
        )];
        let data = build_slide_comments(&comments).unwrap();
        assert!(!data.is_empty());
    }

    #[test]
    fn test_build_slide_comments_empty_author() {
        let comments = vec![SlideComment::new("", "Anonymous comment", 100, 100)];
        let data = build_slide_comments(&comments).unwrap();
        assert!(!data.is_empty());
    }

    #[test]
    fn test_build_slide_comments_empty_text() {
        let comments = vec![SlideComment::new("Author", "", 100, 100)];
        let data = build_slide_comments(&comments).unwrap();
        assert!(!data.is_empty());
    }

    #[test]
    fn test_build_slide_comments_large_coordinates() {
        let comments = vec![SlideComment::new("Author", "Text", 10000, 20000)];
        let data = build_slide_comments(&comments).unwrap();
        assert!(!data.is_empty());
    }

    #[test]
    fn test_build_slide_comments_negative_coordinates() {
        let comments = vec![SlideComment::new("Author", "Text", -100, -200)];
        let data = build_slide_comments(&comments).unwrap();
        assert!(!data.is_empty());
    }
}
