use super::{CellValue, Section, Slide, StructuredData, Table};

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_table_creation() {
        let mut table = Table::new("Test Table".to_string());
        assert_eq!(table.name, "Test Table");
        assert_eq!(table.row_count, 0);
        assert_eq!(table.column_count, 0);

        table.set_cell(0, 0, CellValue::Text("Header 1".to_string()));
        table.set_cell(0, 1, CellValue::Text("Header 2".to_string()));
        table.set_cell(1, 0, CellValue::Number(42.0));
        table.set_cell(1, 1, CellValue::Boolean(true));

        assert_eq!(table.row_count, 2);
        assert_eq!(table.column_count, 2);

        let csv = table.to_csv();
        assert!(csv.contains("Header 1"));
        assert!(csv.contains("42"));
    }

    #[test]
    fn test_cell_value() {
        let text_cell = CellValue::Text("Hello".to_string());
        assert_eq!(text_cell.to_string(), "Hello");
        assert!(!text_cell.is_empty());

        let empty_cell = CellValue::Empty;
        assert_eq!(empty_cell.to_string(), "");
        assert!(empty_cell.is_empty());

        let number_cell = CellValue::Number(std::f64::consts::PI);
        assert_eq!(number_cell.to_string(), "3.141592653589793");
    }

    #[test]
    fn test_slide_creation() {
        let mut slide = Slide::new(0);
        assert_eq!(slide.index, 0);
        assert_eq!(slide.title, None);

        slide.title = Some("Introduction".to_string());
        slide.text_content.push("Point 1".to_string());
        slide.text_content.push("Point 2".to_string());
        slide.notes = Some("Speaker notes".to_string());

        let all_text = slide.all_text();
        assert_eq!(all_text.len(), 4);
        assert_eq!(all_text[0], "Introduction");
        assert_eq!(all_text[3], "Speaker notes");
    }

    #[test]
    fn test_section_creation() {
        let mut section = Section::new(0);
        section.heading = Some("Chapter 1".to_string());
        section.paragraphs.push("First paragraph.".to_string());
        section.paragraphs.push("Second paragraph.".to_string());

        let all_text = section.all_text();
        assert_eq!(all_text.len(), 3);
        assert_eq!(all_text[0], "Chapter 1");
    }

    #[test]
    fn test_structured_data() {
        let mut table = Table::new("Data".to_string());
        table.set_cell(0, 0, CellValue::Text("A".to_string()));

        let mut slide = Slide::new(0);
        slide.title = Some("Title".to_string());

        let mut section = Section::new(0);
        section.heading = Some("Heading".to_string());

        let data = StructuredData {
            tables: vec![table],
            slides: vec![slide],
            sections: vec![section],
        };

        assert!(!data.is_empty());
        let summary = data.summary();
        assert!(summary.contains("Tables: 1"));
        assert!(summary.contains("Slides: 1"));
        assert!(summary.contains("Sections: 1"));
    }
}
