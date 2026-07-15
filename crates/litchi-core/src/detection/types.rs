//! File format type enumeration.

/// Supported file formats that can be detected.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
#[non_exhaustive]
pub enum FileFormat {
    /// Microsoft Word Document (OLE2 format, .doc)
    Doc,
    /// Microsoft Word Document (OOXML format, .docx)
    Docx,
    /// Microsoft PowerPoint Presentation (OLE2 format, .ppt)
    Ppt,
    /// Microsoft PowerPoint Presentation (OOXML format, .pptx)
    Pptx,
    /// Microsoft Excel Spreadsheet (OLE2 format, .xls)
    Xls,
    /// Microsoft Excel Spreadsheet (OOXML format, .xlsx)
    Xlsx,
    /// Microsoft Excel Spreadsheet (Binary OOXML format, .xlsb)
    Xlsb,
    /// Rich Text Format Document (.rtf)
    Rtf,
    /// Apple Pages Document (.pages)
    Pages,
    /// Apple Keynote Presentation (.key)
    Keynote,
    /// Apple Numbers Spreadsheet (.numbers)
    Numbers,
    /// OpenDocument Text (.odt)
    Odt,
    /// OpenDocument Spreadsheet (.ods)
    Ods,
    /// OpenDocument Presentation (.odp)
    Odp,
    /// OpenDocument Drawing (.odg, .otg)
    Odg,
    /// OpenDocument Chart (.odc, .otc)
    Odc,
    /// OpenDocument Formula (.odf, .otf)
    Odf,
    /// OpenDocument Image (.odi, .oti)
    Odi,
    /// OpenDocument Master Document (.odm, .otm)
    Odm,
    /// OpenDocument Web Document (.oth)
    Oth,
    /// OpenDocument Database Front End (.odb)
    Odb,
}
