//! Package-independent validated presentation-structure values.

use super::super::custom_show::List as ShowList;
use super::super::sections::List as SectionList;

/// One entry in `p:sldIdLst`, resolved through the presentation relationship set.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Reference {
    pub slide_id: u32,
    pub relationship_id: String,
    pub part_name: String,
}

/// Validated presentation ordering, custom shows, and modern sections.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Graph {
    pub slides: Vec<Reference>,
    pub custom_shows: ShowList,
    pub sections: SectionList,
}
