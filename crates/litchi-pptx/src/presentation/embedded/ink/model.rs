use litchi_opc::PackURI;

/// Inert metadata for one InkML content part anchored on a slide.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Annotation {
    pub(crate) slide_index: usize,
    pub(crate) index: usize,
    pub(crate) relationship_id: String,
    pub(crate) part_name: PackURI,
    pub(crate) trace_count: usize,
    pub(crate) trace_group_count: usize,
}

impl Annotation {
    pub fn slide_index(&self) -> usize {
        self.slide_index
    }

    pub fn index(&self) -> usize {
        self.index
    }

    pub fn relationship_id(&self) -> &str {
        &self.relationship_id
    }

    pub fn part_name(&self) -> &PackURI {
        &self.part_name
    }

    pub fn trace_count(&self) -> usize {
        self.trace_count
    }

    pub fn trace_group_count(&self) -> usize {
        self.trace_group_count
    }
}

/// Outcome of storing a validated InkML part.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct StoredAnnotation {
    pub relationship_id: String,
    pub part_name: PackURI,
}
