use litchi_opc::PackURI;

/// Inert metadata for one `InkML` content part anchored on a slide.
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
    #[must_use]
    pub fn slide_index(&self) -> usize {
        self.slide_index
    }

    #[must_use]
    pub fn index(&self) -> usize {
        self.index
    }

    #[must_use]
    pub fn relationship_id(&self) -> &str {
        &self.relationship_id
    }

    #[must_use]
    pub fn part_name(&self) -> &PackURI {
        &self.part_name
    }

    #[must_use]
    pub fn trace_count(&self) -> usize {
        self.trace_count
    }

    #[must_use]
    pub fn trace_group_count(&self) -> usize {
        self.trace_group_count
    }
}

/// Outcome of storing a validated `InkML` part.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct StoredAnnotation {
    pub relationship_id: String,
    pub part_name: PackURI,
}
