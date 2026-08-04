//! Immutable semantic values for this document family.

/// A referenced master-document subdocument.
#[derive(Clone, Debug, PartialEq, Eq)]
pub struct Subdocument {
    href: String,
}
impl Subdocument {
    pub fn new(href: impl Into<String>) -> Self {
        Self { href: href.into() }
    }
    pub fn href(&self) -> &str {
        &self.href
    }
}
/// A master-document section.
#[derive(Clone, Debug, PartialEq, Eq)]
pub struct Section {
    name: String,
    children: Vec<Subdocument>,
}
impl Section {
    pub fn new(name: impl Into<String>) -> Self {
        Self {
            name: name.into(),
            children: Vec::new(),
        }
    }
    pub fn name(&self) -> &str {
        &self.name
    }
    pub fn children(&self) -> &[Subdocument] {
        &self.children
    }
    pub fn push(&mut self, child: Subdocument) {
        self.children.push(child);
    }
}
