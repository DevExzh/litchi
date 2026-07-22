//! Inert MS-OVBA project-storage discovery for legacy Word documents.
//!
//! Discovery reads CFB directory names only. It never opens, decompresses,
//! parses, or executes the `PROJECT`, `dir`, `_VBA_PROJECT`, or module streams.

use std::collections::BTreeMap;

/// Directory-only metadata for one candidate MS-OVBA project storage.
///
/// MS-OVBA defines a project root storage containing a `VBA` storage and a
/// `PROJECT` stream. The `VBA` storage in turn requires `_VBA_PROJECT` and
/// `dir` streams. This model reports the observed directory topology without
/// interpreting any stream content.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct VbaProjectStorage {
    project_root_path: Vec<String>,
    vba_storage_path: Vec<String>,
    has_project_stream: bool,
    has_project_wm_stream: bool,
    has_project_lk_stream: bool,
    has_vba_project_stream: bool,
    has_dir_stream: bool,
    candidate_module_stream_names: Vec<String>,
    srp_stream_names: Vec<String>,
}

impl VbaProjectStorage {
    /// Return the CFB path of the MS-OVBA project root storage.
    pub fn project_root_path(&self) -> &[String] {
        &self.project_root_path
    }

    /// Return the CFB path of the `VBA` storage.
    pub fn vba_storage_path(&self) -> &[String] {
        &self.vba_storage_path
    }

    /// Whether the project root has the required `PROJECT` stream.
    pub fn has_project_stream(&self) -> bool {
        self.has_project_stream
    }

    /// Whether the project root has the optional `PROJECTwm` stream.
    pub fn has_project_wm_stream(&self) -> bool {
        self.has_project_wm_stream
    }

    /// Whether the project root has the optional `PROJECTlk` stream.
    pub fn has_project_lk_stream(&self) -> bool {
        self.has_project_lk_stream
    }

    /// Whether the `VBA` storage has the required `_VBA_PROJECT` stream.
    pub fn has_vba_project_stream(&self) -> bool {
        self.has_vba_project_stream
    }

    /// Whether the `VBA` storage has the required compressed `dir` stream.
    pub fn has_dir_stream(&self) -> bool {
        self.has_dir_stream
    }

    /// Return direct `VBA` child streams that may be module streams.
    ///
    /// `_VBA_PROJECT`, `dir`, and optional `__SRP_*` streams are excluded.
    /// The stream bytes are not opened, so these names are candidates rather
    /// than a claim that any particular module has executable source code.
    pub fn candidate_module_stream_names(&self) -> &[String] {
        &self.candidate_module_stream_names
    }

    /// Return optional `__SRP_*` stream names observed in the `VBA` storage.
    ///
    /// MS-OVBA specifies that SRP streams must be ignored. This method exposes
    /// names only; it does not read their content.
    pub fn srp_stream_names(&self) -> &[String] {
        &self.srp_stream_names
    }

    /// Whether both streams required inside the `VBA` storage are present.
    pub fn has_required_vba_streams(&self) -> bool {
        self.has_vba_project_stream && self.has_dir_stream
    }

    /// Whether the required project-root and `VBA` storage names are present.
    ///
    /// This validates directory topology only. It does not validate the binary
    /// stream formats or load any VBA source code.
    pub fn is_structurally_complete(&self) -> bool {
        self.has_project_stream && self.has_required_vba_streams()
    }

    /// Whether directory metadata indicates that the storage may contain macro code.
    ///
    /// This is a conservative presence signal, not code analysis. Macro/module
    /// stream content is never opened, decompressed, parsed, or executed.
    pub fn may_contain_macro_code(&self) -> bool {
        self.is_structurally_complete() && !self.candidate_module_stream_names.is_empty()
    }
}

pub(crate) fn discover_vba_project_storages(
    stream_paths: &[Vec<String>],
) -> Vec<VbaProjectStorage> {
    let mut direct_children = BTreeMap::<Vec<String>, Vec<String>>::new();
    for path in stream_paths {
        for vba_index in 0..path.len().saturating_sub(1) {
            if !path[vba_index].eq_ignore_ascii_case("VBA") || path.len() != vba_index + 2 {
                continue;
            }
            direct_children
                .entry(path[..=vba_index].to_vec())
                .or_default()
                .push(path[vba_index + 1].clone());
        }
    }

    direct_children
        .into_iter()
        .map(|(vba_storage_path, mut children)| {
            sort_case_insensitively(&mut children);
            children.dedup_by(|left, right| left.eq_ignore_ascii_case(right));

            let project_root_path = vba_storage_path[..vba_storage_path.len() - 1].to_vec();
            let has_project_stream = has_direct_stream(stream_paths, &project_root_path, "PROJECT");
            let has_project_wm_stream =
                has_direct_stream(stream_paths, &project_root_path, "PROJECTwm");
            let has_project_lk_stream =
                has_direct_stream(stream_paths, &project_root_path, "PROJECTlk");
            let has_vba_project_stream = children
                .iter()
                .any(|name| name.eq_ignore_ascii_case("_VBA_PROJECT"));
            let has_dir_stream = children.iter().any(|name| name.eq_ignore_ascii_case("dir"));
            let srp_stream_names = children
                .iter()
                .filter(|name| is_srp_stream(name))
                .cloned()
                .collect();
            let candidate_module_stream_names = children
                .into_iter()
                .filter(|name| {
                    !name.eq_ignore_ascii_case("_VBA_PROJECT")
                        && !name.eq_ignore_ascii_case("dir")
                        && !is_srp_stream(name)
                })
                .collect();

            VbaProjectStorage {
                project_root_path,
                vba_storage_path,
                has_project_stream,
                has_project_wm_stream,
                has_project_lk_stream,
                has_vba_project_stream,
                has_dir_stream,
                candidate_module_stream_names,
                srp_stream_names,
            }
        })
        .collect()
}

fn has_direct_stream(stream_paths: &[Vec<String>], parent: &[String], name: &str) -> bool {
    stream_paths.iter().any(|path| {
        path.len() == parent.len() + 1
            && path
                .iter()
                .zip(parent)
                .all(|(component, expected)| component.eq_ignore_ascii_case(expected))
            && path
                .last()
                .is_some_and(|component| component.eq_ignore_ascii_case(name))
    })
}

fn is_srp_stream(name: &str) -> bool {
    name.get(..6)
        .is_some_and(|prefix| prefix.eq_ignore_ascii_case("__SRP_"))
}

fn sort_case_insensitively(names: &mut [String]) {
    names.sort_by(|left, right| {
        left.to_ascii_lowercase()
            .cmp(&right.to_ascii_lowercase())
            .then_with(|| left.cmp(right))
    });
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::doc::Package;
    use litchi_cfb::OleWriter;
    use std::io::Cursor;

    #[test]
    fn discovers_complete_vba_storage_from_directory_names_only() {
        let stream_paths = vec![
            vec!["WordDocument".to_string()],
            vec!["Macros".to_string(), "PROJECT".to_string()],
            vec!["Macros".to_string(), "PROJECTwm".to_string()],
            vec!["Macros".to_string(), "PROJECTlk".to_string()],
            vec![
                "Macros".to_string(),
                "vBa".to_string(),
                "_vba_project".to_string(),
            ],
            vec!["Macros".to_string(), "vBa".to_string(), "DIR".to_string()],
            vec![
                "Macros".to_string(),
                "vBa".to_string(),
                "ThisDocument".to_string(),
            ],
            vec![
                "Macros".to_string(),
                "vBa".to_string(),
                "Module1".to_string(),
            ],
            vec![
                "Macros".to_string(),
                "vBa".to_string(),
                "__sRp_0".to_string(),
            ],
        ];

        let discovered = discover_vba_project_storages(&stream_paths);
        assert_eq!(discovered.len(), 1);
        let project = &discovered[0];
        assert_eq!(project.project_root_path(), ["Macros"]);
        assert_eq!(project.vba_storage_path(), ["Macros", "vBa"]);
        assert!(project.has_project_stream());
        assert!(project.has_project_wm_stream());
        assert!(project.has_project_lk_stream());
        assert!(project.has_required_vba_streams());
        assert!(project.is_structurally_complete());
        assert!(project.may_contain_macro_code());
        assert_eq!(
            project.candidate_module_stream_names(),
            ["Module1", "ThisDocument"]
        );
        assert_eq!(project.srp_stream_names(), ["__sRp_0"]);
    }

    #[test]
    fn reports_incomplete_projects_and_never_opens_macro_streams() {
        let mut writer = OleWriter::new();
        writer
            .create_stream(&["WordDocument"], b"not parsed")
            .unwrap();
        writer
            .create_stream(&["Macros", "VBA", "dir"], b"intentionally invalid")
            .unwrap();
        let mut bytes = Cursor::new(Vec::new());
        writer.write_to(&mut bytes).unwrap();

        let package = Package::from_reader(Cursor::new(bytes.into_inner())).unwrap();
        let projects = package.vba_project_storages();
        assert_eq!(projects.len(), 1);
        assert!(projects[0].has_dir_stream());
        assert!(!projects[0].has_vba_project_stream());
        assert!(!projects[0].has_project_stream());
        assert!(!projects[0].is_structurally_complete());
        assert!(!projects[0].may_contain_macro_code());
    }
}
