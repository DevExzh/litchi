//! Native Word visuals are inspected as inert metadata and package resources.
use litchi_docx::{Package, paragraph::RunContent};
use litchi_opc::PackURI;
use sha2::{Digest, Sha256};

fn path(name: &str) -> std::path::PathBuf {
    std::path::Path::new(env!("CARGO_MANIFEST_DIR"))
        .join("../../test-data/microsoft-office/word")
        .join(name)
}
fn native(name: &str) -> Package {
    Package::open(path(name)).unwrap()
}
fn hash(bytes: &[u8]) -> String {
    Sha256::digest(bytes)
        .iter()
        .map(|b| format!("{b:02x}"))
        .collect()
}

fn members(bytes: &[u8]) -> std::collections::BTreeMap<String, Vec<u8>> {
    let reader = litchi_opc::phys_pkg::OwnedPhysPkgReader::from_bytes(
        bytes.to_vec(),
        litchi_opc::SourceLimits::default(),
    )
    .unwrap();
    reader
        .member_names()
        .unwrap()
        .into_iter()
        .map(|name| {
            let data = reader.read_member(&name).unwrap();
            (name, data)
        })
        .collect()
}

#[test]
fn native_picture_extents_alt_text_and_ole_workbook_are_distinct_resources() {
    let package = native("word-visual-objects.docx");
    let document = package.document().unwrap();
    let opc = package.opc_package();
    let main = opc.main_document_part().unwrap();
    let images = document
        .paragraphs()
        .unwrap()
        .iter()
        .flat_map(|p| p.images().unwrap())
        .collect::<Vec<_>>();
    assert_eq!(images.len(), 2);
    for (id, width, height, expected_hash) in [
        (
            "rId4",
            2743200,
            2057400,
            "b6f8d168d41e0fc296eb1e0e3855d9669584b7528c2ef4966817443d9bc9b929",
        ),
        (
            "rId5",
            1828800,
            1028700,
            "d594d8f832d3ab38d03e0d06316d2fa5e401217e50e21e01814e593d752da3e8",
        ),
    ] {
        let image = images.iter().find(|image| image.r_embed() == id).unwrap();
        assert_eq!((image.width_emu(), image.height_emu()), (width, height));
        assert_eq!(hash(&image.data(opc, main.rels()).unwrap()), expected_hash);
    }
    assert_eq!(
        images
            .iter()
            .find(|i| i.r_embed() == "rId4")
            .unwrap()
            .description(),
        "INLINE_ALT_007: abstract JPEG with overlapping colored curves."
    );
    assert_eq!(
        images
            .iter()
            .find(|i| i.r_embed() == "rId5")
            .unwrap()
            .description(),
        "FLOATING_ALT_007: synthetic PNG poster anchored at middle right with square wrapping."
    );
    let mut objects = Vec::new();
    for paragraph in document.paragraphs().unwrap() {
        for run in paragraph.runs().unwrap() {
            for content in run.contents().unwrap() {
                if let RunContent::EmbeddedObject(object) = content {
                    objects.push(object);
                }
            }
        }
    }
    assert_eq!(objects.len(), 1);
    let object = &objects[0];
    assert_eq!(object.prog_id(), Some("Excel.Sheet.12"));
    assert_eq!(object.object_type(), Some("Embed"));
    assert_eq!(object.draw_aspect(), Some("Content"));
    assert_eq!(object.ole_r_id(), Some("rId7"));
    assert_eq!(object.preview_image().unwrap().r_embed(), "rId6");
    let preview = object.preview_image().unwrap();
    assert_eq!(
        (preview.width_emu(), preview.height_emu()),
        (2489200, 1028700)
    );
    assert_eq!(
        hash(&preview.data(opc, main.rels()).unwrap()),
        "2585e48c4e95e526a3ca2cf32e14354d2bbb43b0d98bff7d4906870071a2c4a3"
    );
    let payload = object.ole_data(opc, main.rels()).unwrap().unwrap();
    assert_eq!(payload.len(), 18372);
    assert_eq!(
        hash(&payload),
        "c4937954a07d192200a9a4cb17d016c345be8405dff727f25e39e2d1cf65e866"
    );
}

#[test]
fn native_doughnut_chart_retains_caches_and_its_separate_workbook() {
    use litchi_drawingml::chart::{TypeGroup, reader};
    let package = native("word-visual-objects.docx");
    let graph = package.chart_graph().unwrap();
    assert_eq!(graph.charts.len(), 1);
    let chart = &graph.charts[0];
    assert_eq!(chart.document_relationship_id, "rId8");
    assert_eq!(chart.part_name, "/word/charts/chart1.xml");
    assert_eq!(chart.styles.len(), 1);
    assert_eq!(chart.color_styles.len(), 1);
    assert_eq!(
        hash(&chart.styles[0].data),
        "c1f9ed46ee58a42fd84eb085275cae0e1393d54b72a5a5fe9ece81381636b69f"
    );
    assert_eq!(
        hash(&chart.color_styles[0].data),
        "04fefba7d31853fa0aa636e5c8b8e308fc3126a9b48a1f4490947fff91d5a92f"
    );
    let workbook = chart.workbook.as_ref().unwrap();
    assert_eq!(workbook.relationship_id, "rId3");
    assert_eq!(
        workbook.part_name,
        "/word/embeddings/Microsoft_Excel_Worksheet1.xlsx"
    );
    assert_eq!(workbook.data.len(), 9894);
    assert_eq!(
        hash(&workbook.data),
        "4e988e3f496ad305c02e39b2e4dc7349867c2c79aa69e0809bc63e99829864a2"
    );
    let parsed = reader::read(
        chart.data.as_slice(),
        &litchi_drawingml::chart::Limits::DEFAULT,
    )
    .unwrap();
    assert_eq!(parsed.plot_area.type_groups.len(), 1);
    let TypeGroup::Doughnut(group) = &parsed.plot_area.type_groups[0] else {
        panic!("expected native doughnut")
    };
    assert_eq!(group.hole_size, 75);
    assert_eq!(group.common.series.len(), 1);
    let series = &group.common.series[0];
    let Some(litchi_drawingml::chart::TitleText::Reference(title)) = &series.title else {
        panic!("expected stored title reference")
    };
    assert_eq!(title.formula, "Sheet1!$B$1");
    assert_eq!(
        series.categories.as_ref().unwrap().values,
        ["1st Qtr", "2nd Qtr", "3rd Qtr", "4th Qtr"]
    );
    assert_eq!(
        series
            .categories
            .as_ref()
            .unwrap()
            .source_ref
            .as_ref()
            .unwrap()
            .formula,
        "Sheet1!$A$2:$A$5"
    );
    assert_eq!(series.values.as_ref().unwrap().values, [8.2, 3.2, 1.4, 1.2]);
    assert_eq!(
        series
            .values
            .as_ref()
            .unwrap()
            .source_ref
            .as_ref()
            .unwrap()
            .formula,
        "Sheet1!$B$2:$B$5"
    );
    let charts = package
        .document()
        .unwrap()
        .paragraphs()
        .unwrap()
        .iter()
        .flat_map(|p| p.charts().unwrap())
        .collect::<Vec<_>>();
    assert_eq!(charts.len(), 1);
    assert_eq!(charts[0].relationship_id(), "rId8");
    assert_eq!(
        (charts[0].width_emu(), charts[0].height_emu()),
        (5486400, 3200400)
    );
}

#[test]
fn native_smartart_process_retains_nodes_and_definition_identities() {
    let package = native("word-vector-diagram.docx");
    let diagrams = package.document().unwrap().smart_arts().unwrap();
    assert_eq!(diagrams.len(), 1);
    let diagram = &diagrams[0];
    assert_eq!(
        diagram.diagram_type(),
        litchi_drawingml::diagram::DiagramType::Process
    );
    assert_eq!(diagram.text(), "ACQUIRE_013\nTRANSFORM_013\nVERIFY_013");
    assert_eq!(diagram.node_tree().len(), 3);
    assert_eq!(diagram.data.points.len(), 18);
    assert_eq!(diagram.data.connections.len(), 18);
    assert_eq!(diagram.data_relationship_id, "rId6");
    assert_eq!(
        diagram.drawing_part_name.as_deref(),
        Some("/word/diagrams/drawing1.xml")
    );
    assert_eq!(diagram.extent, Some((5486400, 3200400)));
    for (definition, id) in [
        (
            diagram.layout.as_ref().unwrap(),
            "urn:microsoft.com/office/officeart/2005/8/layout/process1",
        ),
        (
            diagram.quick_style.as_ref().unwrap(),
            "urn:microsoft.com/office/officeart/2005/8/quickstyle/simple1",
        ),
        (
            diagram.colors.as_ref().unwrap(),
            "urn:microsoft.com/office/officeart/2005/8/colors/accent1_2",
        ),
    ] {
        assert_eq!(definition.unique_id.as_deref(), Some(id));
    }
}

#[test]
fn native_rounded_rectangle_exposes_one_active_textbox_story() {
    let package = native("word-visual-objects.docx");
    let document = package.document().unwrap();
    let boxes = document.text_boxes().unwrap();
    assert_eq!(boxes.len(), 1);
    assert_eq!(
        boxes[0].text(),
        "SHAPE_TEXTBOX_007\nNative rounded rectangle"
    );
    assert_eq!(boxes[0].name.as_deref(), Some("Rounded Rectangle 3"));
    assert!(!boxes[0].is_word_art());
    assert_eq!(boxes[0].anchor, litchi_docx::textbox::TextBoxAnchor::Vml);
    // This reader's MCE profile selects the native VML fallback. It does not
    // qualify the separate general drawing inventory's missing WPS branch.
    assert!(
        package
            .opc_package()
            .get_part(&PackURI::new("/word/document.xml").unwrap())
            .unwrap()
            .blob()
            .windows(11)
            .any(|v| v == b"v:roundrect")
    );
}

#[test]
fn native_chart_hole_edit_preserves_cached_data_workbooks_and_companions() {
    let package = native("word-visual-objects.docx");
    let original = members(&package.to_bytes().unwrap());
    let mut graph = package.chart_graph().unwrap();
    let xml = String::from_utf8(graph.charts[0].data.clone()).unwrap();
    assert_eq!(xml.matches("<c:holeSize val=\"75\"/>").count(), 1);
    graph.charts[0].data = xml
        .replace("<c:holeSize val=\"75\"/>", "<c:holeSize val=\"55\"/>")
        .into_bytes();
    let mut edit = package.edit_chart();
    assert!(edit.replace_graph(graph.clone()).unwrap());
    let commit = edit.commit().unwrap();
    let directory = tempfile::tempdir().unwrap();
    let destination = directory.path().join("chart-edited.docx");
    commit.package().save(&destination).unwrap();
    let reopened = Package::open(&destination).unwrap();
    assert_eq!(reopened.chart_graph().unwrap(), graph);
    let updated = members(&reopened.to_bytes().unwrap());
    assert_eq!(
        updated.keys().collect::<Vec<_>>(),
        original.keys().collect::<Vec<_>>()
    );
    for (name, data) in &original {
        if name != "word/charts/chart1.xml" {
            assert!(updated.get(name) == Some(data), "changed {name}");
        }
    }
    let inverse = commit
        .package()
        .apply_chart(&commit.patch().inverse())
        .unwrap();
    assert!(members(&inverse.package().to_bytes().unwrap()) == original);
    assert!(members(&package.to_bytes().unwrap()) == original);
}

#[test]
fn native_svg_detach_replace_and_inverse_preserve_fallback_and_diagram_parts() {
    use litchi_docx::{
        drawing::DrawingPlacement,
        source_backed::{
            self,
            story_text::Selector,
            svg::{Input, OwnerState, PictureSelector},
        },
    };
    let source = std::fs::read(path("word-vector-diagram.docx")).unwrap();
    let original = members(&source);
    let package = source_backed::Package::from_bytes(&source).unwrap();
    let views = package.svg_pictures(Selector::Main).unwrap();
    assert_eq!(views.len(), 1);
    let view = &views[0];
    let selector = PictureSelector::new(0, 0);
    assert_eq!(view.selector(), selector);
    assert_eq!(view.placement(), DrawingPlacement::Inline);
    assert_eq!(view.owner_state(), OwnerState::Embedded);
    assert_eq!(view.raster_relationship_id(), "rId4");
    assert_eq!(view.raster_part_uri().as_str(), "/word/media/image1.png");
    assert_eq!(
        hash(&original["word/media/image1.png"]),
        "97fefe9283fc468417bf2b339b48197711947bd05cc835e40a31088b4e1e7990"
    );
    let svg = view.svg().unwrap();
    assert_eq!(svg.relationship_id(), "rId5");
    assert_eq!(svg.part_uri().as_str(), "/word/media/image2.svg");
    assert_eq!(
        hash(svg.bytes()),
        "91bdc6ffe9ddcd4ddda5c1e4872f0003b6b0864a0523c19945a4917da6858a95"
    );
    let replacement = std::str::from_utf8(svg.bytes())
        .unwrap()
        .replace("SVG_MARKER_007", "SVG_REVIEW_013");
    assert!(replacement.as_bytes() != svg.bytes());
    for replace in [false, true] {
        let mut edit = package.edit_svg([selector]).unwrap();
        assert!(edit.detach_svg(selector).unwrap());
        let detached = edit.commit().unwrap();
        let detach_inverse = detached.patch().inverse();
        let commit = if replace {
            let mut attach = detached.package().edit_svg([selector]).unwrap();
            attach
                .attach_svg(selector, Input::borrowed(replacement.as_bytes()))
                .unwrap();
            attach.commit().unwrap()
        } else {
            detached
        };
        let mut bytes = Vec::new();
        commit.package().write_to(&mut bytes).unwrap();
        let reopened = source_backed::Package::from_bytes(&bytes).unwrap();
        let after = reopened.svg_pictures(Selector::Main).unwrap();
        assert_eq!(after.len(), 1);
        assert_eq!(after[0].raster_relationship_id(), "rId4");
        if replace {
            assert!(after[0].svg().unwrap().bytes() == replacement.as_bytes());
        } else {
            assert_eq!(after[0].owner_state(), OwnerState::None);
            assert!(after[0].svg().is_none());
        }
        let updated = members(&bytes);
        for (name, data) in &original {
            if ![
                "[Content_Types].xml",
                "word/document.xml",
                "word/_rels/document.xml.rels",
                "word/media/image2.svg",
            ]
            .contains(&name.as_str())
            {
                assert!(updated.get(name) == Some(data), "changed {name}");
            }
        }
        let owner = Package::from_bytes(&bytes).unwrap();
        let diagrams = owner.smart_arts().unwrap();
        assert_eq!(diagrams[0].text(), "ACQUIRE_013\nTRANSFORM_013\nVERIFY_013");
        assert_eq!(diagrams[0].extent, Some((5486400, 3200400)));
        let inverse = commit
            .package()
            .apply_svg(&commit.patch().inverse())
            .unwrap();
        let mut restored = Vec::new();
        if replace {
            inverse
                .package()
                .apply_svg(&detach_inverse)
                .unwrap()
                .package()
                .write_to(&mut restored)
                .unwrap();
        } else {
            inverse.package().write_to(&mut restored).unwrap();
        }
        assert!(restored == source);
    }
    assert!(std::fs::read(path("word-vector-diagram.docx")).unwrap() == source);
}
