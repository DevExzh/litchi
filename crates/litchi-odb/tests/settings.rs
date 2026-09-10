#![allow(
    clippy::unwrap_used,
    reason = "settings integration tests keep assertions concise"
)]

use litchi_odb::{
    ApplicationConnectionSettings, AutoIncrementSettings, BooleanComparisonMode,
    CharacterSetSettings, Connection, DataSourceSetting, DataSourceSettingType, Database,
    DelimiterSettings, DriverSettings, FileDatabaseTarget, Limits, LoginSettings,
    ServerDatabaseTarget, TableFilter, TableSetting, TableTypeFilter,
};

const SOURCE: &str = concat!(
    r#"<?xml version="1.0" encoding="UTF-8"?><office:document-content "#,
    r#"xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" "#,
    r#"xmlns:db="urn:oasis:names:tc:opendocument:xmlns:database:1.0" "#,
    r#"xmlns:xlink="http://www.w3.org/1999/xlink" "#,
    r#"xmlns:sdbc="urn:example:sdbc-driver" "#,
    r#"xmlns:lo="urn:example:odb-extension" office:version="1.4">"#,
    r#"<office:body><office:database><db:data-source>"#,
    r#"<db:connection-data><db:connection-resource xlink:type="simple" xlink:href="sdbc:demo"/>"#,
    r#"<db:login db:user-name="alice" db:is-password-required="true" db:login-timeout="30"/>"#,
    r#"</db:connection-data>"#,
    r#"<db:driver-settings db:show-deleted="true" db:base-dn="ou=people">"#,
    r#"<db:auto-increment db:additional-column-statement="identity" db:row-retrieving-statement="returning id"/>"#,
    r#"<db:delimiter db:field="," db:string="&quot;" db:decimal="." db:thousand=","/>"#,
    r#"<db:character-set db:encoding="UTF-8"/><db:table-settings>"#,
    r#"<db:table-setting db:is-first-row-header-line="true"><db:character-set db:encoding="ISO-8859-1"/></db:table-setting>"#,
    r#"<lo:driver-extension lo:keep="yes"/></db:table-settings></db:driver-settings>"#,
    r#"<db:application-connection-settings db:enable-sql92-check="true" db:boolean-comparison-mode="equal-boolean" db:max-row-count="100">"#,
    r#"<db:table-filter><db:table-include-filter><db:table-filter-pattern>customers*</db:table-filter-pattern></db:table-include-filter>"#,
    r#"<db:table-exclude-filter><db:table-filter-pattern>customers_tmp</db:table-filter-pattern></db:table-exclude-filter></db:table-filter>"#,
    r#"<db:table-type-filter><db:table-type>TABLE</db:table-type><db:table-type>VIEW</db:table-type></db:table-type-filter>"#,
    r#"<db:data-source-settings><db:data-source-setting db:data-source-setting-name="timeout" db:data-source-setting-type="int">"#,
    r#"<db:data-source-setting-value>15</db:data-source-setting-value></db:data-source-setting></db:data-source-settings>"#,
    r#"<lo:application-extension lo:keep="yes"/></db:application-connection-settings>"#,
    r#"</db:data-source></office:database></office:body></office:document-content>"#,
);

fn database() -> Database {
    database_from(SOURCE)
}

fn database_from(source: &str) -> Database {
    Database::from_bytes(
        litchi_odb::Builder::new()
            .content_xml(source)
            .build()
            .unwrap(),
    )
    .unwrap()
}

#[test]
fn reads_all_settings_families_without_execution() {
    let database = database();
    let settings = database.settings().unwrap();
    let login = settings.login().unwrap();
    assert_eq!(login.user_name(), Some("alice"));
    assert_eq!(login.password_required(), Some(true));
    assert_eq!(login.login_timeout(), Some(30));

    let driver = settings.driver().unwrap();
    assert_eq!(driver.show_deleted(), Some(true));
    assert_eq!(driver.base_dn(), Some("ou=people"));
    assert_eq!(
        driver.auto_increment().unwrap().row_retrieving_statement(),
        Some("returning id")
    );
    assert_eq!(driver.delimiter().unwrap().field(), Some(","));
    assert_eq!(driver.character_set().unwrap().encoding(), Some("UTF-8"));
    assert_eq!(driver.table_settings().len(), 1);
    assert_eq!(
        driver.table_settings()[0]
            .character_set()
            .unwrap()
            .encoding(),
        Some("ISO-8859-1")
    );

    let application = settings.application_connection().unwrap();
    assert_eq!(application.enable_sql92_check(), Some(true));
    assert_eq!(
        application.boolean_comparison_mode(),
        Some(BooleanComparisonMode::EqualBoolean)
    );
    assert_eq!(application.max_row_count(), Some(100));
    assert_eq!(
        application.table_filter().unwrap().include(),
        &["customers*".to_string()]
    );
    assert_eq!(
        application.table_filter().unwrap().exclude(),
        &["customers_tmp".to_string()]
    );
    assert_eq!(
        application.table_type_filter().unwrap().table_types(),
        &["TABLE".to_string(), "VIEW".to_string()]
    );
    assert_eq!(
        application.data_source_settings()[0].values(),
        &["15".to_string()]
    );
}

#[test]
fn settings_edit_reopens_and_keeps_unknown_children() {
    let source = database();
    let settings = source.settings().unwrap();
    let replacement = settings
        .clone()
        .with_login(Some(LoginSettings::new().with_user_name("bob")))
        .with_driver(Some(
            DriverSettings::new()
                .with_character_set(Some(CharacterSetSettings::new().with_encoding("UTF-8")))
                .with_auto_increment(Some(
                    AutoIncrementSettings::new().with_row_retrieving_statement("returning key"),
                ))
                .with_delimiter(Some(DelimiterSettings::new().with_field(";")))
                .with_table_settings(vec![TableSetting::new().with_show_deleted(Some(false))]),
        ))
        .with_application_connection(Some(
            ApplicationConnectionSettings::new()
                .with_boolean_comparison_mode(Some(BooleanComparisonMode::EqualInteger))
                .with_table_filter(Some(
                    TableFilter::new().with_include(vec!["orders*".into()]),
                ))
                .with_table_type_filter(Some(
                    TableTypeFilter::new().with_table_types(vec!["TABLE".into()]),
                ))
                .with_data_source_settings(vec![
                    DataSourceSetting::new("fetch-size", DataSourceSettingType::Int)
                        .with_values(vec!["100".into()]),
                ]),
        ));
    let mut edit = source.edit();
    edit.set_settings(replacement.clone()).unwrap();
    let commit = edit.commit().unwrap();
    assert!(commit.changed());
    assert!(
        commit
            .database()
            .content_xml()
            .contains("lo:driver-extension")
    );
    assert!(
        commit
            .database()
            .content_xml()
            .contains("lo:application-extension")
    );
    assert_eq!(commit.database().settings().unwrap(), &replacement);
    assert_eq!(
        commit
            .patch()
            .inverse()
            .apply(commit.database())
            .unwrap()
            .as_bytes(),
        source.as_bytes()
    );

    let reopened = Database::from_bytes(commit.database().as_bytes().to_vec()).unwrap();
    assert_eq!(reopened.settings().unwrap(), &replacement);
}

#[test]
fn fresh_builder_and_changed_publications_reopen_the_typed_projection() {
    let fresh = Database::from_bytes(litchi_odb::Builder::new().build().unwrap()).unwrap();
    assert!(fresh.settings().unwrap().is_empty());

    let mut edit = fresh.edit();
    edit.set_login_settings(Some(LoginSettings::new().with_user_name("fresh-user")))
        .unwrap();
    let commit = edit.commit().unwrap();
    let reopened = Database::from_bytes(commit.database().as_bytes().to_vec()).unwrap();
    assert_eq!(
        reopened.settings().unwrap().login().unwrap().user_name(),
        Some("fresh-user")
    );
}

#[test]
fn settings_limits_refuse_unbounded_values_before_mutation() {
    let database = database();
    let catalog = database
        .catalog_with(Limits::default().with_max_filter_patterns(0))
        .unwrap();
    let error = catalog.settings().unwrap_err();
    assert!(matches!(error, litchi_core::Error::InvalidFormat(_)));
}

#[test]
fn character_set_names_follow_odf_text_encoding_syntax() {
    for encoding in ["", "UTF 8", "8UTF", "UTF:8", "éUTF", " UTF-8", "UTF-8 "] {
        let xml = SOURCE.replace("UTF-8\"/>", &format!("{encoding}\"/>"));
        let malformed = database_from(&xml);
        assert!(
            matches!(
                malformed.settings(),
                Err(litchi_core::Error::InvalidFormat(_))
            ),
            "{encoding:?}"
        );

        for nested in [false, true] {
            let source = database();
            let charset = CharacterSetSettings::new().with_encoding(encoding);
            let driver = if nested {
                DriverSettings::new().with_table_settings(vec![
                    TableSetting::new().with_character_set(Some(charset)),
                ])
            } else {
                DriverSettings::new().with_character_set(Some(charset))
            };
            let replacement = source.settings().unwrap().clone().with_driver(Some(driver));
            let mut edit = source.edit();
            assert!(
                matches!(
                    edit.set_settings(replacement),
                    Err(litchi_core::Error::InvalidFormat(_))
                ),
                "{encoding:?}"
            );
            assert_eq!(
                edit.commit().unwrap().database().as_bytes(),
                source.as_bytes()
            );
        }
    }

    for encoding in ["UTF-8", "ISO_8859-1", "a", "x.custom-1"] {
        let source = database();
        let replacement = source
            .settings()
            .unwrap()
            .clone()
            .with_driver(Some(DriverSettings::new().with_character_set(Some(
                CharacterSetSettings::new().with_encoding(encoding),
            ))));
        let mut edit = source.edit();
        edit.set_settings(replacement).unwrap();
        let commit = edit.commit().unwrap();
        assert_eq!(
            commit
                .database()
                .settings()
                .unwrap()
                .driver()
                .unwrap()
                .character_set()
                .unwrap()
                .encoding(),
            Some(encoding)
        );
    }
}

#[test]
fn authored_setting_values_use_one_aggregate_budget() {
    let database = database();
    let values = vec!["v".to_owned(); 65_535];
    let settings = ApplicationConnectionSettings::new().with_data_source_settings(vec![
        DataSourceSetting::new("many", DataSourceSettingType::String).with_values(values),
    ]);
    let before = database.as_bytes().to_vec();
    let mut edit = database.edit();
    assert!(matches!(
        edit.set_application_connection_settings(Some(settings)),
        Err(litchi_core::Error::InvalidFormat(_))
    ));
    let committed = edit.commit().unwrap();
    assert_eq!(committed.database().as_bytes(), before.as_slice());
}

#[test]
fn tab_heavy_attribute_names_are_rejected_by_the_preflight_budget() {
    let tab_heavy = "\t".repeat(1024);
    let settings = (0..16_000)
        .map(|index| {
            DataSourceSetting::new(format!("{tab_heavy}{index}"), DataSourceSettingType::String)
                .with_values(vec!["v".to_owned()])
        })
        .collect();
    let source = database();
    let before = source.content_xml().to_owned();
    let replacement = source
        .settings()
        .unwrap()
        .clone()
        .with_application_connection(Some(
            ApplicationConnectionSettings::new().with_data_source_settings(settings),
        ));
    let mut edit = source.edit();
    assert!(matches!(
        edit.set_settings(replacement),
        Err(litchi_core::Error::InvalidFormat(_))
    ));
    assert_eq!(edit.staged_content_xml(), before);
}

#[test]
fn settings_splice_empty_containers_and_lexical_extensions() {
    let source = SOURCE
        .replace(
            r#"<db:table-filter><db:table-include-filter><db:table-filter-pattern>customers*</db:table-filter-pattern></db:table-include-filter><db:table-exclude-filter><db:table-filter-pattern>customers_tmp</db:table-filter-pattern></db:table-exclude-filter></db:table-filter>"#,
            r#"<db:table-filter lo:marker="keep"><!--retain--><lo:filter-extension lo:value="x"/></db:table-filter>"#,
        )
        .replace(
            r#"<db:table-type-filter><db:table-type>TABLE</db:table-type><db:table-type>VIEW</db:table-type></db:table-type-filter>"#,
            r#"<db:table-type-filter/>"#,
        );
    let database = database_from(&source);
    let replacement = database
        .settings()
        .unwrap()
        .clone()
        .with_application_connection(Some(
            ApplicationConnectionSettings::new()
                .with_table_filter(Some(
                    TableFilter::new().with_include(vec!["orders*".into()]),
                ))
                .with_table_type_filter(Some(
                    TableTypeFilter::new().with_table_types(vec!["TABLE".into()]),
                )),
        ));
    let mut edit = database.edit();
    edit.set_settings(replacement).unwrap();
    let content = edit.commit().unwrap().database().content_xml().to_owned();
    assert!(content.contains(r#"lo:marker="keep""#));
    assert!(content.contains(r#"<!--retain-->"#));
    assert!(content.contains(r#"<lo:filter-extension lo:value="x"/>"#));
    assert!(content.contains("orders*"));
    assert!(content.contains("<db:table-type>TABLE</db:table-type>"));
}

#[test]
fn settings_replacement_keeps_comments_and_foreign_pattern_children() {
    let source = SOURCE.replace(
        "<db:table-filter-pattern>customers*</db:table-filter-pattern>",
        r#"<db:table-filter-pattern lo:marker="keep"><!--retain-->customers*</db:table-filter-pattern>"#,
    );
    let database = database_from(&source);
    let replacement = database
        .settings()
        .unwrap()
        .clone()
        .with_application_connection(Some(
            ApplicationConnectionSettings::new().with_table_filter(Some(
                TableFilter::new().with_include(vec!["orders*".into()]),
            )),
        ));
    let mut edit = database.edit();
    edit.set_settings(replacement).unwrap();
    let content = edit.commit().unwrap().database().content_xml().to_owned();
    assert!(content.contains(r#"lo:marker="keep""#));
    assert!(content.contains(r#"<!--retain-->"#));
    assert!(content.contains("orders*"));
}

#[test]
fn login_only_edit_leaves_other_families_byte_exact() {
    let database = database();
    let original = database.content_xml().to_owned();
    let driver_start = original.find("<db:driver-settings").unwrap();
    let driver_end = original[driver_start..]
        .find("</db:driver-settings>")
        .map(|offset| driver_start + offset + "</db:driver-settings>".len())
        .unwrap();
    let driver = &original[driver_start..driver_end];
    let mut edit = database.edit();
    edit.set_login_settings(Some(LoginSettings::new().with_user_name("bob")))
        .unwrap();
    let content = edit.commit().unwrap().database().content_xml().to_owned();
    assert!(content.contains(driver));
}

#[test]
fn settings_updates_preserve_unknown_markup_inside_repeated_items() {
    let source = SOURCE
        .replace(
            r#"<db:table-setting db:is-first-row-header-line="true"><db:character-set db:encoding="ISO-8859-1"/></db:table-setting>"#,
            r#"<db:table-setting db:is-first-row-header-line="true" lo:marker="table"><db:character-set db:encoding="ISO-8859-1"/><lo:table-extension/></db:table-setting>"#,
        )
        .replace(
            r#"<db:data-source-setting db:data-source-setting-name="timeout" db:data-source-setting-type="int"><db:data-source-setting-value>15</db:data-source-setting-value></db:data-source-setting>"#,
            r#"<db:data-source-setting db:data-source-setting-name="timeout" db:data-source-setting-type="int" lo:marker="setting"><db:data-source-setting-value>15</db:data-source-setting-value><lo:setting-extension/></db:data-source-setting>"#,
        );
    let database = database_from(&source);
    let replacement = database
        .settings()
        .unwrap()
        .clone()
        .with_driver(Some(
            database
                .settings()
                .unwrap()
                .driver()
                .unwrap()
                .clone()
                .with_table_settings(vec![TableSetting::new().with_show_deleted(Some(false))]),
        ))
        .with_application_connection(Some(
            database
                .settings()
                .unwrap()
                .application_connection()
                .unwrap()
                .clone()
                .with_data_source_settings(vec![
                    DataSourceSetting::new("timeout", DataSourceSettingType::Int)
                        .with_values(vec!["30".into()]),
                ]),
        ));
    let mut edit = database.edit();
    edit.set_settings(replacement).unwrap();
    let content = edit.commit().unwrap().database().content_xml().to_owned();
    assert!(content.contains(r#"lo:marker="table""#));
    assert!(content.contains("<lo:table-extension/>"));
    assert!(content.contains(r#"lo:marker="setting""#));
    assert!(content.contains("<lo:setting-extension/>"));
    assert!(content.contains(">30</db:data-source-setting-value>"));
}

#[test]
fn removing_modeled_filters_rejects_an_invalid_required_container_atomically() {
    let source = SOURCE
        .replace(
            r#"<db:table-filter><db:table-include-filter><db:table-filter-pattern>customers*</db:table-filter-pattern></db:table-include-filter><db:table-exclude-filter><db:table-filter-pattern>customers_tmp</db:table-filter-pattern></db:table-exclude-filter></db:table-filter>"#,
            r#"<db:table-filter><db:table-include-filter lo:marker="keep"><db:table-filter-pattern>customers*</db:table-filter-pattern><lo:extension/></db:table-include-filter></db:table-filter>"#,
        )
        .replace(
            r#"<db:table-type-filter><db:table-type>TABLE</db:table-type><db:table-type>VIEW</db:table-type></db:table-type-filter>"#,
            r#"<db:table-type-filter/>"#,
        );
    let database = database_from(&source);
    let mut application = database
        .settings()
        .unwrap()
        .application_connection()
        .unwrap()
        .clone();
    application = application
        .with_table_filter(None)
        .with_table_type_filter(None);
    let before = database.as_bytes().to_vec();
    let mut edit = database.edit();
    assert!(
        edit.set_application_connection_settings(Some(application))
            .is_err()
    );
    assert!(edit.staged_content_xml().contains(r#"lo:marker="keep""#));
    assert!(edit.staged_content_xml().contains("<lo:extension/>"));
    assert_eq!(
        edit.commit().unwrap().database().as_bytes(),
        before.as_slice()
    );
}

#[test]
fn settings_writer_reuses_an_aliased_database_namespace() {
    let source = SOURCE.replace("xmlns:db=", "xmlns:d=").replace("db:", "d:");
    let database = database_from(&source);
    let mut edit = database.edit();
    edit.set_login_settings(Some(LoginSettings::new().with_user_name("bob")))
        .unwrap();
    let content = edit.commit().unwrap().database().content_xml().to_owned();
    assert!(content.contains("<d:login"));
    assert!(!content.contains("<db:login"));
}

#[test]
fn repeated_child_insertion_uses_the_real_close_tag_after_comment_text() {
    let source = database_from(&SOURCE.replace(
        r#"<db:delimiter db:field="," db:string="&quot;" db:decimal="." db:thousand=","/>"#,
        r#"<!-- producer text contains </db:not-a-real-element> -->"#,
    ));
    let replacement = source.settings().unwrap().clone().with_driver(Some(
        source
            .settings()
            .unwrap()
            .driver()
            .unwrap()
            .clone()
            .with_delimiter(Some(DelimiterSettings::new().with_field(";"))),
    ));
    let mut edit = source.edit();
    edit.set_settings(replacement).unwrap();
    let content = edit.commit().unwrap().database().content_xml().to_owned();
    let comment = "<!-- producer text contains </db:not-a-real-element> -->";
    let comment_end = content.find(comment).unwrap() + comment.len();
    let delimiter = content.find("<db:delimiter").unwrap();
    assert!(delimiter >= comment_end);
    assert_eq!(
        edit_content_settings(&content)
            .driver()
            .unwrap()
            .delimiter()
            .unwrap()
            .field(),
        Some(";")
    );
}

fn edit_content_settings(content: &str) -> litchi_odb::DatabaseSettings {
    database_from(content).settings().unwrap().clone()
}

#[test]
fn settings_reader_rejects_invalid_typed_children_and_empty_required_groups() {
    let text_in_login = SOURCE.replace(
        r#"<db:login db:user-name="alice" db:is-password-required="true" db:login-timeout="30"/>"#,
        r#"<db:login db:user-name="alice">unexpected</db:login>"#,
    );
    assert!(database_from(&text_in_login).settings().is_err());

    let empty_include = SOURCE.replace(
        r#"<db:table-include-filter><db:table-filter-pattern>customers*</db:table-filter-pattern></db:table-include-filter>"#,
        r#"<db:table-include-filter/>"#,
    );
    assert!(database_from(&empty_include).settings().is_err());

    let empty_settings = SOURCE.replace(
        r#"<db:data-source-settings><db:data-source-setting db:data-source-setting-name="timeout" db:data-source-setting-type="int"><db:data-source-setting-value>15</db:data-source-setting-value></db:data-source-setting></db:data-source-settings>"#,
        r#"<db:data-source-settings/>"#,
    );
    assert!(database_from(&empty_settings).settings().is_err());

    let empty_types = SOURCE.replace(
        r#"<db:table-type-filter><db:table-type>TABLE</db:table-type><db:table-type>VIEW</db:table-type></db:table-type-filter>"#,
        r#"<db:table-type-filter/>"#,
    );
    assert!(
        database_from(&empty_types)
            .settings()
            .unwrap()
            .application_connection()
            .unwrap()
            .table_type_filter()
            .unwrap()
            .table_types()
            .is_empty()
    );

    let empty_values = SOURCE.replace(
        r#"<db:data-source-setting db:data-source-setting-name="timeout" db:data-source-setting-type="int"><db:data-source-setting-value>15</db:data-source-setting-value></db:data-source-setting>"#,
        r#"<db:data-source-setting db:data-source-setting-name="timeout" db:data-source-setting-type="int"/>"#,
    );
    assert!(database_from(&empty_values).settings().is_err());

    let duplicate_table_child = SOURCE.replace(
        r#"<db:character-set db:encoding="ISO-8859-1"/>"#,
        r#"<db:delimiter/><db:delimiter/><db:character-set db:encoding="ISO-8859-1"/>"#,
    );
    assert!(database_from(&duplicate_table_child).settings().is_err());
}

#[test]
fn settings_reader_does_not_use_opaque_markup_to_satisfy_one_or_more() {
    for replacement in [
        r#"<db:table-include-filter><!--comment-only--></db:table-include-filter>"#,
        r#"<db:table-exclude-filter><lo:foreign-pattern/></db:table-exclude-filter>"#,
    ] {
        let source = SOURCE
            .replace(
                r#"<db:table-include-filter><db:table-filter-pattern>customers*</db:table-filter-pattern></db:table-include-filter>"#,
                replacement,
            )
            .replace(
                r#"<db:table-exclude-filter><db:table-filter-pattern>customers_tmp</db:table-filter-pattern></db:table-exclude-filter>"#,
                replacement,
            );
        assert!(database_from(&source).settings().is_err());
    }

    let foreign_settings = SOURCE.replace(
        r#"<db:data-source-settings><db:data-source-setting db:data-source-setting-name="timeout" db:data-source-setting-type="int"><db:data-source-setting-value>15</db:data-source-setting-value></db:data-source-setting></db:data-source-settings>"#,
        r#"<db:data-source-settings><!--comment-only--><lo:foreign-setting/></db:data-source-settings>"#,
    );
    assert!(database_from(&foreign_settings).settings().is_err());

    let foreign_value = SOURCE.replace(
        r#"<db:data-source-setting db:data-source-setting-name="timeout" db:data-source-setting-type="int"><db:data-source-setting-value>15</db:data-source-setting-value></db:data-source-setting>"#,
        r#"<db:data-source-setting db:data-source-setting-name="timeout" db:data-source-setting-type="int"><lo:foreign-value/></db:data-source-setting>"#,
    );
    assert!(database_from(&foreign_value).settings().is_err());
}

#[test]
fn settings_reader_requires_typed_connection_resource_attributes_and_empty_content() {
    for target in [
        r#"<db:connection-resource xlink:href="sdbc:demo"/>"#,
        r#"<db:connection-resource xlink:type="extended" xlink:href="sdbc:demo"/>"#,
        r#"<db:connection-resource xlink:type="simple"/>"#,
        r#"<db:connection-resource xlink:type="simple" xlink:href="sdbc:demo"><lo:foreign-child/></db:connection-resource>"#,
    ] {
        let source = SOURCE.replace(
            r#"<db:connection-resource xlink:type="simple" xlink:href="sdbc:demo"/>"#,
            target,
        );
        assert!(database_from(&source).settings().is_err(), "{target}");
    }

    let empty_href = SOURCE.replace(
        r#"<db:connection-resource xlink:type="simple" xlink:href="sdbc:demo"/>"#,
        r#"<db:connection-resource xlink:type="simple" xlink:href=""/>"#,
    );
    assert!(database_from(&empty_href).settings().is_ok());
}

#[test]
fn changed_login_cannot_publish_an_invalid_connection_resource() {
    let malformed = database_from(&SOURCE.replace(
        r#"<db:connection-resource xlink:type="simple" xlink:href="sdbc:demo"/>"#,
        r#"<db:connection-resource xlink:href="sdbc:demo"/>"#,
    ));
    assert!(malformed.settings().is_err());
    let before = malformed.as_bytes().to_vec();
    let mut edit = malformed.edit();
    assert!(
        edit.set_login_settings(Some(LoginSettings::new().with_user_name("bob")))
            .is_err()
    );
    assert_eq!(
        edit.commit().unwrap().database().as_bytes(),
        before.as_slice()
    );
}

#[test]
fn settings_reader_requires_exactly_one_database_description_target() {
    let connection_resource =
        r#"<db:connection-resource xlink:type="simple" xlink:href="sdbc:demo"/>"#;
    for owner in [
        r#"<db:database-description/>"#,
        r#"<db:database-description><lo:foreign-target/></db:database-description>"#,
        r#"<db:database-description><db:file-based-database xlink:type="simple" xlink:href="file:///demo"/><db:server-database db:hostname="db" db:database-name="demo"/></db:database-description>"#,
    ] {
        let source = SOURCE.replace(connection_resource, owner);
        assert!(database_from(&source).settings().is_err(), "{owner}");
    }

    let valid_file = SOURCE.replace(
        connection_resource,
        r#"<db:database-description><db:file-based-database xlink:type="simple" xlink:href="file:///demo" db:media-type="application/octet-stream"/></db:database-description>"#,
    );
    assert!(database_from(&valid_file).settings().is_ok());

    let valid_server = SOURCE.replace(
        connection_resource,
        r#"<db:database-description><db:server-database db:type="sdbc:generic" db:hostname="db" db:database-name="demo"/></db:database-description>"#,
    );
    assert!(database_from(&valid_server).settings().is_ok());
}

#[test]
fn normative_connection_targets_read_and_write_without_opening_drivers() {
    let connection_resource =
        r#"<db:connection-resource xlink:type="simple" xlink:href="sdbc:demo"/>"#;
    let file_source = SOURCE.replace(
        connection_resource,
        r#"<db:database-description><db:file-based-database xlink:type="simple" xlink:href="file:///demo" db:media-type="application/octet-stream" db:extension="db"/></db:database-description>"#,
    );
    let file_database = database_from(&file_source);
    assert!(matches!(
        file_database.catalog().unwrap().connection(),
        Some(Connection::FileTarget(target))
            if target.href() == "file:///demo"
                && target.media_type() == "application/octet-stream"
                && target.extension() == Some("db")
    ));
    assert!(file_database.settings().is_ok());

    let existing_file = file_database.catalog().unwrap().connection().cloned();
    let mut no_op_edit = file_database.edit();
    no_op_edit.set_connection(existing_file).unwrap();
    let no_op = no_op_edit.commit().unwrap();
    assert!(!no_op.changed());
    assert_eq!(no_op.database().as_bytes(), file_database.as_bytes());

    let fresh_source = database();
    let mut fresh_file_edit = fresh_source.edit();
    fresh_file_edit
        .set_connection(Some(Connection::file_with_media_type(
            "file:///fresh",
            "application/vnd.demo",
        )))
        .unwrap();
    let fresh_file = fresh_file_edit.commit().unwrap().into_database();
    assert!(fresh_file.content_xml().contains("xlink:type=\"simple\""));
    assert!(
        fresh_file
            .content_xml()
            .contains("db:media-type=\"application/vnd.demo\"")
    );
    assert!(matches!(
        fresh_file.catalog().unwrap().connection(),
        Some(Connection::FileTarget(target))
            if target.href() == "file:///fresh"
                && target.media_type() == "application/vnd.demo"
    ));
    assert!(fresh_file.settings().is_ok());

    let mut file_edit = file_database.edit();
    file_edit
        .set_connection(Some(Connection::server_with_local_socket_namespace(
            "sdbc:firebird",
            "urn:example:sdbc-driver",
            "/run/firebird.sock",
            None,
        )))
        .unwrap();
    let changed = file_edit.commit().unwrap().into_database();
    assert!(changed.content_xml().contains("db:type=\"sdbc:firebird\""));
    assert!(
        changed
            .content_xml()
            .contains("db:local-socket=\"/run/firebird.sock\"")
    );
    assert!(!changed.content_xml().contains("db:database-name="));
    assert!(matches!(
        changed.catalog().unwrap().connection(),
        Some(Connection::ServerTarget(target))
            if target.database_type() == "sdbc:firebird"
                && target.database_name().is_none()
    ));
    assert!(changed.settings().is_ok());

    let mut host_edit = changed.edit();
    host_edit
        .set_connection(Some(Connection::ServerTarget(
            ServerDatabaseTarget::new("sdbc:generic")
                .with_database_type_namespace("urn:example:sdbc-driver")
                .with_host("db.example.test", Some(5432))
                .with_database_name(Some("ledger".to_owned())),
        )))
        .unwrap();
    let host_database = host_edit.commit().unwrap().into_database();
    assert!(host_database.settings().is_ok());
    assert!(host_database.content_xml().contains("db:port=\"5432\""));

    let opaque_source = SOURCE.replace(
        connection_resource,
        r#"<db:database-description lo:description="retain"><db:file-based-database xlink:type="simple" xlink:href="file:///demo" db:media-type="application/octet-stream" lo:target="retain"/></db:database-description>"#,
    );
    let opaque_database = database_from(&opaque_source);
    let mut opaque_edit = opaque_database.edit();
    opaque_edit
        .set_connection(Some(Connection::server_with_type_namespace(
            "sdbc:generic",
            "urn:example:sdbc-driver",
            "db.example.test",
            None,
            None,
        )))
        .unwrap();
    let opaque_commit = opaque_edit.commit().unwrap();
    let opaque_changed = opaque_commit.database();
    assert!(
        opaque_changed
            .content_xml()
            .contains(r#"lo:description="retain""#)
    );
    assert!(
        opaque_changed
            .content_xml()
            .contains(r#"lo:target="retain""#)
    );
    assert_eq!(
        opaque_commit
            .patch()
            .inverse()
            .apply(opaque_changed)
            .unwrap()
            .as_bytes(),
        opaque_database.as_bytes()
    );

    let local_namespace_source = SOURCE.replace(
        connection_resource,
        r#"<db:database-description><db:file-based-database xmlns:ext="urn:example:local" ext:target="retain" xlink:type="simple" xlink:href="file:///demo" db:media-type="application/octet-stream"/></db:database-description>"#,
    );
    let local_namespace_database = database_from(&local_namespace_source);
    let mut local_namespace_edit = local_namespace_database.edit();
    local_namespace_edit
        .set_connection(Some(Connection::server_with_local_socket_namespace(
            "sdbc:generic",
            "urn:example:sdbc-driver",
            "/run/demo.sock",
            None,
        )))
        .unwrap();
    let local_namespace_changed = local_namespace_edit.commit().unwrap();
    assert!(
        local_namespace_changed
            .database()
            .content_xml()
            .contains(r#"xmlns:ext="urn:example:local" ext:target="retain""#)
    );

    let empty_description_source = SOURCE.replace(
        connection_resource,
        r#"<db:database-description xmlns:ext="urn:example:empty" ext:owner="retain"/>"#,
    );
    let empty_description_database = database_from(&empty_description_source);
    let mut empty_description_edit = empty_description_database.edit();
    empty_description_edit
        .set_connection(Some(Connection::Resource("sdbc:recovered".to_owned())))
        .unwrap();
    let empty_description_changed = empty_description_edit.commit().unwrap();
    assert!(
        empty_description_changed
            .database()
            .content_xml()
            .contains(r#"xmlns:ext="urn:example:empty" ext:owner="retain""#)
    );
}

#[test]
fn connection_resource_conversion_carries_ancestor_extension_bindings() {
    let source = SOURCE
        .replace(r#"xmlns:lo="urn:example:odb-extension" "#, "")
        .replace(r#"<lo:driver-extension lo:keep="yes"/>"#, "")
        .replace(r#"<lo:application-extension lo:keep="yes"/>"#, "")
        .replace(
            r#"<db:connection-resource xlink:type="simple" xlink:href="sdbc:demo"/>"#,
            r#"<db:database-description xmlns:lo="urn:example:ancestor"><db:file-based-database xlink:type="simple" xlink:href="file:///demo" db:media-type="application/octet-stream" lo:target="retain"/></db:database-description>"#,
        );
    let database = database_from(&source);
    let mut edit = database.edit();
    edit.set_connection(Some(Connection::Resource("sdbc:recovered".to_owned())))
        .unwrap();
    let commit = edit.commit().unwrap();
    let content = commit.database().content_xml();
    assert!(content.contains(r#"xmlns:lo="urn:example:ancestor""#));
    assert!(content.contains(r#"lo:target="retain""#));
    assert!(commit.database().settings().is_ok());
}

#[test]
fn connection_target_conversion_carries_locally_owned_xlink_binding() {
    let source = SOURCE
        .replace(r#"xmlns:xlink="http://www.w3.org/1999/xlink" "#, "")
        .replace(
            r#"<db:connection-resource xlink:type="simple" xlink:href="sdbc:demo"/>"#,
            r#"<db:connection-resource xmlns:x="http://www.w3.org/1999/xlink" x:type="simple" x:href="sdbc:demo"/>"#,
        );
    let database = database_from(&source);
    let mut edit = database.edit();
    edit.set_connection(Some(Connection::Resource("sdbc:recovered".to_owned())))
        .unwrap();
    let committed = edit.commit().unwrap().into_database();
    assert!(
        committed
            .content_xml()
            .contains(r#"<db:connection-resource x:href="sdbc:recovered" x:type="simple""#)
    );
    assert!(
        committed
            .content_xml()
            .contains(r#"xmlns:x="http://www.w3.org/1999/xlink""#)
    );
    let reopened = Database::from_bytes(committed.as_bytes().to_vec()).unwrap();
    assert_eq!(
        reopened.catalog().unwrap().connection().unwrap().clone(),
        Connection::Resource("sdbc:recovered".to_owned())
    );
}

#[test]
fn connection_target_content_blocks_all_target_replacements_atomically() {
    let source = SOURCE.replace(
        r#"<db:connection-resource xlink:type="simple" xlink:href="sdbc:demo"/>"#,
        r#"<db:connection-resource xlink:type="simple" xlink:href="sdbc:demo"><lo:foreign-child/></db:connection-resource>"#,
    );
    let database = database_from(&source);
    let before = database.content_xml().to_owned();
    let mut edit = database.edit();
    assert!(
        edit.set_connection(Some(Connection::file_with_media_type(
            "file:///replacement",
            "application/octet-stream",
        )))
        .is_err()
    );
    assert_eq!(edit.staged_content_xml(), before);
    assert_eq!(edit.commit().unwrap().database().content_xml(), before);
}

#[test]
fn reserved_odf_names_are_not_accepted_as_opaque_target_markup() {
    for target in [
        r#"<db:connection-resource xlink:type="simple" xlink:href="sdbc:demo" db:bogus="x"/>"#,
        r#"<db:connection-resource xlink:type="simple" xlink:href="sdbc:demo" lo:keep="x" db:bogus="x"/>"#,
        r#"<db:connection-resource xlink:type="simple" xlink:href="sdbc:demo" xlink:bogus="x"/>"#,
        r#"<db:connection-resource xlink:type="simple" xlink:href="sdbc:demo"><office:bogus/></db:connection-resource>"#,
    ] {
        let database = database_from(&SOURCE.replace(
            r#"<db:connection-resource xlink:type="simple" xlink:href="sdbc:demo"/>"#,
            target,
        ));
        assert!(database.catalog().is_err(), "catalog accepted {target}");
        assert!(database.settings().is_err(), "settings accepted {target}");
        let before = database.content_xml().to_owned();
        let mut edit = database.edit();
        assert!(
            edit.set_connection(Some(Connection::Resource("sdbc:replacement".to_owned())))
                .is_err()
        );
        assert_eq!(edit.staged_content_xml(), before);
    }
}

#[test]
fn server_type_qnames_require_binding_or_explicit_authoring_namespace() {
    let unbound = SOURCE.replace(
        r#"<db:connection-resource xlink:type="simple" xlink:href="sdbc:demo"/>"#,
        r#"<db:database-description><db:server-database db:type="nope:generic" db:hostname="db"/></db:database-description>"#,
    );
    let unbound = database_from(&unbound);
    assert!(unbound.catalog().is_err());
    assert!(unbound.settings().is_err());

    let unbound_database = database();
    let mut unbound_edit = unbound_database.edit();
    assert!(
        unbound_edit
            .set_connection(Some(Connection::server_with_type(
                "sdbc:generic",
                "db.example.test",
                None,
                None,
            )))
            .is_err()
    );

    let database = database();
    let mut edit = database.edit();
    edit.set_connection(Some(Connection::server_with_type_namespace(
        "vendor:generic",
        "urn:example:vendor-db",
        "db.example.test",
        None,
        None,
    )))
    .unwrap();
    let committed = edit.commit().unwrap().into_database();
    assert!(
        committed
            .content_xml()
            .contains(r#"xmlns:vendor="urn:example:vendor-db""#)
    );
    assert_eq!(
        committed.catalog().unwrap().connection().unwrap().clone(),
        Connection::server_with_type_namespace(
            "vendor:generic",
            "urn:example:vendor-db",
            "db.example.test",
            None,
            None,
        )
    );
}

#[test]
fn server_target_without_address_is_normative_and_reversible() {
    let target = ServerDatabaseTarget::new("vendor:generic")
        .with_database_type_namespace("urn:example:vendor-db")
        .with_database_name(Some("ledger".to_owned()));
    let database = database();
    let mut edit = database.edit();
    edit.set_connection(Some(Connection::ServerTarget(target.clone())))
        .unwrap();
    let committed = edit.commit().unwrap().into_database();
    let content = committed.content_xml();
    assert!(content.contains(r#"db:type="vendor:generic""#));
    assert!(content.contains(r#"db:database-name="ledger""#));
    assert!(!content.contains("db:hostname="));
    assert!(!content.contains("db:local-socket="));
    let reopened = Database::from_bytes(committed.as_bytes().to_vec()).unwrap();
    assert_eq!(
        reopened.catalog().unwrap().connection(),
        Some(&Connection::ServerTarget(target))
    );
    assert!(reopened.settings().is_ok());
}

#[test]
fn numeric_attribute_references_preserve_login_control_characters() {
    let source = database_from(&SOURCE.replace(
        r#"db:user-name="alice""#,
        r#"db:user-name="a&#x9;b&#xA;c&#xD;d""#,
    ));
    let expected = "a\tb\nc\rd";
    assert_eq!(
        source.settings().unwrap().login().unwrap().user_name(),
        Some(expected)
    );
    let mut edit = source.edit();
    edit.set_login_settings(Some(
        LoginSettings::new()
            .with_user_name(expected)
            .with_password_required(Some(false)),
    ))
    .unwrap();
    let committed = edit.commit().unwrap().into_database();
    assert!(
        committed
            .content_xml()
            .contains(r#"db:user-name="a&#x9;b&#xA;c&#xD;d""#)
    );
    assert_eq!(
        committed.settings().unwrap().login().unwrap().user_name(),
        Some(expected)
    );
}

#[test]
fn text_node_carriage_returns_are_numeric_and_reopen_equal() {
    let expected = "a\tb\nc\rd";
    let replacement = database()
        .settings()
        .unwrap()
        .clone()
        .with_application_connection(Some(
            ApplicationConnectionSettings::new()
                .with_table_filter(Some(
                    TableFilter::new().with_include(vec![expected.to_owned()]),
                ))
                .with_table_type_filter(Some(
                    TableTypeFilter::new().with_table_types(vec![expected.to_owned()]),
                ))
                .with_data_source_settings(vec![
                    DataSourceSetting::new("control", DataSourceSettingType::String)
                        .with_values(vec![expected.to_owned()]),
                ]),
        ));
    let source = database();
    let mut edit = source.edit();
    edit.set_settings(replacement.clone()).unwrap();
    let committed = edit.commit().unwrap().into_database();
    assert!(committed.content_xml().contains("a\tb\nc&#xD;d"));
    let reopened = Database::from_bytes(committed.as_bytes().to_vec()).unwrap();
    assert_eq!(reopened.settings().unwrap(), &replacement);
}

#[test]
fn connection_target_validation_failure_is_atomic() {
    let database = database();
    let before = database.content_xml().to_owned();
    let mut edit = database.edit();
    let invalid = Connection::ServerTarget(ServerDatabaseTarget::new("not-a-qname"));
    assert!(edit.set_connection(Some(invalid)).is_err());
    assert_eq!(edit.staged_content_xml(), before);
    assert_eq!(edit.commit().unwrap().database().content_xml(), before);

    let mut second = database.edit();
    let oversized = Connection::FileTarget(FileDatabaseTarget::new(
        "file:///demo",
        "x".repeat(1024 * 1024 + 1),
    ));
    assert!(second.set_connection(Some(oversized)).is_err());
    assert_eq!(second.staged_content_xml(), before);
}

#[test]
fn connection_no_op_keeps_legacy_source_without_claiming_typed_validity() {
    let legacy_source = SOURCE.replace(
        r#" office:version="1.4""#,
        "",
    )
    .replace(
        r#"<db:connection-resource xlink:type="simple" xlink:href="sdbc:demo"/>"#,
        r#"<db:database-description><db:file-based-database xlink:href="file:///legacy"/></db:database-description>"#,
    );
    let database = database_from(&legacy_source);
    assert!(matches!(
        database.catalog().unwrap().connection(),
        Some(Connection::File(value)) if value == "file:///legacy"
    ));
    let mut edit = database.edit();
    edit.set_connection(Some(Connection::File("file:///legacy".to_owned())))
        .unwrap();
    let commit = edit.commit().unwrap();
    assert!(!commit.changed());
    assert_eq!(commit.database().as_bytes(), database.as_bytes());
}

#[test]
fn normative_connection_target_grammar_rejects_missing_and_conflicting_fields() {
    let connection_resource =
        r#"<db:connection-resource xlink:type="simple" xlink:href="sdbc:demo"/>"#;
    for target in [
        r#"<db:database-description><db:file-based-database xlink:href="file:///demo" db:media-type="application/octet-stream"/></db:database-description>"#,
        r#"<db:database-description><db:file-based-database xlink:type="simple" xlink:href="file:///demo"/></db:database-description>"#,
        r#"<db:database-description><db:server-database db:hostname="db" db:database-name="demo"/></db:database-description>"#,
        r#"<db:database-description><db:server-database db:type="sdbc:generic" db:hostname="db" db:local-socket="socket"/></db:database-description>"#,
        r#"<db:database-description><db:server-database db:type="sdbc:generic" db:port="5432"/></db:database-description>"#,
        r#"<db:database-description><db:server-database db:type="sdbc:generic" db:port="0" db:hostname="db"/></db:database-description>"#,
        r#"<db:database-description><db:server-database db:type="s dbc:generic" db:local-socket="socket"/></db:database-description>"#,
        r#"<db:database-description><db:server-database db:type="sdbc:generic"><lo:foreign/></db:server-database></db:database-description>"#,
    ] {
        let source = SOURCE.replace(connection_resource, target);
        let database = database_from(&source);
        assert!(database.catalog().is_err(), "catalog accepted {target}");
        assert!(database.settings().is_err(), "settings accepted {target}");
    }
}

#[test]
fn settings_reader_applies_schema_whitespace_collapse_to_tokens() {
    let source = SOURCE
        .replace(
            r#"db:enable-sql92-check="true""#,
            "db:enable-sql92-check=\"&#x20;true&#x9;\"",
        )
        .replace(
            r#"db:boolean-comparison-mode="equal-boolean""#,
            "db:boolean-comparison-mode=\"&#x20;equal-boolean&#x20;\"",
        );
    let settings = database_from(&source).settings().unwrap().clone();
    let application = settings.application_connection().unwrap();
    assert_eq!(application.enable_sql92_check(), Some(true));
    assert_eq!(
        application.boolean_comparison_mode(),
        Some(BooleanComparisonMode::EqualBoolean)
    );

    let numeric = database_from(&SOURCE.replace(
        r#"db:show-deleted="true""#,
        "db:show-deleted=\"&#x9;1&#xA;\"",
    ));
    assert_eq!(
        numeric.settings().unwrap().driver().unwrap().show_deleted(),
        Some(true)
    );
}

#[test]
fn settings_reader_requires_the_normative_connection_owner_and_target() {
    let connection_data = r#"<db:connection-data><db:connection-resource xlink:type="simple" xlink:href="sdbc:demo"/><db:login db:user-name="alice" db:is-password-required="true" db:login-timeout="30"/></db:connection-data>"#;
    let missing_owner = SOURCE.replace(connection_data, "");
    assert!(database_from(&missing_owner).settings().is_err());

    let missing_target = SOURCE.replace(
        r#"<db:connection-resource xlink:type="simple" xlink:href="sdbc:demo"/>"#,
        "",
    );
    assert!(database_from(&missing_target).settings().is_err());

    let duplicate_target = SOURCE.replace(
        r#"<db:connection-resource xlink:type="simple" xlink:href="sdbc:demo"/>"#,
        r#"<db:connection-resource xlink:type="simple" xlink:href="sdbc:demo"/><db:database-description/>"#,
    );
    assert!(database_from(&duplicate_target).settings().is_err());
}

#[test]
fn failed_settings_edit_cannot_republish_a_source_without_connection_data() {
    let connection_data = r#"<db:connection-data><db:connection-resource xlink:type="simple" xlink:href="sdbc:demo"/><db:login db:user-name="alice" db:is-password-required="true" db:login-timeout="30"/></db:connection-data>"#;
    let malformed = database_from(&SOURCE.replace(connection_data, ""));
    let before = malformed.content_xml().to_owned();
    let mut edit = malformed.edit();
    assert!(
        edit.set_login_settings(Some(LoginSettings::new().with_user_name("bob")))
            .is_err()
    );
    let committed = edit.commit().unwrap();
    assert_eq!(committed.database().content_xml(), before);
}

#[test]
fn settings_reader_has_dedicated_table_setting_and_integer_limits() {
    let database = database();
    let catalog = database
        .catalog_with(Limits::default().with_max_table_settings(0))
        .unwrap();
    assert!(catalog.settings().is_err());

    let login_overflow = SOURCE.replace(
        r#"db:login-timeout="30""#,
        r#"db:login-timeout="18446744073709551616""#,
    );
    assert!(database_from(&login_overflow).settings().is_err());

    let row_count_overflow = SOURCE.replace(
        r#"db:max-row-count="100""#,
        r#"db:max-row-count="9223372036854775808""#,
    );
    assert!(database_from(&row_count_overflow).settings().is_err());
}

#[test]
fn settings_reader_bounds_entity_expansion_before_text_growth() {
    let expanded = "&amp;".repeat(5);
    let source = format!(
        r#"<?xml version="1.0"?><office:document-content xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:db="urn:oasis:names:tc:opendocument:xmlns:database:1.0" xmlns:xlink="http://www.w3.org/1999/xlink"><office:body><office:database><db:data-source><db:connection-data><db:connection-resource xlink:href="" xlink:type="simple"/></db:connection-data><db:application-connection-settings><db:table-filter><db:table-include-filter><db:table-filter-pattern>{expanded}</db:table-filter-pattern></db:table-include-filter></db:table-filter></db:application-connection-settings></db:data-source></office:database></office:body></office:document-content>"#
    );
    let database = database_from(&source);
    assert!(
        database
            .catalog_with(Limits::default().with_max_attribute_bytes(4))
            .unwrap()
            .settings()
            .is_err()
    );
}
