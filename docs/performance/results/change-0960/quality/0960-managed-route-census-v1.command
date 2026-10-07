cd /home/zhuhe/litchi-performance-0960/source-candidate-v1
LITCHI_0960_ROUTE_CENSUS_DIR=/home/zhuhe/litchi-performance-0960/route-census-v1 cargo +1.95.0 test --manifest-path /home/zhuhe/litchi-performance-0960/source-candidate-v1/Cargo.toml --locked --offline -p litchi-xlsx --lib managed_fixture_route_census_requires_explicit_external_input -- --ignored --nocapture
