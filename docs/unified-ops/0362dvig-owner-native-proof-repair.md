# 0362dvig — owner-native proof repair validation receipt

This record binds the owner timing/proof repair at source commit
`9d85708aed4f81c81570638c833d7261fb4c28da`.  It records source-backed
correctness and compile checks only.  The `qualified` values in native route
reports are per-operation probe results; they are not a paired performance
qualification claim.

The reduced machine-readable receipt below points to external logs and native
reports.  No measurement artifact is committed here.

```json
{
  "schema": "litchi.drawingml.owner-native-proof-repair.v1",
  "source": {
    "commit": "9d85708aed4f81c81570638c833d7261fb4c28da",
    "probe": {
      "path": "tools/perf-baseline/src/drawingml_raw_probe.rs",
      "bytes": 197929,
      "sha256": "1916e6940b00f719883805e9d1a4567f7947c8cc74c1e37b05b46115dc6f220a"
    },
    "validator": {
      "path": "tools/native_qualification/drawingml_raw_validator.py",
      "bytes": 71129,
      "sha256": "ef1319b3bef23dfe06719ddd4c1259e2b72e42b760a99960d3ad772562b4ef57"
    },
    "manifest": {
      "path": "tools/native_qualification/manifests/drawingml-raw.json",
      "bytes": 11399,
      "sha256": "a94e262d2fd2d9055de1036f2c87a31cc198142002bc27fec92f586eda2283e8"
    },
    "plan": {
      "path": "docs/unified-ops/results/0362dvie/qualification-plan.json",
      "bytes": 38152,
      "sha256": "e0cb370e0399c072de9fe4b141ceff682b87d7936bec5332cedc9a181fc13505",
      "source": "/home/zhuhe/code/litchi-worktrees/uo2-coordinator-metadata-integration"
    }
  },
  "lock_pins": {
    "historical_workspace": {
      "sha256": "52ed0cacac15b5e440349bfc0f4aca2665bcf50d59d4321624e596bb410783c6",
      "bytes": 139591
    },
    "candidate_workspace": {
      "sha256": "4100b1fbb272dd24842f33a2d6f353a3f3a97dfa7e435c4649a0e1a029bfaea1",
      "bytes": 139606
    },
    "current_perf_baseline": {
      "sha256": "eac86284f67136aa856a8dcb6f383178b89ca92b26434a3de1d7408d68f2690c",
      "bytes": 57227
    },
    "historical_minimal_driver_lock": {
      "sha256": "93878f7a5925cb0484a770bdfb7132cb124dd43c7d37a31ad08fe19b0d8dd741",
      "bytes": 29747,
      "note": "diagnostic minimal-driver lock; not the frozen paired-source lock"
    }
  },
  "current_default_all_targets": {
    "argv": [
      "(deleted-coord-artifact)", "check",
      "--manifest-path", "tools/perf-baseline/Cargo.toml", "--all-targets",
      "--locked", "--offline", "--target-dir",
      "/tmp/dml-owner-record-default-alltargets-target"
    ],
    "environment": {
      "LITCHI_CAPPED": "unset", "LITCHI_MEM_MAX": "20G",
      "CARGO_BUILD_JOBS": "1", "CARGO_INCREMENTAL": "0"
    },
    "features": [
      "current-api", "current-xlsx-raw-opc", "rtf-current-api",
      "drawingml-probe", "drawingml-raw-current", "drawingml-owner-current",
      "drawingml-theme-family-current", "drawingml-theme-family-owner-limits-current",
      "drawingml-ink-limits-current", "drawingml-transform-limits-current"
    ],
    "exit": 0,
    "target_bytes": 629177857,
    "target": "/tmp/dml-owner-record-default-alltargets-target",
    "logs": [
      {"path": "/tmp/dml-owner-record-default-alltargets-9d857/stdout.log", "bytes": 0, "sha256": "e3b0c44298fc1c149afbf4c8996fb92427ae41e4649b934ca495991b7852b855"},
      {"path": "/tmp/dml-owner-record-default-alltargets-9d857/stderr.log", "bytes": 13803, "sha256": "677d16897c3b715ecbb8cb239434ddfdee686891c2a14e3010f680dbbd7b737d"},
      {"path": "/tmp/dml-owner-record-default-alltargets-9d857/time.log", "bytes": 42, "sha256": "12540d151b461d45866ddb0034d6634675633b74e1bbab4f83bb97d292dd7ac4"}
    ]
  },
  "historical_compile_checks": [
    {
      "id": "945", "source_revision": "945eb1d6b9ef6c56fa26c82e12758d38c1b9015d",
      "features": ["drawingml-probe"], "exit": {"generate_lockfile": 0, "check": 0},
      "generate_lockfile_argv": ["(deleted-coord-artifact)", "generate-lockfile", "--manifest-path", "/tmp/dml-owner-baselines-v8.9eNoOE/work/945/Cargo.toml", "--offline"],
      "check_argv": ["(deleted-coord-artifact)", "check", "--manifest-path", "/tmp/dml-owner-baselines-v8.9eNoOE/work/945/Cargo.toml", "--locked", "--offline", "--bins", "--features", "drawingml-probe"],
      "lock": {"sha256": "93878f7a5925cb0484a770bdfb7132cb124dd43c7d37a31ad08fe19b0d8dd741", "bytes": 29747},
      "target": "/tmp/dml-owner-baselines-v8.9eNoOE/targets/945",
      "receipt": {"path": "/tmp/dml-owner-baselines-v8.9eNoOE/receipts/945.json", "bytes": 2974, "sha256": "93a8651745278514f475aff40e0fa8e9c18525678fb05349f85304ce9a08d85e"},
      "check_log": {"path": "/tmp/dml-owner-baselines-v8.9eNoOE/work/945/check.stderr.log", "bytes": 7846, "sha256": "aa2819d522795e59e7d73733e33c9a306980607f44fecaeffde14d4540c84e6b"}
    },
    {
      "id": "2f7", "source_revision": "2f7db7ee7fa01893e9099217fc4243687d196f63",
      "features": ["drawingml-probe", "drawingml-owner-current", "drawingml-theme-family-current"], "exit": {"generate_lockfile": 0, "check": 0},
      "generate_lockfile_argv": ["(deleted-coord-artifact)", "generate-lockfile", "--manifest-path", "/tmp/dml-owner-baselines-v8.9eNoOE/work/2f7/Cargo.toml", "--offline"],
      "check_argv": ["(deleted-coord-artifact)", "check", "--manifest-path", "/tmp/dml-owner-baselines-v8.9eNoOE/work/2f7/Cargo.toml", "--locked", "--offline", "--bins", "--features", "drawingml-probe,drawingml-owner-current,drawingml-theme-family-current"],
      "lock": {"sha256": "93878f7a5925cb0484a770bdfb7132cb124dd43c7d37a31ad08fe19b0d8dd741", "bytes": 29747},
      "target": "/tmp/dml-owner-baselines-v8.9eNoOE/targets/2f7",
      "receipt": {"path": "/tmp/dml-owner-baselines-v8.9eNoOE/receipts/2f7.json", "bytes": 3098, "sha256": "814e08fc19847dd8e69d72b8246f1df5964d7e88e569335a565602b4c204557a"},
      "check_log": {"path": "/tmp/dml-owner-baselines-v8.9eNoOE/work/2f7/check.stderr.log", "bytes": 7625, "sha256": "26668e60f1949f7ee5a88764506d4c6a6775d15d283c1845ffb7a2728b6cc0b6"}
    },
    {
      "id": "996", "source_revision": "996651b6b19717a0192dfd4144bccd367fb2ea44",
      "features": ["drawingml-probe", "drawingml-owner-current", "drawingml-theme-family-current"], "exit": {"generate_lockfile": 0, "check": 0},
      "generate_lockfile_argv": ["(deleted-coord-artifact)", "generate-lockfile", "--manifest-path", "/tmp/dml-owner-baselines-v8.9eNoOE/work/996/Cargo.toml", "--offline"],
      "check_argv": ["(deleted-coord-artifact)", "check", "--manifest-path", "/tmp/dml-owner-baselines-v8.9eNoOE/work/996/Cargo.toml", "--locked", "--offline", "--bins", "--features", "drawingml-probe,drawingml-owner-current,drawingml-theme-family-current"],
      "lock": {"sha256": "93878f7a5925cb0484a770bdfb7132cb124dd43c7d37a31ad08fe19b0d8dd741", "bytes": 29747},
      "target": "/tmp/dml-owner-baselines-v8.9eNoOE/targets/996",
      "receipt": {"path": "/tmp/dml-owner-baselines-v8.9eNoOE/receipts/996.json", "bytes": 3098, "sha256": "6f6e9fec2b5305103fdeeeae363867bcbb87c2a829011c41c1edb792596f7f72"},
      "check_log": {"path": "/tmp/dml-owner-baselines-v8.9eNoOE/work/996/check.stderr.log", "bytes": 7625, "sha256": "9c99e7739bfeaaf154e4e71f2a20b4a8b34a26b4425f1ce72c0330225f3a1571"}
    },
    {
      "id": "11728", "source_revision": "11728db894181ffbcdbcea3ae0008d51fdd2d8af",
      "features": ["drawingml-probe", "drawingml-owner-current"], "exit": {"generate_lockfile": 0, "check": 0},
      "generate_lockfile_argv": ["(deleted-coord-artifact)", "generate-lockfile", "--manifest-path", "/tmp/dml-owner-baselines-v8.9eNoOE/work/11728/Cargo.toml", "--offline"],
      "check_argv": ["(deleted-coord-artifact)", "check", "--manifest-path", "/tmp/dml-owner-baselines-v8.9eNoOE/work/11728/Cargo.toml", "--locked", "--offline", "--bins", "--features", "drawingml-probe,drawingml-owner-current"],
      "lock": {"sha256": "93878f7a5925cb0484a770bdfb7132cb124dd43c7d37a31ad08fe19b0d8dd741", "bytes": 29747},
      "target": "/tmp/dml-owner-baselines-v8.9eNoOE/targets/11728",
      "receipt": {"path": "/tmp/dml-owner-baselines-v8.9eNoOE/receipts/11728.json", "bytes": 3043, "sha256": "d382c4d0969bfde1cd62ff4f8871808cab606a818795e32bbe5844d96c417df9"},
      "check_log": {"path": "/tmp/dml-owner-baselines-v8.9eNoOE/work/11728/check.stderr.log", "bytes": 7641, "sha256": "e9e2d4890c5e6edf2f612c917965d92db6e2285e5d5d259f02e1acb7ed72c413"}
    }
  ],
  "owner_routes": {
    "runtime_argv_template": ["{binary}", "--lane", "{lane}", "--endpoint", "owner", "--owner-route", "{route}", "--role", "{role}", "--mode", "normal", "--fixture", "tools/native_qualification/fixtures/drawingml_raw_fixtures.json", "--source-revision", "{source_revision}", "--iterations", "1", "--warmups", "3", "--blocks", "5", "--samples", "20", "--target-ms", "500", "--minimum-ms", "150", "--json", "{report}"],
    "route_result": {
      "945": {
        "native_paths": {"chart": ["Snapshot::edit -> Edit::commit", "Patch::apply"], "theme-family": ["Snapshot::edit -> Edit::commit", "Patch::apply"], "ink-actions": ["Edit::from_profile -> Edit::commit", "Patch::apply"], "transform": ["Snapshot::edit -> Edit::commit", "Patch::apply"]},
        "reports": {
          "chart": {"path": "/tmp/dml-owner-baselines-v8.9eNoOE/runtime/945/drawingml-chart.json", "bytes": 50859, "sha256": "2fcb20aeab263fe62a7e2d4524c29566cc532b9f6c61b2b985e98f6880f00332"},
          "theme-family": {"path": "/tmp/dml-owner-baselines-v8.9eNoOE/runtime/945/drawingml-theme-family.json", "bytes": 52959, "sha256": "78b4bad94f463c7c642b2a519b8c1efc38d795781b8a6a565c5e8689904611c9"},
          "ink-actions": {"path": "/tmp/dml-owner-baselines-v8.9eNoOE/runtime/945/drawingml-ink-actions.json", "bytes": 52865, "sha256": "7a549949ddcd8845056e16d8ec4a31ce2e3602ce4e3f090d0d620ad1d1de2aaf"},
          "transform": {"path": "/tmp/dml-owner-baselines-v8.9eNoOE/runtime/945/drawingml-transform.json", "bytes": 52399, "sha256": "632a529fa72db50d5b021be76c907f8ae6c6bde266480d7d9d9880ca73243dda"}
        }
      },
      "2f7": {
        "native_paths": {"chart": ["Package::edit -> Edit::commit", "Package::apply"], "theme-family": ["Package::edit -> Edit::commit", "Package::apply"], "ink-actions": ["Prepared::edit -> Edit::commit", "Prepared::apply"], "transform": ["Package::edit -> Edit::commit", "Package::apply"]},
        "reports": {
          "chart": {"path": "/tmp/dml-owner-baselines-v8.9eNoOE/runtime/2f7/drawingml-chart.json", "bytes": 50810, "sha256": "23816c058632816bdc84bf4c926e4f514d9c6821bf0440a20b124bd92d46f046"},
          "theme-family": {"path": "/tmp/dml-owner-baselines-v8.9eNoOE/runtime/2f7/drawingml-theme-family.json", "bytes": 52889, "sha256": "c23d65298fcb9c18a01f4984ee4e0543dd88ef48458f7df274f7cb2d87c9be16"},
          "ink-actions": {"path": "/tmp/dml-owner-baselines-v8.9eNoOE/runtime/2f7/drawingml-ink-actions.json", "bytes": 52480, "sha256": "9a4794195d957df0e4dcc29273057ba38f835cad1180b05a96e5cf00243e3c37"},
          "transform": {"path": "/tmp/dml-owner-baselines-v8.9eNoOE/runtime/2f7/drawingml-transform.json", "bytes": 52305, "sha256": "869ebc9a2d01738f662ed9a86cabd98e4c0551eb975fa1fc217b4a67a09112c9"}
        }
      },
      "996": {
        "native_paths": {"chart": ["Package::edit -> Edit::commit", "Package::apply"], "theme-family": ["Package::edit -> Edit::commit", "Package::apply"], "ink-actions": ["Prepared::edit -> Edit::commit", "Prepared::apply"], "transform": ["Package::edit -> Edit::commit", "Package::apply"]},
        "reports": {
          "chart": {"path": "/tmp/dml-owner-baselines-v8.9eNoOE/runtime/996/drawingml-chart.json", "bytes": 50810, "sha256": "2af8cb5a84e34d62430a9c5b48e42f939c564bc223ab66d57121039aff265927"},
          "theme-family": {"path": "/tmp/dml-owner-baselines-v8.9eNoOE/runtime/996/drawingml-theme-family.json", "bytes": 52889, "sha256": "ffac9e5619ac66092ff4b75b8c3d29afdaa0496a16a96718659bc62ea5dd10b3"},
          "ink-actions": {"path": "/tmp/dml-owner-baselines-v8.9eNoOE/runtime/996/drawingml-ink-actions.json", "bytes": 52480, "sha256": "adb7822996022108e83a4ca4beeb8e985a5bad9ea653f387f6e6a7be7029992d"},
          "transform": {"path": "/tmp/dml-owner-baselines-v8.9eNoOE/runtime/996/drawingml-transform.json", "bytes": 52305, "sha256": "629afbd8079682c42819494cc87e51c74a33ccaad859f3b25decf484485b9222"}
        }
      },
      "11728": {
        "native_paths": {"chart": ["Package::edit -> Edit::commit", "Package::apply"], "theme-family": ["Package::edit -> Edit::commit", "Package::apply"], "ink-actions": ["Prepared::edit -> Edit::commit", "Prepared::apply"], "transform": ["Package::edit -> Edit::commit", "Package::apply"]},
        "reports": {
          "chart": {"path": "/tmp/dml-owner-baselines-v8.9eNoOE/runtime/11728/drawingml-chart.json", "bytes": 50810, "sha256": "8b3af3d337fcc1c8b4c2e0433477079a9dbb354fd5628b9f41df7a66df63a296"},
          "theme-family": {"path": "/tmp/dml-owner-baselines-v8.9eNoOE/runtime/11728/drawingml-theme-family.json", "bytes": 52889, "sha256": "5ca5fa1f4a065866b5bbf5a42b91873b6ca45f45abe3340b00f7dd8761057ab2"},
          "ink-actions": {"path": "/tmp/dml-owner-baselines-v8.9eNoOE/runtime/11728/drawingml-ink-actions.json", "bytes": 52480, "sha256": "05e2be9ee8f03a9c09d2f88dce3102a1ef9f15dd721edcd7a0a88cf3281862bb"},
          "transform": {"path": "/tmp/dml-owner-baselines-v8.9eNoOE/runtime/11728/drawingml-transform.json", "bytes": 52305, "sha256": "57a65b984639613009a7bdfebd72aba529abcc549d4191d66eb3f7efaabbb326"}
        }
      },
      "current": {
        "source_revision": "9d85708aed4f81c81570638c833d7261fb4c28da",
        "binary": {"path": "/tmp/dml-owner-current-proof-run-final/debug/drawingml-raw-native", "bytes": 810061936, "sha256": "afbcd055b9111f0371335dd4364f25b831178e8a5a69afc7dfb18490bf366417"},
        "native_paths": {"chart": ["Package::edit -> Edit::commit", "Package::apply"], "theme-family": ["Package::edit -> Edit::commit", "Package::apply"], "ink-actions": ["Prepared::edit -> Edit::commit", "Prepared::apply"], "transform": ["Package::edit -> Edit::commit", "Package::apply"]},
        "reports": {
          "chart": {"path": "/tmp/dml-owner-current-proof-source-9d857/chart.json", "bytes": 50809, "sha256": "5232c699b28e825901073e63a6d1d8f23a7f2f15f2f1a60b94e2f19d560129c7"},
          "theme-family": {"path": "/tmp/dml-owner-current-proof-source-9d857/theme-family.json", "bytes": 52388, "sha256": "15e9d07e577a87e6fc4ba43fba94bbc9578658035f3b077555214c17f59fcc43"},
          "ink-actions": {"path": "/tmp/dml-owner-current-proof-source-9d857/ink-actions.json", "bytes": 52479, "sha256": "095fef30e65df6b2c4d56c802dbc41ab4f8263e16f4591a7783c3dca0dbe7577"},
          "transform": {"path": "/tmp/dml-owner-current-proof-source-9d857/transform.json", "bytes": 52304, "sha256": "68a062c7119a68f74da936f465a858b5c6c87fac440a42ecc40931ce74b66b67"}
        }
      },
      "per_report_invariants": {"rows": 5, "mutation_verified": true, "unchanged_remainder": true, "stale_reason_verified": true}
    }
  },
  "paired_build_attempts": [
    {
      "candidate_revision": "ef3d904ea7c58cda1b7e2a7921d7ae53eff0f76b",
      "runner_argv": ["python3", "tools/native_qualification/drawingml_raw_paired_runner.py", "--manifest", "tools/native_qualification/manifests/drawingml-raw.json", "--phase", "build"],
      "receipt": {"path": "/tmp/dml-paired-ef3d/output-build/build-receipt.json", "bytes": 16317, "sha256": "99163ac23831523303ea4224fbbfed821a87f43227137f7d0edc2928be344efe"},
      "exit": 2, "result": "refused before profile build by quiet-host guard; unrelated cargo/rustc jobs were active"
    },
    {
      "candidate_revision": "9d85708aed4f81c81570638c833d7261fb4c28da",
      "runner_argv": ["python3", "tools/native_qualification/drawingml_raw_paired_runner.py", "--manifest", "tools/native_qualification/manifests/drawingml-raw.json", "--phase", "build"],
      "receipt": {"path": "/tmp/dml-paired-ef3d/output-build2/build-receipt.json", "bytes": 269042, "sha256": "6d01c45ff4583734201440647b97daec53fc7dea50d3d16e122cbb40efb8f553"},
      "exit": 2, "result": "refused before profile build by quiet-host guard; unrelated cargo/rustc/rustfmt jobs were active"
    }
  ],
  "qualification": {
    "historical_owner_routes_executed": 16,
    "current_owner_routes_executed": 4,
    "ordinary_default_all_targets": "passed",
    "paired_build": "not_run",
    "abba": "not_run",
    "performance_claim": "not_claimed",
    "full_frozen_52ed_source_pair": "required for final paired qualification"
  }
}
```

The historical generated-driver checks use lock `93878f…` as explicitly
marked above; they are compile diagnostics, not substitutes for the frozen
52ed/4100 paired-source locks.  The two paired-runner receipts preserve the
quiet-host refusals rather than converting them into a qualification result.
