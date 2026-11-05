# 0362ul152 — bound-policy test integration and baseline admission

The common protocol installation is prepared against P
`4ab42bb9678769930d7c17173827fe5693370ab6`, with immutable candidate C
`f8cd398a770f7363f628be40edb3ce4219566c16`. P's bindings remain pending
while the two source-generation jobs read that exact prebinding commit.
No Cargo or native measurement is started by this preparation.

## DrawingML installation tested

An isolated sparse P checkout receives the generated literal binding, the
ready/bound manifest projection and the test migration. Both complete
DrawingML test modules pass: 27 tests. The initial run finds one stale compile
adapter assertion: a bound policy rejects the wrong candidate revision,
rather than reporting that the policy is incomplete. The corrected assertion
retains that rejection, and explicit pending fixtures retain the earlier
missing-binding coverage.

The combined five-file patch remains external at
`(deleted-coord-artifact)/combined-integration.patch`:
143,584 bytes, SHA-256
`4d9616546aac92b0ba5b86c1b333e819786588d4fcce0989ef96181d8766ec78`.
Root independently applies it to copied source files and verifies the exact
five-file scope, Python syntax, literal/JSON equality and manifest projection.
The manifest differs only by readiness and the reviewed candidate binding;
its existing role-lock fields remain intact. The root review is 1,064 bytes,
SHA-256 `e0e96fe1f3228da5a0e869da6be5db7ad24f83d0ed173e77047e8ee4167fa1fc`.
The temporary worktree is removed; the registered count returns to 55.

## Other policy consumers

Current-native's proposed test migration requires the frozen candidate,
three role locks and both ready manifests. It retains incomplete-policy
negatives through an explicit incomplete policy shared by the policy,
preparer and validator. Six explicit-incomplete negatives and a synthetic
complete-candidate check pass. Its full integrated suite still requires the
real installed binding.

Root returns the first signed test patch because its production-bound test
could skip and its prepare negative still depended on production being
pending. The successor requires a bound production policy without a skip and
uses an explicit temporary pending manifest with a mocked missing binding
for both negative paths. It retains the output nonmutation assertion.

| External migration | Bytes | SHA-256 |
|---|---:|---|
| `current-native-final-binding-test-migration-v1/0001-migrate-current-native-policy-tests.patch` | 15,506 | `581e726e639089465d8080f0fdf323dd5b1cbe6b6d3906de5d6a61810624da5d` |
| `signed-native-v4-signed-test-migration-v1-20261025/test_signed_native_policy_v2.patch` | 6,006 | `e574db5095276dbd49d68d5ac190d2ae8058396d7fad3b0cd422007cf4698f74` |

P's sparse checkout omitted eleven fixtures needed by the signed policy
tests. Root restores their exact tracked Git blobs from P's own HEAD, checks
every blob/size/SHA-256 and confirms unchanged HEAD, Git status and all three
locks. This materializes 197,887 bytes and changes no tracked content. The
existing signed policy suite then passes all 14 tests. The hydration receipt
is 4,656 bytes, SHA-256
`c4a98adb299f9b013175977ef51b8178bc13fb99a6e3b71cfda7bd065730629d`.

The prose-only audit also removes stale claims that production is pending.
Its signed test module docstring overlaps the functional test patch and must
be applied only once. The common-P installation must run the current-native
installer first, while P is clean; the literal bindings, manifest changes,
consumer tests and applicable prose changes then belong in the same commit.

## DrawingML baseline admission

The immutable v5 recipe admits the exact preparation status of the retained
baselines: three unstaged tracked changes, six untracked adapters and the
ignored root lock. It checks all three role locks and eight adapter inputs.
Extra or missing changes, incorrect locks and changed adapters are refused.
Its candidate binding and generated outer driver remain byte-identical to v4.
Independent review passes ten source-only checks. The unchanged driver's
execution seam is not rerun under the live build workload.

Root independently compares the old `52ed0cac…` and approved `4100b1fb…`
locks as TOML. The package identity set and all 518 registry package entries
are unchanged; only `litchi-ops` gains the `litchi-vba` dependency. This
matches the captured transition from source commit
`0dcc767c68967f765412379f0234ed92c21fbc1a`; no Cargo resolution occurs.
Root also exercises positive admission for all eleven overlay inputs using a
temporary copy with the approved lock. The actual four historical roots
remain unchanged. Six additional physical baseline roots still need to be
materialized after the source-only plan is reviewed.

The v5 recipe is 77,923 bytes, SHA-256
`7a1d6470d4249c957f59962d7640382ca81b53a42fe0d4f62937a0c579e1fa6a`.
The root lock comparison receipt is 799 bytes, SHA-256
`5be4f7ddd8e409614590be0a88595cf8ccf85600c44879a808cef14c3e138fe7`;
positive overlay evidence is 3,459 bytes, SHA-256
`a8806a3b0f064ab69fadcfdf975ecc7e05ead662041cee50f9835765b5611e54`.

## Live qualification

Root verifies seventeen completed final65 rows, including all seven compiler
supplements and ten full-gate rows. Logs, Cargo traces, exits, source/lock
identities and post-exit process observations agree. The partial verification
receipt is 2,280 bytes, SHA-256
`432d75c90944c154e0f6ecf7750d7263f877c1938dea4a4da2a7409f3156fa1e`.
The full gate and both source-generation services remain live. A constructed
signed tar is not accepted until generation and verification finish.

Actual native builds, fidelity and performance, raw census capture/review,
remaining gates, final trunk integration and cleanup are still required.
Audit 22 remains the only deferral. Nothing is pushed.
