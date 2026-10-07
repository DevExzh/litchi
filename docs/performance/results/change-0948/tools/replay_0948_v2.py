"""Portable replay-only adapter for the frozen 0948 v4 reducer.

The frozen reducer is strict about invocation custody, but it derives absolute
argv and cwd values from its current packet directory.  That is correct for
the original capture and intentionally fails after a packet is relocated.  This
adapter verifies the exact frozen reducer bytes, binds the historical capture
root from the retained authorization receipt, and patches only the reducer's
expected invocation strings.  All files read by the reducer remain under the
relocated packet; the historical root is metadata used for comparison only.

This adapter never launches a workload.  Its negative-control action is
create-only and checks that tampered historical cwd/argv and a missing physical
artifact are rejected.
"""

from __future__ import annotations

import argparse
import hashlib
import importlib.util
import json
from pathlib import Path
import sys
from typing import Any, Callable


ADAPTER_PATH = Path(__file__).absolute()
PACKET = ADAPTER_PATH.parent.parent
V4_RELATIVE = "tools/reduce_0948_v4.py"
V4_SHA256 = "f960a94cad1544ac7f64c8a5006eb172bf5db849735b576fba58a6338a5b1d17"
AUTH_RELATIVE = "receipts/root-qualification-go-v1.json"
AUTH_SHA256 = "59e12d460778f5552604c08b612e353172792317d8f75ea026dc6bf653d03613"
PLAN_RELATIVE = "plan-v3.json"
PLAN_SHA256 = "e4bb028993ec25b0b538616153a09b17fdfb26b91c8255336eab9393d87de73f"
HISTORICAL_ROOT = "/home/zhuhe/litchi-performance-0948"
NEGATIVE_RECEIPT_RELATIVE = "receipts/replay-adapter-negative-controls-v1.json"
SCHEMA = "litchi.performance.0948.portable-replay-adapter.v1"
NEGATIVE_SCHEMA = "litchi.performance.0948.portable-replay-negative-controls.v1"


class ReplayError(RuntimeError):
    """A relocated packet failed a replay binding or custody check."""


def require(condition: bool, message: str) -> None:
    if not condition:
        raise ReplayError(message)


def sha256_file(path: Path) -> str:
    require(path.is_file() and not path.is_symlink(),
            f"replay binding is not a regular file: {path}")
    return hashlib.sha256(path.read_bytes()).hexdigest()


def read_json(path: Path) -> Any:
    require(path.is_file() and not path.is_symlink(),
            f"replay JSON is not a regular file: {path}")

    def reject_duplicates(pairs: list[tuple[str, Any]]) -> dict[str, Any]:
        result: dict[str, Any] = {}
        for key, value in pairs:
            require(key not in result, f"duplicate JSON key {key!r} in {path}")
            result[key] = value
        return result

    try:
        return json.loads(path.read_text(encoding="utf-8"),
                          object_pairs_hook=reject_duplicates)
    except (OSError, json.JSONDecodeError) as error:
        raise ReplayError(f"cannot read replay JSON {path}: {error}") from error


def packet_file(relative: str, expected_sha256: str) -> Path:
    path = PACKET / relative
    require(path.is_file() and not path.is_symlink(),
            f"missing relocated replay binding {path}")
    observed = sha256_file(path)
    require(observed == expected_sha256,
            f"{relative} SHA-256 changed: {observed} != {expected_sha256}")
    return path


def bind_historical_root() -> tuple[str, dict[str, Any]]:
    """Read only the retained authorization receipt under the new packet root."""
    require(ADAPTER_PATH.is_file() and not ADAPTER_PATH.is_symlink()
            and ADAPTER_PATH == ADAPTER_PATH.resolve(),
            f"replay adapter path is not a real file: {ADAPTER_PATH}")
    require(PACKET.is_dir() and not PACKET.is_symlink()
            and PACKET == PACKET.resolve(),
            f"replay packet root is not a real directory: {PACKET}")
    authorization_path = packet_file(AUTH_RELATIVE, AUTH_SHA256)
    # Bind the exact frozen plan before importing or dispatching the reducer.
    # The reducer performs its own descriptor checks, but this adapter's
    # relocation contract must fail closed if the plan itself was substituted.
    packet_file(PLAN_RELATIVE, PLAN_SHA256)
    authorization = read_json(authorization_path)
    require(isinstance(authorization, dict)
            and authorization.get("schema") == "litchi.performance.0948.root-qualification-go.v1"
            and authorization.get("status") == "authorized",
            "historical root authorization schema/status changed")
    require(authorization.get("argv") == [
        "python3", "-B", "tools/capture_0948_v3.py", "qualification"],
            "historical root authorization argv changed")
    historical = authorization.get("cwd")
    require(isinstance(historical, str)
            and historical == HISTORICAL_ROOT
            and Path(historical).is_absolute()
            and Path(historical).as_posix() == historical
            and ".." not in Path(historical).parts,
            "historical capture root is not the bound normalized absolute path")
    current = str(PACKET.resolve())
    require(current != historical,
            "replay packet still has the historical capture root; relocation test is not active")
    return historical, {
        "path": AUTH_RELATIVE,
        "sha256": AUTH_SHA256,
        "schema": authorization["schema"],
        "argv": authorization["argv"],
        "cwd": historical,
    }


def load_frozen_reducer() -> Any:
    """Verify v4 bytes before importing any reducer code."""
    require(PACKET.is_dir() and not PACKET.is_symlink()
            and PACKET == PACKET.resolve(),
            f"replay packet root is not a real directory: {PACKET}")
    path = packet_file(V4_RELATIVE, V4_SHA256)
    spec = importlib.util.spec_from_file_location("litchi_0948_frozen_reducer_v4", path)
    require(spec is not None and spec.loader is not None,
            "cannot load frozen reducer v4")
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    return module


class HistoricalCwd:
    """Physical relocated directory with the historical cwd string."""

    __slots__ = ("physical", "historical")

    def __init__(self, physical: Path, historical: str) -> None:
        self.physical = physical
        self.historical = historical

    def is_dir(self) -> bool:
        return self.physical.is_dir()

    def is_symlink(self) -> bool:
        return self.physical.is_symlink()

    def __str__(self) -> str:
        return self.historical

    def __repr__(self) -> str:
        return f"HistoricalCwd({self.historical!r}, physical={self.physical!r})"


def rewrite_packet_prefix(value: str, packet: Path, historical: str) -> str:
    """Rewrite only absolute paths rooted at the relocated packet."""
    current = str(packet.resolve())
    if value == current:
        return historical
    prefix = current + "/"
    if value.startswith(prefix):
        return historical + value[len(current):]
    return value


def install_relocation_adapter(reducer: Any, historical: str) -> Callable[[], None]:
    """Patch only v4's expected command/cwd renderers, then return a restore."""
    packet = reducer.PACKET.resolve()
    original_command_for = reducer.command_for
    original_path_value = reducer.path_value

    def relocated_path_value(value: Any, label: str, base: Path = packet, *,
                             allow_absolute: bool = False) -> Any:
        physical = original_path_value(value, label, base=base,
                                        allow_absolute=allow_absolute)
        if label == "workload_cwd":
            return HistoricalCwd(physical, historical)
        return physical

    def historical_command_for(plan: dict[str, Any], job: dict[str, Any],
                               folder: Path) -> list[str]:
        physical_command = original_command_for(plan, job, folder)
        require(all(isinstance(item, str) for item in physical_command),
                "frozen reducer produced a non-string argv entry")
        return [rewrite_packet_prefix(item, packet, historical)
                for item in physical_command]

    reducer.path_value = relocated_path_value
    reducer.command_for = historical_command_for

    def restore() -> None:
        reducer.path_value = original_path_value
        reducer.command_for = original_command_for

    return restore


def invocation_check(record: dict[str, Any], expected_argv: list[str],
                     historical: str) -> None:
    require(record.get("argv") == expected_argv
            and record.get("cwd") == historical,
            "historical invocation custody mismatch")


def write_once(path: Path, value: Any) -> None:
    require(not path.exists() and not path.is_symlink(),
            f"refusing to overwrite replay receipt {path}")
    path.parent.mkdir(parents=True, exist_ok=True)
    with path.open("x", encoding="utf-8") as handle:
        json.dump(value, handle, indent=2, sort_keys=True, ensure_ascii=False)
        handle.write("\n")


def negative_controls(reducer: Any, historical: str, authorization: dict[str, Any],
                      receipt_path: Path) -> dict[str, Any]:
    """Run pure custody negative controls without changing retained artifacts."""
    plan = reducer.load_plan()
    job = reducer.jobs("qualification")[0]
    folder = PACKET / "qualification" / reducer.child_name(job)
    unmodified_command = reducer.command_for(plan, job, folder)
    unmodified_workload = reducer.path_value(plan["workload_cwd"], "workload_cwd",
                                              allow_absolute=True)
    unmodified_cwd = str(unmodified_workload)
    restore = install_relocation_adapter(reducer, historical)
    try:
        expected = reducer.command_for(plan, job, folder)
    finally:
        restore()

    sample_path = PACKET / "qualification" / reducer.child_name(job) / "start.json"
    sample = reducer.read_json(sample_path)
    require(isinstance(sample, dict), "negative-control sample start is not an object")
    require(sample.get("cwd") == historical and sample.get("cwd") != unmodified_cwd,
            "unmodified v4 relocation did not expose the historical cwd mismatch")
    require(sample.get("argv") != unmodified_command,
            "unmodified v4 relocation did not expose the historical argv mismatch")
    invocation_check(sample, expected, historical)

    tampered_cwd = dict(sample)
    tampered_cwd["cwd"] = historical + "/tampered"
    try:
        invocation_check(tampered_cwd, expected, historical)
    except ReplayError:
        cwd_result = "rejected"
    else:
        raise ReplayError("negative control accepted tampered historical cwd")

    tampered_argv = dict(sample)
    tampered_argv["argv"] = list(expected)
    tampered_argv["argv"][-1] = expected[-1] + ".tampered"
    try:
        invocation_check(tampered_argv, expected, historical)
    except ReplayError:
        argv_result = "rejected"
    else:
        raise ReplayError("negative control accepted tampered historical argv")

    missing = PACKET / ".replay-negative-missing-artifact"
    require(not missing.exists() and not missing.is_symlink(),
            f"negative-control path unexpectedly exists: {missing}")
    try:
        reducer.digest(missing)
    except reducer.ReviewError:
        missing_result = "rejected"
    else:
        raise ReplayError("negative control accepted missing relocated artifact")

    receipt = {
        "schema": NEGATIVE_SCHEMA,
        "status": "pass",
        "adapter": {"path": ADAPTER_PATH.relative_to(PACKET).as_posix(),
                    "sha256": sha256_file(ADAPTER_PATH)},
        "frozen_reducer": {"path": V4_RELATIVE, "sha256": V4_SHA256},
        "plan": {"path": PLAN_RELATIVE, "sha256": PLAN_SHA256},
        "historical_authorization": authorization,
        "historical_root": historical,
        "relocated_packet_root": str(PACKET.resolve()),
        "unmodified_v4_relocation": {
            "status": "expected-failure",
            "reason": "v4 renders absolute command and cwd expectations from its current PACKET",
            "observed_cwd": {"captured": sample["cwd"], "unmodified_v4": unmodified_cwd},
            "observed_argv_mismatch": True,
        },
        "controls": {
            "tampered_historical_cwd": cwd_result,
            "tampered_historical_argv": argv_result,
            "missing_relocated_artifact": missing_result,
        },
    }
    write_once(receipt_path, receipt)
    return receipt


def replay(reducer: Any, historical: str, action: str, output: Path) -> dict[str, Any]:
    require(not output.exists() and not output.is_symlink(),
            f"refusing to overwrite replay output {output}")
    plan = reducer.load_plan()
    restore = install_relocation_adapter(reducer, historical)
    try:
        result = reducer.analysis(plan, action, output)
    finally:
        restore()
    return result


def main() -> None:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("action", choices=["qualification", "native", "allocator", "all",
                                            "negative-controls"])
    parser.add_argument("--output", type=Path,
                        help="create-only replay analysis output for a reduction action")
    parser.add_argument("--receipt", type=Path,
                        help="create-only negative-control receipt")
    args = parser.parse_args()

    historical, authorization = bind_historical_root()
    reducer = load_frozen_reducer()
    if args.action == "negative-controls":
        receipt = args.receipt or (PACKET / NEGATIVE_RECEIPT_RELATIVE)
        result = negative_controls(reducer, historical, authorization, receipt)
        print(json.dumps({"status": result["status"], "receipt": str(receipt)}, sort_keys=True))
        return

    require(args.output is not None, "replay action requires --output")
    result = replay(reducer, historical, args.action, args.output.absolute())
    print(json.dumps({"status": result["status"], "stage": result["stage"],
                      "historical_root": historical}, sort_keys=True))


if __name__ == "__main__":
    try:
        main()
    except (ReplayError, AssertionError, OSError, ValueError) as error:
        print(f"0948 portable replay failed: {error}", file=sys.stderr)
        raise SystemExit(1)
