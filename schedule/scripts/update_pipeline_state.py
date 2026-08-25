from datetime import datetime, timezone
from pathlib import Path
import argparse
import json
import os
import sys


DEFAULT_STATE_FILE = "schedule/data/pipeline-state.json"


def utc_now():
    return datetime.now(timezone.utc).isoformat()


def default_state():
    return {
        "status": "healthy",
        "errorLevel": None,
        "errorCode": None,
        "firstDetected": None,
        "lastDetected": None,
        "lastNotified": None,
    }


def load_state(path):
    if not path.exists() or path.stat().st_size == 0:
        return default_state()

    with path.open(
        "r",
        encoding="utf-8",
    ) as file:
        return json.load(file)


def save_state(path, state):
    path.parent.mkdir(
        parents=True,
        exist_ok=True,
    )

    temp_path = path.with_suffix(
        ".json.tmp"
    )

    with temp_path.open(
        "w",
        encoding="utf-8",
    ) as file:
        json.dump(
            state,
            file,
            indent=2,
            ensure_ascii=False,
        )

        file.write("\n")

    temp_path.replace(path)


def emit_output(name, value):
    output_path = os.environ.get(
        "GITHUB_OUTPUT"
    )

    if output_path:
        with open(
            output_path,
            "a",
            encoding="utf-8",
        ) as file:
            file.write(
                f"{name}={value}\n"
            )

    print(
        f"{name}={value}"
    )


def handle_healthy(previous_state):
    """
    Healthy -> Healthy:
        No state-file change.
        GitHub Actions history is our heartbeat log.

    Failed -> Healthy:
        This is a meaningful recovery transition.
        Update persistent state and request notification.
    """

    if previous_state.get("status") == "failed":
        now = utc_now()

        new_state = {
            "status": "healthy",
            "errorLevel": None,
            "errorCode": None,
            "firstDetected": None,
            "lastDetected": now,
            "lastNotified":
                previous_state.get(
                    "lastNotified"
                ),
        }

        return (
            new_state,
            "recovered",
            True,
            True,
        )

    return (
        previous_state,
        "healthy",
        False,
        False,
    )


def handle_failure(
    previous_state,
    level,
    code,
):
    """
    Same failure -> Same failure:
        Do not rewrite the state file.

    Healthy -> Failed, or one failure -> different failure:
        Record the new incident and request notification.
    """

    same_incident = (
        previous_state.get("status")
        == "failed"
        and
        previous_state.get(
            "errorLevel"
        )
        == level
        and
        previous_state.get(
            "errorCode"
        )
        == code
    )

    if same_incident:
        return (
            previous_state,
            "continuing_failure",
            False,
            False,
        )

    now = utc_now()

    new_state = {
        "status": "failed",
        "errorLevel": level,
        "errorCode": code,
        "firstDetected": now,
        "lastDetected": now,
        "lastNotified": None,
    }

    return (
        new_state,
        "new_failure",
        True,
        True,
    )


def main():
    parser = argparse.ArgumentParser(
        description=(
            "Track Blue Mountain Pilates "
            "schedule pipeline health."
        )
    )

    parser.add_argument(
        "--state",
        choices=[
            "healthy",
            "failed",
        ],
        required=True,
    )

    parser.add_argument(
        "--level",
        type=int,
        choices=[
            1,
            2,
        ],
    )

    parser.add_argument(
        "--code",
    )

    parser.add_argument(
        "--state-file",
        default=DEFAULT_STATE_FILE,
    )

    args = parser.parse_args()

    if args.state == "failed":
        if args.level is None:
            parser.error(
                "--level is required "
                "when state=failed"
            )

        if not args.code:
            parser.error(
                "--code is required "
                "when state=failed"
            )

    state_path = Path(
        args.state_file
    )

    previous_state = load_state(
        state_path
    )

    if args.state == "healthy":
        (
            new_state,
            transition,
            notify,
            state_changed,
        ) = handle_healthy(
            previous_state
        )

    else:
        (
            new_state,
            transition,
            notify,
            state_changed,
        ) = handle_failure(
            previous_state,
            args.level,
            args.code,
        )

    if state_changed:
        save_state(
            state_path,
            new_state,
        )

    print()
    print(
        "Pipeline state:"
    )

    print(
        "Previous:",
        previous_state.get(
            "status"
        ),
    )

    print(
        "Current:",
        new_state["status"],
    )

    print(
        "Transition:",
        transition,
    )

    if new_state["status"] == "failed":
        print(
            "Level:",
            new_state[
                "errorLevel"
            ],
        )

        print(
            "Code:",
            new_state[
                "errorCode"
            ],
        )

    print(
        "Notify:",
        "yes"
        if notify
        else "no",
    )

    print(
        "Persistent state changed:",
        "yes"
        if state_changed
        else "no",
    )

    emit_output(
        "transition",
        transition,
    )

    emit_output(
        "notify",
        "true"
        if notify
        else "false",
    )

    emit_output(
        "status",
        new_state["status"],
    )

    emit_output(
        "state_changed",
        "true"
        if state_changed
        else "false",
    )

    emit_output(
        "error_level",
        (
            str(
                new_state[
                    "errorLevel"
                ]
            )
            if new_state[
                "errorLevel"
            ] is not None
            else ""
        ),
    )

    emit_output(
        "error_code",
        new_state.get(
            "errorCode"
        ) or "",
    )


if __name__ == "__main__":
    try:
        main()

    except Exception as error:
        print(
            "::error::"
            "PIPELINE STATE MANAGER FAILURE: "
            f"{type(error).__name__}: "
            f"{error}"
        )

        sys.exit(1)