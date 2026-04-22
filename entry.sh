#!/bin/bash

VENV_PATH=./.venv
ENTRYPOINT_SH="../simple_vuln_automater/entrypoint.sh"

if [ -d "$VENV_PATH" ]; then
    echo "Activating virtual environment..."
    source "$VENV_PATH/bin/activate"

    ${ENTRYPOINT_SH} --config_root ../simple_vuln_tester --app_folder ../simple_vuln_scanner --fullscan

    deactivate
else
    echo "Error: Virtual environment directory ${VENV_PATH} not found."
    exit 1
fi