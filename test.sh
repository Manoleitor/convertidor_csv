#!/bin/bash
export DYLD_LIBRARY_PATH="/opt/homebrew/lib:$DYLD_LIBRARY_PATH"
source .venv/bin/activate && python3 -m unittest test_process_signatures.py "$@"