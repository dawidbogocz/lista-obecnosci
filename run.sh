#!/bin/bash
# Launch the Attendance Sheet app
cd "$(dirname "$0")"
exec /tmp/attendance_venv/bin/python3 lista_obecnosci.py "$@"