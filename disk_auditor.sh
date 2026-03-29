#!/bin/bash
# Script 3: Disk and Permission Auditor

DIRS=("/etc" "/var/log" "/home" "/usr/bin" "/tmp")

echo "======================"
echo "Directory Audit Report"
echo "======================"

for DIR in "${DIRS[@]}"; do
    if [ -d "$DIR" ]; then
        # $1=permissions, $3=owner, $4=group
        PERMS=$(ls -ld "$DIR" | awk '{print "Perms: "$1" | Owner: "$3":"$4}')
        SIZE=$(du -sh "$DIR" 2>/dev/null | cut -f1)
        echo "$DIR => $PERMS | Size: $SIZE"
    else
        echo "$DIR does not exist on this system"
    fi
done

echo "----------------------"
echo "Python Project Audit:"
# Checking python config or library directory
PYTHON_LIB="/usr/lib/python3" # A common directory for python3
if [ -d "$PYTHON_LIB" ]; then
    PY_PERMS=$(ls -ld "$PYTHON_LIB" | awk '{print "Perms: "$1" | Owner: "$3":"$4}')
    echo "$PYTHON_LIB => $PY_PERMS"
else
    echo "Directory $PYTHON_LIB does not exist on this system."
fi
