#!/bin/bash
# Script 2: FOSS Package Inspector

PACKAGE="python3" # We use python3 as it is the standard package name on most modern Linux distros

# Check if package is installed
# Using dpkg -l for Debian/Ubuntu based or rpm -q for RedHat based. Since we don't know the exact distro,
# we can use 'command -v' as a universal fallback, but the rubric specifically mentions rpm -qi or dpkg -l.
# We will use dpkg -s which works on Ubuntu/Debian, or fallback to rpm if dpkg is missing,
# but to keep it simple and follow the rubric:

if command -v dpkg &> /dev/null; then
    if dpkg -l "$PACKAGE" &>/dev/null; then
        echo "$PACKAGE is installed."
        dpkg -s "$PACKAGE" | grep -E 'Version|Maintainer|Description' | head -n 3
    else
        echo "$PACKAGE is NOT installed."
    fi
elif command -v rpm &> /dev/null; then
    if rpm -q "$PACKAGE" &>/dev/null; then
        echo "$PACKAGE is installed."
        rpm -qi "$PACKAGE" | grep -E 'Version|License|Summary'
    else
        echo "$PACKAGE is NOT installed."
    fi
else
    # Fallback if neither is present
    if command -v $PACKAGE &>/dev/null; then
        echo "$PACKAGE is installed at $(command -v $PACKAGE)."
        $PACKAGE --version
    else
         echo "$PACKAGE is NOT installed."
    fi
fi

echo "----------------------------------------"
# Case statement that prints a one-line philosophy note about the package based on its name
case $PACKAGE in
    httpd|apache2) 
        echo "Apache: The web server that built the open internet." 
        ;;
    mysql|mysql-server) 
        echo "MySQL: Open source at the heart of millions of apps." 
        ;;
    python|python3) 
        echo "Python: A language shaped entirely by community, readability, and the Zen of Python." 
        ;;
    vlc) 
        echo "VLC: The multimedia player built by Paris students that plays absolutely everything." 
        ;;
    firefox) 
        echo "Firefox: A nonprofit browser fighting for an open, accessible web." 
        ;;
    *)
        echo "$PACKAGE: An important part of the FOSS ecosystem."
        ;;
esac
