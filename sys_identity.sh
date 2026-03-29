#!/bin/bash
# Script 1: System Identity Report
# Author: [Your Name] | Course: Open Source Software

# --- Variables ---
STUDENT_NAME="Shree Nigam"        # Fill in your name
SOFTWARE_CHOICE="Python"             # Fill in your chosen software

# --- System info ---
KERNEL=$(uname -r)
USER_NAME="shree"
UPTIME=$(uptime -p)
DISTRO=$(cat /etc/os-release | grep -w "PRETTY_NAME" | cut -d= -f2 | tr -d '"')
# Getting date in a clear format
CURRENT_DATE=$(date +"%Y-%m-%d %H:%M:%S")
# Adding license message
LICENSE_MSG="This operating system is primarily covered by the GNU General Public License (GPL)."

# --- Display ---
echo "================================"
echo " Open Source Audit — $STUDENT_NAME"
echo "================================"
echo "Kernel     : $KERNEL"
echo "Distro     : $DISTRO"
echo "User       : $USER_NAME"
echo "Home Dir   : /home/shree"
echo "Uptime     : $UPTIME"
echo "Date/Time  : $CURRENT_DATE"
echo "================================"
echo "License    : $LICENSE_MSG"
echo "Chosen     : $SOFTWARE_CHOICE"
