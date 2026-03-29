# Python: Open Source Audit Project

This repository contains five shell scripts and a comprehensive report outline covering the "Open Source Audit" capstone project for the chosen open source software: **Python**.

## Developer Information
- **Student Name**: Shree Nigam
- **Registration Number**: 24BAI10688
- **Chosen Software**: Python

## Project Shell Scripts
All scripts are designed to run on a standard Linux environment.

1. **`sys_identity.sh`**: Displays system information including kernel version, OS distro, logged-in user, and uptime using basic variables and command substitution.
2. **`package_inspector.sh`**: Checks if `python3` is installed using package managers (`dpkg` or `rpm`) and utilizes a `case` statement to echo a philosophical description of various FOSS packages.
3. **`disk_auditor.sh`**: Iterates through an array of directories calculating size and permissions via `ls -ld`, `awk`, and `du` commands.
4. **`log_analyzer.sh`**: Reads a log file iteratively using a `while read` loop to count keyword occurrences (e.g., 'error') and displays the last 5 matching strings.
5. **`manifesto.sh`**: Interactively asks the user three questions to generate a custom Open Source Philosophy manifesto, writing it directly into a newly created `.txt` file.

## Instructions to Run
1. Clone this repository locally or navigate into this folder on your Linux terminal.
2. Grant absolute execution privileges to all files:
   ```bash
   chmod +x *.sh
   ```
3. Run any file using `./script_name.sh`. For example:
   ```bash
   ./sys_identity.sh
   ```
*(Note: Script 4 (`log_analyzer.sh`) requires a file path argument, for example: `./log_analyzer.sh /var/log/syslog error`).*

## Dependencies
- Standard Bash shell (`/bin/bash`).
- Common Linux utilities: `awk`, `cut`, `grep`, `tail`, `dpkg` or `rpm`, `du`. No extra installation required out-of-the-box on mainstream distributions.
