#!/bin/bash
# Script 4: Log File Analyzer
# Usage: ./log_analyzer.sh /var/log/syslog test (or another keyword)

LOGFILE=$1
KEYWORD=${2:-"error"}
COUNT=0

# Default keyword is 'error'
if [ -z "$LOGFILE" ]; then
    echo "Usage: $0 <path_to_log_file> [keyword]"
    exit 1
fi

if [ ! -f "$LOGFILE" ]; then
    echo "Error: File $LOGFILE not found."
    exit 1
fi

# Do-while style retry if the file is empty
if [ ! -s "$LOGFILE" ]; then
    echo "The file is completely empty."
    read -p "Do you want to specify a different file? (y/n) " RETRY
    if [ "$RETRY" = "y" ]; then
        read -p "Enter new filename: " LOGFILE
        if [ ! -f "$LOGFILE" ] || [ ! -s "$LOGFILE" ]; then
            echo "File still invalid or empty. Exiting."
            exit 1
        fi
    else
        exit 1
    fi
fi

while IFS= read -r LINE; do
    if echo "$LINE" | grep -iq "$KEYWORD"; then
        COUNT=$((COUNT + 1))
    fi
done < "$LOGFILE"

echo "======================"
echo "Log Analysis Summary"
echo "======================"
echo "Keyword '$KEYWORD' found $COUNT times in $LOGFILE"
echo ""
echo "Last 5 matching lines:"
tail -n 1000 "$LOGFILE" | grep -i "$KEYWORD" | tail -n 5
