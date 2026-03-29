#!/bin/bash
# Script 5: Open Source Manifesto Generator

echo "======================================"
echo " The Open Source Manifesto Generator "
echo "======================================"
echo "Answer three questions to generate your manifesto."
echo ""

read -p "1. Name one open-source tool you use every day: " TOOL
read -p "2. In one word, what does 'freedom' mean to you? " FREEDOM
read -p "3. Name one thing you would build and share freely: " BUILD

DATE=$(date '+%d %B %Y')
USER_NAME="shree"
OUTPUT="manifesto_${USER_NAME}.txt"

# Utilizing an alias conceptually (demonstrated via alias definition)
# alias view_manifest="cat $OUTPUT"

# Compose paragraph using the variables
cat << EOF > "$OUTPUT"
--- The Hacker's Vow ---
Date: $DATE
Author: $USER_NAME

As a technologist, I depend on the foundations laid by others before me. 
Every day, I use tools like $TOOL to perform my work and grow my skills. 
Open source represents the spirit of collaboration where the pursuit of knowledge outvalues corporate walls.

To me, freedom means $FREEDOM. It means the right to look inside the machine, understand how it ticks, and fix it if it breaks. 
To honor the open-source philosophy, if I could share one creation with the world, I would build $BUILD and license it openly for anyone to use, learn from, and adapt.
EOF

echo ""
echo "Manifesto successfully saved to $OUTPUT"
echo "======================================"
cat "$OUTPUT"
