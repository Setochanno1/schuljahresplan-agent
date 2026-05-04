#!/bin/bash

APP_NAME="Schuljahresplan-Agent"
INSTALL_DIR="$HOME/.local/bin"
ICON_DIR="$HOME/.local/share/icons"
DESKTOP_DIR="$HOME/.local/share/applications"

mkdir -p "$INSTALL_DIR"
mkdir -p "$ICON_DIR"
mkdir -p "$DESKTOP_DIR"

cp "$APP_NAME" "$INSTALL_DIR/$APP_NAME"
chmod +x "$INSTALL_DIR/$APP_NAME"

cp "schuljahresplan-agent.png" "$ICON_DIR/schuljahresplan-agent.png"

cat > "$DESKTOP_DIR/schuljahresplan-agent.desktop" <<EOF
[Desktop Entry]
Name=Schuljahresplan Agent
Comment=Schuljahrespläne automatisch erstellen und bearbeiten
Exec=$INSTALL_DIR/$APP_NAME
Icon=$ICON_DIR/schuljahresplan-agent.png
Type=Application
Categories=Education;Office;
Terminal=false
EOF

chmod +x "$DESKTOP_DIR/schuljahresplan-agent.desktop"

echo "Installation abgeschlossen."
echo "Du findest den Schuljahresplan Agent jetzt im App-Menü."
