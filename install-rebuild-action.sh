#!/bin/bash
# Installs a macOS Quick Action (right-click menu) to rebuild the plugin data.
# Run once: bash install-rebuild-action.sh

WORKFLOW_DIR="$HOME/Library/Services/Rebuild VOD Slides.workflow"
PLUGIN_DIR="$(cd "$(dirname "$0")" && pwd)"

mkdir -p "$WORKFLOW_DIR/Contents"

cat > "$WORKFLOW_DIR/Contents/Info.plist" << 'PLIST'
<?xml version="1.0" encoding="UTF-8"?>
<!DOCTYPE plist PUBLIC "-//Apple//DTD PLIST 1.0//EN" "http://www.apple.com/DTDs/PropertyList-1.0.dtd">
<plist version="1.0">
<dict>
	<key>NSServices</key>
	<array>
		<dict>
			<key>NSMenuItem</key>
			<dict>
				<key>default</key>
				<string>Rebuild VOD Slides</string>
			</dict>
			<key>NSMessage</key>
			<string>runWorkflowAsService</string>
		</dict>
	</array>
</dict>
</plist>
PLIST

cat > "$WORKFLOW_DIR/Contents/document.wflow" << WFLOW
<?xml version="1.0" encoding="UTF-8"?>
<!DOCTYPE plist PUBLIC "-//Apple//DTD PLIST 1.0//EN" "http://www.apple.com/DTDs/PropertyList-1.0.dtd">
<plist version="1.0">
<dict>
	<key>AMApplicationBuild</key>
	<string>523</string>
	<key>AMApplicationVersion</key>
	<string>2.10</string>
	<key>AMDocumentVersion</key>
	<string>2</string>
	<key>actions</key>
	<array>
		<dict>
			<key>action</key>
			<dict>
				<key>AMAccepts</key>
				<dict>
					<key>Container</key>
					<string>List</string>
					<key>Optional</key>
					<true/>
					<key>Types</key>
					<array>
						<string>com.apple.cocoa.string</string>
					</array>
				</dict>
				<key>AMActionVersion</key>
				<string>2.0.3</string>
				<key>AMApplication</key>
				<array>
					<string>Automator</string>
				</array>
				<key>AMBundleIdentifier</key>
				<string>com.apple.RunShellScript</string>
				<key>AMName</key>
				<string>Run Shell Script</string>
				<key>AMParameters</key>
				<dict>
					<key>COMMAND_STRING</key>
					<string>cd "$PLUGIN_DIR" &amp;&amp; /usr/bin/python3 rebuild.py 2&gt;&amp;1 | /usr/bin/osascript -e 'on run argv
set theText to ""
repeat with aLine in argv
set theText to theText &amp; aLine &amp; return
end repeat
display dialog theText with title "VOD Slide Builder" buttons {"OK"} default button "OK"
end run' -</string>
					<key>CheckedForUserDefaultShell</key>
					<true/>
					<key>inputMethod</key>
					<integer>0</integer>
					<key>shell</key>
					<string>/bin/bash</string>
					<key>source</key>
					<string></string>
				</dict>
				<key>AMProvides</key>
				<dict>
					<key>Container</key>
					<string>List</string>
					<key>Types</key>
					<array>
						<string>com.apple.cocoa.string</string>
					</array>
				</dict>
			</dict>
		</dict>
	</array>
	<key>connectors</key>
	<dict/>
	<key>workflowMetaData</key>
	<dict>
		<key>workflowTypeIdentifier</key>
		<string>com.apple.Automator.servicesMenu</string>
	</dict>
</dict>
</plist>
WFLOW

echo "Installed! Right-click anywhere in Finder → Quick Actions → 'Rebuild VOD Slides'"
echo "Or find it in System Settings → Extensions → Finder."
