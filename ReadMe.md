# Ribbon-Erstellung für Office-Dateien

## Der manuelle Weg
Die Office-Datei umbenennen auf zip.

Aus der Zip-Datei mus die Datei  \\_rels\\.rels exthrahiert werden.

Anschließned muss dieser Verweis am Ende der Auflistung eingefügt werden.

`<Relationship Type="http://schemas.microsoft.com/office/2007/relationships/ui/extensibility" Target="/customUI/customUI14.xml" Id="demoID" />`

Die Id muss eindeutig sein.

Danach wird die Datei in der Zip Datei mit der bestehenden ausgetauscht.

Danach wird ein Ordner customUI angelegt und eine customUI14.xml dort abgelegt.

Beispiel für die Datei customUI14.xml

`<customUI xmlns="http://schemas.microsoft.com/office/2009/07/customui">

	<ribbon>
		<tabs>
			<tab id="customTab" label="Contoso" insertAfterMso="TabHome">
				<group idMso="GroupClipboard" />
				<group idMso="GroupFont" />
				<group id="customGroup" label="Contoso Tools">
					<button id="customButton1" label="ConBold" size="large" onAction="conBoldSub" imageMso="Bold" />
					<button id="customButton2" label="ConItalic" size="large" onAction="conItalicSub" imageMso="Italic" />
					<button id="customButton3" label="ConUnderline" size="large" onAction="conUnderlineSub" imageMso="Underline" />
				</group>
				<group idMso="GroupEnterDataAlignment" />
				<group idMso="GroupEnterDataNumber" />
				<group idMso="GroupQuickFormatting" />
			</tab>
		</tabs>
	</ribbon>
</customUI>`

Der Ordner customUI wird dann in die Zip-Datei kopiert.

Zum Ende wird die Zip-Date wieder auf das Original umbenannt,

# Mit Werkzeugen

Zu erst bietet sich das kostenfreie Werkzeug OfficeCustomUIEditor von Microsoft an. https://github.com/OfficeDev/office-custom-ui-editor.

Ein Nachfolger ist Office RibbonX Editor https://github.com/fernandreu/office-ribbonx-editor

Als Excellösung bietet sich der RibbonX Visual Designer für 32bit an. https://www.andypope.info/vba/ribboneditor.htm