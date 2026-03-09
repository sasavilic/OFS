EXT = OFS
UNOPKG = unopkg

.PHONY: build install reinstall uninstall clean

build: $(EXT).oxt

$(EXT).oxt: META-INF/manifest.xml description.xml Addons.xcu OFS.xcs ProtocolHandler.xcu python/ofs.py dialogs/ LICENSE
	zip -r $(EXT).oxt META-INF/ description.xml Addons.xcu OFS.xcs ProtocolHandler.xcu python/ofs.py dialogs/ LICENSE

install: build
	$(UNOPKG) remove com.fortunacommerc.ofs 2>/dev/null || true
	$(UNOPKG) add --suppress-license $(EXT).oxt
	@echo "Restart LibreOffice to apply changes."

reinstall: build
	$(UNOPKG) add --force --suppress-license $(EXT).oxt
	@echo "Restart LibreOffice to apply changes."

uninstall:
	$(UNOPKG) remove com.fortunacommerc.ofs 2>/dev/null || true

clean:
	rm -f $(EXT).oxt
