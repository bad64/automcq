.PHONY: install
install:
	@cp automcq.py /usr/local/bin/automcq
	@echo "Installed to /usr/local/bin/automcq"
	
.PHONY: uninstall
uninstall:
	@rm /usr/local/bin/automcq
	@echo "Uninstalled automcq"