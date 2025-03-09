LANG_PATH := $(APPDATA)/ppm/languages

.PHONY: build

build:
	mkdir -p "$(LANG_PATH)"
	cp -r languages/* "$(LANG_PATH)"

