LANG_PATH := $(APP_DATA)/ppm/languages

.PHONY: build

build:
	mkdir -p "$(LANG_PATH)"
	cp -r languages/* "$(LANG_PATH)"
