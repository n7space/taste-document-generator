.PHONY: build test run install

CONFIG ?= Release
TARGET_FRAMEWORK ?= net10.0
PROJECT := $(abspath src/TasteDocumentGenerator/TasteDocumentGenerator.csproj)
TEST_PROJECT := $(abspath tests/TasteDocumentGenerator.Tests/TasteDocumentGenerator.Tests.csproj)
BIN_DIR := $(HOME)/.local/bin
SHIM := $(BIN_DIR)/TasteDocumentGenerator
OUTPUT_DLL := $(abspath src/TasteDocumentGenerator/bin/$(CONFIG)/$(TARGET_FRAMEWORK)/TasteDocumentGenerator.dll)

build:
	dotnet build "$(PROJECT)" -c $(CONFIG)

test:
	dotnet test "$(TEST_PROJECT)" -c $(CONFIG) -l:"console;verbosity=normal"

run: build
	dotnet run --project "$(PROJECT)" -c $(CONFIG) -- gui

install: build
	mkdir -p $(BIN_DIR)
	printf '%s\n%s\n' '#!/bin/sh' 'exec dotnet "$(OUTPUT_DLL)" "$$@"' > $(SHIM)
	chmod +x $(SHIM)
