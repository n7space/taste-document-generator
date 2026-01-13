# TASTE Document Generator

## General

TASTE Document Generator (TDG), created as a part of "Model-Based Execution Platform for Space Applications" project (contract 4000146882/24/NL/KK) financed by the European Space Agency.

TDG generates documents by parsing the provided source "template" file, invoking the embedded commands, and merging the resulting documents. Base requirements are provided in MBEP-N7S-EP-SRS, while the overall design is documented in MBEP-N7S-EP-SDD.

## Installation

The recommended way to install TASTE Document Generator locally is via the provided `Makefile` targets:

```bash
make build            # compile the Avalonia application (default Release)
make test             # optional: run the unit + integration suite
make install          # install a shim under ~/.local/bin/TasteDocumentGenerator
```

`make install` depends on `make build`, then creates `~/.local/bin/TasteDocumentGenerator`, a lightweight wrapper that runs the previously built `src/TasteDocumentGenerator/bin/<Config>/net10.0/TasteDocumentGenerator.dll`. Ensure `~/.local/bin` is on your `PATH` so the `TasteDocumentGenerator` command is available from any shell.

## Configuration

The GUI can load its inputs from an XML settings file. The default layout is available at [data/taste-document-generator-settings.xml](data/taste-document-generator-settings.xml) and contains:

- `<InputInterfaceViewPath>`: path to the Interface View (defaults to `interfaceview.xml`).
- `<InputDeploymentViewPath>`: path to the Deployment View (defaults to `deploymentview.dv.xml`).
- `<InputOpus2ModelPath>`: optional OPUS2 model file.
- `<InputTemplatePath>`: source `.docx` template.
- `<OutputFilePath>`: destination document path.
- `<InputTemplateDirectoryPath>`: base directory for relative template lookups.
- `<Target>`: logical target (e.g., `ASW`, `CubeSat`).
- `<DoOpenDocument>`: boolean flag indicating whether the generated document should open automatically after success.

You can point the GUI to a custom settings file by launching it with `TasteDocumentGenerator gui --configuration-path /path/to/settings.xml` or by setting the `TDG_SETTINGS_PATH` environment variable before launching the app.

## Running

The assumed use case is for the TASTE Document Generator to be invoked from Space Creator. However, if TDG is to be used manually, the following command line interface, as documented in the built-in help, is available:

```bash
TasteDocumentGenerator generate --help

TasteDocumentGenerator 1.0.0

	-t, --template-path       Required. Input template file path

	-i, --interface-view      (Default: interfaceview.xml) Input Interface View file path

	-d, --deployment-view     (Default: deploymentview.dv.xml) Input Deployment View file path

	-p, --opus2-model-path    Input OPUS2 model file path

	-o, --output-path         Required. Output file path

	--target                  (Default: ASW) Target system

	--template-directory      (Default: ) Template directory path

	--template-processor      (Default: template-processor) Template processor binary to execute [to support testing]

	--help                    Display this help screen.

	--version                 Display version information.
```

To launch the GUI directly, use `make run` or `TasteDocumentGenerator gui`. To generate a document from the CLI, supply the `generate` verb with the desired paths, for example:

```bash
TasteDocumentGenerator generate \
	-t data/test_in_simple.docx \
	-i interfaceview.xml \
	-d deploymentview.dv.xml \
	-o output.docx \
	--target CubeSat
```

## Frequently Asked Questions (FAQ)

None

## Troubleshooting

None