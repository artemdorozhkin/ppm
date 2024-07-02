# PPM - VBA Package Manager

\[eng\]\[[rus](README_ru.md)\]

`ppm` is a package manager developed for VBA and with VBA, providing a command-line interface (CLI-like) through the Immediate Window in the VBA IDE. It's created to help developers manage their VBA projects by organizing code, dependencies, and facilitating common tasks.

## Features

Currently, `ppm` supports a few commands:

- [`init`](#init): Initialises the package.
- [`publish`](#publish): Uploads the project to the server or local registry.
- [`install`](#install): Installs packages with dependencies into the project.
- [`uninstall`](#uninstall): Removes packages with dependencies from the project if other packages do not use them..
- [`export`](#export): Exports modules to the project root folder.
- [`sync`](#sync): Synchronises the project modules with the root folder.
- [`config`](#config): Manages configurations.
- [`help`](#help): Provides usage assistance and descriptions for commands.

## Commands

### init

#### Usage:
ppm "init \[options\]"

ppm "init"

Creates a 'package' module with basic package information.

#### Flags:

-y|--yes Skips the dialogue and sets default values.

-n|--name Sets the value for the project name.

**Example**:
```vb
ppm "init -y"
```

**Result**:
```json
// package.bas
'@Folder("PearPMProject")
'{
'  "name": "PearPM",
'  "version": "1.0.0",
'  "description": "",
'  "author": "",
'  "git": ""
'}
```

**Example**:
```vb
ppm "init -n MyPack -y"
```

**Result**:
```json
// package.bas
'@Folder("PearPMProject")
'{
'  "name": "MyPack",
'  "version": "1.0.0",
'  "description": "",
'  "author": "",
'  "git": ""
'}
```


### publish

#### Usage:
ppm "publish \[options\]"

Uploads the project to the server or local registry.

#### Flags:
-l|--local Publishes the package to the local registry.

-r|--registry Specifies the registry path or URL.

**Example**:
```vb
ppm "publish -l"
```

**Result**:
Publishes the package to the local registry.

**Example**:
```vb
ppm "publish -r http://example.com/registry"
```

**Result**:
Publishes the package to the specified registry URL.

### install

#### Usage:
ppm "install \[options\] \[package\[@version\]\]"

Installs packages with dependencies into the project.

#### Flags:
-l|--local Installs packages and dependencies from the local registry.
-r|--registry Specifies the registry path or URL.

**Example**:
```vb
ppm "install lodash"
```

**Result**:
Installs the latest version of lodash from the default registry.

**Example**:
```vb
ppm "install lodash@4.17.21 -l"
```

**Result**:
Installs version 4.17.21 of lodash from the local registry.

### uninstall

#### Usage:
ppm "uninstall \[package\]"

Removes packages with dependencies from the project if other packages do not use them.

**Example**:
```vb
ppm "uninstall pstrings"
```

**Result**:
Removes the pstrings package from the project.

### export

#### Usage:
ppm "export \[options\]"

Exports modules to the project root folder.

#### Flags:
-e|--encoding Specifies the encoding for exported files.

-s|--save-struct Saves the RubberDuck structure when exporting a project.

-p|--path Specifies the folder to export to.

--no-clear Does not delete files from the last export.

**Example**:
```vb
ppm "export -p ./dist -e UTF-8"
```

**Result**:
Exports project files to the './dist' directory with UTF-8 encoding.

### sync

#### Usage:
ppm "sync"

Synchronises the project modules with the root folder.

#### Example:
```vb
ppm "sync"
```

**Result**:
Synchronises all the project modules with files from the root folder.

### config

#### Usage:
ppm "config \<subcommand\> \[options\]"

Manages the ppm configuration file.

#### Flags:
-g|--global Uses global config.

-l|--location Specifies config location .

**Example**:
```vb
ppm "config set key=value"
```

**Result**:
Sets the config value for the given key.

**Example**:
```vb
ppm "config get key"
```

**Result**:
Gets the config value for the given key.

**Example**:
```vb
ppm "config delete key"
```

**Result**:
Deletes the config entry for the given key.

**Example**:
```vb
ppm "config edit"
```

**Result**:
Opens the config file for editing.

### version

#### Использование:
ppm "version \<new version | major | minor | patch\>"

Sets the new version of the package.

**Example**:
```vb
ppm "version 1.1.1"

' Output: v1.1.1
```

**Result**:
```json
// package.bas
'@Folder("PearPMProject")
'{
'  "name": "PearPM",
'  "version": "1.1.1",
'  "description": "",
'  "author": "",
'  "git": ""
'}
```

**Example**:
```vb
ppm "version patch"

' Output: v1.1.2
```

**Result**:
```json
// package.bas
'@Folder("PearPMProject")
'{
'  "name": "PearPM",
'  "version": "1.1.2",
'  "description": "",
'  "author": "",
'  "git": ""
'}
```

**Example**:
```vb
ppm "version minor"

' Output: v1.2.0
```

**Result**:
```json
// package.bas
'@Folder("PearPMProject")
'{
'  "name": "PearPM",
'  "version": "1.2.0",
'  "description": "",
'  "author": "",
'  "git": ""
'}
```

**Example**:
```vb
ppm "version major"

' Output: v2.0.0
```

**Result**:
```json
// package.bas
'@Folder("PearPMProject")
'{
'  "name": "PearPM",
'  "version": "2.0.0",
'  "description": "",
'  "author": "",
'  "git": ""
'}
```

### help

#### Usage:
ppm "help \[command\]"

Provides usage assistance and descriptions for commands.

#### Example:
```vb
ppm "help init"
```

**Result**:
Shows information about the `init` command.

**Generic Example**:
```vb
ppm "help"
```

**Result**:
Shows general usage information and available commands.

## Contribution

Contributions to `ppm` are welcomed. To contribute, please submit a pull request with a detailed explanation of your proposed changes or enhancements.

## License

The `ppm` VBA package manager is open-source software, and it is available under the MIT License.
