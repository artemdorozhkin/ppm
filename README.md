# PPM - VBA Package Manager

\[eng\]\[[rus](README_ru.md)\]

![GitHub-Mark-Light](./assets/white.jpg#gh-light-mode-only)![GitHub-Mark-Dark](./assets/black.jpg#gh-dark-mode-only)

[Buy me a coffee](https://donate.stream/donate_668a7791e2948)

`ppm` is a package manager developed for VBA and with VBA, providing a command-line interface (CLI-like) through the Immediate Window in the VBA IDE. It's created to help developers manage their VBA projects by organizing code, dependencies, and facilitating common tasks.

## Commands list

Currently, `ppm` supports a few commands:

- [`class`](#class): Manages the project classes (.cls).
- [`config`](#config): Manages configurations.
- [`export`](#export): Exports modules to the project root folder.
- [`help`](#help): Provides usage assistance and descriptions for commands.
- [`init`](#init): Initialises the package.
- [`install`](#install): Installs packages with dependencies into the project.
- [`module`](#module): Manages the project modules (.bas).
- [`ref`](#ref): Manages the project references.
- [`publish`](#publish): Uploads the project to the server or local registry.
- [`search`](#search): Search for the package on the server or locally.
- [`sync`](#sync): Synchronises the project modules with the root folder.
- [`uninstall`](#uninstall): Removes packages with dependencies from the project if other packages do not use them..
- [`version`](#version): Sets the new version of the package.

## Commands

### class

[`⤴️commands list`](#commands-list)

#### Usage:

ppm "class \[subcommand\] \<path\> \[options\]"

Manages the project classes (.cls).

#### Flags:

-c|--create-constructor Create module constructor for the class.

**Example**:

```vb
ppm "class add NewClass"
' or just
ppm "cls NewClass"
```

**Result**:
Adds the NewClass class to the project.

**Example**:

```vb
ppm "class add NewClass --create-constructor"
' or just
ppm "cls NewClass -c"
```

**Result**:
Adds the NewClass class and module constructor to the project.
If class already exists, adds only module constructor.

**Example**:

```vb
ppm "class move /path/Someclass"
' or just
ppm "cls mv /path/Someclass"
```

**Result**:
Moves Someclass and module constructor to RD directory ‘path’.

**Example**:

```vb
ppm "class delete /path/Someclass"
' or just
ppm "cls delete Someclass"
```

**Result**:
Deletes Someclass and module constructor.

### config

[`⤴️commands list`](#commands-list)

#### Usage:

ppm "config \<subcommand\> \[options\]"

Manages the ppm configuration file.

#### Flags:

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

**Example**:

```vb
ppm "config list"
```

**Result**:
Prints all keys and values from the config file.

### export

[`⤴️commands list`](#commands-list)

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

### help

[`⤴️commands list`](#commands-list)

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

### init

[`⤴️commands list`](#commands-list)

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

### install

[`⤴️commands list`](#commands-list)

#### Usage:

ppm "install \[options\] \[package\[@version\]\]"

Installs packages with dependencies into the project.

#### Flags:

-l|--local Installs packages and dependencies from the local registry.
-r|--registry Specifies the registry path or URL.

**Example**:

```vb
ppm "install pstrings"
```

**Result**:
Installs the latest version of pstrings from the default registry.

**Example**:

```vb
ppm "install pstrings@4.17.21 -l"
```

**Result**:
Installs version 4.17.21 of pstrings from the local registry.

### module

[`⤴️commands list`](#commands-list)

#### Usage:

ppm "module \[subcommand\] \<path\>"

Manages the project modules (.bas).

**Example**:

```vb
ppm "module add NewModule"
' or just
ppm "m NewModule"
```

**Result**:
Adds the NewModule module to the project.

**Example**:

```vb
ppm "module move /path/SomeModule"
' or just
ppm "m mv /path/SomeModule"
```

**Result**:
Moves SomeModule to RD directory ‘path’.

**Example**:

```vb
ppm "module delete /path/SomeModule"
' or just
ppm "m delete SomeModule"
```

**Result**:
Deletes SomeModule.

### publish

[`⤴️commands list`](#commands-list)

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

### search

[`⤴️commands list`](#commands-list)

#### Usage:

ppm "search \<package name\>"

Search for the package on the server or locally.

#### Flags:

-l|--local Search for the package in the local registry.

-r|--registry Specifies the registry path or URL.

#### Example:

```vb
ppm "search array --local"
```

**Result**:

```vb
NAME      | VERSION | AUTHOR         | DESCRIPTION          |
ArrayList | 0.5.0   | artemdorozhkin | ppm is a package ... |
PArrays   | 2.0.0   |                |                      |
```

### sync

[`⤴️commands list`](#commands-list)

#### Usage:

ppm "sync"

Synchronises the project modules with the root folder.

#### Example:

```vb
ppm "sync"
```

**Result**:
Synchronises all the project modules with files from the root folder.

### uninstall

[`⤴️commands list`](#commands-list)

#### Usage:

ppm "uninstall \[package\]"

Removes packages with dependencies from the project if other packages do not use them.

**Example**:

```vb
ppm "uninstall pstrings"
```

**Result**:
Removes the pstrings package from the project.

### version

[`⤴️commands list`](#commands-list)

#### Usage:

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

## Contribution

Contributions to `ppm` are welcomed. To contribute, please submit a pull request with a detailed explanation of your proposed changes or enhancements.

## License

The `ppm` VBA package manager is open-source software, and it is available under the MIT License.
