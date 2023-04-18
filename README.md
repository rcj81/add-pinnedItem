# Add-PinnedItem

This powershell script adds a pinned item.

## Introduction

This script uses Windows API to add a folder as a pinned item. The standard Windows API will add a pinned item and remove it if it is already present. Unfortunately, it does not know how to distinguish between the two beforehand, and it does not return a value. This script will attempt to add a pinned item. If it finds that the item was removed, it will add it back. This is designed to be non-interactive for use with login scripts and group policies.

## Installation

To 'install' this script, download the .ps1 file and run in powershell.

## Parameters

$filePath is the path of the folder you want to pin. This is not flagged as mandatory, but since it is needed, the script will not process if no path is provided.
$chatty is a flag that will write data to the screen when turned on. It is not needed, and may write to the screen depending on how the script is run.

## Usage

Run the powershell script and include the $filePath parameter.

## Contributing

If you would like to contribute to this project, please fork the repository and submit a pull request.

## License

You are free to use this as you like. Please credit the author or link to the repository if possible.
