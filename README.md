# Parse Media Library

## Overview

This project contains a Python script designed to catalog media files, specifically movies and TV shows, into a structured Excel spreadsheet. The script scans a designated directory, extracts pertinent details from the media files' names, and populates an Excel file with this data, creating a comprehensive media library catalog.

## Features

- **File Scanning**: Recursively scans the specified directory for media files.
- **Detail Extraction**: Parses file names to extract titles, release years, and season/episode information.
- **Excel Integration**: Outputs the collected media information into an Excel file, organizing the library in a tabular format.
- **Duplication Handling**: Includes logic to handle and avoid duplicate entries in the catalog.
- **Size and Modification Time**: Captures file size and last modified time for additional file management capabilities.

## How It Works

The script operates by walking through the file system starting from a given root directory. It identifies media files (currently `.mp4` files) and uses regular expressions to extract metadata from the file names. This metadata typically includes the title of the movie or TV show, the year of release, and, for TV shows, the season and episode numbers.

Once extracted, the script checks for duplicates to ensure each media entry is unique within the catalog. It then appends this information to an Excel file, creating a neatly organized media library. The script also adjusts the width of the columns in the Excel file for better readability, excluding specific columns like file size and modification time from auto-adjustment.

## Examples

![File Directory](img/img(1).png)
![Spreadsheet](img/img(2).png)

## Setup

To run this script, you will need:

- Python 3.x
- pandas (`pip install pandas`)
- openpyxl (`pip install openpyxl`)

## Usage

1. Ensure you have the required Python version and libraries installed.
2. Place the script in a directory of your choice.
3. Modify the `media_directory` variable in the script to point to your media files' location.
4. Run the script using `python PARSE_MEDIA.py`.

## Contributing

Contributions to this project are welcome. Please follow the standard fork-and-pull request workflow on GitHub.

## License

This project is licensed under the MIT License - see the LICENSE.md file for details.

## Disclaimer

This script is for personal use and educational purposes only. Please ensure you have the rights to access and manage the files you use with this script.
