# RepoData Utilities

Utilities for working with data from the [RepoData](https://github.com/tanseyem/RepoData) project.

## Requirements

*   Python 2.7
*   geopy
*   openpyxl

## Installation

1. Download or clone repository: `git clone git@github.com:RockefellerArchiveCenter/RepoDataUtilities`.
2. Install requirements by running `pip install -r requirements.txt`.

## Usage

### Converting Data

`python convert.py source [-o format]`

* source is a valid `.xlsx` file
* format is the output format. Choices are `csv`, `json` and `geojson`. If no format option is specified the script will convert the source file to all formats.

Converted files will be created in the same directory as the source file.

## Contributing

Pull requests welcome!

## Authors

Hillel Arnold

## License

Code is released under the MIT License. See `LICENSE.md` for further information.
