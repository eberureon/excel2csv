# excel2csv

Command Line Tool to convert Excel Files to CSV with pipes as delimiter

## Installation

Use Rust's internal package manager `cargo` to install excel2csv:

```
git clone https://github.com/ebelleon/excel2csv.git
cd excel2csv
cargo install --path .
```

## Usage

```
Convert Excel to CSV

USAGE:
    excel2csv.exe [OPTIONS] --input <input>

FLAGS:
    -h, --help       Prints help information
    -V, --version    Prints version information

OPTIONS:
    -d, --delimiter <delimiter>    Custom Delimiter [default: |]
    -i, --input <input>            Excel to be converted
    -o, --output <output>          Destination for Converted CSV with file ending [default: input Name]
```
