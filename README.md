# mkreport-create

## Foreword
Sometimes, we need to report our progress when we do porting,refactoring etc. This python script was created to make reports in xlsx format.

## Dependencies
Python 3
mkreport-create uses XlsxWriter module to work with xlsx format.
use:
    pip3 install XlsxWriter

## Usage

### mkreport-create

usage: mkreport-create.py [-h] -makeout MAKEOUT [--pvs PVS] [--out OUT]

optional arguments:
  -h, --help        show this help message and exit
  --pvs PVS         PVS csv report
  --out OUT         Output file

required arguments:
  -makeout MAKEOUT  GNU make output


### summary-diff
usage: summary-diff.py [-h] -makeout1 MAKEOUT1 -makeout2 MAKEOUT2 [-pvs1 PVS1]
                       [-pvs2 PVS2] [--out OUT]

optional arguments:
  -h, --help          show this help message and exit
  -pvs1 PVS1          PVS csv report for the previous state
  -pvs2 PVS2          PVS csv report for the current state
  --out OUT           Output file

required arguments:
  -makeout1 MAKEOUT1  GNU make output for the previous state
  -makeout2 MAKEOUT2  GNU make output for the current state



