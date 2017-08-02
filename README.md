# pruef-auswert

## Usage

The main program is pruef_auswert.py. It requires as input the xls input files and the points describing the Exam evaluation:

    max_pts - maximum reachable points
    pmax    - points for 1.0
    pmin    - points for 4.0

Here is an example: `./pruef_auswert.py file1.xls file2.xls 35 25 15`

I have implemented also a very short manual that can be reached easily from the terminal:

    ./pruef_auswert.py --help

## Dependencies

Two packages that does not come with Ubuntu preinstalled are xlutils and xlrd. To install them:

    pip install xlutils xlrd
