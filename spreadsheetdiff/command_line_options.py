import argparse
import logging
import os
import sys
from pathlib import Path


def parse_command_line_opts():
    parser = argparse.ArgumentParser(description="A tool to compare two excel files "
                                                 "with annotation of the differences.")
    parser.add_argument("-i", "--input-files",
                        nargs=2,
                        help="Two paths to the Excel files (.xlsx or .ods format) "
                             "to be compared with each other.",
                        required=True)

    parser.add_argument("-o", "--out-dir",
                        nargs=1,
                        help="Path to the output directory.",
                        required=True)

    parser.add_argument("-b", "--bold", action='store_true',
                        help="Displays differences in resulting spreadsheet "
                             "in bold text style.", required=False)

    parser.add_argument("-c", "--bg-color",
                        nargs=1,
                        help="Displays differences in resulting spreadsheet "
                             "in a specific color. The color has to be given in quoted "
                             "hex color code (e.g. '#ff0000') or color name "
                             "(e.g. red).",
                        required=False)

    parser.add_argument("-v", "--verbose", action='store_true',
                        help="Increase output verbosity to debug level.",
                        required=False)

    parser.add_argument("-q", "--quiet", action='store_true',
                        help="Decrease output verbosity to warning level. "
                             "Ignores -v flag.", required=False)

    args = parser.parse_args()
    input_files = args.input_files
    out_dir = ''.join(args.out_dir)
    bold = args.bold
    highlight = ''.join(args.bg_color) if args.bg_color else None
    verbose = args.verbose
    quiet = args.quiet

    style = []

    if "-h" in sys.argv[1:] or "--help" in sys.argv[1:]:
        parser.print_help()
        exit(0)
    else:
        if input_files and len(input_files) == 2:
            path_1 = Path(input_files[0])
            path_2 = Path(input_files[1])
            if not (path_1.is_file() and path_2.is_file()):
                input_files = None
        if out_dir and not os.path.exists(out_dir):
            os.makedirs(out_dir)
        if bold:
            style.append(("bold", True))
        if highlight:
            style.append(("bg_color", f"{highlight}"))
        if not verbose:
            logging.getLogger().setLevel(logging.INFO)
        if quiet:
            logging.getLogger().setLevel(logging.WARNING)

    return input_files, out_dir, style
