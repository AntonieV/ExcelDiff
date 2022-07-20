
import argparse
import logging
import sys

# pip install odfpy
# pd.read_excel("the_document.ods", engine="odf")


def main(argv):
    print("hello world2!")
    parser = argparse.ArgumentParser(description="A tool to compare two excel files "
                                                 "with annotation of the differences.")
    parser.add_argument("-i", "--input-files",
                        type=list,
                        nargs=2,
                        help="Two paths to the Excel files to "
                             "be compared with each other.",
                        required=True)
    parser.add_argument("-o", "--out-dir", type=str,
                        nargs=1,
                        default='',
                        help="Path to the output directory.", required=True)
    parser.add_argument("-v", "-- verbose", type=str, nargs="+",
                        help="...", required=False)
    args = parser.parse_args()

    # if len(args.manufacturer) != len(args.date):
    #     logging.error(
    #         "..."
    #     )
    #     exit(1)


if __name__ == "__main__":
    print("hello world")
    try:
        main(sys.argv)
    except Exception as e:
        logging.exception(e)
