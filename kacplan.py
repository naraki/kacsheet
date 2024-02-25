#!/usr/bin/env python3
# kacplan.py : 道場予約リスト変換

import sys
import argparse
import kacworkbook as kwb

if __name__ == "__main__":
    conf = {}
    parser = argparse.ArgumentParser(
        prog=sys.argv[0],
        usage=f"python {sys.argv[0]} ", 
        description="予約抽選結果抽出スクリプト.",
        epilog="end", # ヘルプの後に表示
        add_help=True, # -h/–-helpオプションの追加
    )  
    parser.add_argument('excel', type=str, help="excel file of reserve plan")
    parser.add_argument('csv', type=str, help="output csv file of reserve plan")
    parser.add_argument('-v', '--verbose', action='store_true', help="verbose mode")

    args = parser.parse_args()

    if args.verbose:
        conf['verbose'] = True
    else:
        conf['verbose'] = False

    wb = kwb.KacWorkBook(conf, args.excel)
    wb.get_sheet_lists()
    wb.output_csv(args.csv)