# -*- coding: UTF-8 -*-

import argparse

if __name__ == '__main__':
    parser = argparse.ArgumentParser()
    parser.add_argument("--excel", type=str, required=False,
                default='./test.xlsx', help='excel path')
    args = parser.parse_args()

    print("[ARGS] excel(%s)"%(args.excel))
