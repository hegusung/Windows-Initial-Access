#!/usr/bin/python3
import argparse
import json
from lolbas_parser import list_tags, lolbas_to_json

def main():
    parser = argparse.ArgumentParser(description='LolBas_to_json')
    parser.add_argument("--list-tags", action='store_true', help='List LOLBas tags', dest='list_tags')
    parser.add_argument("--to-json", type=str, help='Put lolbas as a json with paths and tags', dest='to_json')
    parser.add_argument("--path", type=str, help='Path to LOLBas yml directory', required=True, dest='path')

    args = parser.parse_args()

    if args.list_tags:
        print("LOLBas tags:")
        for tag in list_tags(args.path):
            print(" - %s" % tag)
    elif args.to_json:
        lolbas_json = lolbas_to_json(args.path)

        f = open(args.to_json, 'w')
        f.write(json.dumps(lolbas_json))
        f.close()

        print("LOLBas json written to %s" % args.to_json)
    else:
        print("No action provided")



if __name__ == '__main__':
    main()
