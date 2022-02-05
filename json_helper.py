#!/usr/bin/env python
# coding: utf-8

import json
import os


def write_json_file(filename='data.json', json_data={}):
    with open(filename, 'w') as outfile:
        json.dump(json_data, outfile, indent=4, sort_keys=True)


def read_json_file(filename='data.json'):
    exists = os.path.isfile(filename)
    if exists:
        with open(filename) as json_file:
            data = json.load(json_file)
    #         data = json.dumps(data, indent=4, sort_keys=True)
        return data
    else:
        return {}
