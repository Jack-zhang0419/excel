import json

with open('config.json') as config_file:
    # print("read configure")
    config = json.load(config_file)

HEADER_COLOR = config['header_color'] if 'header_color' in config else "FFFFFF"
BLOCK_COLORS = config['block_colors'] if 'block_colors' in config else []
COLUMN_TYPES = config['column_types'] if 'column_types' in config else []


def defined_column_type(column_number):
    filter_column_types = [
        x for x in COLUMN_TYPES if x["column_number"] == column_number
    ]
    if len(filter_column_types) == 1:
        return filter_column_types[0]
    else:
        return None
