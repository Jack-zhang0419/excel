import json

with open('config.json') as config_file:
    # print("read configure")
    data = json.load(config_file)

HEADER_COLOR = data['header_color']
BLOCK_COLORS = data['block_colors']
