import re

OFFSET_LETTER_NUMBER = 65


def _parse_file_name_(orginal_file_name):
    """
    return: source_sheet_no, sequence_no, source_sheet_no, source_block_no, orginal_file_name
    for example: 0, 0, 0, 0, A0.xlsx
    if source block = -1, means: copy all rows except header
    """
    file_name = orginal_file_name
    if '.' in file_name:
        file_name = file_name.split('.')[0]  # remove ext

    if '-' not in file_name:
        file_name = f"{file_name}-{file_name}"  # normalize file name

    matched = re.match(r"(\w)(\d+)-(\w)(\d*)", file_name)
    if matched:
        groups = matched.groups()
        # convert to sheet_no A|a -> 0, B|b -> 1
        return ord(groups[0].upper()) - OFFSET_LETTER_NUMBER, int(
            groups[1]), ord(groups[2].upper()) - OFFSET_LETTER_NUMBER, int(
                groups[3]) if groups[3] else -1, orginal_file_name
    else:
        raise ValueError(
            f"{orginal_file_name} not matched file_name standard format, please rename it, and try again."
        )


def parse_file_names(file_list: list[str]):
    parsed_list = []
    for file in file_list:
        parsed = _parse_file_name_(file)
        found = [
            x for x in parsed_list if x[0] == parsed[0] and x[1] == parsed[1]
        ]
        if len(found) > 0:
            raise ValueError(f"duplicate sequence_no found in {file}")

        parsed_list.append(parsed)

    # sorted by sheet_no, and then sorted by sequence_no
    sorted_files = sorted(parsed_list, key=lambda x: (x[0], x[1]))

    return sorted_files
