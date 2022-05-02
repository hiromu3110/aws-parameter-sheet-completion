import json
import re
import logging
import sys
import boto3
import openpyxl
import jsonpath_ng.ext

logger = logging.getLogger(__name__)
logger.setLevel(logging.INFO)

session = None
clients = dict()
upper_case = re.compile(r'([A-Z])')


# converts a string from camel case to snake case
def to_snake(camel):
    return upper_case.sub(r'_\1', camel)[1:].lower()


# checks whether the action is non-destructive
def is_safe_action(action_name):
    if action_name.startswith('Get'):
        return True
    if action_name.startswith('Describe'):
        return True
    if action_name.startswith('List'):
        return True
    return False


# fetch a single value with JSONPath
def get_value(json, path):
    values = get_values(json, path)
    if 0 < len(values):
        # NOTE: returns the first element
        return values[0]
    else:
        return None


# fetch multiple values with JSONPath
def get_values(json_string, path):
    logger.info('JSONPath string: ' + path)
    values = [x.value for x in jsonpath_ng.ext.parse(path).find(json_string)]
    logger.info('Parsed values: ' + json.dumps(values))
    return values


# resolves the placeholders
def resolve_placeholders(template, symbol, start_cell):
    logger.debug(f'{start_cell.coordinate}: template string is {template}')
    result = template
    cell = start_cell
    for i in range(1, 1 + template.count(symbol)):
        # e.g., %1 -> foo, %2 -> bar
        result = result.replace(f'{symbol}{i}', cell.value)
        cell = cell.offset(column=1)
    logger.debug(f'{cell.coordinate}: resolved string is {result}')
    return result


# calls the API
def invoke(api_name, region_name, action_name, request_params):
    client = clients.get((api_name, region_name))
    if client is None:
        try:
            client = session.client(api_name, region_name=region_name)
            clients[(api_name, region_name)] = client
        except Exception as e:
            logger.error('API name or region name is not valid')
            logger.error('API name: ' + api_name)
            logger.error('region name: ' + region_name)
            raise e

    try:
        method_name = to_snake(action_name)
        method = getattr(client, method_name)
    except Exception as e:
        logger.error('no such action: ' + action_name)
        raise e

    if not is_safe_action(action_name):
        raise Exception('Performing unsafe action was detected: ' + action_name)

    try:
        request = json.loads(request_params)
    except Exception as e:
        logger.error('parameter string is not valid JSON: ' + request_params)
        raise e

    try:
        logger.info('API: ' + api_name)
        logger.info('Region: ' + region_name)
        logger.info('Action: ' + action_name)
        logger.info('Request: ' + json.dumps(request, indent=2))
        response = method(**request)
        logger.info('Response: ' + json.dumps(response, indent=2))
        return response
    except Exception as e:
        logger.error('request is not accepted: ' + json.dumps(request, indent=2))
        raise e


# reads API parameters
def read_api_params(start_cell):
    cell = start_cell

    api_name = cell.value
    logger.debug(f'{cell.coordinate}: API name is [{api_name}]')
    cell = cell.offset(column=1)

    region_name = cell.value
    logger.debug(f'{cell.coordinate}: region name is [{region_name}]')
    cell = cell.offset(column=1)

    action_name = cell.value
    logger.debug(f'{cell.coordinate}: action name is [{action_name}]')
    cell = cell.offset(column=1)

    req_params = cell.value
    logger.debug(f'{cell.coordinate}: request params is {req_params}')
    cell = cell.offset(column=1)

    req_params = resolve_placeholders(req_params, '%', cell)
    logger.debug(f'resolved request params is {req_params}')

    return dict(
        api_name=api_name,
        region_name=region_name,
        action_name=action_name,
        request_params=req_params
    )


# reads the row and builds a JSONPath string
def read_path(start_cell):
    # if a row is [foo, bar, baz], this method returns $.foo.bar.baz
    string = ''

    cell = start_cell
    while cell.value:
        string += '.' + cell.value
        cell = cell.offset(column=1)

    if string:
        logger.debug('JSONPath: ' + string)
        return '$' + string
    else:
        raise Exception('Path is not found: ' + start_cell.coordinate)


# seeks a next column to process
def seek_column_symbol(symbol, start_cell):
    max_column = start_cell.parent.max_column

    cell = start_cell
    while cell.column <= max_column:
        logger.debug(f'seeking [{symbol}] in {cell.coordinate}')
        if cell.value == symbol:
            logger.debug(f'[{symbol}] is found in {cell.coordinate}')
            return cell.offset(column=1)
        cell = cell.offset(column=1)
    logger.info(f'[{symbol}] is not found in {cell.coordinate}')
    raise Exception(f'Symbol [{symbol}] is not found: {start_cell.column_letter}')


# writes a value
def write_value(cell, value):
    if value is None:
        cell.number_format = openpyxl.styles.numbers.FORMAT_GENERAL
        cell.value = '=NA()'
    else:
        cell.number_format = openpyxl.styles.numbers.FORMAT_TEXT
        cell.value = str(value)


def process_worksheet(worksheet):
    response = None

    for row in worksheet.iter_rows(min_col=1, max_col=1):
        for cell in row:
            logger.debug(f'current row is {cell.row}')
            if cell.value == '%':
                input_cell = seek_column_symbol('%%', cell)
                request = read_api_params(input_cell)
                response = invoke(**request)
            elif cell.value == '#':
                input_cell = seek_column_symbol('##', cell)
                path = read_path(input_cell)
                if '%' in path:
                    param_cell = seek_column_symbol('###', cell)
                    path = resolve_placeholders(path, '%', param_cell)
                value = get_value(response, path)
                output_cell = seek_column_symbol('####', cell)
                write_value(output_cell, value)


def process_workbook(src_filename, dst_filename):
    wb = openpyxl.load_workbook(src_filename)
    ws = wb['TargetSheets']
    target_sheets = None
    for col in ws.iter_cols(values_only=True):
        target_sheets = col
        break
    try:
        for sheet_name in target_sheets:
            ws = wb[sheet_name]
            process_worksheet(ws)
        wb.save(dst_filename)
    except Exception as e:
        raise Exception(f'Failed: {sheet_name}') from e


if __name__ == '__main__':
    sh = logging.StreamHandler(sys.stdout)
    logger.addHandler(sh)

    session = boto3.Session(profile_name='default')
    identity = session.client('sts').get_caller_identity()
    logger.info('GetCallerIdentity: ' + json.dumps(identity, indent=2))
    process_workbook('Book1.xlsx', 'Book2.xlsx')
