"""Hello Analytics Reporting API V4."""

# from apiclient.discovery import build
from pprint import pprint
import calendar
from datetime import datetime

from googleapiclient.discovery import build
from oauth2client.service_account import ServiceAccountCredentials

SCOPES = ['https://www.googleapis.com/auth/analytics.readonly']
KEY_FILE_LOCATION = 'named-chariot.json'
VIEW_ID = '238950825'


# VIEW_ID = '140264369'


def initialize_analyticsreporting():
    """Initializes an Analytics Reporting API V4 service object.

    Returns:
      An authorized Analytics Reporting API V4 service object.
    """
    credentials = ServiceAccountCredentials.from_json_keyfile_name(
        KEY_FILE_LOCATION, SCOPES)

    # Build the service object.
    analytics = build('analyticsreporting', 'v4', credentials=credentials)

    return analytics


def get_report_mtd(analytics, request_date):
    """Queries the Analytics Reporting API V4.

    Args:
      analytics: An authorized Analytics Reporting API V4 service object.
    Returns:
      The Analytics Reporting API V4 response.
    """
    return analytics.reports().batchGet(
        body={
            'reportRequests': [
                {
                    'viewId': VIEW_ID,
                    'dateRanges': [request_date_to_range(request_date)],
                    'metrics': [
                        {'expression': 'ga:sessions'},
                        {'expression': 'ga:pageviews'},
                        {'expression': 'ga:users'}
                    ],
                    'dimensions': [
                        {'name': 'ga:year'},
                        {'name': 'ga:month'}
                    ]
                }]
        }
    ).execute()


def get_report_ytd(analytics, request_date):
    """Queries the Analytics Reporting API V4.

    Args:
      analytics: An authorized Analytics Reporting API V4 service object.
    Returns:
      The Analytics Reporting API V4 response.
    """
    return analytics.reports().batchGet(
        body={
            'reportRequests': [
                {
                    'viewId': VIEW_ID,
                    'dateRanges': [request_date_to_range(request_date)],
                    'metrics': [
                        {'expression': 'ga:sessions'},
                        {'expression': 'ga:pageviews'},
                        {'expression': 'ga:users'}
                    ],
                    'dimensions': [
                        {'name': 'ga:year'}

                    ]
                }]
        }
    ).execute()


def print_response_mtd(response):
    """Parses and prints the Analytics Reporting API V4 response.

    Args:
      response: An Analytics Reporting API V4 response.
    """
    source_data = {"sessions": [],
                   "pageviews": [],
                   "users": []}

    for report in response.get('reports', []):
        columnHeader = report.get('columnHeader', {})
        dimensionHeaders = columnHeader.get('dimensions', [])
        metricHeaders = columnHeader.get('metricHeader', {}).get('metricHeaderEntries', [])

        for row in report.get('data', {}).get('rows', []):
            dimensions = row.get('dimensions', [])
            dateRangeValues = row.get('metrics', [])
            the_date = f'{dimensions[0]}{dimensions[1]}'
            # print(the_date)

            for metricHeader, value in zip(metricHeaders, dateRangeValues[0]['values']):
                this_key = remove_ga(metricHeader["name"])
                this_value = (the_date, int(value))
                # print(this_key)
                # print(this_value)
                source_data[this_key].append(this_value)

    source_data['Session'] = source_data['sessions']
    del source_data['sessions']
    source_data['Touchpoint'] = source_data['pageviews']
    del source_data['pageviews']
    source_data['Active users'] = source_data['users']
    del source_data['users']

    # pprint(source_data)
    return source_data


def print_response_ytd(response):
    """Parses and prints the Analytics Reporting API V4 response.

    Args:
      response: An Analytics Reporting API V4 response.
    """
    source_data = {"sessions": [],
                   "pageviews": [],
                   "users": []}

    for report in response.get('reports', []):
        columnHeader = report.get('columnHeader', {})
        dimensionHeaders = columnHeader.get('dimensions', [])
        metricHeaders = columnHeader.get('metricHeader', {}).get('metricHeaderEntries', [])

        for row in report.get('data', {}).get('rows', []):
            dimensions = row.get('dimensions', [])
            dateRangeValues = row.get('metrics', [])
            the_date = f'{dimensions[0]}'
            # print(the_date)

            for metricHeader, value in zip(metricHeaders, dateRangeValues[0]['values']):
                this_key = remove_ga(metricHeader["name"])
                this_value = (the_date, int(value))
                # print(this_key)
                # print(this_value)
                source_data[this_key].append(this_value)

    source_data['Session'] = source_data['sessions']
    del source_data['sessions']
    source_data['Touchpoint'] = source_data['pageviews']
    del source_data['pageviews']
    source_data['Active users'] = source_data['users']
    del source_data['users']

    # pprint(source_data)
    return source_data


def remove_ga(ga):
    if str(ga)[:3] == "ga:":
        return str(ga)[3:]
    else:
        return str(ga)


def request_date_to_range(request_date):
    max_day = \
        calendar.monthrange(datetime.strptime(request_date, '%Y%m').year,
                            datetime.strptime(request_date, '%Y%m').month)[1]
    start_date = datetime.strptime(f"{datetime.strptime(request_date, '%Y%m').year}0101", '%Y%m%d').strftime('%Y-%m-%d')
    end_date = datetime.strptime(request_date + str(max_day), '%Y%m%d').strftime('%Y-%m-%d')
    return {'startDate': start_date, 'endDate': end_date}


def date_merge_into_lillyChina(new_Dict, mtd_or_ytd):
    ytd_lillyChina2020 = {'Session': [('2021', 41930)], 'Touchpoint': [('2021', 100069)],
                          'Active users': [('2021', 27829)]}
    mtd_lillyChina2020 = {
        'Session': [('202101', 12446), ('202102', 9514), ('202103', 14596), ('202104', 5371), ('202105', 3)],
        'Touchpoint': [('202101', 30803), ('202102', 23299), ('202103', 33455), ('202104', 12509), ('202105', 3)],
        'Active users': [('202101', 8300), ('202102', 6731), ('202103', 10868), ('202104', 4176), ('202105', 3)]}

    data_to_merge = {}
    if mtd_or_ytd.lower() == 'mtd':
        data_to_merge = mtd_lillyChina2020
    elif mtd_or_ytd.lower() == 'ytd':
        data_to_merge = ytd_lillyChina2020

    for ak, av in new_Dict.items():
        if ak not in data_to_merge:
            # print(f'{ak} not in data_to_merge')
            data_to_merge[ak] = av
        else:
            # print(f'{ak} in data_to_merge')
            for e in av:
                data_to_merge[ak].append(e)

    # print(data_to_merge)
    merged_data = {}
    for bk, bv in data_to_merge.items():
        my_set = {x[0] for x in bv}
        # print(my_set)
        my_sums = [(i, sum(x[1] for x in bv if x[0] == i)) for i in my_set]
        merged_data[bk] = my_sums
    # print(merged_data)
    return merged_data


def run_mtd(request_date):
    analytics = initialize_analyticsreporting()
    # print(analytics)

    response = get_report_mtd(analytics, request_date)

    # pprint(response.get('reports'))

    # res = date_merge_into_lillyChina(print_response_mtd(response), 'mtd')
    res = print_response_mtd(response)
    return res


def run_ytd(request_date):
    analytics = initialize_analyticsreporting()
    # print(analytics)
    response = get_report_ytd(analytics, request_date)
    # pprint(response.get('reports'))

    # res = date_merge_into_lillyChina(print_response_ytd(response), 'ytd')
    res = print_response_ytd(response)

    return res


if __name__ == '__main__':
    mtd = run_mtd('202201')
    ytd = run_ytd('202201')
    print(mtd)
    print(ytd)
