import argparse
import json
import os
import sys
from pprint import pprint

import HTMLTable
from application import Application


# sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf8')


def factory(classname):
    import application
    cls = getattr(application, classname)
    return cls


ap = argparse.ArgumentParser()
# ap.add_argument("-d", "--debug", required=False, default='0',
#                 help="enable debug mode")
ap.add_argument("-r", "--request_date", required=False,
                help="6 digital number for request date e.g. 202102")
ap.add_argument("-x", "--xlsx", required=False,
                help="source data excel file name in path sourceData")
args = vars(ap.parse_args())

if __name__ == '__main__':

    # app = 'eOrdering'
    request_date = None
    xlsx_file = None

    with open('config.json') as f:
        config = json.load(f)
    apps = config["applications"]
    apps_to_run = config["apps to run"]
    if args["request_date"]:
        request_date = args["request_date"]
    else:
        request_date = config["request date"]

    if args["xlsx"]:
        xlsx_file = args["xlsx"]
    else:
        xlsx_file = config["xlsx sd"]

    arrowUp = '''
        <span style='font-size:19;color:#00B050'>&#9650;</span><span style='color:black'>
        '''
    arrowDown = '''
        <span style='font-size:19;color:#ED7D31'>&#9660;</span><span style='color:black'>
        '''
    html_output = None
    task_summary = {'success': [], 'failed': []}
    percentage_summary = {}

    for app in apps:
        if app not in apps_to_run:
            continue
        app_setting = config["app setting"][app]
        # print(app_setting)
        app_cls = factory(app_setting["class"])
        sd_type = app_setting["sd type"]
        cred_index = ''
        if 'cred_index' in app_setting:
            cred_index = app_setting["cred_index"]
            # print(cred_index)
        keyDataName = app_setting["keyData"]
        instance = app_cls(request_date, cred_index=cred_index)

        try:
            print(f'[{app}] Start to process.....')
            # setattr(instance, 'app', app)
            # setattr(instance, 'cred_index', cred_index)
            # print(f'[{app}] Setting user cred.....')
            # instance.select_user_creds(cred_index=cred_index)
            print(f"[{app}] Loading source data...")
            source_data = None
            if sd_type == "db" or sd_type == "api":
                source_data = instance.load_source_data()
            elif sd_type == "xlsx":
                source_file = instance.find_file(xlsx_file, 'sourceData')
                source_data = instance.load_source_file(source_file)
            print(f"[{app}] Before to html")
            pprint(source_data)

            if source_data is None:
                raise Exception(f"[{app}]:source_data is None")

            print(f"[{app}] Generating trending image...")
            key_data = instance.get_key_data(source_data, keyDataName[0])
            print(f'{keyDataName}:{key_data}')
            trend_res = instance.draw_trending(app, key_data)
            if trend_res is None:
                raise Exception(f"[ERROR] [{app}]: Failed to generate trending image...")

            percentage, trend_img_path = trend_res[0], trend_res[1]
            percentage_summary[app] = (keyDataName[0], percentage)

            # trend_img = instance.get_abs_path(trend_img_path) + "\\" + trend_img_path.split('\\')[1]
            # print(trend_img_path)
            # print(trend_img)

            print(f"[{app}] Creating report html...")
            html_table_data = instance.html_table_data(source_data)
            print(html_table_data)

            app_ta = app_setting["TA"]
            YTD = app_setting["YTD"]
            MTD = app_setting["MTD"]
            UTD = app_setting["UTD"]

            appDisplayName = app_setting["displayName"]

            arrow = ''
            if percentage > 0:
                arrow = arrowUp
            elif percentage < 0:
                arrow = arrowDown

            # trend_img = f'<img border="0" width="215" src="{trend_img}">'
            trend_img_tag = instance.to_html_tag('', 'img', {'border': '0', 'width': '215', 'src': f'{trend_img_path}'})
            # print(trend_img_tag)

            # Fixed size area
            app_header = HTMLTable.TableCell(instance.to_div('Application', is_header=True),
                                             attribs={'bgcolor': '9CC2E5', 'rowspan': 2})
            ta_header = HTMLTable.TableCell(instance.to_div('Target Audience', is_header=True),
                                            attribs={'bgcolor': '9CC2E5', 'rowspan': 2})
            trend_header_text = ''
            if request_date[-2:] == '01':
                trend_header_text = f'{instance.last_month_of(request_date)[0:4]}/{instance.last_month_of(request_date)[-2:]} - {request_date[0:4]}/{request_date[-2:]}</br>{keyDataName[0]} Trend '
            else:
                key_data_len = len(list(filter(lambda d: d[0][0:4] == request_date[0:4], key_data)))
                print(f"{key_data_len}")
                trend_header_text = f'{instance.last_month_of(request_date)[0:4]}/{int(request_date[-2:]) - key_data_len + 1} - {request_date[0:4]}/{request_date[-2:]}</br>{keyDataName[0]} Trend '
                # trend_header_text = f'{instance.last_month_of(request_date)[0:4]}/01 - {request_date[0:4]}/{request_date[-2:]}</br>{keyDataName[0]} Trend '
            trend_header = HTMLTable.TableCell(instance.to_div(trend_header_text, is_header=True),
                                               attribs={'bgcolor': '9CC2E5', 'rowspan': 2})
            app_data = HTMLTable.TableCell(instance.to_div(f'{appDisplayName}'))
            ta_data = HTMLTable.TableCell(instance.to_div(f'{app_ta}'))
            trend_data = HTMLTable.TableCell(instance.to_div(trend_img_tag))

            header_row = [app_header, ta_header, trend_header]
            name_row = []
            data_row = [app_data, ta_data, trend_data]

            if len(YTD) > 0:
                YTD_header = HTMLTable.TableCell(instance.to_div(f'YTD {request_date[0:4]}', is_header=True),
                                                 attribs={'bgcolor': '9CC2E5', 'colspan': len(YTD)})
                header_row.append(YTD_header)

                for e, d in zip(YTD, html_table_data["YTD_data"]):
                    name_row.append(HTMLTable.TableCell(instance.to_div(e, is_header=True), bgcolor='9CC2E5'))
                    data_row.append(instance.to_div(d))
            if len(MTD) > 0:
                MTD_header = HTMLTable.TableCell(
                    instance.to_div(f'MTD {request_date[0:4]}-{request_date[-2:]}', is_header=True),
                    attribs={'bgcolor': '9CC2E5', 'colspan': len(MTD)})
                header_row.append(MTD_header)
                for e, d in zip(MTD, html_table_data["MTD_data"]):
                    name_row.append(HTMLTable.TableCell(instance.to_div(e, is_header=True), bgcolor='9CC2E5'))
                    if e in keyDataName:
                        data_row.append(instance.to_div(d + arrow))
                    else:
                        data_row.append(instance.to_div(d))
            if len(UTD) > 0:
                UTD_header = HTMLTable.TableCell(
                    instance.to_div(f'UTD (Since launch from {app_setting["launchFrom"]})', is_header=True),
                    attribs={'bgcolor': '9CC2E5', 'colspan': len(UTD)})
                header_row.append(UTD_header)
                for e, d in zip(UTD, html_table_data["UTD_data"]):
                    name_row.append(HTMLTable.TableCell(instance.to_div(e, is_header=True), bgcolor='9CC2E5'))
                    data_row.append(instance.to_div(d))

            header_row.append(
                HTMLTable.TableCell(instance.to_div('Remarks / Comments', is_header=True),
                                    attribs={'bgcolor': '9CC2E5', 'rowspan': 2}))
            data_row.append(instance.to_div(' '))

            t = HTMLTable.Table(header_row=header_row)
            t.rows.append(name_row)
            t.rows.append(data_row)
            html_code = str(t)
            # print(html_code)

            html_output = instance.html_output(app, html_code)

            task_summary['success'].append(app)
        except Exception as e:
            exc_type, exc_obj, exc_tb = sys.exc_info()
            fname = os.path.split(exc_tb.tb_frame.f_code.co_filename)[1]
            print(
                f"[ERROR] [{app}]: {e} on {fname} {exc_tb.tb_lineno}"
                f"\nSkipping [{app}] and go to next application... "
            )
            task_summary['failed'].append(app)
            continue
        else:
            pass
        finally:
            pass

        for k, v in percentage_summary.items():
            print(k + ' ' + v[0] + ': Compared to last month ' + str(v[1]) + '%')
    html_output_filename, html_output_path = html_output[0], html_output[1]

    percentage_summary = {k: v for k, v in
                          sorted(percentage_summary.items(), key=lambda item: item[1][1], reverse=True)}
    percentage_summary_html = '</br>'.join(
        [k + ' ' + v[0] + ': Compared to last month ' + str(v[1]) + '%' for k, v in percentage_summary.items()])
    Application.prepend_to_html("App Percentage Summary:", percentage_summary_html, filename=html_output_filename,
                                path=html_output_path)

    task_summary_html = '</br>'.join([k + ': ' + str(l) for k, l in task_summary.items()])
    Application.prepend_to_html("Auto KPI summary:", task_summary_html, filename=html_output_filename,
                                path=html_output_path)
