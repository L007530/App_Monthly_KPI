import json
import sys
import os
from datetime import datetime, date
from collections import defaultdict

# sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf8')
from pprint import pprint

import matplotlib.pyplot as plt
import pandas as pd
import pyodbc


class Application:
    def __init__(self, request_date, db_conn_timeout='30', app=''):
        self.request_date = request_date
        self.db_conn_timeout = db_conn_timeout
        self.app = app

    @staticmethod
    def find_file(name, path):
        for root, dirs, files in os.walk(path):
            if name in files:
                return os.path.join(root, name)
            else:
                # print(f"[FileNotFoundError] Cannot find '{name}' under '{path}'")
                return None

    # @staticmethod
    # def get_abs_path(file):
    #     return str(pathlib.Path(file).parent.absolute())
    #     # return os.path.dirname(os.path.abspath(file))

    @staticmethod
    def get_absolute_path(file_path):
        """
        This function takes a file path as input and returns its absolute path.

        :param file_path: The path of the file (string)
        :return: The absolute path of the file (string)
        """
        # Get the absolute path of the file
        absolute_path = f"{os.path.abspath(file_path)}"

        # Return the absolute path
        return absolute_path

    @staticmethod
    def create_path(*args):
        return os.path.join(*args)

    @staticmethod
    def prepend_to_html(title, prepend_data, filename, path='output'):
        # html=find_file(name, path)
        # with open(f"{path}\\{filename}", 'r+') as f:
        with open(Application.create_path(path, f"{filename}"), 'r+') as f:
            content = f.read()
            f.seek(0, 0)
            f.write(Application.to_html_tag(title, tag='div', attribs={
                'style': 'font-size:20px;font-weight:bold;margin-top:20px'})
                    + '</br>' + str(prepend_data).rstrip('\r\n') + '</br></br>' + content)

    def html_output(self, app, html_data, ext_info=None, filename='HTML App Monthly report',
                    path='output'):
        # with open(f"{path}\\{filename}_{self.request_date}.html", "a") as html:
        with open(Application.create_path(path, f"{filename}_{self.request_date}.html"), "a") as html:
            html.writelines(
                Application.to_html_tag('&#10146;' + app, attribs={
                    'style': 'font-size:17px;font-weight:bold;margin-bottom:10px;margin-top:20px'}))
            # html.write(html_data)
            for line in html_data:
                # print(line)
                html.writelines(line)
            if ext_info:
                html.writelines(f'\n{ext_info}')
            html.writelines('</br></br>')
        return f"{filename}_{self.request_date}.html", path

    @staticmethod
    def last_month_of(yyyymm):
        if len(yyyymm) != 6:
            print("[ERROR]Date format not correct, e.g. '202012' is accepted")
            return None
        else:
            try:
                m = int(yyyymm[-2:])
                y = int(yyyymm[:4])
                dt = date(y, m, 1)
                last_dt = dt - pd.DateOffset(months=1)
                return str(last_dt).split('-')[0] + str(last_dt).split('-')[1]
            except ValueError:
                print('[ValueError] Date argument not correct')
                return None

    @staticmethod
    def number_with_comma(number):
        try:
            return "{:,}".format(int(number))
        except ValueError:
            print('[ValueError] Only number(integer) is accepted - Method: number_with_comma()')
            return None

    @staticmethod
    def to_div(text, is_header=False):
        attribs = {'style': 'text-align:center'}
        if is_header:
            attribs = {'style': 'text-align:center;font-weight:bold'}
        return Application.to_html_tag(text, 'div', attribs=attribs)

    @staticmethod
    def to_html_tag(text, tag='div', attribs=None):
        if attribs is None:
            attribs = {}
        atts = [k + '="' + v + '" ' for k, v in attribs.items()]
        atts_str = ''
        for att in atts:
            atts_str += att
        content = text
        if text is None:
            content = ''
        # print(atts_str)
        code = f'''
            <{tag} {atts_str}>{content}</{tag}>
            '''
        return code

    @staticmethod
    def create_dir_if_not_exist(path):
        if not os.path.exists(path):
            os.makedirs(path)
        # elif os.path.exists(path):
        #     print(f'Existed {path}, did nothing')

    def get_key_data(self, source_data, key_name):
        if source_data is None:
            return None
        return source_data[key_name]

    def draw_trending(self, app, key_source_data):
        if self.request_date is None or key_source_data is None:
            return None
        try:
            if str(datetime.strptime(self.request_date, "%Y%m").month) == '1' or \
                    str(datetime.strptime(self.request_date, "%Y%m").month) == '01':
                last_month = self.last_month_of(self.request_date)

                x = [100, 200]
                y = [e[1] for e in key_source_data if
                     datetime.strptime(e[0], "%Y%m") == datetime.strptime(last_month, "%Y%m") or
                     datetime.strptime(e[0], "%Y%m") == datetime.strptime(self.request_date, "%Y%m")]
                if x != y:
                    x = x[:len(y)]
                plt.plot(x, y, linestyle='-', marker='o', ms=30, mfc='r', mec='r', linewidth=6)
                plt.axis('off')
                figure = plt.gcf()
                figure.set_size_inches(24, 3)
                # img_path = f"output\\trending"
                img_path = Application.create_path("output", "trending")
                Application.create_dir_if_not_exist(img_path)
                img_name = f"{app}_{self.request_date}.png"
                # img_path_re = f"trending\\{img_name}"
                img_path_re = Application.create_path("trending", f"{img_name}")
                # plt.savefig(f"{img_path}\\{img_name}", dpi=100, bbox_inches='tight')
                plt.savefig(Application.create_path(img_path, img_name), dpi=100, bbox_inches='tight')

                if y[-2] == 0:
                    percentage = 0
                else:
                    percentage = (y[-1] / y[-2] - 1) * 100
                plt.clf()
                return round(percentage, 2), img_path_re
            else:
                y = [e[1] for e in key_source_data if
                     datetime.strptime(e[0], "%Y%m") <= datetime.strptime(self.request_date, "%Y%m") and
                     datetime.strptime(e[0], "%Y%m").year == datetime.strptime(self.request_date,
                                                                               "%Y%m").year]

                x = [i for i in range(1, int(datetime.strptime(self.request_date, "%Y%m").month) + 1)]
                if x != y:
                    x = x[:len(y)]
                # print(f'y={len(y)}')
                # print(f'x={len(x)}')
                plt.plot(x, y, linestyle='-', marker='o', ms=25, mfc='r', mec='r', linewidth=6)
                plt.axis('off')
                figure = plt.gcf()
                figure.set_size_inches(24, 3)
                # img_path = f"output\\trending"
                img_path = Application.create_path("output", "trending")
                Application.create_dir_if_not_exist(img_path)
                img_name = f"{app}_{self.request_date}.png"
                # img_path_re = f"trending\\{img_name}"
                img_path_re = Application.create_path("trending", f"{img_name}")
                plt.savefig(f"{img_path}\\{img_name}", dpi=100, bbox_inches='tight')
                plt.savefig(Application.create_path(f"{img_path}", f"{img_name}"), dpi=100, bbox_inches='tight')

                if y[-2] == 0:
                    percentage = 100
                else:
                    percentage = (y[-1] / y[-2] - 1) * 100
                plt.clf()
                return round(percentage, 2), img_path_re
        except Exception as e:
            print(e)
            return None

    def load_source_file(self, source_file):
        if source_file is None:
            return None
        try:
            df = pd.read_excel(source_file, sheet_name=self.app, header=0, index_col=0)
            source_date = {
                k: sorted(list(filter(lambda e: e[0] != 'nan', [(str(p)[:6], q) for p, q in v.items()])),
                          key=lambda ele: datetime.strptime((ele[0]), '%Y%m'),
                          reverse=False)
                for k, v in df.to_dict().items()}
            return source_date
        except Exception as e:
            print(str(e))
            return None

    @staticmethod
    def load_user_creds(cred_file_path, cred_index="cred1"):
        cred_file = Application.find_file(cred_file_path, '.')
        if cred_file is None:
            return None
        with open(cred_file) as f:
            creds = json.load(f)
        cred = creds[cred_index]
        return cred

    def print_error(self, error_msg, fname, exc_tb_line):
        print(f"[ERROR] [{self.app}]: {error_msg} on {fname} {exc_tb_line}", file=sys.stderr)


class EOrdering(Application):
    # def __init__(self, request_date, cred_index='cred3'):
    #     super().__init__(request_date)
    #     self.cred_index = cred_index
    #     self.svr_name = 'GH3DMZDBPHA.xh3.lilly.com,1433'
    #     self.db_name = 'eOrdering'
    #     self.UID = Application.load_user_creds('creds.json', cred_index=cred_index)['UID']
    #     self.pwd = Application.load_user_creds('creds.json', cred_index=cred_index)['pwd']

    def __init__(self, request_date, cred_index=''):
        super().__init__(request_date)
        self.cred_index = cred_index
        self.app = "eOrdering"

    # def load_source_data(self):
    #     sql_order = '''
    #             SELECT
    #             concat(datepart(yyyy, [request_date]), datepart(mm, [request_date])) as 'request_month',
    #             COUNT(id) as 'orders_number'
    #             FROM [e_order_app].[app_form_instance]
    #             where [is_deleted]=0 and [status] <> -99
    #             group by concat(datepart(yyyy, [request_date]), datepart(mm, [request_date]))
    #         '''
    #     sql_vendor = '''
    #                     select count([id]) from [e_order_sys].[sys_user] where [is_deleted]=0 and [vendor_code] is not null and [vendor_code]<>''
    #                     '''
    #     try:
    #         conn = pyodbc.connect('Driver={SQL Server};'
    #                               f'Server={self.svr_name};'
    #                               f'Database={self.db_name};'
    #                               f'UID={self.UID};'
    #                               f'PWD={self.pwd};'
    #                               f'Trusted_Connection=no;'
    #                               f'timeout={self.db_conn_timeout};')
    #         conn.timeout = int(self.db_conn_timeout)
    #         cursor = conn.cursor()
    #         cursor.execute(sql_order)
    #         orders_number = [d for d in cursor]
    #         orders_number.sort(key=lambda date: datetime.strptime(date[0], "%Y%m"))
    #         source_data = {"Orders number": orders_number}
    #
    #         cursor.execute(sql_vendor)
    #         source_data["Total vendor number"] = [d for d in cursor][0][0]
    #
    #         max_date_of_source_data = source_data["Orders number"][-1][0]
    #         min_date_of_source_data = source_data["Orders number"][0][0]
    #         if datetime.strptime(max_date_of_source_data, "%Y%m") < datetime.strptime(self.request_date, "%Y%m"):
    #             raise Exception(
    #                 f"[ERROR]Request date must be (<=) equal to or earlier than: {max_date_of_source_data}")
    #         elif datetime.strptime(min_date_of_source_data, "%Y%m") > datetime.strptime(self.request_date, "%Y%m"):
    #             raise Exception(
    #                 f"[ERROR]Request date must be (>=) equal to or later than: {min_date_of_source_data}")
    #         return source_data
    #     except Exception as e:
    #         print(str(e))
    #         return None

    # def get_key_data(self, source_data, key_name):
    #     if source_data is None:
    #         return None
    #     # key_name = "Orders number"
    #     return source_data[key_name]

    def html_table_data(self, source_data):
        if source_data is None:
            return None

        request_mtd_source_data = [e[1] for e in source_data["Orders number"] if
                                   datetime.strptime(e[0], "%Y%m") == datetime.strptime(
                                       self.request_date,
                                       "%Y%m")]

        request_utd_source_data = [e[1] for e in source_data["Orders number (UTD)"] if
                                   datetime.strptime(e[0], "%Y%m") == datetime.strptime(
                                       self.request_date,
                                       "%Y%m")]

        request_ytd_source_data = [e[1] for e in source_data["Orders number (YTD)"] if
                                   datetime.strptime(e[0], "%Y%m") == datetime.strptime(
                                       self.request_date,
                                       "%Y%m")]

        total_vendor_number = [e[1] for e in source_data["Total vendor number"] if
                               datetime.strptime(e[0], "%Y%m") == datetime.strptime(
                                   self.request_date,
                                   "%Y%m")]

        YTD_data = [self.number_with_comma(total_vendor_number[0]), self.number_with_comma(request_ytd_source_data[0])]
        MTD_data = [self.number_with_comma(total_vendor_number[0]), self.number_with_comma(request_mtd_source_data[0])]
        UTD_data = [self.number_with_comma(total_vendor_number[0]), self.number_with_comma(request_utd_source_data[0])]

        return {"YTD_data": YTD_data, "MTD_data": MTD_data, "UTD_data": UTD_data}


class LCCP(Application):
    def __init__(self, request_date, cred_index=''):
        super().__init__(request_date)
        self.cred_index = cred_index
        self.app = "LCCP"

    # def load_source_file(self, source_file):
    #     if source_file is None:
    #         return None
    #     try:
    #         df = pd.read_excel(source_file, sheet_name=self.app, header=0, index_col=0)
    #         source_date = {
    #             k: sorted([(str(p)[:6], q) for p, q in v.items()], key=lambda e: datetime.strptime((e[0]), '%Y%m'),
    #                       reverse=False)
    #             for k, v in df.to_dict().items()}
    #         return source_date
    #     except Exception as e:
    #         print(str(e))
    #         return None
    #
    # def get_key_data(self, source_data, key_name):
    #     if source_data is None:
    #         return None
    #     # key_name = "Patient testing record"
    #     return source_data[key_name]

    def html_table_data(self, source_data):
        if source_data is None:
            return None
        request_utd_New_HCP_enrollment = [e[1] for e in source_data["New HCP enrollment"] if
                                          datetime.strptime(e[0], "%Y%m") <= datetime.strptime(self.request_date,
                                                                                               "%Y%m")]

        request_ytd_New_HCP_enrollment = [e[1] for e in source_data["New HCP enrollment"] if
                                          datetime.strptime(e[0], "%Y%m") <= datetime.strptime(self.request_date,
                                                                                               "%Y%m") and
                                          datetime.strptime(e[0], "%Y%m").year == datetime.strptime(self.request_date,
                                                                                                    "%Y%m").year]
        request_utd_New_Patient_enrollment = [e[1] for e in source_data["New Patient enrollment"] if
                                              datetime.strptime(e[0], "%Y%m") <= datetime.strptime(self.request_date,
                                                                                                   "%Y%m")]

        request_ytd_New_Patient_enrollment = [e[1] for e in source_data["New Patient enrollment"] if
                                              datetime.strptime(e[0], "%Y%m") <= datetime.strptime(self.request_date,
                                                                                                   "%Y%m") and
                                              datetime.strptime(e[0], "%Y%m").year == datetime.strptime(
                                                  self.request_date,
                                                  "%Y%m").year]

        request_utd_Patient_testing_record = [e[1] for e in source_data["Patient testing record"] if
                                              datetime.strptime(e[0], "%Y%m") <= datetime.strptime(self.request_date,
                                                                                                   "%Y%m")]

        request_ytd_Patient_testing_record = [e[1] for e in source_data["Patient testing record"] if
                                              datetime.strptime(e[0], "%Y%m") <= datetime.strptime(self.request_date,
                                                                                                   "%Y%m") and
                                              datetime.strptime(e[0], "%Y%m").year == datetime.strptime(
                                                  self.request_date,
                                                  "%Y%m").year]

        YTD_data = [self.number_with_comma(sum(request_ytd_New_HCP_enrollment)),
                    self.number_with_comma(sum(request_ytd_New_Patient_enrollment)),
                    self.number_with_comma(sum(request_ytd_Patient_testing_record))]
        MTD_data = [self.number_with_comma(request_ytd_New_HCP_enrollment[-1]),
                    self.number_with_comma(request_ytd_New_Patient_enrollment[-1]),
                    self.number_with_comma(request_ytd_Patient_testing_record[-1])]
        UTD_data = [self.number_with_comma(sum(request_utd_New_HCP_enrollment)),
                    self.number_with_comma(sum(request_utd_New_Patient_enrollment)),
                    self.number_with_comma(sum(request_utd_Patient_testing_record))]

        return {"YTD_data": YTD_data, "MTD_data": MTD_data, "UTD_data": UTD_data}


# EMSL current
class MLWechat(Application):
    def __init__(self, request_date, cred_index='cred1'):
        super().__init__(request_date)
        self.cred_index = cred_index
        self.svr_name = 'GH3DMZDBPHA.XH3.LILLY.COM,1433'
        self.db_name = 'mlwechat'
        self.UID = Application.load_user_creds('creds.json', cred_index=self.cred_index)['UID']
        self.pwd = Application.load_user_creds('creds.json', cred_index=self.cred_index)['pwd']
        self.app = "MLWechat"

    def load_source_data(self):
        sql = f'''
                    select 
            count(isnull(push.seid,hit.seid) ) Content_covered_SE_number ,
            sum(isnull(NoOfPush,0) ) Content_update_number ,
            sum(isnull(NoOfHits,0)) Touchpoint, 
             isnull(push.ym,hit.ym) ym
        
            from 
            (
                SELECT
                a.SEID,
                COUNT(DISTINCT a.MessageHeadID) NoOfPush,
                CONVERT(NVARCHAR(6), [SubmittedOn],112) as ym
                FROM [dbo].[MessageHead] a INNER JOIN [dbo].[User] b ON a.UserID=b.UserID
                INNER JOIN [dbo].[MessageCategory] c ON a.MessageHeadID=c.MessageHeadID
                INNER JOIN [dbo].[Message] d ON c.MessageCategoryID=d.MessageCategoryID
                INNER JOIN [dbo].[SE] e ON a.SEID=e.SEID
                WHERE  e.Deleted=0 AND b.Deleted=0 AND a.[Status] IN (2,3,4) 
                AND CHARINDEX('_test',b.Name,1)=0
                AND  CHARINDEX('_test',e.Name,1)=0
                GROUP BY 
                a.SEID,
                CONVERT(NVARCHAR(6), [SubmittedOn],112)
                --order by CONVERT(NVARCHAR(6), [SubmittedOn],112) desc
            ) push
            full outer join
            (
                SELECT 
                b.SEID,
                COUNT(a.PaperID) NoOfHits,
                CONVERT(NVARCHAR(6), [ReadOn],112) ym					
                FROM [dbo].[PaperReadTrack] a 
                JOIN [dbo].[WechatSE] b ON a.WechatSEID=b.WechatSEID
                JOIN [dbo].[SE] c ON b.SEID=c.SEID
                JOIN [dbo].[vw_User] d ON b.UserID=d.UserID 
                WHERE 
                    CHARINDEX('_test',c.Name,1)=0
                    AND CHARINDEX('_test',d.Name,1)=0  
                group by b.SEID,CONVERT(NVARCHAR(6), [ReadOn],112)
            ) hit on push.seid=hit.SEID and push.ym=hit.ym
            where 
            isnull(push.ym,hit.ym) >=202001
            group by isnull(push.ym,hit.ym)  
            order by isnull(push.ym,hit.ym)  

        '''
        try:
            conn = pyodbc.connect('Driver={SQL Server};'
                                  f'Server={self.svr_name};'
                                  f'Database={self.db_name};'
                                  f'UID={self.UID};'
                                  f'PWD={self.pwd};'
                                  f'Trusted_Connection=no;'
                                  f'timeout={self.db_conn_timeout};')
            conn.timeout = int(self.db_conn_timeout)
            cursor = conn.cursor()
            cursor.execute(sql)
            dt = [d for d in cursor]
            dt.sort(key=lambda date: datetime.strptime(date[3], "%Y%m"))
            dt = list(filter(lambda x: (x[3] <= self.request_date), dt))
            pprint(dt)
            source_data = {"Content covered SE number": [(x[3], x[0]) for x in dt]}
            source_data['Content update number'] = [(x[3], x[1]) for x in dt]
            source_data['Touchpoint'] = [(x[3], x[2]) for x in dt]
        except Exception as e:
            exc_type, exc_obj, exc_tb = sys.exc_info()
            fname = os.path.split(exc_tb.tb_frame.f_code.co_filename)[1]
            Application.print_error(self, error_msg=e, fname=fname, exc_tb_line=exc_tb.tb_lineno)
            return None
        else:
            return source_data

    def get_key_data(self, source_data, key_name):
        if source_data is None:
            return None
        return source_data[key_name]

    def html_table_data(self, source_data):
        if source_data is None:
            return None
        # request_utd_New_HCP_enrollment = [e[1] for e in source_data["New HCP enrollment"] if
        #                                   datetime.strptime(e[0], "%Y%m") <= datetime.strptime(self.request_date,
        #                                                                                        "%Y%m")]
        try:
            request_ytd_Content_covered_SE_number = [e[1] for e in source_data["Content covered SE number"] if
                                                     datetime.strptime(e[0], "%Y%m") <= datetime.strptime(
                                                         self.request_date,
                                                         "%Y%m") and
                                                     datetime.strptime(e[0], "%Y%m").year == datetime.strptime(
                                                         self.request_date,
                                                         "%Y%m").year]
            # request_utd_New_Patient_enrollment = [e[1] for e in source_data["New Patient enrollment"] if
            #                                       datetime.strptime(e[0], "%Y%m") <= datetime.strptime(self.request_date,
            #                                                                                            "%Y%m")]

            request_ytd_Content_update_number = [e[1] for e in source_data["Content update number"] if
                                                 datetime.strptime(e[0], "%Y%m") <=
                                                 datetime.strptime(self.request_date, "%Y%m") and
                                                 datetime.strptime(e[0], "%Y%m").year == datetime.strptime(
                                                     self.request_date, "%Y%m").year]

            # request_utd_Patient_testing_record = [e[1] for e in source_data["Patient testing record"] if
            #                                       datetime.strptime(e[0], "%Y%m") <= datetime.strptime(self.request_date,
            #                                                                                            "%Y%m")]

            request_ytd_Touchpoint = [e[1] for e in source_data["Touchpoint"] if
                                      datetime.strptime(e[0], "%Y%m") <= datetime.strptime(self.request_date,
                                                                                           "%Y%m") and
                                      datetime.strptime(e[0], "%Y%m").year == datetime.strptime(
                                          self.request_date,
                                          "%Y%m").year]

            YTD_data = [self.number_with_comma(sum(request_ytd_Content_covered_SE_number)),
                        self.number_with_comma(sum(request_ytd_Content_update_number)),
                        self.number_with_comma(sum(request_ytd_Touchpoint))]
            MTD_data = [self.number_with_comma(request_ytd_Content_covered_SE_number[-1]),
                        self.number_with_comma(request_ytd_Content_update_number[-1]),
                        self.number_with_comma(request_ytd_Touchpoint[-1])]
            # UTD_data = [self.number_with_comma(sum(request_utd_New_HCP_enrollment)),
            #             self.number_with_comma(sum(request_utd_New_Patient_enrollment)),
            #             self.number_with_comma(sum(request_utd_Patient_testing_record))]
        except Exception as e:
            exc_type, exc_obj, exc_tb = sys.exc_info()
            fname = os.path.split(exc_tb.tb_frame.f_code.co_filename)[1]
            Application.print_error(self, error_msg=e, fname=fname, exc_tb_line=exc_tb.tb_lineno)
            return None
        else:
            return {"YTD_data": YTD_data, "MTD_data": MTD_data, "UTD_data": []}


class MLWechat_xlsx(MLWechat, Application):
    def __init__(self, request_date, cred_index=''):
        super().__init__(request_date)
        self.cred_index = cred_index
        self.app = "MLWechat"


class RAPID(Application):
    def __init__(self, request_date, cred_index=''):
        super().__init__(request_date)
        self.cred_index = cred_index
        self.app = "RAPID"

    def html_table_data(self, source_data):
        if source_data is None:
            return None
        request_ytd_Total_Report_number = \
            list(filter(lambda e: e[0] == self.request_date, source_data["Total Report number"]))[0][1]
        # print("request_ytd_Total_Report_number=" + str(request_ytd_Total_Report_number))
        request_mtd_Total_Report_number = request_ytd_Total_Report_number

        request_ytd_Access_Number = [e[1] for e in source_data["Access Number"] if
                                     datetime.strptime(e[0], "%Y%m") <= datetime.strptime(self.request_date,
                                                                                          "%Y%m") and
                                     datetime.strptime(e[0], "%Y%m").year == datetime.strptime(
                                         self.request_date,
                                         "%Y%m").year]

        request_ytd_Active_users = list(filter(lambda e: e[0] == \
                                                         self.request_date, source_data["Active users(YTD)"]))[0][1]
        request_mtd_Active_users = list(filter(lambda e: e[0] == \
                                                         self.request_date, source_data["Active users"]))[0][1]

        YTD_data = [self.number_with_comma(request_ytd_Total_Report_number),
                    self.number_with_comma(sum(request_ytd_Access_Number)),
                    self.number_with_comma(request_ytd_Active_users)]
        MTD_data = [self.number_with_comma(request_mtd_Total_Report_number),
                    self.number_with_comma(request_ytd_Access_Number[-1]),
                    self.number_with_comma(request_mtd_Active_users)]

        return {"YTD_data": YTD_data, "MTD_data": MTD_data, "UTD_data": []}


class Chatbot_Abandoned(Application):
    def html_table_data(self, source_data):
        if source_data is None:
            return None
        request_ytd_Satisfactory_Number = [e[1] for e in source_data["Satisfactory Number"] if
                                           datetime.strptime(e[0], "%Y%m") <= datetime.strptime(self.request_date,
                                                                                                "%Y%m") and
                                           datetime.strptime(e[0], "%Y%m").year == datetime.strptime(
                                               self.request_date,
                                               "%Y%m").year]
        request_mtd_Satisfactory_Number = request_ytd_Satisfactory_Number[-1]

        request_ytd_Not_Rate_Number = [e[1] for e in source_data["Not Rate Number"] if
                                       datetime.strptime(e[0], "%Y%m") <= datetime.strptime(self.request_date,
                                                                                            "%Y%m") and
                                       datetime.strptime(e[0], "%Y%m").year == datetime.strptime(
                                           self.request_date,
                                           "%Y%m").year]
        request_mtd_Not_Rate_Number = request_ytd_Not_Rate_Number[-1]

        request_ytd_Unsatisfactory_Number = [e[1] for e in source_data["Unsatisfactory Number"] if
                                             datetime.strptime(e[0], "%Y%m") <= datetime.strptime(self.request_date,
                                                                                                  "%Y%m") and
                                             datetime.strptime(e[0], "%Y%m").year == datetime.strptime(
                                                 self.request_date,
                                                 "%Y%m").year]
        request_mtd_Unsatisfactory_Number = request_ytd_Unsatisfactory_Number[-1]

        request_ytd_Conversation_Number = [e[1] for e in source_data["Conversation Number"] if
                                           datetime.strptime(e[0], "%Y%m") <= datetime.strptime(self.request_date,
                                                                                                "%Y%m") and
                                           datetime.strptime(e[0], "%Y%m").year == datetime.strptime(
                                               self.request_date,
                                               "%Y%m").year]
        request_mtd_Conversation_Number = request_ytd_Conversation_Number[-1]

        request_ytd_Satisfaction_Rate = (sum(request_ytd_Satisfactory_Number) + sum(request_ytd_Not_Rate_Number)) \
                                        / sum(request_ytd_Conversation_Number)
        request_mtd_Satisfaction_Rate = \
            list(filter(lambda e: e[0] == self.request_date, source_data["Satisfaction Rate"]))[0][1]

        YTD_data = [self.number_with_comma(sum(request_ytd_Satisfactory_Number)),
                    self.number_with_comma(sum(request_ytd_Not_Rate_Number)),
                    self.number_with_comma(sum(request_ytd_Unsatisfactory_Number)),
                    self.number_with_comma(sum(request_ytd_Conversation_Number)),
                    f'{round(float(request_ytd_Satisfaction_Rate), 3) * 100}%']
        MTD_data = [self.number_with_comma(request_mtd_Satisfactory_Number),
                    self.number_with_comma(request_mtd_Not_Rate_Number),
                    self.number_with_comma(request_mtd_Unsatisfactory_Number),
                    self.number_with_comma(request_mtd_Conversation_Number),
                    f'{round(float(request_mtd_Satisfaction_Rate), 3) * 100}%']

        return {"YTD_data": YTD_data, "MTD_data": MTD_data, "UTD_data": []}


class Chatbot(Application):
    def html_table_data(self, source_data):
        if source_data is None:
            return None

        request_ytd_Active_User = [e[1] for e in source_data["Active User(YTD)"] if
                                   datetime.strptime(e[0], "%Y%m") == datetime.strptime(
                                       self.request_date,
                                       "%Y%m")]
        request_mtd_Active_User = [e[1] for e in source_data["Active User"] if
                                   datetime.strptime(e[0], "%Y%m") == datetime.strptime(
                                       self.request_date,
                                       "%Y%m")]

        request_ytd_Conversation_Number = [e[1] for e in source_data["Conversation Number(YTD)"] if
                                           datetime.strptime(e[0], "%Y%m") == datetime.strptime(
                                               self.request_date,
                                               "%Y%m")]
        request_mtd_Conversation_Number = [e[1] for e in source_data["Conversation Number"] if
                                           datetime.strptime(e[0], "%Y%m") == datetime.strptime(
                                               self.request_date,
                                               "%Y%m")]

        request_ytd_Satisfaction_Rate = [e[1] for e in source_data["Satisfaction Rate(YTD)"] if
                                         datetime.strptime(e[0], "%Y%m") == datetime.strptime(
                                             self.request_date,
                                             "%Y%m")]
        request_mtd_Satisfaction_Rate = [e[1] for e in source_data["Satisfaction Rate"] if
                                         datetime.strptime(e[0], "%Y%m") == datetime.strptime(
                                             self.request_date,
                                             "%Y%m")]
        request_ytd_Accuracy_Rate = [e[1] for e in source_data["Accuracy Rate(YTD)"] if
                                     datetime.strptime(e[0], "%Y%m") == datetime.strptime(
                                         self.request_date,
                                         "%Y%m")]
        request_mtd_Accuracy_Rate = [e[1] for e in source_data["Accuracy Rate"] if
                                     datetime.strptime(e[0], "%Y%m") == datetime.strptime(
                                         self.request_date,
                                         "%Y%m")]
        request_ytd_Reply_Rate = [e[1] for e in source_data["Reply Rate(YTD)"] if
                                  datetime.strptime(e[0], "%Y%m") == datetime.strptime(
                                      self.request_date,
                                      "%Y%m")]
        request_mtd_Reply_Rate = [e[1] for e in source_data["Reply Rate"] if
                                  datetime.strptime(e[0], "%Y%m") == datetime.strptime(
                                      self.request_date,
                                      "%Y%m")]

        YTD_data = [self.number_with_comma(request_ytd_Active_User[-1]),
                    self.number_with_comma(request_ytd_Conversation_Number[-1]),
                    f'{round(float(request_ytd_Satisfaction_Rate[-1]), 3) * 100}%',
                    f'{round(float(request_ytd_Accuracy_Rate[-1]), 3) * 100}%',
                    f'{round(float(request_ytd_Reply_Rate[-1]), 3) * 100}%']
        MTD_data = [self.number_with_comma(request_mtd_Active_User[-1]),
                    self.number_with_comma(request_mtd_Conversation_Number[-1]),
                    f'{round(float(request_mtd_Satisfaction_Rate[-1]), 3) * 100}%',
                    f'{round(float(request_mtd_Accuracy_Rate[-1]), 3) * 100}%',
                    f'{round(float(request_mtd_Reply_Rate[-1]), 3) * 100}%']

        return {"YTD_data": YTD_data, "MTD_data": MTD_data, "UTD_data": []}


class Chatbot_HCP(Chatbot, Application):
    def __init__(self, request_date, cred_index=''):
        super().__init__(request_date)
        self.cred_index = cred_index
        self.app = "Chatbot_HCP"

    # def html_table_data(self, source_data):
    #     if source_data is None:
    #         return None
    #     request_ytd_Active_User = \
    #         list(filter(lambda e: e[0] == self.request_date, source_data["Active User(YTD)"]))[0][1]
    #     request_mtd_Active_User = \
    #         list(filter(lambda e: e[0] == self.request_date, source_data["Active User"]))[0][1]
    #
    #     request_ytd_Conversation_Number = \
    #         list(filter(lambda e: e[0] == self.request_date,
    #                     source_data["Conversation Number(YTD)"]))[0][1]
    #     request_mtd_Conversation_Number = \
    #         list(filter(lambda e: e[0] == self.request_date, source_data["Conversation Number"]))[0][1]
    #
    #     request_ytd_Satisfaction_Rate = \
    #         list(filter(lambda e: e[0] == self.request_date,
    #                     source_data["Satisfaction Rate(YTD)"]))[0][1]
    #     request_mtd_Satisfaction_Rate = \
    #         list(filter(lambda e: e[0] == self.request_date, source_data["Satisfaction Rate"]))[0][1]
    #
    #     YTD_data = [self.number_with_comma(request_ytd_Active_User),
    #                 self.number_with_comma(request_ytd_Conversation_Number),
    #                 f'{request_ytd_Satisfaction_Rate * 100}%']
    #     MTD_data = [self.number_with_comma(request_mtd_Active_User),
    #                 self.number_with_comma(request_mtd_Conversation_Number),
    #                 f'{request_mtd_Satisfaction_Rate * 100}%']
    #
    #     return {"YTD_data": YTD_data, "MTD_data": MTD_data, "UTD_data": []}


class Chatbot_LCCP(Application):
    def __init__(self, request_date, cred_index=''):
        super().__init__(request_date)
        self.cred_index = cred_index
        self.app = "Chatbot_LCCP"

    def html_table_data(self, source_data):
        if source_data is None:
            return None
        request_ytd_Active_User = \
            list(filter(lambda e: e[0] == self.request_date, source_data["Active User(YTD)"]))[0][1]
        request_mtd_Active_User = \
            list(filter(lambda e: e[0] == self.request_date, source_data["Active User"]))[0][1]

        request_ytd_Conversation_Number = \
            list(filter(lambda e: e[0] == self.request_date,
                        source_data["Conversation Number(YTD)"]))[0][1]
        request_mtd_Conversation_Number = \
            list(filter(lambda e: e[0] == self.request_date, source_data["Conversation Number"]))[0][1]

        request_ytd_Accuracy = \
            list(filter(lambda e: e[0] == self.request_date,
                        source_data["Satisfaction (Only Count in Scope)(YTD)"]))[0][1]
        request_mtd_Accuracy = \
            list(filter(lambda e: e[0] == self.request_date, source_data["Satisfaction (Only Count in Scope)"]))[0][1]

        YTD_data = [self.number_with_comma(request_ytd_Active_User),
                    self.number_with_comma(request_ytd_Conversation_Number),
                    f'{request_ytd_Accuracy * 100}%']
        MTD_data = [self.number_with_comma(request_mtd_Active_User),
                    self.number_with_comma(request_mtd_Conversation_Number),
                    f'{request_mtd_Accuracy * 100}%']

        return {"YTD_data": YTD_data, "MTD_data": MTD_data, "UTD_data": []}


class Chatbot_HR(Chatbot, Application):
    def __init__(self, request_date, cred_index=''):
        super().__init__(request_date)
        self.cred_index = cred_index
        self.app = "Chatbot_HR"


class Chatbot_MOE(Chatbot, Application):
    def __init__(self, request_date, cred_index=''):
        super().__init__(request_date)
        self.cred_index = cred_index
        self.app = "Chatbot_MOE"


class Chatbot_SFE(Chatbot, Application):
    def __init__(self, request_date, cred_index=''):
        super().__init__(request_date)
        self.cred_index = cred_index
        self.app = "Chatbot_SFE"


class Chatbot_LISHAN(Chatbot, Application):
    def __init__(self, request_date, cred_index=''):
        super().__init__(request_date)
        self.cred_index = cred_index
        self.app = "Chatbot_LISHAN"


class Chatbot_FINANCE(Chatbot, Application):
    def __init__(self, request_date, cred_index=''):
        super().__init__(request_date)
        self.cred_index = cred_index
        self.app = "Chatbot_FINANCE"


class Chatbot_IT(Chatbot, Application):
    def __init__(self, request_date, cred_index=''):
        super().__init__(request_date)
        self.cred_index = cred_index
        self.app = "Chatbot_IT"


class Chatbot_Procurement(Chatbot, Application):
    def __init__(self, request_date, cred_index=''):
        super().__init__(request_date)
        self.cred_index = cred_index
        self.app = "Chatbot_Procurement"


class WeChatEnt(Application):
    def __init__(self, request_date, cred_index='cred2'):
        super().__init__(request_date)
        self.cred_index = cred_index
        self.svr_name = 'GH3DMZDBPHA.xh3.lilly.com,1433'
        self.db_name = 'Wechat_MGT_PLT'
        self.UID = Application.load_user_creds('creds.json', cred_index=cred_index)['UID']
        self.pwd = Application.load_user_creds('creds.json', cred_index=cred_index)['pwd']
        self.app = "WeChatEnt"

    def load_source_data(self):
        # 频道访问量
        sql_channels_access_number = f"""
                                    select CONVERT(nvarchar(6), AccessDate , 112) as 'YYYYMM',
                                    sum(AccessCount) as 'Channels Access Number'
                                    FROM dbo.AppAccessReport 
                                    where CONVERT(nvarchar(6), AccessDate , 112)>= '{self.request_date[0:4]}01' and CONVERT(nvarchar(6), AccessDate , 112)<= '{self.request_date}'
                                    group by CONVERT(nvarchar(6), AccessDate , 112)

                                    """
        # 频道数量
        sql_channels = f" SELECT count( [Id]) 'Channels' FROM [Wechat_MGT_PLT].[dbo].[SysWechatConfig] where IsDeleted=0"

        # 活跃用户数 month
        sql_Active_users_m = f"""
                            SELECT count([GroupBy1].[K1]) as 'active user'
                            FROM ( SELECT [Extent1].[UserId] AS [K1], COUNT(1) AS [A1] FROM [dbo].[UserBehavior] AS [Extent1]
                            where CONVERT(nvarchar(6), [Extent1].[CreatedTime] , 112)>= '{self.request_date}' and CONVERT(nvarchar(6), [Extent1].[CreatedTime] , 112)<= '{self.request_date}'
                            AND ( NOT ((9 = [Extent1].[ContentType]) AND ([Extent1].[ContentType] IS NOT NULL)))
                            AND ( NOT (([Extent1].[UserId] IS NULL) OR (( CAST(LEN([Extent1].[UserId]) AS int)) = 0)))
                            GROUP BY [Extent1].[UserId] ) AS [GroupBy1]
                            """

        # 活跃用户数 year
        sql_Active_users_y = f"""
                                    SELECT count([GroupBy1].[K1]) as 'active user'
                                    FROM ( SELECT [Extent1].[UserId] AS [K1], COUNT(1) AS [A1] FROM [dbo].[UserBehavior] AS [Extent1]
                                    where CONVERT(nvarchar(6), [Extent1].[CreatedTime] , 112)>= '{self.request_date[0:4]}01' and CONVERT(nvarchar(6), [Extent1].[CreatedTime] , 112)<= '{self.request_date}'
                                    AND ( NOT ((9 = [Extent1].[ContentType]) AND ([Extent1].[ContentType] IS NOT NULL)))
                                    AND ( NOT (([Extent1].[UserId] IS NULL) OR (( CAST(LEN([Extent1].[UserId]) AS int)) = 0)))
                                    GROUP BY [Extent1].[UserId] ) AS [GroupBy1]
                                    """

        try:
            conn = pyodbc.connect('Driver={SQL Server};'
                                  f'Server={self.svr_name};'
                                  f'Database={self.db_name};'
                                  f'UID={self.UID};'
                                  f'PWD={self.pwd};'
                                  f'Trusted_Connection=no;'
                                  f'timeout={self.db_conn_timeout};')
            conn.timeout = int(self.db_conn_timeout)
            cursor = conn.cursor()
            cursor.execute(sql_channels_access_number)
            channels_access_number = [d for d in cursor]
            channels_access_number.sort(key=lambda date: datetime.strptime(date[0], "%Y%m"))
            source_data = {"Channels Access Number": channels_access_number}

            cursor.execute(sql_channels)
            source_data["Channels"] = [d for d in cursor][0][0]

            cursor.execute(sql_Active_users_m)
            source_data["Active users"] = [d for d in cursor][0][0]

            cursor.execute(sql_Active_users_y)
            source_data["Active users (year)"] = [d for d in cursor][0][0]

            # max_date_of_source_data = source_data["Orders number"][-1][0]
            # min_date_of_source_data = source_data["Orders number"][0][0]
            # if datetime.strptime(max_date_of_source_data, "%Y%m") < datetime.strptime(self.request_date, "%Y%m"):
            #     raise Exception(
            #         f"[ERROR]Request date must be (<=) equal to or earlier than: {max_date_of_source_data}")
            # elif datetime.strptime(min_date_of_source_data, "%Y%m") > datetime.strptime(self.request_date, "%Y%m"):
            #     raise Exception(
            #         f"[ERROR]Request date must be (>=) equal to or later than: {min_date_of_source_data}")
            return source_data
        except Exception as e:
            print(str(e))
            return None

    def html_table_data(self, source_data):
        if source_data is None:
            return None
        request_ytd_Channels = source_data["Channels"]
        # list(filter(lambda e: e[0] == self.request_date, source_data["Channels(YTD)"]))[0][1]

        request_mtd_Channels = source_data["Channels"]
        # list(filter(lambda e: e[0] == self.request_date, source_data["Channels"]))[0][1]

        request_ytd_Channels_Access_Number = [e[1] for e in source_data["Channels Access Number"] if
                                              datetime.strptime(e[0], "%Y%m") <= datetime.strptime(self.request_date,
                                                                                                   "%Y%m") and
                                              datetime.strptime(e[0], "%Y%m").year == datetime.strptime(
                                                  self.request_date,
                                                  "%Y%m").year]
        # request_mtd_Channels_Access_Number = \
        #     list(filter(lambda e: e[0] == self.request_date, source_data["Channels Access Number"]))[0][1]

        request_ytd_Active_users = source_data["Active users (year)"]
        # list(filter(lambda e: e[0] == self.request_date, source_data["Active users(YTD)"]))[0][1]
        request_mtd_Active_users = source_data["Active users"]
        # list(filter(lambda e: e[0] == self.request_date, source_data["Active users"]))[0][1]

        YTD_data = [self.number_with_comma(request_ytd_Channels),
                    self.number_with_comma(sum(request_ytd_Channels_Access_Number)),
                    self.number_with_comma(request_ytd_Active_users)]
        MTD_data = [self.number_with_comma(request_mtd_Channels),
                    self.number_with_comma(request_ytd_Channels_Access_Number[-1]),
                    self.number_with_comma(request_mtd_Active_users)]

        return {"YTD_data": YTD_data, "MTD_data": MTD_data, "UTD_data": []}


class WeChatEnt_xlsx(WeChatEnt, Application):
    def __init__(self, request_date, cred_index=''):
        super().__init__(request_date)
        self.cred_index = cred_index
        self.app = "WeChatEnt"

    def html_table_data(self, source_data):
        if source_data is None:
            return None
        try:
            request_ytd_Active_users = \
                list(filter(lambda e: e[0] == self.request_date, source_data["Active users (year)"]))[0][1]
            request_mtd_Active_users = \
                list(filter(lambda e: e[0] == self.request_date, source_data["Active users"]))[0][1]

            request_mtd_Channels = \
                list(filter(lambda e: e[0] == self.request_date, source_data["Channels"]))[0][1]
            request_ytd_Channels = request_mtd_Channels

            request_ytd_Channels_Access_Number = [e[1] for e in source_data["Channels Access Number"] if
                                                  datetime.strptime(e[0], "%Y%m") <= datetime.strptime(
                                                      self.request_date,
                                                      "%Y%m") and
                                                  datetime.strptime(e[0], "%Y%m").year == datetime.strptime(
                                                      self.request_date,
                                                      "%Y%m").year]

            YTD_data = [self.number_with_comma(request_ytd_Channels),
                        self.number_with_comma(sum(request_ytd_Channels_Access_Number)),
                        self.number_with_comma(request_ytd_Active_users)]
            MTD_data = [self.number_with_comma(request_mtd_Channels),
                        self.number_with_comma(request_ytd_Channels_Access_Number[-1]),
                        self.number_with_comma(request_mtd_Active_users)]
        except Exception as e:
            exc_type, exc_obj, exc_tb = sys.exc_info()
            fname = os.path.split(exc_tb.tb_frame.f_code.co_filename)[1]
            Application.print_error(self, error_msg=e, fname=fname, exc_tb_line=exc_tb.tb_lineno)
            return None
        else:
            return {"YTD_data": YTD_data, "MTD_data": MTD_data, "UTD_data": []}


class Coterie(Application):
    def __init__(self, request_date, cred_index=''):
        super().__init__(request_date)
        self.cred_index = cred_index
        self.app = "Coterie"

    # def load_source_file(self, source_file):
    #     if source_file is None:
    #         return None
    #     try:
    #         df = pd.read_excel(source_file, sheet_name=self.app, header=0, index_col=0)
    #         source_date = {
    #             k: sorted([(str(p)[:6], q) for p, q in v.items()], key=lambda ele: datetime.strptime((ele[0]), '%Y%m'),
    #                       reverse=False)
    #             for k, v in df.to_dict().items()}
    #         return source_date
    #     except Exception as e:
    #         print(str(e))
    #         return None

    def html_table_data(self, source_data):
        if source_data is None:
            return None
        request_mtd_Total_Vendor_number = \
            list(filter(lambda e: e[0] == self.request_date, source_data["Total Vendor number"]))[0][1]
        request_ytd_Total_Vendor_number = request_mtd_Total_Vendor_number
        request_utd_Total_Vendor_number = request_mtd_Total_Vendor_number

        request_ytd_Payment_number = [e[1] for e in source_data["Payment number"] if
                                      datetime.strptime(e[0], "%Y%m") <= datetime.strptime(self.request_date,
                                                                                           "%Y%m") and
                                      datetime.strptime(e[0], "%Y%m").year == datetime.strptime(
                                          self.request_date,
                                          "%Y%m").year]

        request_mtd_Payment_number = \
            list(filter(lambda e: e[0] == self.request_date, source_data["Payment number"]))[0][1]
        request_utd_Payment_number = [e[1] for e in source_data["Payment number"] if
                                      datetime.strptime(e[0], "%Y%m") <= datetime.strptime(self.request_date, "%Y%m")]

        request_ytd_Active_users = \
            list(filter(lambda e: e[0] == self.request_date, source_data["Active users(YTD)"]))[0][1]
        request_mtd_Active_users = \
            list(filter(lambda e: e[0] == self.request_date, source_data["Active users"]))[0][1]
        request_utd_Active_users = \
            list(filter(lambda e: e[0] == self.request_date, source_data["Active users(UTD)"]))[0][1]

        YTD_data = [self.number_with_comma(request_ytd_Total_Vendor_number),
                    self.number_with_comma(sum(request_ytd_Payment_number)),
                    self.number_with_comma(request_ytd_Active_users), ]
        MTD_data = [self.number_with_comma(request_mtd_Total_Vendor_number),
                    self.number_with_comma(request_mtd_Payment_number),
                    self.number_with_comma(request_mtd_Active_users)]
        UTD_data = [self.number_with_comma(request_utd_Total_Vendor_number),
                    self.number_with_comma(sum(request_utd_Payment_number)),
                    self.number_with_comma(request_utd_Active_users)]

        return {"YTD_data": YTD_data, "MTD_data": MTD_data, "UTD_data": UTD_data}


class AcronymBot(Application):
    def __init__(self, request_date, cred_index=''):
        super().__init__(request_date)
        self.cred_index = cred_index
        self.app = "AcronymBot"

    # def load_source_file(self, source_file):
    #     if source_file is None:
    #         return None
    #     try:
    #         df = pd.read_excel(source_file, sheet_name=self.app, header=0, index_col=0)
    #         source_date = {
    #             k: sorted([(CHIEF Cloudstr(p), q) for p, q in v.items()], key=lambda ele: datetime.strptime((ele[0]), '%Y%m'),
    #                       reverse=False)
    #             for k, v in df.to_dict().items()}
    #         return source_date
    #     except Exception as e:
    #         print(str(e))
    #         return None

    def html_table_data(self, source_data):
        if source_data is None:
            return None
        request_ytd_Total_Number_Searches = [e[1] for e in source_data["Total Number Searches"] if
                                             datetime.strptime(e[0], "%Y%m") <= datetime.strptime(self.request_date,
                                                                                                  "%Y%m") and
                                             datetime.strptime(e[0], "%Y%m").year == datetime.strptime(
                                                 self.request_date,
                                                 "%Y%m").year]
        request_mtd_Total_Number_Searches = request_ytd_Total_Number_Searches[-1]

        request_ytd_Total_Number_Users = \
            list(filter(lambda e: e[0] == self.request_date, source_data["Total Number Users(YTD)"]))[0][1]
        request_mtd_Total_Number_Users = \
            list(filter(lambda e: e[0] == self.request_date, source_data["Total Number Users"]))[0][1]

        request_ytd_Total_Success_Rate = \
            list(filter(lambda e: e[0] == self.request_date, source_data["Total Success Rate(YTD)"]))[0][1]
        request_mtd_Total_Success_Rate = \
            list(filter(lambda e: e[0] == self.request_date, source_data["Total Success Rate"]))[0][1]

        YTD_data = [self.number_with_comma(sum(request_ytd_Total_Number_Searches)),
                    self.number_with_comma(request_ytd_Total_Number_Users),
                    f"{request_ytd_Total_Success_Rate * 100}%"]
        MTD_data = [self.number_with_comma(request_mtd_Total_Number_Searches),
                    self.number_with_comma(request_mtd_Total_Number_Users),
                    f"{request_mtd_Total_Success_Rate * 100}%"]

        return {"YTD_data": YTD_data, "MTD_data": MTD_data, "UTD_data": []}


class IPatient(Application):
    def __init__(self, request_date, cred_index=''):
        super().__init__(request_date)
        self.cred_index = cred_index
        self.app = "iPatient"

    def html_table_data(self, source_data):
        if source_data is None:
            return None
        request_ytd_Subscribed_users = \
            list(filter(lambda e: e[0] == self.request_date, source_data["Subscribed users(YTD)"]))[0][1]
        request_mtd_Subscribed_users = \
            list(filter(lambda e: e[0] == self.request_date, source_data["Subscribed users"]))[0][1]
        request_utd_Subscribed_users = \
            list(filter(lambda e: e[0] == self.request_date, source_data["Subscribed users(UTD)"]))[0][1]

        request_ytd_Touchpoint = \
            list(filter(lambda e: e[0] == self.request_date, source_data["Touchpoint(YTD)"]))[0][1]
        request_mtd_Touchpoint = \
            list(filter(lambda e: e[0] == self.request_date, source_data["Touchpoint"]))[0][1]
        request_utd_Touchpoint = \
            list(filter(lambda e: e[0] == self.request_date, source_data["Touchpoint(UTD)"]))[0][1]

        request_mtd_Active_users = \
            list(filter(lambda e: e[0] == self.request_date, source_data["Active users"]))[0][1]

        request_mtd_TP_Active_user = \
            list(filter(lambda e: e[0] == self.request_date, source_data["TP/Active user"]))[0][1]

        YTD_data = [self.number_with_comma(request_ytd_Subscribed_users),
                    self.number_with_comma(request_ytd_Touchpoint)]
        MTD_data = [self.number_with_comma(request_mtd_Subscribed_users),
                    self.number_with_comma(request_mtd_Touchpoint),
                    self.number_with_comma(request_mtd_Active_users),
                    self.number_with_comma(request_mtd_TP_Active_user)]
        UTD_data = [self.number_with_comma(request_utd_Subscribed_users),
                    self.number_with_comma(request_utd_Touchpoint)]

        return {"YTD_data": YTD_data, "MTD_data": MTD_data, "UTD_data": UTD_data}


class LillyMedical(Application):
    def __init__(self, request_date, cred_index=''):
        super().__init__(request_date)
        self.cred_index = cred_index
        self.app = "LillyMedical"

    def html_table_data(self, source_data):
        if source_data is None:
            return None
        request_ytd_Total_page_visit = [e[1] for e in source_data["Total page visit"] if
                                        datetime.strptime(e[0], "%Y%m") <= datetime.strptime(self.request_date,
                                                                                             "%Y%m") and
                                        datetime.strptime(e[0], "%Y%m").year == datetime.strptime(
                                            self.request_date, "%Y%m").year]
        request_mtd_Total_page_visit = request_ytd_Total_page_visit[-1]
        request_utd_Total_page_visit = [e[1] for e in source_data["Total page visit"] if
                                        datetime.strptime(e[0], "%Y%m") <= datetime.strptime(self.request_date, "%Y%m")]

        request_ytd_FAQ_page_visit = [e[1] for e in source_data["FAQ page visit"] if
                                      datetime.strptime(e[0], "%Y%m") <= datetime.strptime(self.request_date,
                                                                                           "%Y%m") and
                                      datetime.strptime(e[0], "%Y%m").year == datetime.strptime(
                                          self.request_date, "%Y%m").year]
        request_mtd_FAQ_page_visit = request_ytd_FAQ_page_visit[-1]
        request_utd_FAQ_page_visit = [e[1] for e in source_data["FAQ page visit"] if
                                      datetime.strptime(e[0], "%Y%m") <= datetime.strptime(self.request_date, "%Y%m")]

        YTD_data = [self.number_with_comma(sum(request_ytd_Total_page_visit)),
                    self.number_with_comma(sum(request_ytd_FAQ_page_visit))]
        MTD_data = [self.number_with_comma(request_mtd_Total_page_visit),
                    self.number_with_comma(request_mtd_FAQ_page_visit)]
        UTD_data = [self.number_with_comma(sum(request_utd_Total_page_visit)),
                    self.number_with_comma(sum(request_utd_FAQ_page_visit))]

        return {"YTD_data": YTD_data, "MTD_data": MTD_data, "UTD_data": UTD_data}


class LillyMedical_MiniProgram(LillyMedical, Application):
    def __init__(self, request_date, cred_index=''):
        super().__init__(request_date)
        self.cred_index = cred_index
        self.app = "LillyMedical_MiniProgram"


# class Chief(Application):
#     def __init__(self, request_date, cred_index=''):
#         super().__init__(request_date)
#         self.cred_index = cred_index
#         self.app = "CHIEF"
#
#     def html_table_data(self, source_data):
#         if source_data is None:
#             return None
#         request_ytd_Forms = \
#             list(filter(lambda e: e[0] == self.request_date, source_data["Forms"]))[0][1]
#         request_mtd_Forms = \
#             list(filter(lambda e: e[0] == self.request_date, source_data["Forms"]))[0][1]
#
#         request_ytd_Submit_records = [e[1] for e in source_data["Submit records"] if
#                                       datetime.strptime(e[0], "%Y%m") <= datetime.strptime(self.request_date,
#                                                                                            "%Y%m") and
#                                       datetime.strptime(e[0], "%Y%m").year == datetime.strptime(
#                                           self.request_date, "%Y%m").year]
#         request_mtd_Submit_records = \
#             list(filter(lambda e: e[0] == self.request_date, source_data["Submit records"]))[0][1]
#
#         request_ytd_Approval_records = [e[1] for e in source_data["Approval records"] if
#                                         datetime.strptime(e[0], "%Y%m") <= datetime.strptime(self.request_date,
#                                                                                              "%Y%m") and
#                                         datetime.strptime(e[0], "%Y%m").year == datetime.strptime(
#                                             self.request_date, "%Y%m").year]
#         request_mtd_Approval_records = \
#             list(filter(lambda e: e[0] == self.request_date, source_data["Approval records"]))[0][1]
#
#         YTD_data = [self.number_with_comma(request_ytd_Forms),
#                     self.number_with_comma(sum(request_ytd_Submit_records)),
#                     self.number_with_comma(sum(request_ytd_Approval_records))]
#         MTD_data = [self.number_with_comma(request_mtd_Forms),
#                     self.number_with_comma(request_mtd_Submit_records),
#                     self.number_with_comma(request_mtd_Approval_records)]
#
#         return {"YTD_data": YTD_data, "MTD_data": MTD_data, "UTD_data": []}


class ChiefCloud(Application):
    def __init__(self, request_date, cred_index=''):
        super().__init__(request_date)
        self.cred_index = cred_index
        self.app = "CHIEF Cloud"

    def html_table_data(self, source_data):
        if source_data is None:
            return None
        request_ytd_Forms = \
            list(filter(lambda e: e[0] == self.request_date, source_data["Forms"]))[0][1]
        request_mtd_Forms = \
            list(filter(lambda e: e[0] == self.request_date, source_data["Forms"]))[0][1]
        request_utd_Forms = \
            list(filter(lambda e: e[0] == self.request_date, source_data["Forms"]))[0][1]

        request_ytd_Submit_records = [e[1] for e in source_data["Submit records"] if
                                      datetime.strptime(e[0], "%Y%m") <= datetime.strptime(self.request_date,
                                                                                           "%Y%m") and
                                      datetime.strptime(e[0], "%Y%m").year == datetime.strptime(
                                          self.request_date, "%Y%m").year]
        request_mtd_Submit_records = \
            list(filter(lambda e: e[0] == self.request_date, source_data["Submit records"]))[0][1]
        request_utd_Submit_records = [e[1] for e in source_data["Submit records"] if
                                      datetime.strptime(e[0], "%Y%m") <= datetime.strptime(self.request_date,
                                                                                           "%Y%m")]

        request_ytd_Approval_records = [e[1] for e in source_data["Approval records"] if
                                        datetime.strptime(e[0], "%Y%m") <= datetime.strptime(self.request_date,
                                                                                             "%Y%m") and
                                        datetime.strptime(e[0], "%Y%m").year == datetime.strptime(
                                            self.request_date, "%Y%m").year]
        request_mtd_Approval_records = \
            list(filter(lambda e: e[0] == self.request_date, source_data["Approval records"]))[0][1]
        request_utd_Approval_records = [e[1] for e in source_data["Approval records"] if
                                        datetime.strptime(e[0], "%Y%m") <= datetime.strptime(self.request_date, "%Y%m")]

        YTD_data = [self.number_with_comma(request_ytd_Forms),
                    self.number_with_comma(sum(request_ytd_Submit_records)),
                    self.number_with_comma(sum(request_ytd_Approval_records))]
        MTD_data = [self.number_with_comma(request_mtd_Forms),
                    self.number_with_comma(request_mtd_Submit_records),
                    self.number_with_comma(request_mtd_Approval_records)]
        UTD_data = [self.number_with_comma(request_utd_Forms),
                    self.number_with_comma(sum(request_utd_Submit_records)),
                    self.number_with_comma(sum(request_utd_Approval_records))]

        return {"YTD_data": YTD_data, "MTD_data": MTD_data, "UTD_data": UTD_data}


class IDoctor(Application):
    def __init__(self, request_date, cred_index=''):
        super().__init__(request_date)
        self.cred_index = cred_index
        self.app = "iDoctor"

    def html_table_data(self, source_data):
        if source_data is None:
            return None
        request_ytd_HCP_Subscribed_users = [e[1] for e in source_data["Subscribed users"] if
                                            datetime.strptime(e[0], "%Y%m") <= datetime.strptime(self.request_date,
                                                                                                 "%Y%m") and
                                            datetime.strptime(e[0], "%Y%m").year == datetime.strptime(
                                                self.request_date, "%Y%m").year]
        request_mtd_HCP_Subscribed_users = \
            list(filter(lambda e: e[0] == self.request_date, source_data["Subscribed users"]))[0][1]
        request_utd_HCP_Subscribed_users = \
            list(filter(lambda e: e[0] == self.request_date, source_data["Subscribed users(UTD)"]))[0][1]

        request_ytd_Touchpoint = [e[1] for e in source_data["Touchpoint"] if
                                  datetime.strptime(e[0], "%Y%m") <= datetime.strptime(self.request_date,
                                                                                       "%Y%m") and
                                  datetime.strptime(e[0], "%Y%m").year == datetime.strptime(
                                      self.request_date, "%Y%m").year]
        request_mtd_Touchpoint = \
            list(filter(lambda e: e[0] == self.request_date, source_data["Touchpoint"]))[0][1]
        request_utd_Touchpoint = [e[1] for e in source_data["Touchpoint"] if
                                  datetime.strptime(e[0], "%Y%m") <= datetime.strptime(self.request_date,
                                                                                       "%Y%m")]

        request_ytd_Active_users = \
            list(filter(lambda e: e[0] == self.request_date, source_data["Active users(YTD)"]))[0][1]
        request_mtd_Active_users = \
            list(filter(lambda e: e[0] == self.request_date, source_data["Active users"]))[0][1]
        request_utd_Active_users = \
            list(filter(lambda e: e[0] == self.request_date, source_data["Active users(UTD)"]))[0][1]

        YTD_data = [self.number_with_comma(sum(request_ytd_HCP_Subscribed_users)),
                    self.number_with_comma(sum(request_ytd_Touchpoint)),
                    self.number_with_comma(request_ytd_Active_users)]
        MTD_data = [self.number_with_comma(request_mtd_HCP_Subscribed_users),
                    self.number_with_comma(request_mtd_Touchpoint),
                    self.number_with_comma(request_mtd_Active_users)]
        UTD_data = [self.number_with_comma(request_utd_HCP_Subscribed_users),
                    self.number_with_comma(sum(request_utd_Touchpoint)),
                    self.number_with_comma(request_utd_Active_users)]
        return {"YTD_data": YTD_data, "MTD_data": MTD_data, "UTD_data": UTD_data}


class IDoctor_rep(Application):
    def __init__(self, request_date, cred_index=''):
        super().__init__(request_date)
        self.cred_index = cred_index
        self.app = "iDoctor_rep"

    def html_table_data(self, source_data):
        if source_data is None:
            return None
        request_ytd_HCP_Touchpoint = [e[1] for e in source_data["HCP Touchpoint"] if
                                      datetime.strptime(e[0], "%Y%m") <= datetime.strptime(self.request_date,
                                                                                           "%Y%m") and
                                      datetime.strptime(e[0], "%Y%m").year == datetime.strptime(
                                          self.request_date, "%Y%m").year]
        request_mtd_HCP_Touchpoint = \
            list(filter(lambda e: e[0] == self.request_date, source_data["HCP Touchpoint"]))[0][1]
        request_utd_HCP_Touchpoint = [e[1] for e in source_data["HCP Touchpoint"] if
                                      datetime.strptime(e[0], "%Y%m") <= datetime.strptime(self.request_date,
                                                                                           "%Y%m")]

        request_ytd_Rep_Touchpoint = [e[1] for e in source_data["Rep Touchpoint"] if
                                      datetime.strptime(e[0], "%Y%m") <= datetime.strptime(self.request_date,
                                                                                           "%Y%m") and
                                      datetime.strptime(e[0], "%Y%m").year == datetime.strptime(
                                          self.request_date, "%Y%m").year]
        request_mtd_Rep_Touchpoint = \
            list(filter(lambda e: e[0] == self.request_date, source_data["Rep Touchpoint"]))[0][1]
        request_utd_Rep_Touchpoint = [e[1] for e in source_data["Rep Touchpoint"] if
                                      datetime.strptime(e[0], "%Y%m") <= datetime.strptime(self.request_date,
                                                                                           "%Y%m")]

        request_ytd_Active_users = 'N/A'
        request_mtd_Active_users = 'N/A'
        request_utd_Active_users = 'N/A'

        YTD_data = [self.number_with_comma(sum(request_ytd_HCP_Touchpoint)),
                    self.number_with_comma(sum(request_ytd_Rep_Touchpoint)),
                    request_ytd_Active_users]
        MTD_data = [self.number_with_comma(request_mtd_HCP_Touchpoint),
                    self.number_with_comma(request_mtd_Rep_Touchpoint),
                    request_mtd_Active_users]
        UTD_data = [self.number_with_comma(sum(request_utd_HCP_Touchpoint)),
                    self.number_with_comma(sum(request_utd_Rep_Touchpoint)),
                    request_utd_Active_users]
        return {"YTD_data": YTD_data, "MTD_data": MTD_data, "UTD_data": UTD_data}


class INurse(Application):
    def __init__(self, request_date, cred_index=''):
        super().__init__(request_date)
        self.cred_index = cred_index
        self.app = "iNurse"

    def html_table_data(self, source_data):
        if source_data is None:
            return None
        request_ytd_Subscribed_users = \
            list(filter(lambda e: e[0] == self.request_date, source_data["Subscribed users(YTD)"]))[0][1]
        request_mtd_Subscribed_users = \
            list(filter(lambda e: e[0] == self.request_date, source_data["Subscribed users"]))[0][1]
        request_utd_Subscribed_users = \
            list(filter(lambda e: e[0] == self.request_date, source_data["Subscribed users(UTD)"]))[0][1]

        request_mtd_Active_users = \
            list(filter(lambda e: e[0] == self.request_date, source_data["Active users"]))[0][1]

        request_ytd_Touchpoint = [e[1] for e in source_data["Touchpoint"] if
                                  datetime.strptime(e[0], "%Y%m") <= datetime.strptime(self.request_date,
                                                                                       "%Y%m") and
                                  datetime.strptime(e[0], "%Y%m").year == datetime.strptime(
                                      self.request_date, "%Y%m").year]
        request_mtd_Touchpoint = \
            list(filter(lambda e: e[0] == self.request_date, source_data["Touchpoint"]))[0][1]
        request_utd_Touchpoint = [e[1] for e in source_data["Touchpoint"] if
                                  datetime.strptime(e[0], "%Y%m") <= datetime.strptime(self.request_date,
                                                                                       "%Y%m")]

        request_ytd_Enrollment_Patients = [e[1] for e in source_data["Enrollment Patients"] if
                                           datetime.strptime(e[0], "%Y%m") <= datetime.strptime(self.request_date,
                                                                                                "%Y%m") and
                                           datetime.strptime(e[0], "%Y%m").year == datetime.strptime(
                                               self.request_date, "%Y%m").year]
        request_mtd_Enrollment_Patients = \
            list(filter(lambda e: e[0] == self.request_date, source_data["Enrollment Patients"]))[0][1]
        request_utd_Enrollment_Patients = [e[1] for e in source_data["Enrollment Patients"] if
                                           datetime.strptime(e[0], "%Y%m") <= datetime.strptime(self.request_date,
                                                                                                "%Y%m")]

        YTD_data = [self.number_with_comma(request_ytd_Subscribed_users),
                    self.number_with_comma(sum(request_ytd_Touchpoint)),
                    self.number_with_comma(sum(request_ytd_Enrollment_Patients))]
        MTD_data = [self.number_with_comma(request_mtd_Subscribed_users),
                    self.number_with_comma(request_mtd_Active_users),
                    self.number_with_comma(request_mtd_Touchpoint),
                    self.number_with_comma(request_mtd_Enrollment_Patients)]
        UTD_data = [self.number_with_comma(request_utd_Subscribed_users),
                    self.number_with_comma(sum(request_utd_Touchpoint)),
                    self.number_with_comma(sum(request_utd_Enrollment_Patients))]
        return {"YTD_data": YTD_data, "MTD_data": MTD_data, "UTD_data": UTD_data}


class SmartSalesToolSSR(Application):
    def __init__(self, request_date, cred_index=''):
        super().__init__(request_date)
        self.cred_index = cred_index
        # self.svr_name = 'GH3SQLDBPRD1' # SQL 2012
        self.svr_name = 'GH3SQLDBPRD1\PRD2022'  # SQL 2016

    def load_source_data(self):
        sql_touchpoint = f'''
        SELECT concat(Year(log_When), MONTH(log_When)) as "YYYYMM", count(1) as 'Touchpoints'
        FROM [China_SSR].[dbo].[OperationLog]
        WHERE [log_Where] <> 'PC' and [Action] <> 'CreateDeviceToken'
        and convert(varchar(10),log_when,111) < = '{self.request_date[:4]}/{self.request_date[-2:]}/32'
        GROUP BY  Year(log_When), MONTH(log_When)
        order by YYYYMM desc
        '''

        sql_active_users_ytd = f'''
                SELECT count(distinct log_who) as 'Active Users YTD'
                FROM [China_SSR].[dbo].[OperationLog]
                WHERE [log_Where] <> 'PC' and [Action] <> 'CreateDeviceToken'
                and convert(varchar(10),log_when,111) > = '{self.request_date[:4]}/01/00'
                and convert(varchar(10),log_when,111) < = '{self.request_date[:4]}/{self.request_date[-2:]}/32'
                '''

        sql_active_users_mtd = f'''
                        SELECT count(distinct log_who) as 'Active Users MTD'
                        FROM [China_SSR].[dbo].[OperationLog]
                        WHERE [log_Where] <> 'PC' and [Action] <> 'CreateDeviceToken'
                        and convert(varchar(10),log_when,111) > = '{self.request_date[:4]}/{self.request_date[-2:]}/00'
                        and convert(varchar(10),log_when,111) < = '{self.request_date[:4]}/{self.request_date[-2:]}/32'
                        '''

        sql_active_users_utd = f'''
                        SELECT count(distinct log_who) as 'Active Users UTD'
                        FROM [China_SSR].[dbo].[OperationLog]
                        WHERE [log_Where] <> 'PC' and [Action] <> 'CreateDeviceToken'
                        and convert(varchar(10),log_when,111) > = '2000/01/00'
                        and convert(varchar(10),log_when,111) < = '{self.request_date[:4]}/{self.request_date[-2:]}/32'
                        '''

        try:
            db_name = 'China_SSR'
            conn = pyodbc.connect('Driver={SQL Server};'
                                  f'Server={self.svr_name};'
                                  f'Database={db_name};'
                                  f'Trusted_Connection=yes;'
                                  f'timeout={self.db_conn_timeout};')
            conn.timeout = int(self.db_conn_timeout)
            cursor = conn.cursor()
            cursor.execute(sql_touchpoint)
            access_number = [d for d in cursor]
            access_number.sort(key=lambda date: datetime.strptime(date[0], "%Y%m"))
            source_data = {"Access Number": access_number}

            cursor.execute(sql_active_users_mtd)
            active_users_mtd = [d for d in cursor]
            source_data["Active Users"] = {"MTD": active_users_mtd[0][0]}

            cursor.execute(sql_active_users_ytd)
            active_users_ytd = [d for d in cursor]
            source_data["Active Users"]["YTD"] = active_users_ytd[0][0]

            cursor.execute(sql_active_users_utd)
            active_users_utd = [d for d in cursor]
            source_data["Active Users"]["UTD"] = active_users_utd[0][0]

            max_date_of_access_number = source_data["Access Number"][-1][0]
            min_date_of_access_number = source_data["Access Number"][0][0]
            if datetime.strptime(max_date_of_access_number, "%Y%m") < datetime.strptime(self.request_date, "%Y%m"):
                raise Exception(
                    f"[ERROR]Request date must be (<=) equal to or earlier than: {max_date_of_access_number}")
            elif datetime.strptime(min_date_of_access_number, "%Y%m") > datetime.strptime(self.request_date, "%Y%m"):
                raise Exception(
                    f"[ERROR]Request date must be (>=) equal to or later than: {min_date_of_access_number}")

            return source_data
        except Exception as e:
            print(str(e))
            return None

    def html_table_data(self, source_data):
        if source_data is None:
            return None
        request_utd_access_number = [e[1] for e in source_data["Access Number"] if
                                     datetime.strptime(e[0], "%Y%m") <= datetime.strptime(self.request_date, "%Y%m")]

        request_ytd_access_number = [e[1] for e in source_data["Access Number"] if
                                     datetime.strptime(e[0], "%Y%m") <= datetime.strptime(self.request_date,
                                                                                          "%Y%m") and
                                     datetime.strptime(e[0], "%Y%m").year == datetime.strptime(self.request_date,
                                                                                               "%Y%m").year]
        request_ytd_active_users = source_data["Active Users"]["YTD"]
        request_mtd_active_users = source_data["Active Users"]["MTD"]
        request_utd_active_users = source_data["Active Users"]["UTD"]

        YTD_data = [self.number_with_comma(sum(request_ytd_access_number)),
                    self.number_with_comma(request_ytd_active_users)]
        MTD_data = [self.number_with_comma(request_ytd_access_number[-1]),
                    self.number_with_comma(request_mtd_active_users)]
        UTD_data = [self.number_with_comma(sum(request_utd_access_number)),
                    self.number_with_comma(request_utd_active_users)]

        return {"YTD_data": YTD_data, "MTD_data": MTD_data, "UTD_data": UTD_data}


class SmartSalesToolSDF(Application):
    def __init__(self, request_date, cred_index=''):
        super().__init__(request_date)
        self.cred_index = cred_index
        # self.svr_name = 'GH3SQLDBPRD1' # SQL 2012
        self.svr_name = 'GH3SQLDBPRD1\PRD2022'  # SQL 2016

    def load_source_data(self):
        sql_touchpoint = f'''
        	SELECT concat(Year(log_When), MONTH(log_When)) as "YYYYMM", Count(1) as 'Touchpoints'
            FROM China_SDF.dbo.Operation_Log 
            where 1=1 
            --and convert(varchar(10),log_when,111) > = '{self.request_date[:4]}/{self.request_date[-2:]}/00' 
            and convert(varchar(10),log_when,111) < = '{self.request_date[:4]}/{self.request_date[-2:]}/32'
            GROUP BY  Year(log_When), MONTH(log_When)
            order by YYYYMM desc
        '''

        sql_active_users_ytd = f'''
                	SELECT count(distinct log_who) as 'Active Users YTD'
                    FROM China_SDF.dbo.Operation_Log 
                    where 1=1 
                    and convert(varchar(10),log_when,111) > = '{self.request_date[:4]}/01/00' 
                    and convert(varchar(10),log_when,111) < = '{self.request_date[:4]}/{self.request_date[-2:]}/32'
                '''

        sql_active_users_mtd = f'''
                    SELECT count(distinct log_who) as 'Active Users MTD'
                    FROM China_SDF.dbo.Operation_Log 
                    where 1=1 
                    and convert(varchar(10),log_when,111) > = '{self.request_date[:4]}/{self.request_date[-2:]}/00' 
                    and convert(varchar(10),log_when,111) < = '{self.request_date[:4]}/{self.request_date[-2:]}/32'
                    '''

        sql_active_users_utd = f'''
                    SELECT count(distinct log_who) as 'Active Users MTD'
                    FROM China_SDF.dbo.Operation_Log 
                    where 1=1 
                    and convert(varchar(10),log_when,111) > = '2000/01/00' 
                    and convert(varchar(10),log_when,111) < = '{self.request_date[:4]}/{self.request_date[-2:]}/32'
                    '''

        try:
            db_name = 'China_SSR'
            conn = pyodbc.connect('Driver={SQL Server};'
                                  f'Server={self.svr_name};'
                                  f'Database={db_name};'
                                  f'Trusted_Connection=yes;'
                                  f'timeout={self.db_conn_timeout};')
            conn.timeout = int(self.db_conn_timeout)
            cursor = conn.cursor()
            cursor.execute(sql_touchpoint)
            access_number = [d for d in cursor]
            access_number.sort(key=lambda date: datetime.strptime(date[0], "%Y%m"))
            source_data = {"Access Number": access_number}

            cursor.execute(sql_active_users_mtd)
            active_users_mtd = [d for d in cursor]
            source_data["Active Users"] = {"MTD": active_users_mtd[0][0]}

            cursor.execute(sql_active_users_ytd)
            active_users_ytd = [d for d in cursor]
            source_data["Active Users"]["YTD"] = active_users_ytd[0][0]

            cursor.execute(sql_active_users_utd)
            active_users_utd = [d for d in cursor]
            source_data["Active Users"]["UTD"] = active_users_utd[0][0]

            max_date_of_access_number = source_data["Access Number"][-1][0]
            min_date_of_access_number = source_data["Access Number"][0][0]
            if datetime.strptime(max_date_of_access_number, "%Y%m") < datetime.strptime(self.request_date, "%Y%m"):
                raise Exception(
                    f"[ERROR]Request date must be (<=) equal to or earlier than: {max_date_of_access_number}")
            elif datetime.strptime(min_date_of_access_number, "%Y%m") > datetime.strptime(self.request_date, "%Y%m"):
                raise Exception(
                    f"[ERROR]Request date must be (>=) equal to or later than: {min_date_of_access_number}")

            return source_data
        except Exception as e:
            print(str(e))
            return None

    def html_table_data(self, source_data):
        if source_data is None:
            return None
        request_utd_access_number = [e[1] for e in source_data["Access Number"] if
                                     datetime.strptime(e[0], "%Y%m") <= datetime.strptime(self.request_date, "%Y%m")]

        request_ytd_access_number = [e[1] for e in source_data["Access Number"] if
                                     datetime.strptime(e[0], "%Y%m") <= datetime.strptime(self.request_date,
                                                                                          "%Y%m") and
                                     datetime.strptime(e[0], "%Y%m").year == datetime.strptime(self.request_date,
                                                                                               "%Y%m").year]
        request_ytd_active_users = source_data["Active Users"]["YTD"]
        request_mtd_active_users = source_data["Active Users"]["MTD"]
        request_utd_active_users = source_data["Active Users"]["UTD"]

        YTD_data = [self.number_with_comma(sum(request_ytd_access_number)),
                    self.number_with_comma(request_ytd_active_users)]
        MTD_data = [self.number_with_comma(request_ytd_access_number[-1]),
                    self.number_with_comma(request_mtd_active_users)]
        UTD_data = [self.number_with_comma(sum(request_utd_access_number)),
                    self.number_with_comma(request_utd_active_users)]

        return {"YTD_data": YTD_data, "MTD_data": MTD_data, "UTD_data": UTD_data}


class TrainingGamification(Application):
    def __init__(self, request_date, cred_index='cred2'):
        super().__init__(request_date)
        self.cred_index = cred_index
        self.svr_name = 'GH3DMZDBPHA.XH3.LILLY.COM,1433'
        self.db_name = 'GamificationWeChat'
        self.UID = Application.load_user_creds('creds.json', cred_index=cred_index)['UID']
        self.pwd = Application.load_user_creds('creds.json', cred_index=cred_index)['pwd']

    def load_source_data(self):
        sql_sign_in_number = '''
            With TestUser as (
            --欣百达             
            SELECT *, Brand = N'欣百达' 
              FROM [GamificationWeChat].[dbo].[g_user]
              where brand_id = 1  and team_id in (1001,1004,1005)
            
              Union All
            
            --希爱力             
              SELECT *, Brand = N'希爱力'      
              FROM [GamificationWeChat].[dbo].[g_user]
              where brand_id = 2  and team_id in (1002,1062,1075)
            
              Union All
            
            --艾乐明      
              SELECT *, Brand = N'艾乐明'      
              FROM [GamificationWeChat].[dbo].[g_user]
              where brand_id = 3  and team_id in (1074,1081)
            
              Union All
            
            --拓咨
              SELECT *, Brand = N'拓咨' 
              FROM [GamificationWeChat].[dbo].[g_user]
              where brand_id = 4  and team_id in (1076,1323)
            
              Union All
            
            --V来计划  PRD环境
              SELECT *, Brand = N'V来计划' 
              FROM [GamificationWeChat].[dbo].[g_user]
              where brand_id = 7  and team_id in (1327)
            
             /****** QA, 排除
            --V来计划  QA环境
              SELECT *, Brand = N'V来计划' 
              FROM [GamificationWeChat].[dbo].[g_user]
              where brand_id = 7  and team_id not in (1080)
              ******/
            
            )
              
             /****** Brand1的每月签到记录数 – 汇总  ******/
              Select t.[Month],t.brandid, 
              Count(t.id) As Cnt
              from 
              ( --签到详细数据
              SELECT s.*, u.th_user_id, u.real_name, u.team_id,t.team_name, CONVERT(varchar(6), s.sign_date, 112) as Month,u.brand_id as brandid
              FROM [GamificationWeChat].[dbo].[g_user_sign_log] s
              left join [GamificationWeChat].[dbo].[g_user] u on s.user_id = u.id
              left join [GamificationWeChat].[dbo].[g_user_team] t on u.team_id = t.id
              where 1 =1
              and u.id not in (
                   Select id from TestUser
            )
            
              and u.brand_id <>99
              ) T
              group by t.[Month],t.brandid
              order by 1,2
        '''
        sql_access_number = '''
            With TestUser as (
            --欣百达             
            SELECT *, Brand = N'欣百达' 
              FROM [GamificationWeChat].[dbo].[g_user]
              where brand_id = 1  and team_id in (1001,1004,1005)
            
              Union All
            
            --希爱力             
              SELECT *, Brand = N'希爱力'      
              FROM [GamificationWeChat].[dbo].[g_user]
              where brand_id = 2  and team_id in (1002,1062,1075)
            
              Union All
            
            --艾乐明      
              SELECT *, Brand = N'艾乐明'      
              FROM [GamificationWeChat].[dbo].[g_user]
              where brand_id = 3  and team_id in (1074,1081)
            
              Union All
            
            --拓咨
              SELECT *, Brand = N'拓咨' 
              FROM [GamificationWeChat].[dbo].[g_user]
              where brand_id = 4  and team_id in (1076,1323)
            
              Union All
            
            --V来计划  PRD环境
              SELECT *, Brand = N'V来计划' 
              FROM [GamificationWeChat].[dbo].[g_user]
              where brand_id = 7  and team_id in (1327)
            
             /****** QA, 排除
            --V来计划  QA环境
              SELECT *, Brand = N'V来计划' 
              FROM [GamificationWeChat].[dbo].[g_user]
              where brand_id = 7  and team_id not in (1080)
              ******/
            )
            
            ,PK as (
            SELECT pk.id, pk_way pk_way,  
                          pk.status  win_str, 
                         pk.send_user_id,pk.accept_user_id,u.real_name send_user_name, us.real_name accept_user_name,CONVERT(varchar(30),pk.create_time,120) create_time,pk.win_user_id, u.brand_id
                from g_user_pk_personal pk
                         LEFT JOIN g_user u on pk.send_user_id = u.id 
                         LEFT JOIN g_user us on us.id =pk.accept_user_id  
                        where 1=1 --全部结果
            
             /******  排除非销售员工  ******/
               and pk.send_user_id not in (
                   Select id from TestUser
            )
            
            )
            
            --Select * from PK
            
            , Training AS (     
            /*训练统计2*/
            SELECT  t.team_name teamName, u.real_name realName, u.th_user_id thUserId,
                                    l.level_show_name levelShowName, q.question_name questionName,
                                    q.question_field questionField,q.answer_correct answerCorrect,
                                    r.reply reply,r.is_correct isCorrect,CONVERT(varchar(30),r.create_time,120) create_Time,
                                    r.id  
                  ,t.team_name --团队名
                  , u.brand_id
                        FROM g_user_train_result r 
                        LEFT JOIN g_brand_question q on r.qid = q.id 
                        LEFT JOIN g_brand_level l on l.level_id = q.brand_level_id 
                        LEFT JOIN g_user u on r.user_id = u.id 
                        LEFT JOIN g_user_team t on u.team_id = t.id 
                        where 1=1 --全部结果
            
             /******  排除非销售员工  ******/
               and r.user_id not in ( 
                   Select id from TestUser
            )
            
            )
            
            --Select COUNT(*) from Training
            
            , Temp AS (
            Select N'2挑战答题数(1.发起方PK)' AS Title,  COUNT(1) *10 AS Cnt,create_time, brand_id
            from PK
            where win_str>0                  --所有有效答题，记录发起人答题数，包含和机器人答题
            Group by create_time, brand_id
            union all
            Select N'2挑战答题数(2.接受方PK-过滤机器人)' AS Title,  COUNT(1) *10 AS Cnt,create_time, brand_id
            from PK
            where win_str>0 and accept_user_id not in (0,10,11,12,13,14,15,16,17,18)    --所有有效答题，过滤机器人，记录接受者答题数
            Group by create_time, brand_id
            union all
            
            Select N'1训练答题数' AS Title,  1 AS Cnt,create_time , brand_id
            from Training
            --Group by create_time, brand_id
            union all
            SELECT N'3.6用户访问装备库详情' AS Title,1 AS Cnt,ul.create_time, U.brand_id
            FROM [GamificationWeChat].[dbo].[g_user_log] ul
            inner JOIN [g_user] U ON U.id = ul.user_id
            where action_type = 6
            
             /******  排除非销售员工  ******/
            and ul.user_id not in ( Select id from TestUser )
            
            --group by create_time, U.brand_id
            union all
            SELECT N'3.1用户登录游戏' AS Title,1 AS Cnt,ul.create_time, U.brand_id
            FROM [GamificationWeChat].[dbo].[g_user_log] ul
            inner JOIN [g_user] U ON U.id = ul.user_id
            where action_type = 1
            
             /******  排除非销售员工  ******/
            and ul.user_id not in ( Select id from TestUser )
            
            --group by create_time, U.brand_id
            union all
            SELECT N'3.2用户访问训练场' AS Title,1 AS Cnt,ul.create_time, U.brand_id
            FROM [GamificationWeChat].[dbo].[g_user_log] ul
            inner JOIN [g_user] U ON U.id = ul.user_id
            where action_type = 2
            
             /******  排除非销售员工  ******/
            and ul.user_id not in ( Select id from TestUser )
            
            --group by create_time, T.brand_id
            union all
            SELECT N'3.3用户访问训练场收藏夹' AS Title,1 AS Cnt,ul.create_time, U.brand_id
            FROM [GamificationWeChat].[dbo].[g_user_log] ul
            inner JOIN [g_user] U ON U.id = ul.user_id
            where action_type = 3
            
             /******  排除非销售员工  ******/
            and ul.user_id not in ( Select id from TestUser )
            
            --group by create_time, T.brand_id
            union all
            SELECT N'3.4用户访问排行榜' AS Title,1 AS Cnt,ul.create_time, U.brand_id
            FROM [GamificationWeChat].[dbo].[g_user_log] ul
            inner JOIN [g_user] U ON U.id = ul.user_id
            where action_type = 4
            
             /******  排除非销售员工  ******/
            and ul.user_id not in ( Select id from TestUser )
            
            --group by create_time, T.brand_id
            union all
            SELECT N'3.5用户访问装备库列表' AS Title,1 AS Cnt,ul.create_time, U.brand_id
            FROM [GamificationWeChat].[dbo].[g_user_log] ul
            inner JOIN [g_user] U ON U.id = ul.user_id
            where action_type = 5
            
             /******  排除非销售员工  ******/
            and ul.user_id not in ( Select id from TestUser )
            
            --group by create_time, U.brand_id
            union all
            SELECT N'3.5用户访问装备库列表' AS Title,1 AS Cnt,ul.create_time, U.brand_id
            FROM [GamificationWeChat].[dbo].[g_user_log] ul
            inner JOIN [g_user] U ON U.id = ul.user_id
            where action_type = 5
            
             /******  排除非销售员工  ******/
            and ul.user_id not in ( Select id from TestUser )
            
            --group by create_time, U.brand_id
            union all
            SELECT N'3.8用户进入个人竞技场' AS Title,1 AS Cnt,ul.create_time, U.brand_id
            FROM [GamificationWeChat].[dbo].[g_user_log] ul
            inner JOIN [g_user] U ON U.id = ul.user_id
            where action_type = 8
            
             /******  排除非销售员工  ******/
            and ul.user_id not in ( Select id from TestUser )
            
            --group by create_time, U.brand_id
            union all
            SELECT N'3.9用户进入团队竞技场' AS Title,1 AS Cnt,ul.create_time, U.brand_id
            FROM [GamificationWeChat].[dbo].[g_user_log] ul
            inner JOIN [g_user] U ON U.id = ul.user_id
            where action_type = 9
            
             /******  排除非销售员工  ******/
            and ul.user_id not in ( Select id from TestUser )
            )
            
            Select CONVERT(varchar(6),create_time,112) as 'YYYYMM',Title, Sum(Cnt) AS Cnt,brand_id
            from Temp 
            where 1=1 --create_time between '2019-05-01' And '2019-06-01'
            --create_time>='2019-05-01' 
            --and 
            --and create_time<'2023-06-01'
            --create_time >'2019-05-01' --开始时间 :startDate
            --and create_time <'2019-06-01' --结束时间 :endDate
            --and CONVERT(varchar(6),create_time,112) = 202003
            group by Title,CONVERT(varchar(7),create_time,120),CONVERT(varchar(6),create_time,112),brand_id
            order by 1,2,4
        '''
        sql_user_number = '''
            Select brand_id, Count(*) as Cnt
            FROM [GamificationWeChat].[dbo].[g_user]
            where ((brand_id=1 and team_id not in (1001,1004,1005))
            or (brand_id=2 and team_id not in (1002,1062,1075))
            or (brand_id=3 and team_id not in (1074,1082,1081))
            or (brand_id=4 and team_id not in (1076,1323))
            or (brand_id=7 and team_id not in (1327)))
            group by brand_id
            order by 1
        '''

        try:
            conn = pyodbc.connect('Driver={SQL Server};'
                                  f'Server={self.svr_name};'
                                  f'Database={self.db_name};'
                                  f'UID={self.UID};'
                                  f'PWD={self.pwd};'
                                  f'Trusted_Connection=no;'
                                  f'timeout={self.db_conn_timeout};')
            conn.timeout = int(self.db_conn_timeout)
            cursor = conn.cursor()
            cursor.execute(sql_user_number)
            user_number = [d for d in cursor]
            user_number_total = sum([e[1] for e in user_number])

            source_data = {"User number": user_number_total}

            cursor = conn.cursor()
            cursor.execute(sql_sign_in_number)
            sign_in_number = [d for d in cursor]
            sign_in_number.sort(key=lambda date: datetime.strptime(date[0], "%Y%m"))
            source_data["Sign in number"] = sign_in_number

            cursor = conn.cursor()
            cursor.execute(sql_access_number)
            access_number = [d for d in cursor]
            # keep brand_id equals  to 3 or 4 or 7
            access_number = [e for e in access_number if e[3] == 3 or e[3] == 4 or e[3] == 7]
            # keep title equals 1, 2(1), 2(2), 3.4 and 3.6
            access_number = [e for e in access_number if
                             e[1][0] == '1' or e[1][0] == '2' or e[1][0:3] == '3.4' or e[1][0:3] == '3.6']
            access_number.sort(key=lambda date: datetime.strptime(date[0], "%Y%m"))
            source_data["Access number"] = access_number
            pprint(source_data["Access number"])

            max_date_of_source_data = source_data["Access number"][-1][0]
            min_date_of_source_data = source_data["Access number"][0][0]
            if datetime.strptime(max_date_of_source_data, "%Y%m") < datetime.strptime(self.request_date, "%Y%m"):
                raise Exception(
                    f"[ERROR]Request date must be (<=) equal to or earlier than: {max_date_of_source_data}")
            elif datetime.strptime(min_date_of_source_data, "%Y%m") > datetime.strptime(self.request_date, "%Y%m"):
                raise Exception(
                    f"[ERROR]Request date must be (>=) equal to or later than: {min_date_of_source_data}")

            return source_data


        except Exception as e:
            print(e)
            return None

    def sum_groupby_month(self, data_list, index, value):
        result = defaultdict(int)
        for e in data_list:
            result[e[index]] += e[value]
        result = list(result.items())
        print(result)
        return result

    def get_key_data(self, source_data, key_name):
        if source_data is None:
            return None
        # key_name = "Access number"
        return self.sum_groupby_month(source_data[key_name], 0, 2)

    def html_table_data(self, source_data):
        if source_data is None:
            return None

        request_ytd_source_data_sign_in_number = [e[1] for e in
                                                  self.sum_groupby_month(source_data["Sign in number"], 0, 2) if
                                                  datetime.strptime(e[0], "%Y%m") <= datetime.strptime(
                                                      self.request_date,
                                                      "%Y%m") and
                                                  datetime.strptime(e[0], "%Y%m").year == datetime.strptime(
                                                      self.request_date,
                                                      "%Y%m").year]

        # request_utd_source_data = [e[1] for e in source_data["User number"] if
        #                            datetime.strptime(e[0], "%Y%m") <= datetime.strptime(self.request_date, "%Y%m")]

        request_ytd_source_data_access_number = [e[1] for e in self.get_key_data(source_data, "Access number") if
                                                 datetime.strptime(e[0], "%Y%m") <= datetime.strptime(self.request_date,
                                                                                                      "%Y%m") and
                                                 datetime.strptime(e[0], "%Y%m").year == datetime.strptime(
                                                     self.request_date,
                                                     "%Y%m").year]
        user_number = self.number_with_comma(source_data["User number"])

        print(f"request_ytd_source_data_user_number: {sum(request_ytd_source_data_sign_in_number)}")
        print(f"request_ytd_source_data_access_number: {sum(request_ytd_source_data_access_number)}")
        print(f"sign_in_number: {user_number}")
        YTD_data = [user_number, self.number_with_comma(sum(request_ytd_source_data_access_number)),
                    self.number_with_comma(sum(request_ytd_source_data_sign_in_number))
                    ]
        MTD_data = [user_number, self.number_with_comma(request_ytd_source_data_access_number[-1]),
                    self.number_with_comma(request_ytd_source_data_sign_in_number[-1]), ]
        # UTD_data = [total_vendor_number, self.number_with_comma(sum(request_utd_source_data)),source_data["Sign in number"]]

        return {"YTD_data": YTD_data, "MTD_data": MTD_data, "UTD_data": None}


class LillyChina(Application):
    def __init__(self, request_date, cred_index=''):
        super().__init__(request_date)
        self.cred_index = cred_index
        self.app = "LillyChina"

    def load_source_data(self):
        import googleAnalyticsAPI as ga
        try:
            if self.request_date[4:] == '01':
                source_data_last = ga.run_mtd(request_date=Application.last_month_of(self.request_date))
                source_data_current = ga.run_mtd(request_date=self.request_date)
                for k, v in source_data_last.items():
                    # source_data_last = {k: v[-1]}
                    source_data_current[k].append(v[-1])

                # pprint(source_data_current)
                source_data = source_data_current
            else:
                source_data = ga.run_mtd(request_date=self.request_date)
            return source_data
        except Exception as e:
            print(f"{e}")
            return None

    def html_table_data(self, source_data):
        if source_data is None:
            return None
        import googleAnalyticsAPI as ga
        source_data_ytd = ga.run_ytd(request_date=self.request_date)

        request_ytd_session = source_data_ytd["Session"][0][1]
        request_ytd_touchpoint = source_data_ytd["Touchpoint"][0][1]
        request_ytd_active_users = source_data_ytd["Active users"][0][1]

        YTD_data = [self.number_with_comma(request_ytd_session),
                    self.number_with_comma(request_ytd_touchpoint),
                    self.number_with_comma(request_ytd_active_users)]

        request_mtd_session = [e[1] for e in source_data["Session"]]
        request_mtd_touchpoint = [e[1] for e in source_data["Touchpoint"]]
        request_mtd_active_users = [e[1] for e in source_data["Active users"]]

        MTD_data = [self.number_with_comma(request_mtd_session[-1]),
                    self.number_with_comma(request_mtd_touchpoint[-1]),
                    self.number_with_comma(request_mtd_active_users[-1])]

        return {"YTD_data": YTD_data, "MTD_data": MTD_data, "UTD_data": []}


class LillyChina_xlsx(LillyChina, Application):
    def __init__(self, request_date, cred_index=''):
        super().__init__(request_date)
        self.cred_index = cred_index
        self.app = "LillyChina"

    def html_table_data(self, source_data):
        if source_data is None:
            return None
        pass


class BAW(Application):
    def __init__(self, request_date, cred_index=''):
        super().__init__(request_date)
        self.cred_index = cred_index
        self.app = "BAW"

    def html_table_data(self, source_data):
        if source_data is None:
            return None
        request_ytd_Finished_forms = \
            list(filter(lambda e: e[0] == self.request_date, source_data["Finished forms(YTD)"]))[0][1]
        request_mtd_Finished_forms = \
            list(filter(lambda e: e[0] == self.request_date, source_data["Finished forms"]))[0][1]

        request_ytd_On_going_forms = \
            list(filter(lambda e: e[0] == self.request_date, source_data["On-going forms(YTD)"]))[0][1]
        request_mtd_On_going_forms = \
            list(filter(lambda e: e[0] == self.request_date, source_data["On-going forms"]))[0][1]

        request_mtd_forms = \
            list(filter(lambda e: e[0] == self.request_date, source_data["Forms"]))[0][1]
        request_ytd_forms = request_mtd_forms

        YTD_data = [self.number_with_comma(request_ytd_forms),
                    self.number_with_comma(request_ytd_Finished_forms),
                    self.number_with_comma(request_ytd_On_going_forms)]
        MTD_data = [self.number_with_comma(request_mtd_forms),
                    self.number_with_comma(request_mtd_Finished_forms),
                    self.number_with_comma(request_mtd_On_going_forms)]

        return {"YTD_data": YTD_data, "MTD_data": MTD_data, "UTD_data": []}
