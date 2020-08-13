import openpyxl
from datetime import datetime
import pyodbc


# Connect to the Microsoft SQL Database
con = pyodbc.connect(Trusted_Connection='no',
                     driver='{SQL Server}',
                     server='192.168.15.32',
                     database='Alpha_Live',
                     UID='pladis_dba',
                     PWD='BigFlats')
cursor = con.cursor()


def tuple_instance_write_destination(number_of_rows, factory_talk_sheet, table_name):
    report_tuple = ()  # Equipment OPC readouts.
    n = 3  # First row is (indicator, value), Second & Third rows are datetime.
    while n <= number_of_rows:
        instance = factory_talk_sheet['B' + str(n)].value  # Set instance as unique indicator value.
        if n == 3:
            instance = datetime.now().strftime('%Y-%m-%d %H:%M')  # Datetime is the instance
        report_tuple = report_tuple + (instance,)
        n += 1  # Increase row
    cursor.execute('USE [Alpha_Live] INSERT INTO [dbo].' + table_name + ' VALUES ' + str(report_tuple))
    con.commit()


# Open workbook and worksheets
data_src_workbook = 'L:/Live Alpha Readings/192.168.15.27_Data Report.xlsm'
factory_talk_wb = openpyxl.load_workbook(data_src_workbook, keep_vba=False, data_only=True, keep_links=False,
                                         read_only=True)
factory_talk_tank_farm = factory_talk_wb['Tank_Farm']
factory_talk_KGM_080 = factory_talk_wb['KGM_080']
factory_talk_TTB_015 = factory_talk_wb['TTB_015']
factory_talk_M5_090 = factory_talk_wb['M5_090']
factory_talk_TT_100 = factory_talk_wb['TT_100']
factory_talk_M5_140 = factory_talk_wb['M5_140']
factory_talk_TT_150 = factory_talk_wb['TT_150']
factory_talk_DBS_080 = factory_talk_wb['DBS_080']
factory_talk_DFR_031 = factory_talk_wb['DFR_031']
factory_talk_HCM_273 = factory_talk_wb['HCM_273']
factory_talk_HCM_274 = factory_talk_wb['HCM_274']
factory_talk_TTM_147 = factory_talk_wb['TTM_147']
factory_talk_Chocotech_Kitchen = factory_talk_wb['Chocotech_Kitchen']

tuple_instance_write_destination(27, factory_talk_tank_farm, '[Tank Farm]')
tuple_instance_write_destination(56, factory_talk_KGM_080, '[KGM 080]')
tuple_instance_write_destination(29, factory_talk_TTB_015, '[TTB 015]')
tuple_instance_write_destination(26, factory_talk_M5_090, '[M5 090]')
tuple_instance_write_destination(26, factory_talk_TT_100, '[TT 100]')
tuple_instance_write_destination(42, factory_talk_M5_140, '[M5 140]')
tuple_instance_write_destination(26, factory_talk_TT_150, '[TT 150]')
tuple_instance_write_destination(16, factory_talk_DBS_080, '[DBS 080]')
tuple_instance_write_destination(19, factory_talk_DFR_031, '[DFR 031]')
tuple_instance_write_destination(31, factory_talk_HCM_273, '[HCM 273]')
tuple_instance_write_destination(31, factory_talk_HCM_274, '[HCM 274]')
tuple_instance_write_destination(14, factory_talk_TTM_147, '[TTM 147]')
tuple_instance_write_destination(45, factory_talk_Chocotech_Kitchen, '[Chocotech Kitchen]')
