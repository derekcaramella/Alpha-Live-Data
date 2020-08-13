import openpyxl
from datetime import datetime
import pyodbc
import numpy as np
import pandas as pd

# Connect to the Microsoft SQL Database
con = pyodbc.connect(Trusted_Connection='no',
                     driver='{SQL Server}',
                     server='192.168.15.32',
                     database='Alpha_Live',
                     UID='pladis_dba',
                     PWD='BigFlats')
cursor = con.cursor()


def tuple_instance_write_destination(number_of_rows, factory_talk_sheet, table_name, instance_tuple_start,
                                     instance_tuple_end):
    standard_tuple = ()  # Datetime, Shift, Workstation, {Bars/min Out, Pieces Per Pack, Trays/min Discharge}
    instance_tuple = ()  # Individual unique indicators
    report_tuple = ()  # Reporting indicators: Availability, OEE, Buffer Fill Level
    observation_tuple = ()  # Tuple that is inserted in the SQL statement

    if table_name in ['[DBS 080]', '[DFR 031]', '[HCM 273]', '[HCM 274]', '[TTM 147]'] and datetime.now().strftime('%H:%M') \
            not in ['06:00', '18:00']:  # If it is a packaging location and not tge start of shift.
        sql_top = pd.read_sql_query('SELECT TOP (1) * FROM [dbo].' + table_name + ' ORDER BY [Timestamp] DESC;',
                                    con).to_numpy()  # Obtain the last entry.

        n = 3  # First row is (indicator, value), Second & Third rows are datetime.
        while n <= number_of_rows:
            instance = factory_talk_sheet['B' + str(n)].value  # Set instance as unique indicator value.
            if n < instance_tuple_start:  # These are the Standard tuples: datetime, shift, workstation...
                if n == 3:
                    instance = datetime.now().strftime('%Y-%m-%d %H:%M')  # Datetime is the instance
                    standard_tuple = standard_tuple + (
                        instance,)  # Standard tuple equals standing standard tuple and datetime.
                else:
                    standard_tuple = standard_tuple + (
                        instance,)  # Standard tuple equals standing standard tuple and unique value.
            elif instance_tuple_start <= n <= instance_tuple_end:  # These are the unique tuples: Row count, rejections...
                instance_tuple = instance_tuple + (
                    instance,)  # Instance tuple equals standing instance tuple and unique value.
            else:
                report_tuple = report_tuple + (instance,)  # Report tuple equals standing report tuple and unique value.
            n += 1  # Increase row

        if table_name == '[DBS 080]':  # Provided that the workstation is the DBS
            flow = np.subtract(instance_tuple, sql_top[0][
                                               3:instance_tuple_end])  # Flow is the difference between the Instance tuple and the applicable SQL request Instance tuple
        elif table_name == '[DFR 031]':  # Provided that the workstation is the DFR
            flow = np.subtract(instance_tuple, sql_top[0][
                                               3:instance_tuple_end - 2])  # Flow is the difference between the Instance tuple and the applicable SQL request Instance tuple
        elif table_name in ['[HCM 273]', '[HCM 274]']:  # Provided that the workstation a SIG
            flow = np.subtract(instance_tuple, sql_top[0][
                                               5:instance_tuple_end - 2])  # Flow is the difference between the Instance tuple and the applicable SQL request Instance tuple
            print(instance_tuple)
            print(sql_top[0][5:instance_tuple_end - 2])
            print(flow)
        else:
            flow = np.subtract(instance_tuple, sql_top[0][
                                               6:instance_tuple_end - 2])  # Flow (TTM) is the difference between the Instance tuple and the applicable SQL request Instance tuple

        observation_tuple = standard_tuple + tuple(
            flow) + report_tuple  # Observation tuple is the aggregation of standard, flow, and report tuples

    # All "wet" locations
    else:
        n = 3  # First row is (indicator, value), Second & Third rows are datetime.
        while n <= number_of_rows:
            instance = factory_talk_sheet['B' + str(n)].value  # Set instance as unique indicator value.
            if n == 3:
                instance = datetime.now().strftime('%Y-%m-%d %H:%M')  # Datetime is the instance
            observation_tuple = observation_tuple + (
                instance,)  # Observation tuple equals standing observation tuple and unique value.
            n += 1  # Increase row

    cursor.execute('USE [Alpha_Live] INSERT INTO [dbo].' + table_name + ' VALUES' + str(observation_tuple))
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

tuple_instance_write_destination(27, factory_talk_tank_farm, '[Tank Farm]', 0, 0)
tuple_instance_write_destination(56, factory_talk_KGM_080, '[KGM 080]', 0, 0)
tuple_instance_write_destination(29, factory_talk_TTB_015, '[TTB 015]', 0, 0)
tuple_instance_write_destination(26, factory_talk_M5_090, '[M5 090]', 0, 0)
tuple_instance_write_destination(26, factory_talk_TT_100, '[TT 100]', 0, 0)
tuple_instance_write_destination(42, factory_talk_M5_140, '[M5 140]', 0, 0)
tuple_instance_write_destination(26, factory_talk_TT_150, '[TT 150]', 0, 0)
tuple_instance_write_destination(16, factory_talk_DBS_080, '[DBS 080]', 6, 16)
tuple_instance_write_destination(19, factory_talk_DFR_031, '[DFR 031]', 6, 17)
tuple_instance_write_destination(31, factory_talk_HCM_273, '[HCM 273]', 8, 29)
tuple_instance_write_destination(31, factory_talk_HCM_274, '[HCM 274]', 8, 29)
tuple_instance_write_destination(14, factory_talk_TTM_147, '[TTM 147]', 9, 12)
tuple_instance_write_destination(45, factory_talk_Chocotech_Kitchen, '[Chocotech Kitchen]', 0, 0)
