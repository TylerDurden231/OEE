"""!Module to calculate TRC Prodution OEE
This module get, using mysql, the information of all done tests of current morning shift and day before afternoon shift.
With all this information, this module calculate availability, performance, quality and oee respectivitly for all slots in each working hour, for the day in each slot and for day in each hour.
A graphic will be build for each done calcule.
In the end its send a mail with all information.  
**Version -** 1.0.0  
**Date -** 08.09.2020.  
**Author -** ESS_HFA (João Lemos & João Gomes)  
@warning Host must have connection to TRC_Alverca Network to be able to connect to database and connection to HFA Network to be able to send mails. 
"""

import smtplib
import os
from email.message import EmailMessage
from email.mime.image import MIMEImage
from datetime import datetime, timedelta
import mysql.connector
import matplotlib.pyplot as plt
import numpy as np
import csv
import zipfile
import requests
import urllib3
import json
from xml.etree import ElementTree
import xmltodict

##
#Number of expected tests in one hour for each model of STBs
STB_Expec_Perf = {"Test": 1, "Pace DCR7151":4, "Pace DCR8151":4, "Pace DSR7151":4, "Pace DSR8151":4, "Thomson DCI7211":4, "Intek S61NV":4, "Pace ZD4500ZNO":4, "Intek DTA":4, "KAON":4, "Arris ZC4430KNO":5}
##
#Number of expected tests in one hour for each model of HGWs
HGW_Expec_Perf = {"HITRON HUB 1.0":12, "HITRON HUB 3.0 v3":12, "HITRON HUB FTTH 2.0":12, "HITRON HUB FTTH":12, "THOMSON EMTA 2.0":12, "HITRON HUB 3.0 v2":12, "HITRON HUB 3.0 v2 ESD":12, "HITRON HUB 3.0":12, "HITRON HUB 3.0 v1 ESD":12, "HITRON HUB 3.0 v1": 12, "HUB 4.0": 12, "GS WIFI": 12, "ARRIS HUB 4.0": 12}
##
#All working hours
hours = ['16h', '17h', '18h', '19h', '20h', '21h', '22h', '23h', '00h', '7h', '8h', '9h', '10h', '11h', '12h', '13h', '14h', '15h']
##
#Working hours of afternoon shift
aftern_shift = ['16h', '17h', '18h', '19h', '20h', '21h', '22h', '23h', '0h']
##
#Working hours of morning shift
morning_shift = ['7h', '8h', '9h', '10h', '11h', '12h', '13h', '14h', '15h']
##
#Number of Working Slots of STB
slots_stb = ["11", "12", "13", "14", "15", "16", "17", "18", "21", "22", "23", "24", "25", "26", "27", "28", "31", "32", "33", "34", "35", "36", "37", "38", "41", "42", "43", "44", "45", "46", "47", "48", "51", "52", "53", "54", "55", "56", "57", "58", "61", "62", "63", "64", "65", "66", "67", "68", "71", "72", "73", "74", "75", "76", "77", "78"]
##
#Number of Working Slots of HGW
slots_hgw = ["01", "02", "03", "04", "05", "06", "07", "08", "09", "10", "11", "12"]
##
#Files to attach to mail
files = ['HGW_Graphs_Hours_Day.png', 
'HGW_Graphs_Slots_Day.png', 
'STB_Graphs_Hours_Day.png', 
'STB_Graphs_Slots_Day.png', 
'OEE_Day_HGW.png', 
'OEE_Day_STB.png',
'OEE_Morning_Shift_STB.png',
'OEE_Afternoon_Shift_STB.png',
'OEE_Morning_Shift_HGW.png',
'OEE_Afternoon_Shift_HGW.png',  
'HGW_Data.csv', 
'STB_Data.csv', 
'Numeric_Parameters.csv', 
'graphs.zip']

##
#STB Data File name
#stb_csv = "STB_Data.csv"
##
#STB Data File name
#hgw_csv = "HGW_Data.csv"

#Credentials for database connection
##
#Database IP address
default_address = "192.168.10.200"
##
#Database access user
user = "read-only_trc"
##
#Database access password 
pwd = "v4pP<2_"
##
#Database name
db = "hgw"


class Connect_to_db:
    """!Class that allows comunicate with Database
    Class with several functions to establish communication with database and send querys and receive the respective answers
    """
    #constructor of object
    def __init__(self, hospedeiro= default_address, utilizador=user, password=pwd, db=db):
        """!Creates a new Connect_to_db object.
        This Connect_to_db Object is being used to define database IP address, access user, access password and database name
        @param hospedeiro Database IP address
        @param utilizador Database User
        @param password Database Password
        @param db Database Name
        ###Example
        ~~~.py 
        object = Connect_to_db
        ~~~
        """ 
        self.hospedeiro = hospedeiro
        self.utilizador = utilizador
        self.password = password
        self.db = db

    #connects the object self to the variables passed in its incialization. 
    ## JL- Connects to DB.
    def connect(self):
        """!Connects to Database
        Use parameters passed when was created Connect_to_db object to establish connection to database
        @returns self.mydb database connection
        """
        self.mydb = mysql.connector.connect(
        host = str(self.hospedeiro),
        user = str(self.utilizador),
        passwd = str(self.password),
        database = str(self.db)
        )
        return self.mydb

    def get_perf_qual_by_slot(self, platform, hour, date_of_today, date_of_yesterday, slot):
        """!Send query
        Send query to get all tests made in a defined hour (model, endtime and result)
        @param platform name of table to check
        @param hour Hour to check
        @param date_of_today Date of day we want to check morning shift
        @param date_of_yesterday Date of day before to check afternoon shift
        @param slot Slot to check
        @warning platform parameter must be **"HGW"** or **"STB"**
        @returns result List of tuples where each tuple is a test done
        """
        hours = ":00:00"
        if hour > 16 and hour <= 23:
            day = date_of_yesterday
        else:
            day = date_of_today
        start_hour_time = str(hour) + ":00:00"
        end_hour_time = str(hour + 1) + ":00:00"

        start_time = day + " " + start_hour_time
        end_time =  day + " " + end_hour_time

        table = "hgw.testresult" + platform.lower()

        #sql = 'SELECT modeltxt, count(*) as "Tests", MAX(endtime) as "EndTime", COUNT(CASE When result = "OK" THEN 1 ELSE NULL END) as "OKs" FROM {} WHERE endtime BETWEEN "{}" AND "{}" AND slot = {} group by modeltxt;'.format(table, start_time, end_time, slot)
        #sql = 'SELECT modeltxt, MAX(endtime) as "EndTime", COUNT(CASE When result = "OK" THEN 1 ELSE NULL END) as "OK" FROM {} WHERE endtime BETWEEN "{}" AND "{}" AND slot = {} group by endtime;'.format(table, start_time, end_time, slot)
        sql = 'SELECT modeltxt, endtime as "EndTime", result as "Result" FROM {} WHERE endtime BETWEEN "{}" AND "{}" AND slot = {};'.format(table, start_time, end_time, slot)

        #execute query
        mycursor = self.mydb.cursor()
        mycursor.execute(sql)

        result = mycursor.fetchall()

        return result
    
    def get_perf_qual_all_in_one(self, platform, date_of_yesterday, date_of_today, week_day):
        """!Send query
        Send query to get all tests made (slot, model, endtime and result)
        @param platform name of table to check
        @param date_of_today Date of day we want to check morning shift
        @param date_of_yesterday Date of day before to check afternoon shift
        @warning platform parameter must be **"HGW"** or **"STB"**
        @returns result List of tuples where each tuple is a test done
        """
        if week_day != 0:
            hour_time = "16:00:00"

            start_time = date_of_yesterday + " " + hour_time
            end_time =  date_of_today + " " + hour_time

            table = "hgw.testresult" + platform.lower()

            #sql = 'SELECT slot, modeltxt, MAX(endtime) as "EndTime", COUNT(CASE When result = "OK" THEN 1 ELSE NULL END) as "OK" FROM {} WHERE endtime BETWEEN "{}" AND "{}" group by endtime order by slot ASC, endtime ASC;'.format(table, start_time, end_time)
            sql = 'SELECT slot, modeltxt, endtime as "EndTime", result as "Result" FROM {} WHERE endtime BETWEEN "{}" AND "{}" order by slot ASC, endtime ASC;'.format(table, start_time, end_time)
        else:
            hour_time = "16:00:00"

            date_yes_plus_one = datetime.today() - timedelta(2)
            date_yes_plus_one = str(date_yes_plus_one.year) + "-" + str(date_yes_plus_one.month) + "-" +str(date_yes_plus_one.day)

            start_time_1 = date_of_yesterday + " " + hour_time
            end_time_1 =  date_yes_plus_one + " 02:00:00" 

            start_time_2 = date_of_today + " 07:00:00"
            end_time_2 =  date_of_today + " " + hour_time

            table = "hgw.testresult" + platform.lower()

            #sql = 'SELECT slot, modeltxt, MAX(endtime) as "EndTime", COUNT(CASE When result = "OK" THEN 1 ELSE NULL END) as "OK" FROM {} WHERE endtime BETWEEN "{}" AND "{}" group by endtime order by slot ASC, endtime ASC;'.format(table, start_time, end_time)
            sql = 'SELECT slot, modeltxt, endtime as "EndTime", result as "Result" FROM {} WHERE (endtime BETWEEN "{}" AND "{}") or (endtime BETWEEN "{}" AND "{}") order by slot ASC, endtime ASC;'.format(table, start_time_1, end_time_1, start_time_2, end_time_2)

        #print(sql)
        mycursor = self.mydb.cursor()
        mycursor.execute(sql)

        result = mycursor.fetchall()

        return result

    #close connection to mysql
    # JL - Close mysql connection
    def close_connection(self):
        """!Close Database connection
        Close Database connections created when function connect() is called
        """
        try:
            self.mydb.close()
            return 2
        except:
            return -1

def create_dictonary(plataform):
    """!Creates a dictonary
    Its created a dictonary to save all values of Availability, Performance, Quality and OEE for each hour in each slot
    @param plataform Platform which dictonary will be create for
    @warning platform parameter must be **"HGW"** or **"STB"**
    @returns dict_stb In case of choosen platform be "STB". 
    @returns dict_hgw In case of choosen platform be "HGW". 
    """ 
    if plataform.lower() == "stb":
        slot = 11
        dict_stb = {}

        for i in range(56):
            dict_stb["Slot_" + str(slot)] = {}
            slot += 1
            if str(slot)[1] == "9":
                slot = int(slot) + 2

        for pos in dict_stb:
            hour = 7
            for h in range(18):
                dict_stb[pos][str(hour) + "h"] = {"Disp":0, "Perf":0, "Qual":0, "OEE":0}
                hour += 1
                if hour == 24:
                    hour = 0
                elif hour == 1:
                    hour = 7
        return dict_stb

    elif plataform.lower() == "hgw":
        slot = 1       
        dict_hgw = {}

        for i in range(12):
            dict_hgw["Slot_" + str(slot)] = {}
            slot += 1

        for pos in dict_hgw:
            hour = 7
            for h in range(18):
                dict_hgw[pos][str(hour) + "h"] = {"Disp":0, "Perf":0, "Qual":0, "OEE":0}
                hour += 1
                if hour == 24:
                    hour = 0
                elif hour == 1:
                    hour = 7
        return dict_hgw

def create_dictonary_slot_day(plataform):
    """!Creates a dictonary
    Its created a dictonary to save all values of Availability, Performance, Quality and OEE for each slot
    @param plataform Platform which dictonary will be create for
    @warning platform parameter must be **"HGW"** or **"STB"**
    @returns dict_stb In case of choosen platform be "STB". 
    @returns dict_hgw In case of choosen platform be "HGW". 
    """ 
    if plataform.lower() == "stb":
        slot = 11
        dict_stb = {}

        for i in range(56):
            dict_stb["Slot_" + str(slot)] = {"Disp":0, "Perf":0, "Qual":0, "OEE":0}
            slot += 1
            if str(slot)[1] == "9":
                slot = int(slot) + 2

        return dict_stb

    elif plataform.lower() == "hgw":
        slot = 1       
        dict_hgw = {}

        for i in range(12):
            dict_hgw["Slot_" + str(slot)] = {"Disp":0, "Perf":0, "Qual":0, "OEE":0}
            slot += 1

        return dict_hgw

def create_dictonary_hour_day():
    """!Creates a dictonary
    Its created a dictonary to save all values of Availability, Performance, Quality and OEE for each hour
    @param plataform Platform which dictonary will be create for
    @warning platform parameter must be **"HGW"** or **"STB"**
    @returns dict_hour Returns a created dictonary
    """ 
    hour = 7
    dict_hour = {}

    for i in range(len(hours)):
        dict_hour[str(hour) + "h"] = {"Disp":0, "Perf":0, "Qual":0, "OEE":0}
        hour += 1
        
        if hour == 24:
            hour = 0
    return dict_hour

def create_dictonary_for_csv(plataform):
    """!Creates a dictonary
    Its created a dictonary to save all values of Number of tests, number of OKs, Number of expected tests, Availability, Performance, Quality and OEE for each hour.
    This dictory is used to build csv files
    @param plataform Platform which dictonary will be create for
    @warning platform parameter must be **"HGW"** or **"STB"**
    @returns dict_hour Returns a created dictonary
    """ 
    if plataform.lower() == "stb":
        slot = 11
        dict_stb = {}

        for i in range(56):
            dict_stb["Slot_" + str(slot)] = {}
            slot += 1
            if str(slot)[1] == "9":
                slot = int(slot) + 2

        for pos in dict_stb:
            hour = 7
            for h in range(18):
                dict_stb[pos][str(hour) + "h"] = {"N_Testes":0, "OKs":0, "N_Testes_Expextaveis":0, "Disp":0, "Perf":0, "Qual":0, "OEE":0}
                hour += 1
                if hour == 24:
                    hour = 0
                elif hour == 1:
                    hour = 7
        return dict_stb

    elif plataform.lower() == "hgw":
        slot = 1       
        dict_hgw = {}

        for i in range(12):
            dict_hgw["Slot_" + str(slot)] = {}
            slot += 1

        for pos in dict_hgw:
            hour = 7
            for h in range(18):
                dict_hgw[pos][str(hour) + "h"] = {"N_Testes":0, "OKs":0, "N_Testes_Expextaveis":0, "Disp":0, "Perf":0, "Qual":0, "OEE":0}
                hour += 1
                if hour == 24:
                    hour = 0
                elif hour == 1:
                    hour = 7
        return dict_hgw

def create_unavailability_dict(plataform):
    """!Creates a dictonary
    Its created a dictonary to save all all times of unavailability of each slot
    @param plataform Platform which dictonary will be create for
    @warning platform parameter must be **"HGW"** or **"STB"**
    @returns unavai_dict Dictonary with list for each key(slot) 
    """ 
    plataform = plataform.lower()
    unavai_dict = {}

    if plataform == 'stb':
        slot = 11
        for i in range(56):
            unavai_dict["Slot_" + str(slot)] = []
            slot += 1
            if str(slot)[1] == "9":
                slot = int(slot) + 2
    elif plataform == 'hgw':
        slot = 1
        for i in range(12):
            unavai_dict["Slot_" + str(slot)] = []
            slot += 1

    return unavai_dict

def fullfill_dict_for_csv(platform, result_dict, results_temp):
    """!Fulfill Dictonary 
    Fulfill dictonary with all values of Number of tests, number of OKs, Number of expected tests, Availability, Performance, Quality and OEE for each hour.
    The values used to calculate the desire values to put on this dictonary are on results_temp argument
    This dictonary is used to build csv files
    @param platform Platform which dictonary will be create for 
    @param result_dict Dictonary created in function create_dictonary_for_csv()
    @param results_temp List of results got with a query sended in function Connect_to_db().get_perf_qual_all_in_one()
    @warning platform parameter must be **"HGW"** or **"STB"**
    @returns result_dict Returns a created dictonary
    """ 
    platform = platform.lower()    
    disponibility_slot_h = 60 
    disponibility_slot_h_prev = 60
    dispo_fin = disponibility_slot_h/disponibility_slot_h_prev
    
    platform = platform.lower()

    if platform == 'stb':
        slot = 11
        for pos in result_dict: 
            hour = 7
            for hora in range(18):
                expect = 0
                results = []
                qual_temp = 0
                for i in range(len(results_temp)):
                    slot_number = pos.replace("Slot_","")
                    if results_temp[i][0] == slot_number and results_temp[i][2].hour == hour:
                        list_temp = list(results_temp[i])
                        del list_temp[0]
                        test_tuple = tuple(list_temp)
                        results.append(test_tuple)

                if len(results) == 0:
                    pref = 0
                    qual = 0
                    oee = 0
                else:
                    model_tested = {}
                    expect = 0                       
                    for i in range(len(results)):
                        if i == 0:
                            if results[i][2] == 'OK':
                                qual_temp = 1
                            if i == len(results)-1:
                                model_tested[results[i][0]] = 60 
                            else:
                                model_tested[results[i][0]] = results[i][1].minute                            
                        else:
                            if results[i][2] == 'OK':
                                qual_temp += 1
                            if results[i][0] in model_tested.keys():
                                if i == len(results)-1:
                                    model_tested[results[i][0]] += 60 - results[i-1][1].minute
                                else:
                                    model_tested[results[i][0]] += results[i][1].minute - results[i-1][1].minute
                            else:
                                if i == len(results)-1:
                                    model_tested[results[i][0]] = 60 - results[i-1][1].minute
                                else:
                                    model_tested[results[i][0]] = results[i][1].minute - results[i-1][1].minute                            
                                
                    for key in model_tested:
                        expect += (model_tested[key]*STB_Expec_Perf[key])/60
                    
                    pref = round(len(results)/expect,4)
                    qual = round(qual_temp/len(results),4)
                    oee =  round((dispo_fin*pref*qual),4)
                
                result_dict[pos][str(hour) + "h"]["N_Testes"] = int(len(results))
                result_dict[pos][str(hour) + "h"]["OKs"] = int(qual_temp)
                result_dict[pos][str(hour) + "h"]["N_Testes_Expextaveis"] = expect
                result_dict[pos][str(hour) + "h"]["Disp"] = round(dispo_fin*100,4)
                result_dict[pos][str(hour) + "h"]["Perf"] = round(pref*100,4)
                result_dict[pos][str(hour) + "h"]["Qual"] = round(qual*100,4)
                result_dict[pos][str(hour) + "h"]["OEE"] = round(oee*100,4)
                
                hour += 1
                if hour == 24:
                    hour = 0
                elif hour == 1:
                    hour = 7

            slot += 1
            if str(slot)[1] == "9":
                slot = int(slot) + 2
    
    elif platform == 'hgw':
        slot = 1
        for pos in result_dict:
            hour = 7
            for hora in range(18):
                expect = 0
                results = []
                qual_temp = 0
                for i in range(len(results_temp)):
                    slot_number = pos.replace("Slot_","")
                    if len(slot_number) == 1:
                        slot_number = "0" + slot_number
                    if results_temp[i][0] == slot_number and results_temp[i][2].hour == hour:
                        list_temp = list(results_temp[i])
                        del list_temp[0]
                        test_tuple = tuple(list_temp)
                        results.append(test_tuple)
               
                if len(results) == 0:
                    pref = 0
                    qual = 0
                    oee = 0
                else:
                    model_tested = {}
                    expect = 0                       
                    for i in range(len(results)):
                        if i == 0:
                            if results[i][2] == 'OK':
                                qual_temp = 1
                            if i == len(results)-1:
                                model_tested[results[i][0]] = 60 
                            else:
                                model_tested[results[i][0]] = results[i][1].minute                            
                        else:
                            if results[i][2] == 'OK':
                                qual_temp += 1
                            if results[i][0] in model_tested.keys():
                                if i == len(results)-1:
                                    model_tested[results[i][0]] += 60 - results[i-1][1].minute
                                else:
                                    model_tested[results[i][0]] += results[i][1].minute - results[i-1][1].minute
                            else:
                                if i == len(results)-1:
                                    model_tested[results[i][0]] = 60 - results[i-1][1].minute
                                else:
                                    model_tested[results[i][0]] = results[i][1].minute - results[i-1][1].minute                            
                                
                    for key in model_tested:
                        expect += (model_tested[key]*HGW_Expec_Perf[key])/60

                    pref = round(len(results)/expect,4)
                    qual = round(qual_temp/len(results),4)
                    oee =  round((dispo_fin*pref*qual),4)

                result_dict[pos][str(hour) + "h"]["N_Testes"] = int(len(results))
                result_dict[pos][str(hour) + "h"]["OKs"] = int(qual_temp)
                result_dict[pos][str(hour) + "h"]["N_Testes_Expextaveis"] = expect
                result_dict[pos][str(hour) + "h"]["Disp"] = round(dispo_fin*100,4)
                result_dict[pos][str(hour) + "h"]["Perf"] = round(pref*100,4)
                result_dict[pos][str(hour) + "h"]["Qual"] = round(qual*100,4)
                result_dict[pos][str(hour) + "h"]["OEE"] = round(oee*100,4)
                
                hour += 1
                if hour == 24:
                    hour = 0
                elif hour == 1:
                    hour = 7

            slot += 1

    return result_dict    

def fullfill_dict(platform, result_dict, results_temp, unvai_slot):
    """!Fulfill Dictonary 
    Fulfill dictonary with all values of Availability, Performance, Quality and OEE for each hour in each slot.
    The values used to calculate the desire values to put on this dictonary are on results_temp argument
    This dictonary is used to build all slot/hour graphics
    @param platform Platform which dictonary will be create for 
    @param result_dict Dictonary created in function create_dictonary()
    @param results_temp List of results got with a query sended in function Connect_to_db().get_perf_qual_all_in_one()
    @warning platform parameter must be **"HGW"** or **"STB"**
    @returns result_dict Returns a created dictonary
    """ 
    platform = platform.lower()    
    disponibility_slot_h = 60 
    disponibility_slot_h_prev = 60
    
    platform = platform.lower()
    ignore_test = False

    #print(results_temp)
    if platform == 'stb':
        slot = 11
        for pos in result_dict: 
            hour = 7
            stopped_list = []
            for i, v in enumerate(unvai_slots[pos]):      
                new_list = unvai_slots[pos][i].split('_')
                new_list[0] = datetime.strptime(new_list[0], '%Y-%m-%d %H:%M:%S')
                new_list[1] = datetime.strptime(new_list[1], '%Y-%m-%d %H:%M:%S')
                stopped_list.append(new_list)
            print(stopped_list)
            for hora in range(18):
                results = []
                qual_temp = 0
                for i in range(len(results_temp)):
                    slot_number = pos.replace("Slot_","")
                    if results_temp[i][0] == slot_number and results_temp[i][2].hour == hour:
                        for j, k in enumerate(stopped_list):
                            if results_temp[i][2].hour >= stopped_list[j][0].hour and results_temp[i][2].hour <= stopped_list[j][1].hour:
                                if results_temp[i][2].hour == stopped_list[j][0].hour and results_temp[i][2].hour != stopped_list[j][1].hour:
                                    if results_temp[i][2].minute > stopped_list[j][0].minute:
                                        ignore_test = True
                                elif results_temp[i][2].hour == stopped_list[j][0].hour and results_temp[i][2].hour == stopped_list[j][1].hour:
                                    if results_temp[i][2].minute > stopped_list[j][0].minute and results_temp[i][2].minute <= stopped_list[j][1].minute:
                                        ignore_test = True
                                elif results_temp[i][2].hour != stopped_list[j][0].hour and results_temp[i][2].hour == stopped_list[j][1].hour:
                                    if results_temp[i][2].minute <= stopped_list[j][1].minute:
                                        ignore_test = True
                                elif results_temp[i][2].hour != stopped_list[j][0].hour and results_temp[i][2].hour != stopped_list[j][1].hour:
                                    ignore_test = True
                                else:
                                    print("!!!!!!!!!!!!!!Error 1!!!!!!!!!!!!!!!!!")
                        if ignore_test == False:
                            list_temp = list(results_temp[i])
                            del list_temp[0]
                            test_tuple = tuple(list_temp)
                            results.append(test_tuple)
                        else:
                            ignore_test = False
                stop_time = 0
                stopped_times = {}
                for j, k in enumerate(stopped_list):
                    if hour >= stopped_list[j][0].hour and hour <= stopped_list[j][1].hour:
                        if hour == stopped_list[j][0].hour and hour == stopped_list[j][1].hour:
                            stop_time += stopped_list[j][1].minute - stopped_list[j][0].minute
                            stopped_times["start_" + str(j)] = stopped_list[j][0].minute
                            stopped_times["stop_" + str(j)] = stopped_list[j][1].minute
                        elif hour == stopped_list[j][0].hour and hour != stopped_list[j][1].hour:
                            stop_time += 60 - stopped_list[j][0].minute
                            stopped_times["start_" + str(j)] = stopped_list[j][0].minute
                        elif hour != stopped_list[j][0].hour and hour == stopped_list[j][1].hour:
                            stop_time += stopped_list[j][1].minute
                            stopped_times["stop_" + str(j)] = stopped_list[j][1].minute
                        elif hour != stopped_list[j][0].hour and hour != stopped_list[j][1].hour:
                            stop_time += 60
                        else:
                            print("!!!!!!!!!!!!!!Error 2!!!!!!!!!!!!!!!!!")
                print(stopped_times)
                dispo_fin = round((60 - stop_time)/disponibility_slot_h_prev,4)

                print(pos + "_Hour_" + str(hour) + "--" + str(results))
                if len(results) == 0:
                    pref = 0
                    qual = 0
                    oee = 0
                else:
                    model_tested = {}
                    expect = 0                       
                    for i in range(len(results)):
                        if i == 0:
                            if results[i][2] == 'OK':
                                qual_temp = 1
                            if i == len(results)-1:
                                model_tested[results[i][0]] = 60 - stop_time
                            else:
                                if len(stopped_times) == 0:
                                    model_tested[results[i][0]] = results[i][1].minute
                                elif list(stopped_times.keys())[0].startswith("stop"):
                                    model_tested[results[i][0]] = results[i][1].minute - stopped_times[list(stopped_times.keys())[0]]
                                elif results[i][1].minute < stopped_times[list(stopped_times.keys())[0]]:
                                    model_tested[results[i][0]] = results[i][1].minute
                                elif results[i][1].minute > stopped_times[list(stopped_times.keys())[0]]:
                                    model_tested[results[i][0]] = results[i][1].minute - stopped_times[list(stopped_times.keys())[1]] + stopped_times[list(stopped_times.keys())[0]]
                                else:
                                    print("!!!!!!!!!!!!!!Error 3!!!!!!!!!!!!!!!!!")
                        else:
                            if results[i][2] == 'OK':
                                qual_temp += 1
                            if results[i][0] in model_tested.keys():
                                if i == len(results)-1:
                                    if len(stopped_times) == 0:
                                        model_tested[results[i][0]] += 60 - results[i-1][1].minute
                                    elif list(stopped_times.keys())[len(stopped_times)-1].startswith("stop"):
                                        if results[i-1][1].minute > stopped_times[list(stopped_times.keys())[len(stopped_times)-1]]:
                                            model_tested[results[i][0]] += 60 - results[i-1][1].minute
                                        elif results[i][1].minute > stopped_times[list(stopped_times.keys())[len(stopped_times)-1]]:
                                            model_tested[results[i][0]] += 60 - stopped_times[list(stopped_times.keys())[len(stopped_times)-1]]
                                        else:
                                            model_tested[results[i][0]] += 60 - stopped_times[list(stopped_times.keys())[len(stopped_times)-1]] + stopped_times[list(stopped_times.keys())[len(stopped_times)-2]] - results[i-1][1].minute
                                    elif list(stopped_times.keys())[len(stopped_times)-1].startswith("start"):
                                        if len(stopped_times) > 1:
                                            if results[i-1][1].minute > stopped_times[list(stopped_times.keys())[len(stopped_times)-2]]:
                                                model_tested[results[i][0]] += stopped_times[list(stopped_times.keys())[len(stopped_times)-1]] - results[i-1][1].minute
                                            else:
                                                model_tested[results[i][0]] += stopped_times[list(stopped_times.keys())[len(stopped_times)-1]] - stopped_times[list(stopped_times.keys())[len(stopped_times)-2]] + stopped_times[list(stopped_times.keys())[len(stopped_times)-3]] - results[i-1][1].minute
                                        else:
                                            model_tested[results[i][0]] += stopped_times[list(stopped_times.keys())[len(stopped_times)-1]] - results[i-1][1].minute
                                    else:
                                        else:
                                            print("!!!!!!!!!!!!!!Error 4!!!!!!!!!!!!!!!!!")
                                else:
                                    if len(stopped_times) == 0:
                                        model_tested[results[i][0]] += results[i][1].minute - results[i-1][1].minute
                                    else:
                                        middle_stop = False
                                        for x, y in stopped_times.items():
                                            if middle_stop:
                                                mstop = stopped_times[x]
                                                break
                                            if stopped_times[x] < results[i][1].minute and stopped_times[x] > results[i-1][1].minute:
                                                middle_stop = True
                                                mstart = stopped_times[x]
                                                continue
                                        if middle_stop:
                                            model_tested[results[i][0]] += results[i][1].minute - mstop + mstart - results[i-1][1].minute
                                        else:
                                            model_tested[results[i][0]] += results[i][1].minute - results[i-1][1].minute

                            else:
                                if i == len(results)-1:
                                    if len(stopped_times) == 0:
                                        model_tested[results[i][0]] = 60 - results[i-1][1].minute
                                    elif list(stopped_times.keys())[len(stopped_times)-1].startswith("stop"):
                                        if results[i-1][1].minute > list(stopped_times.keys())[len(stopped_times)-1]:
                                            model_tested[results[i][0]] = 60 - results[i-1][1].minute
                                        elif results[i][1].minute > list(stopped_times.keys())[len(stopped_times)-1]:
                                            model_tested[results[i][0]] = 60 - list(stopped_times.keys())[len(stopped_times)-1]
                                        else:
                                            model_tested[results[i][0]] = 60 - list(stopped_times.keys())[len(stopped_times)-1] + list(stopped_times.keys())[len(stopped_times)-2] - results[i-1][1].minute
                                    elif list(stopped_times.keys())[len(stopped_times)-1].startswith("start"):
                                        if len(stopped_times) > 1:
                                            if results[i-1][1].minute > list(stopped_times.keys())[len(stopped_times)-2]:
                                                model_tested[results[i][0]] = list(stopped_times.keys())[len(stopped_times)-1] - results[i-1][1].minute
                                            else:
                                                model_tested[results[i][0]] = list(stopped_times.keys())[len(stopped_times)-1] - list(stopped_times.keys())[len(stopped_times)-2] + list(stopped_times.keys())[len(stopped_times)-3] - results[i-1][1].minute
                                        else:
                                            model_tested[results[i][0]] = list(stopped_times.keys())[len(stopped_times)-1] - results[i-1][1].minute
                                    else:
                                    print("!!!!!!!!!!!!!!Error 5!!!!!!!!!!!!!!!!!")
                                else:
                                    if len(stopped_times) == 0:
                                        model_tested[results[i][0]] = results[i][1].minute - results[i-1][1].minute   
                                    else:
                                        middle_stop = False
                                        for x, y in stopped_times.items():
                                            if middle_stop:
                                                mstop = stopped_times[x]
                                                break
                                            if stopped_times[x] < results[i][1].minute and stopped_times[x] > results[i-1][1].minute:
                                                middle_stop = True
                                                mstart = stopped_times[x]
                                                continue
                                        if middle_stop:
                                            model_tested[results[i][0]] = results[i][1].minute - mstop + mstart - results[i-1][1].minute
                                        else:
                                            model_tested[results[i][0]] = results[i][1].minute - results[i-1][1].minute

                        #print(model_tested)
                    for key in model_tested:
                        expect += (model_tested[key]*STB_Expec_Perf[key])/(60)
                    
                    pref = round(len(results)/expect,4)
                    qual = round(qual_temp/len(results),4)
                    oee =  round((dispo_fin*pref*qual),4)
                #print("Availability: " + str(dispo_fin) + " Preformance: " + str(pref) + " Quality: " + str(qual) + " OEE: " + str(oee))

                result_dict[pos][str(hour) + "h"]["Disp"] = round(dispo_fin*100,4)
                result_dict[pos][str(hour) + "h"]["Perf"] = round(pref*100,4)
                result_dict[pos][str(hour) + "h"]["Qual"] = round(qual*100,4)
                result_dict[pos][str(hour) + "h"]["OEE"] = round(oee*100,4)
                
                hour += 1
                if hour == 24:
                    hour = 0
                elif hour == 1:
                    hour = 7
           
            slot += 1
            if str(slot)[1] == "9":
                slot = int(slot) + 2
    
    elif platform == 'hgw':
        slot = 1
        for pos in result_dict:
            hour = 7
            stopped_list = []
            for i, v in enumerate(unvai_slots[pos]):      
                new_list = unvai_slots[pos][i].split('_')
                new_list[0] = datetime.strptime(new_list[0], '%Y-%m-%d %H:%M:%S')
                new_list[1] = datetime.strptime(new_list[1], '%Y-%m-%d %H:%M:%S')
                stopped_list.append(new_list)
            for hora in range(18):
                results = []
                qual_temp = 0
                for i in range(len(results_temp)):
                    slot_number = pos.replace("Slot_","")
                    if len(slot_number) == 1:
                        slot_number = "0" + slot_number
                    if results_temp[i][0] == slot_number and results_temp[i][2].hour == hour:
                        for j, k in enumerate(stopped_list):
                            if results_temp[i][2].hour >= stopped_list[j][0].hour and results_temp[i][2].hour <= stopped_list[j][1].hour:
                                if results_temp[i][2].hour == stopped_list[j][0].hour and results_temp[i][2].hour != stopped_list[j][1].hour:
                                    if results_temp[i][2].minute > stopped_list[j][0].minute:
                                        ignore_test = True
                                elif results_temp[i][2].hour == stopped_list[j][0].hour and results_temp[i][2].hour == stopped_list[j][1].hour:
                                    if results_temp[i][2].minute > stopped_list[j][0].minute and results_temp[i][2].minute <= stopped_list[j][1].minute:
                                        ignore_test = True
                                elif results_temp[i][2].hour != stopped_list[j][0].hour and results_temp[i][2].hour == stopped_list[j][1].hour:
                                    if results_temp[i][2].minute <= stopped_list[j][1].minute:
                                        ignore_test = True
                                elif results_temp[i][2].hour != stopped_list[j][0].hour and results_temp[i][2].hour != stopped_list[j][1].hour:
                                    ignore_test = True
                                else:
                                    print("!!!!!!!!!!!!!!Error 1!!!!!!!!!!!!!!!!!")
                        if ignore_test == False:
                            list_temp = list(results_temp[i])
                            del list_temp[0]
                            test_tuple = tuple(list_temp)
                            results.append(test_tuple)
                        else:
                            ignore_test = False
                stop_time = 0
                stopped_times = {}
                for j, k in enumerate(stopped_list):
                    if hour >= stopped_list[j][0].hour and hour <= stopped_list[j][1].hour:
                        if hour == stopped_list[j][0].hour and hour == stopped_list[j][1].hour:
                            stop_time += stopped_list[j][1].minute - stopped_list[j][0].minute
                            stopped_times["start_" + str(j)] = stopped_list[j][0].minute
                            stopped_times["stop_" + str(j)] = stopped_list[j][1].minute
                        elif hour == stopped_list[j][0].hour and hour != stopped_list[j][1].hour:
                            stop_time += 60 - stopped_list[j][0].minute
                            stopped_times["start_" + str(j)] = stopped_list[j][0].minute
                        elif hour != stopped_list[j][0].hour and hour == stopped_list[j][1].hour:
                            stop_time += stopped_list[j][1].minute
                            stopped_times["stop_" + str(j)] = stopped_list[j][1].minute
                        elif hour != stopped_list[j][0].hour and hour != stopped_list[j][1].hour:
                            stop_time += 60
                        else:
                            print("!!!!!!!!!!!!!!Error 2!!!!!!!!!!!!!!!!!")

                dispo_fin = round((60 - stop_time)/disponibility_slot_h_prev,4)

                if len(results) == 0:
                    pref = 0
                    qual = 0
                    oee = 0
                else:
                    model_tested = {}
                    expect = 0                       
                    for i in range(len(results)):
                        if i == 0:
                            if results[i][2] == 'OK':
                                qual_temp = 1
                            if i == len(results)-1:
                                model_tested[results[i][0]] = 60 - stop_time
                            else:
                                if len(stopped_times) == 0:
                                    model_tested[results[i][0]] = results[i][1].minute
                                elif list(stopped_times.keys())[0].startswith("stop"):
                                    model_tested[results[i][0]] = results[i][1].minute - stopped_times[list(stopped_times.keys())[0]]
                                elif results[i][1].minute < stopped_times[list(stopped_times.keys())[0]]:
                                    model_tested[results[i][0]] = results[i][1].minute
                                elif results[i][1].minute > stopped_times[list(stopped_times.keys())[0]]:
                                    model_tested[results[i][0]] = results[i][1].minute - stopped_times[list(stopped_times.keys())[1]] + stopped_times[list(stopped_times.keys())[0]]    
                                else:
                                    print("!!!!!!!!!!!!!!Error 3!!!!!!!!!!!!!!!!!")
                        else:
                            if results[i][2] == 'OK':
                                qual_temp += 1
                            if results[i][0] in model_tested.keys():
                                if i == len(results)-1:
                                    if len(stopped_times) == 0:
                                        model_tested[results[i][0]] += 60 - results[i-1][1].minute
                                    elif list(stopped_times.keys())[len(stopped_times)-1].startswith("stop"):
                                        if results[i-1][1].minute > stopped_times[list(stopped_times.keys())[len(stopped_times)-1]]:
                                            model_tested[results[i][0]] += 60 - results[i-1][1].minute
                                        elif results[i][1].minute > stopped_times[list(stopped_times.keys())[len(stopped_times)-1]]:
                                            model_tested[results[i][0]] += 60 - stopped_times[list(stopped_times.keys())[len(stopped_times)-1]]
                                        else:
                                            model_tested[results[i][0]] += 60 - stopped_times[list(stopped_times.keys())[len(stopped_times)-1]] + stopped_times[list(stopped_times.keys())[len(stopped_times)-2]] - results[i-1][1].minute
                                    elif list(stopped_times.keys())[len(stopped_times)-1].startswith("start"):
                                        if len(stopped_times) > 1:
                                            if results[i-1][1].minute > stopped_times[list(stopped_times.keys())[len(stopped_times)-2]]:
                                                model_tested[results[i][0]] += stopped_times[list(stopped_times.keys())[len(stopped_times)-1]] - results[i-1][1].minute
                                            else:
                                                model_tested[results[i][0]] += stopped_times[list(stopped_times.keys())[len(stopped_times)-1]] - stopped_times[list(stopped_times.keys())[len(stopped_times)-2]] + stopped_times[list(stopped_times.keys())[len(stopped_times)-3]] - results[i-1][1].minute
                                        else:
                                            model_tested[results[i][0]] += stopped_times[list(stopped_times.keys())[len(stopped_times)-1]] - results[i-1][1].minute
                                    else:
                                        print("!!!!!!!!!!!!!!Error 4!!!!!!!!!!!!!!!!!")
                                else:
                                    if len(stopped_times) == 0:
                                        model_tested[results[i][0]] += results[i][1].minute - results[i-1][1].minute
                                    else:
                                        middle_stop = False
                                        for x, y in stopped_times.items():
                                            if middle_stop:
                                                mstop = stopped_times[x]
                                                break
                                            if stopped_times[x] < results[i][1].minute and stopped_times[x] > results[i-1][1].minute:
                                                middle_stop = True
                                                mstart = stopped_times[x]
                                                continue
                                        if middle_stop:
                                            model_tested[results[i][0]] += results[i][1].minute - mstop + mstart - results[i-1][1].minute
                                        else:
                                            model_tested[results[i][0]] += results[i][1].minute - results[i-1][1].minute

                            else:
                                if i == len(results)-1:
                                    if len(stopped_times) == 0:
                                        model_tested[results[i][0]] = 60 - results[i-1][1].minute
                                    elif list(stopped_times.keys())[len(stopped_times)-1].startswith("stop"):
                                        if results[i-1][1].minute > list(stopped_times.keys())[len(stopped_times)-1]:
                                            model_tested[results[i][0]] = 60 - results[i-1][1].minute
                                        elif results[i][1].minute > list(stopped_times.keys())[len(stopped_times)-1]:
                                            model_tested[results[i][0]] = 60 - list(stopped_times.keys())[len(stopped_times)-1]
                                        else:
                                            model_tested[results[i][0]] = 60 - list(stopped_times.keys())[len(stopped_times)-1] + list(stopped_times.keys())[len(stopped_times)-2] - results[i-1][1].minute
                                    elif list(stopped_times.keys())[len(stopped_times)-1].startswith("start"):
                                        if len(stopped_times) > 1:
                                            if results[i-1][1].minute > list(stopped_times.keys())[len(stopped_times)-2]:
                                                model_tested[results[i][0]] = list(stopped_times.keys())[len(stopped_times)-1] - results[i-1][1].minute
                                            else:
                                                model_tested[results[i][0]] = list(stopped_times.keys())[len(stopped_times)-1] - list(stopped_times.keys())[len(stopped_times)-2] + list(stopped_times.keys())[len(stopped_times)-3] - results[i-1][1].minute
                                        else:
                                            model_tested[results[i][0]] = list(stopped_times.keys())[len(stopped_times)-1] - results[i-1][1].minute
                                    else:
                                        print("!!!!!!!!!!!!!!Error 5!!!!!!!!!!!!!!!!!")
                                else:
                                    if len(stopped_times) == 0:
                                        model_tested[results[i][0]] = results[i][1].minute - results[i-1][1].minute   
                                    else:
                                        middle_stop = False
                                        for x, y in stopped_times.items():
                                            if middle_stop:
                                                mstop = stopped_times[x]
                                                break
                                            if stopped_times[x] < results[i][1].minute and stopped_times[x] > results[i-1][1].minute:
                                                middle_stop = True
                                                mstart = stopped_times[x]
                                                continue
                                        if middle_stop:
                                            model_tested[results[i][0]] = results[i][1].minute - mstop + mstart - results[i-1][1].minute
                                        else:
                                            model_tested[results[i][0]] = results[i][1].minute - results[i-1][1].minute
                            
                    for key in model_tested:
                        expect += (model_tested[key]*HGW_Expec_Perf[key])/60

                    pref = round(len(results)/expect,4)
                    qual = round(qual_temp/len(results),4)
                    oee =  round((dispo_fin*pref*qual),4)

                result_dict[pos][str(hour) + "h"]["Disp"] = round(dispo_fin*100,4)
                result_dict[pos][str(hour) + "h"]["Perf"] = round(pref*100,4)
                result_dict[pos][str(hour) + "h"]["Qual"] = round(qual*100,4)
                result_dict[pos][str(hour) + "h"]["OEE"] = round(oee*100,4)
                
                hour += 1
                if hour == 24:
                    hour = 0
                elif hour == 1:
                    hour = 7

            slot += 1

    return result_dict    

def fullfill_dict_slot_day(platform, result_dict, results_temp):
    """!Fulfill Dictonary 
    Fulfill dictonary with all values of Availability, Performance, Quality and OEE for each hour.
    The values used to calculate the desire values to put on this dictonary are on results_temp argument
    This dictonary is used to build graphics of OEE parameters of one day for all slots
    @param platform Platform which dictonary will be create for 
    @param result_dict Dictonary created in function create_dictonary_slot_day()
    @param results_temp List of results got with a query sended in function Connect_to_db().get_perf_qual_all_in_one()
    @warning platform parameter must be **"HGW"** or **"STB"**
    @returns result_dict Returns a created dictonary
    """
    platform = platform.lower()
    disponibility_slot_h = 60 
    disponibility_slot_h_prev = 60
    dispo_fin = disponibility_slot_h/disponibility_slot_h_prev
    slot = 11
    corres_hour = {'16h':1, '17h':2, '18h':3, '19h':4, '20h':5, '21h':6, '22h':7, '23h':8, '00h':9, '7h':10, '8h':11, '9h':12, '10h':13, '11h':14, '12h':15, '13h':16, '14h':17, '15h':18}
    #print(results_temp)
    for pos in result_dict: 
        results = []
        qual_temp = 0
        for i in range(len(results_temp)):
            slot_number = pos.replace("Slot_","")
            if len(slot_number) == 1:
                slot_number = "0" + slot_number
            if results_temp[i][0] == slot_number and str(results_temp[i][2].hour) + "h" in hours:
                list_temp = list(results_temp[i])
                del list_temp[0]
                test_tuple = tuple(list_temp)
                results.append(test_tuple)
        #print("*****" + pos + "----" + str(results))

        if len(results) == 0:
            pref = 0
            qual = 0
            oee = 0
        else:
            model_tested = {}
            expect = 0                       
            for i in range(len(results)):
                if i == 0:
                    if results[i][2] == 'OK':
                        qual_temp = 1
                    if i == len(results)-1:
                        model_tested[results[i][0]] = 60*len(hours)
                    else:
                        model_tested[results[i][0]] = ((corres_hour[str(results[i][1].hour) + "h"] - 1) *60) + results[i][1].minute                            
                else:
                    if results[i][2] == 'OK':
                        qual_temp += 1
                    if results[i][0] in model_tested.keys():
                        if i == len(results)-1:
                            model_tested[results[i][0]] += (60*len(hours)) - (((corres_hour[str(results[i-1][1].hour) + "h"] - 1) * 60) + results[i-1][1].minute)
                        else:
                            model_tested[results[i][0]] += (((corres_hour[str(results[i][1].hour) + "h"] - 1) * 60) + results[i][1].minute) - (((corres_hour[str(results[i-1][1].hour) + "h"] - 1) * 60) + results[i-1][1].minute)
                    else:
                        if i == len(results)-1:
                            model_tested[results[i][0]] = (60*len(hours)) - (((corres_hour[str(results[i-1][1].hour) + "h"] - 1) * 60) + results[i-1][1].minute)
                        else:
                            model_tested[results[i][0]] = (((corres_hour[str(results[i][1].hour) + "h"] - 1) * 60) + results[i][1].minute) - (((corres_hour[str(results[i-1][1].hour) + "h"] - 1) * 60) + results[i-1][1].minute)                            
            #print(model_tested)
            if platform == "stb":
                for key in model_tested:
                    expect += (model_tested[key]*STB_Expec_Perf[key]*18)/(60*len(hours))
            elif platform == "hgw":
                for key in model_tested:
                    expect += (model_tested[key]*HGW_Expec_Perf[key]*18)/(60*len(hours))
            
            # print("Lenght - " + str(len(results)))
            # print("Expect - " + str(expect))
            # print("Qual - " + str(qual_temp))
            pref = round(len(results)/expect,4)
            qual = round(qual_temp/len(results),4)
            oee =  round((dispo_fin*pref*qual),4)
            
        result_dict[pos]["Disp"] = round(dispo_fin*100,4)
        result_dict[pos]["Perf"] = round(pref*100,4)
        result_dict[pos]["Qual"] = round(qual*100,4)
        result_dict[pos]["OEE"] = round(oee*100,4)  

        slot += 1
        if platform == "stb":
            if str(slot)[1] == "9":
                slot = int(slot) + 2       
        
    return result_dict 

def fullfill_dict_old(platform, result_dict, date_yesterday, date_today):
    """!Fulfill Dictonary 
    Fulfill dictonary with all values of Availability, Performance, Quality and OEE for each hour.
    The values used to calculate the desire values to put on this dictonary are in variable results, result of several querys using function Connect_to_db().get_perf_qual_by_slot()
    This dictonary is used to build all slot/hour graphics
    @param platform Platform which dictonary will be create for 
    @param result_dict Dictonary created in function create_dictonary()
    @param date_yesterday 
    @param date_today 
    @see Connect_to_db().get_perf_qual_by_slot()
    @warning platform parameter must be **"HGW"** or **"STB"**
    @returns result_dict Returns a created dictonary
    """ 
    platform = platform.lower()    
    disponibility_slot_h = 60 
    disponibility_slot_h_prev = 60
    dispo_fin = disponibility_slot_h/disponibility_slot_h_prev
    
    platform = platform.lower()

    db = Connect_to_db()
    db.connect()

    if platform == 'stb':
        slot = 11
        for pos in result_dict:
            hour = 7
            for hora in range(18):
                results = db.get_perf_qual_by_slot(platform, hour, date_today, date_yesterday, slot)
                print(str(pos) + "_" + str(hora) + str(results))
                qual_temp = 0
                if len(results) == 0:
                    pref = 0
                    qual = 0
                    oee = 0
                else:
                    model_tested = {}
                    expect = 0                       
                    for i in range(len(results)):
                        if i == 0:
                            if results[i][2] == 'OK':
                                qual_temp = 1
                            if i == len(results)-1:
                                model_tested[results[i][0]] = 60 
                            else:
                                model_tested[results[i][0]] = results[i][1].minute                            
                        else:
                            if results[i][2] == 'OK':
                                qual_temp += 1
                            if results[i][0] in model_tested.keys():
                                if i == len(results)-1:
                                    model_tested[results[i][0]] += 60 - results[i-1][1].minute
                                else:
                                    model_tested[results[i][0]] += results[i][1].minute - results[i-1][1].minute
                            else:
                                if i == len(results)-1:
                                    model_tested[results[i][0]] = 60 - results[i-1][1].minute
                                else:
                                    model_tested[results[i][0]] = results[i][1].minute - results[i-1][1].minute                            
                                
                    for key in model_tested:
                        expect += (model_tested[key]*STB_Expec_Perf[key])/60

                    pref = round(len(results)/expect,4)
                    qual = round(qual_temp/len(results),4)
                    oee =  round((dispo_fin*pref*qual),4)
                    
                result_dict[pos][str(hour) + "h"]["Disp"] = round(dispo_fin*100,4)
                result_dict[pos][str(hour) + "h"]["Perf"] = round(pref*100,4)
                result_dict[pos][str(hour) + "h"]["Qual"] = round(qual*100,4)
                result_dict[pos][str(hour) + "h"]["OEE"] = round(oee*100,4)
                
                hour += 1
                if hour == 24:
                    hour = 0
                elif hour == 1:
                    hour = 7

            slot += 1
            if str(slot)[1] == "9":
                slot = int(slot) + 2
    elif platform == 'hgw':
        slot = 1
        for pos in result_dict:
            for hora in range(19):
                results = db.get_perf_qual_by_slot(platform, hour, date_today, date_yesterday, slot)
                print(str(pos) + "_" + str(hora) + str(results))
                qual_temp = 0
                if len(results) == 0:
                    pref = 0
                    qual = 0
                    oee = 0
                else:
                    model_tested = {}
                    expect = 0                       
                    for i in range(len(results)):
                        if i == 0:
                            if results[i][2] == 'OK':
                                qual_temp = 1
                            if i == len(results)-1:
                                model_tested[results[i][0]] = 60 
                            else:
                                model_tested[results[i][0]] = results[i][1].minute                            
                        else:
                            if results[i][2] == 'OK':
                                qual_temp += 1
                            if results[i][0] in model_tested.keys():
                                if i == len(results)-1:
                                    model_tested[results[i][0]] += 60 - results[i-1][1].minute
                                else:
                                    model_tested[results[i][0]] += results[i][1].minute - results[i-1][1].minute
                            else:
                                if i == len(results)-1:
                                    model_tested[results[i][0]] = 60 - results[i-1][1].minute
                                else:
                                    model_tested[results[i][0]] = results[i][1].minute - results[i-1][1].minute                            
                                
                    for key in model_tested:
                        expect += (model_tested[key]*STB_Expec_Perf[key])/60

                    pref = round(len(results)/expect,4)
                    qual = round(qual_temp/len(results),4)
                    oee =  round((dispo_fin*pref*qual),4)
                    
                result_dict[pos][str(hour) + "h"]["Disp"] = round(dispo_fin*100,4)
                result_dict[pos][str(hour) + "h"]["Perf"] = round(pref*100,4)
                result_dict[pos][str(hour) + "h"]["Qual"] = round(qual*100,4)
                result_dict[pos][str(hour) + "h"]["OEE"] = round(oee*100,4)
                
                hour += 1
                if hour == 24:
                    hour = 0
                elif hour == 1:
                    hour = 7

            slot += 1

    db.close_connection()
    
    return result_dict

def fullfill_unavailability_dict(platform, result_dict, data_fim, data_inicio):
    # url = "http://192.168.10.200/v1/avaria/all?api_key=uadmin&comentarioFilter=&dataFimFilter=" + data_fim + "+16:00:00&dataInicioFilter=" + data_inicio + "+16:00:00&estadoFilter=&plataformaFilter=&recordsOffset=0&recordsPerPage=1000&slotFilter=&withRelatedRecords=true"
    # print(url)
    # headers = {'accept' : 'text/html,application/xhtml+xml,application/xml;q=0.9,image/avif,image/webp,image/apng,*/*;q=0.8,application/signed-exchange;v=b3;q=0.9'
    #             }
    # r = requests.get(url, headers=headers)
    # dictsf = json.dumps(xmltodict.parse(r.text))
    # dictsf = json.loads(dictsf)
    # #print(dictsf)

    # for i, v in enumerate(dictsf['response']['items']):
    #     print(dictsf['response']['items'][i])

    lista = ["2020-09-22 07:00:00_2020-09-22 07:14:00", "2020-09-22 07:20:00_2020-09-22 07:25:00"]

    result_dict["Slot_11"] = lista

    return result_dict

def check_color(value):
    """!Set Background Color for mail 
    Set background color of each numeric cell of html tables sended in mail accordind to value passed as argument
    @param value Value that will be written on table cell
    @returns Returns value hexadecimal of color
    """ 
    if value < 80:
        return "#ff0000"
    elif value >= 80 and value < 90:
        return "#ffff33"
    else:
        return "#33ff33"

def mail_message(values_stb, values_hgw):
    """!Build HTML message to sent in mail
    Build HTML message to sent in mail. Use values of Availability, Performance, Quality and OEE to fill cretaed HTML tables.
    @param values_stb Dictonary of OEE parameters values of STB plataform
    @param values_hgw Dictonary of OEE parameters values of HGW plataform
    @returns message Returns HTML message
    """ 
    color_disp = check_color(values_stb["Disponibilidade"])
    color_perf = check_color(values_stb["Performance"])
    color_qual = check_color(values_stb["Quality"])
    color_oee = check_color(values_stb["OEE"])

    message = """\
        <!DOCTYPE html>
        <html>
            <body>
                <style>
                table, th, td {
                border: 1px solid black;
                border-collapse: collapse
                }
                </style>
                <h1 style="color:SlateGray;">STB</h1>"""

    message = message + """\
                <table style="width:100%">
                    <tr>
                        <th colspan="4" style="background-color:#999966">STB</th>
                    </tr>
                    <tr>
                        <th style="background-color:#d6d6c2">Disponibilidade</th>
                        <th style="background-color:#d6d6c2">Performance</th>
                        <th style="background-color:#d6d6c2">Qualidade</th>
                        <th style="background-color:#d6d6c2">OEE</th>
                    </tr>
                    <tr>
                        <td style="text-align: center; background-color:""" + color_disp +"""">""" + str(values_stb["Disponibilidade"]) + """</td>
                        <td style="text-align: center; background-color:""" + color_perf + """">""" + str(values_stb["Performance"]) + """</td>
                        <td style="text-align: center; background-color:""" + color_qual + """">""" + str(values_stb["Quality"]) + """</td>
                        <td style="text-align: center; background-color:""" + color_oee + """">""" + str(values_stb["OEE"]) + """</td>
                    </tr>
                </table><br>"""

    color_disp = check_color(values_hgw["Disponibilidade"])
    color_perf = check_color(values_hgw["Performance"])
    color_qual = check_color(values_hgw["Quality"])
    color_oee = check_color(values_hgw["OEE"])

    message = message + """\
    <h1 style="color:SlateGray;">HGW</h1>"""

    message = message + """\
                <table style="width:100%">
                    <tr>
                        <th colspan="4" style="background-color:#999966">HGW</th>
                    </tr>
                    <tr>
                        <th style="background-color:#d6d6c2">Disponibilidade</th>
                        <th style="background-color:#d6d6c2">Performance</th>
                        <th style="background-color:#d6d6c2">Qualidade</th>
                        <th style="background-color:#d6d6c2">OEE</th>
                    </tr>
                    <tr>
                        <td style="text-align: center; background-color:""" + color_disp +"""">""" + str(values_hgw["Disponibilidade"]) + """</td>
                        <td style="text-align: center; background-color:""" + color_perf + """">""" + str(values_hgw["Performance"]) + """</td>
                        <td style="text-align: center; background-color:""" + color_qual + """">""" + str(values_hgw["Quality"]) + """</td>
                        <td style="text-align: center; background-color:""" + color_oee + """">""" + str(values_hgw["OEE"]) + """</td>
                    </tr>
                </table><br>"""

    message = message + """\
            </body>
        </html>
        """ 
    
    return message

def send_mail(mensage):
    """!Send mail with all information
    Create a SMTP connection to desire server and sends an email to desire contacts with message passed as argument and after attached documents declared in files variable.
    @see mail_message()
    @param mensage HTML message to send in body message
    """ 
    contacts = ['joao.gomes@hfa.pt']

    user = "hfa.notificacoes@hfa.pt"
    pwd = "Janela1;"

    msg = EmailMessage()
    msg['Subject'] = 'OEE Plataforma STB/HGW Alverca'
    msg['From'] = user
    msg['To'] = ','.join(contacts)
    msg.set_content('How about dinner at 6pm this saturday?')

    msg.add_alternative(mensage, subtype='html')

    for fich in files:
        with open(fich, 'rb') as f:
            file_data = f.read()
            file_name = f.name

        msg.add_attachment(file_data, maintype = 'application', subtype = 'octet-stream', filename = file_name)

    with smtplib.SMTP('svrexchange.hfa.pt', 587) as smtp:
        smtp.ehlo()
        smtp.starttls()
        smtp.ehlo()
        
        smtp.login(user, pwd)
        smtp.send_message(msg)

def make_graph(x_axis, y_axis, platform, graph_name, pic_name):
    """!Create Graphic and save it
    Create a line grafic with values passed in x_axis for x axis and y_axis for y axis
    @param x_axis List of values for x axis
    @param y_axis List of values for y axis
    @param platform Platform which dictonary will be create for 
    @param graph_name Name of graph to create
    @param pic_name Name of picture where graph will be save
    @warning platform parameter must be **"HGW"** or **"STB"**
    @warning The lenght of x_axis list must be the same of y_axis list
    """ 
    platform = platform.upper()

    graph_name = graph_name.replace("_","")

    plt.style.use('dark_background')
    #Set the size of graph in inches
    plt.figure(figsize=(13,6))
    #Create Line_Graph with (x, y)
    plt.plot(x_axis, y_axis)
    #Create DOT_Graph with (x, y)
    plt.scatter(x_axis, y_axis, color='red')
    #Graph title
    plt.title(graph_name)
    #Graph label x axis
    plt.xlabel('Hours')
    #Graph label y axis
    plt.ylabel('OEE (%)')
    #Create min, max and intervals for x axis
    plt.yticks(np.arange(0, 120, 20))
    #Create limits for graph y axis 
    plt.ylim(0, 120)

    #Add all values in all intercections
    for i, v in enumerate(y_axis):
        plt.text(i, v+5, str(v), rotation=90, color='blue', ha="center")

    # pic_name = graph_name.replace("/", "_")
    # pic_name = platform + "_" + pic_name

    plt.savefig('C:\\Users\\joao.gomes\\Desktop\\OEE\\Graphs\\' + str(pic_name))

    plt.close()
    #plt.show()

def make_4_graphs_in_one(dict_graph_name_and_axis, xname, platform, pic_name):
    """!Create 4 Graphics and save them in one picture
    Create 4 bar grafic with values passed in dict_graph_name_and_axis dictonary. This dictonary have 4 keys and their values are a list with two list, fist one has the values of x axis and second the values of y axis.
    @param dict_graph_name_and_axis dictonary of 4 Keys. Each value has a list of two lists that represent the values of x axis and y axis respectevly
    @param xname Name of x axys
    @param platform Platform which dictonary will be create for 
    @param pic_name Name of picture where graph will be save
    @warning platform parameter must be **"HGW"** or **"STB"**
    """ 
    platform = platform.upper()
    
    fig = plt.figure()

    plt.figure(figsize=(20,10))

    plt.style.use('classic')

    plt.suptitle(pic_name.replace(".png",""), fontsize=5, color='white')

    j = 221
    for key, axis in dict_graph_name_and_axis.items():
        ax = plt.subplot(j)
        #Create Line_Graph with (x, y)
        #plt.plot(dict_graph_name_and_axis[key][0], dict_graph_name_and_axis[key][1])
        barlist = plt.bar(dict_graph_name_and_axis[key][0], dict_graph_name_and_axis[key][1])
        #Create DOT_Graph with (x, y)
        plt.scatter(dict_graph_name_and_axis[key][0], dict_graph_name_and_axis[key][1], color='red')
        #Graph title
        plt.title(key)
        #Graph label x axis
        plt.xlabel(xname)
        #Graph label y axis
        plt.ylabel(str(key) + '(%)')

        if xname == "Slots":
            if platform == "STB":
                for x in range(56):
                    if x < 8:
                        barlist[x].set_color('red') 
                    elif x >= 8 and x < 16:
                        barlist[x].set_color('green') 
                    elif x >= 16 and x < 24:
                        barlist[x].set_color('blue') 
                    elif x >= 24 and x < 32:
                        barlist[x].set_color('black') 
                    elif x >= 32 and x < 40:
                        barlist[x].set_color('orange') 
                    elif x >= 40 and x < 48:
                        barlist[x].set_color('yellow') 
                    elif x >= 48 and x < 56:
                        barlist[x].set_color('brown') 
            elif platform == "HGW":
                for x in range(12):
                    if x < 4:
                        barlist[x].set_color('red') 
                    elif x >= 4 and x < 8:
                        barlist[x].set_color('green') 
                    elif x >= 8 and x < 12:
                        barlist[x].set_color('blue') 

        if xname == "Slots" and platform == "STB":
            ax.tick_params(labelsize = 8.5, labelrotation = 90)
        
        #Create min, max and intervals for x axis
        plt.yticks(np.arange(0, 120, 20))
        #Create limits for graph y axis 
        plt.ylim(0, 120)
        if xname == "Slots" and platform == "STB":
            #Add all values in all intercections
            for i, v in enumerate(dict_graph_name_and_axis[key][1]):
                plt.text(i, v+5, str(v), rotation=90, color='blue', ha="center", fontsize = 8.5) 
        else:
            #Add all values in all intercections
            for i, v in enumerate(dict_graph_name_and_axis[key][1]):
                plt.text(i, v+5, str(v), rotation=90, color='blue', ha="center")

        j += 1

    plt.savefig('' + str(pic_name))

def create_slot_hour_graphs(platform, values):
    """!Create a list with OEE values for each slot
    Create a list from dictonary passed in values argument. 
    The script will go through values dictonary and save values for each hour in list OEE. Then will call make_graph() function to make graph of values of that slot. The list is reseted and will be filled with tha next slot. This loop stops with the last slot.
    @param platform Platform which dictonary will be create for 
    @param values Dictonary with all values of Availability, Performance, Quality and OEE for each hour in each slot. Created in fullfill_dict()
    @see make_graph()
    @warning platform parameter must be **"HGW"** or **"STB"**
    """ 
    for slot, slot_dict in values.items():
        OEE = []
        index = 0
        for hour in slot_dict:
            if values[slot][hour]['OEE'] > 100:
                oee_value = 100
            else:
                oee_value = values[slot][hour]['OEE']
            if hour in aftern_shift:
                OEE.insert(index,oee_value)
                index += 1
            else:
                OEE.append(oee_value) 
        pic_name = platform.upper() + "_" + slot.replace("_","") + " OEE_h (%)"
        #print(OEE)
        make_graph(hours, OEE, platform, str(slot) + " OEE/h (%)", pic_name)

def create_slot_hour_graphs_4graphs(platform, values):
    """!Save a picture of 4 created graphics (Availability, Performance, Quality, OEE) for each slot
    Use values of dictonary passed in values argument, created in fullfill_dict(), to create a dictonary "graphics" with Keys that represent each slot and its values is a dictonary with 4 keys (Disponibilidade, Performance, Quality and OEE) and their values is a list of two lists ([x axis], [y axis])
    This graphics dictonary is used to build 4 graphics (Availability, Performance, Quality, OEE) for slot and save them in one picture
    @param platform Platform which dictonary will be create for 
    @param values Dictonary with all values of Availability, Performance, Quality and OEE for each hour in each slot. Created in fullfill_dict()
    @warning platform parameter must be **"HGW"** or **"STB"**
    """ 
    graphics = {}
    for slot, slot_dict in values.items():  
        index = 0
        OEE = []
        disp = []
        perf = []
        qual = []
        for hour in slot_dict:
            if values[slot][hour]['OEE'] > 100:
                oee_value = 100
            else:
                oee_value = values[slot][hour]['OEE']
            if hour in aftern_shift:

                OEE.insert(index,round(oee_value,1))
                disp.insert(index,round(values[slot][hour]["Disp"]))
                perf.insert(index,round(values[slot][hour]["Perf"],1))
                qual.insert(index,round(values[slot][hour]["Qual"],1))
                index += 1
            else:
                OEE.append(round(oee_value,1)) 
                disp.append(round(values[slot][hour]["Disp"]))
                perf.append(round(values[slot][hour]["Perf"],1))
                qual.append(round(values[slot][hour]["Qual"],1))
        print(perf)
        
        #print(graphics)

        graphics[slot] = {"Disponibilidade":[hours, disp], "Performance":[hours, perf], "Quality":[hours, qual], "OEE":[hours, OEE]}
 
    for key, grafs in graphics.items():
        j = 221
        fig = plt.figure()

        plt.figure(figsize=(20,10))

        plt.style.use('classic')

        plt.suptitle(key.replace("_","") + "_Graphs", fontsize=32, color='white')
        for graf in grafs:
            max_y_value = 120   
            interval = 20       
            plt.subplot(j)
            #Create Bar_Graph with (x, y)
            plt.bar(graphics[key][graf][0], graphics[key][graf][1])
            #Create Line_Graph with (x, y)
            #plt.plot(graphics[key][graf][0], graphics[key][graf][1], color='red')           
            #Create DOT_Graph with (x, y)
            plt.scatter(graphics[key][graf][0], graphics[key][graf][1], color='red')
            #Graph title
            plt.title(key.replace("_","") + "_" + str(graf))
            #Graph label x axis
            plt.xlabel('Hours')
            #Graph label y axis
            plt.ylabel(str(graf) + '(%)')
            if graf == "Performance":   
                if max(graphics[key][graf][1]) > 100:
                    max_y_value = max(graphics[key][graf][1]) + 20   
                    interval = max_y_value/6
            #Create min, max and intervals for x axis
            plt.yticks(np.arange(0, max_y_value, interval))
            #Create limits for graph y axis 
            plt.ylim(0, max_y_value)
            #Add all values in all intercections
            for i, v in enumerate(graphics[key][graf][1]):
                plt.text(i, v+5, str(v), rotation=90, color='blue', ha="center") 

            j += 1

        pic_name = platform.upper() + "_" + key.replace("_","")
        plt.savefig('C:\\Users\\joao.gomes\\Desktop\\OEE\\Graphs\\' + str(pic_name))
        plt.close()

def create_oee_graph_by_day(platform, values):
    """!Create a list with OEE values of day in each hour
    Create a list with OEE values of day in each hour from dictonary passed in values argument. 
    The script will go through values dictonary and save values for each hour of all slots in list OEE_Day. Then will call make_graph() function to make graph of values of this list.
    @param platform Platform which dictonary will be create for 
    @param values Dictonary with all values of Availability, Performance, Quality and OEE for each hour in each slot. Created in fullfill_dict()
    @see make_graph()
    @warning platform parameter must be **"HGW"** or **"STB"**
    """ 
    hour = 16
    OEE_Day = []
    for hora in range(18):
        hour_oee_temp = 0
        slots = 0
        for slot in values:
            hour_oee_temp += values[slot][str(hour) + "h"]['OEE']
            slots += 1
        hour_oee = round((hour_oee_temp/slots),2)
        if hour_oee > 100:
            hour_oee = 100
        
        OEE_Day.append(hour_oee)

        hour += 1
        if hour == 24:
            hour = 0
        elif hour == 1:
            hour = 7

    pic_name = str(date_today) + "_" + platform.upper()  + "_OEE_h_Day (%)"
    #print(OEE_Day)
    make_graph(hours, OEE_Day, platform, "OEE/h-Day (%)", pic_name)

def create_graphs_by_day(platform, values):
    """!Create lists with availability, performance, quality and OEE values of day in each hour
    Create a list with availability, performance, quality and OEE values of day in each hour from dictonary passed in values argument. 
    The script will go through values dictonary and save values for each hour of all slots in lists. 
    Then will create a dictonary with this lists and call make_4_graphs_in_one() function to make 4 graphs of values of this lists and save it in one picture.
    @param platform Platform which dictonary will be create for 
    @param values Dictonary with all values of Availability, Performance, Quality and OEE for each hour in each slot. Created in fullfill_dict()
    @see make_4_graphs_in_one()
    @warning platform parameter must be **"HGW"** or **"STB"**
    """ 
    hour = 16
    Disp_Day = []
    Perf_day = []
    Qual_Day = []
    OEE_Day = []
    for hora in range(18):
        hour_oee_temp = 0
        hour_disp_temp = 0
        hour_perf_temp = 0
        hour_qual_temp = 0
        slots = 0
        for slot in values:
            hour_disp_temp += values[slot][str(hour) + "h"]['Disp']
            hour_perf_temp += values[slot][str(hour) + "h"]['Perf']
            hour_qual_temp += values[slot][str(hour) + "h"]['Qual']
            #hour_oee_temp += values[slot][str(hour) + "h"]['OEE']
            slots += 1

        hour_disp = round((hour_disp_temp/slots),2)
        hour_perf = round((hour_perf_temp/slots),2)
        hour_qual = round((hour_qual_temp/slots),2)
        hour_oee = (hour_disp/100)*(hour_perf/100)*(hour_qual/100)
        hour_oee = round(hour_oee,2)*100

        if hour_oee >= 100:
            hour_oee = 100
        if hour_disp == 100:
            Disp_Day.append(int(hour_disp))
        else:
            Disp_Day.append(hour_disp)
        Perf_day.append(hour_perf)
        Qual_Day.append(hour_qual)
        OEE_Day.append(round(hour_oee,2))

        hour += 1
        if hour == 24:
            hour = 0
        elif hour == 1:
            hour = 7
    
    graphics = {"Disponibilidade":[hours, Disp_Day], "Performance":[hours, Perf_day], "Quality":[hours, Qual_Day], "OEE":[hours, OEE_Day]}

    pic_name = str(date_today) + "_" + platform.upper() + "_Graphs_Hours_Day.png"

    make_4_graphs_in_one(graphics, "Hours", platform, pic_name)

def values_for_shift(values, shift):
    """!Create a Dictonary with OEE params for a shift. 
    Use values of dictonary 'values' and calculate the average of availability, performance, quality and OEE for a shift.
    @param Dictonary with all values. Builded by fullfill_dict() function
    @param shift Shift of the day
    @see fullfill_dict()
    @returns values_shift Returns a dictonary with values of availability, performance, quality and OEE of shift day
    """
    shift = shift.lower()
    if shift == "manha":
        hour = 7
    elif shift == "tarde":
        hour = 16
    Disp_hour = []
    Perf_hour = []
    Qual_hour = []
    OEE_hour = []
    for hora in range(9):
        hour_oee_temp = 0
        hour_disp_temp = 0
        hour_perf_temp = 0
        hour_qual_temp = 0
        slots = 0
        for slot in values:
            hour_disp_temp += values[slot][str(hour) + "h"]['Disp']
            hour_perf_temp += values[slot][str(hour) + "h"]['Perf']
            hour_qual_temp += values[slot][str(hour) + "h"]['Qual']
            hour_oee_temp += values[slot][str(hour) + "h"]['OEE']
            slots += 1

        hour_disp = round((hour_disp_temp/slots),2)
        hour_perf = round((hour_perf_temp/slots),2)
        hour_qual = round((hour_qual_temp/slots),2)
        hour_oee = round((hour_oee_temp/slots),2)

        if hour_oee >= 100:
            hour_oee = 100
        if hour_disp == 100:
            Disp_hour.append(int(hour_disp))
        else:
            Disp_hour.append(hour_disp)
        Perf_hour.append(hour_perf)
        Qual_hour.append(hour_qual)
        OEE_hour.append(hour_oee)

        hour += 1
        if hour == 24:
            hour = 0
        elif hour == 1:
            hour = 7

    Disp_shift = round(sum(Disp_hour)/len(Disp_hour),2)
    Perf_shift = round(sum(Perf_hour)/len(Perf_hour),2)
    Qual_shift = round(sum(Qual_hour)/len(Qual_hour),2)
    OEE_shift = round(sum(OEE_hour)/len(OEE_hour),2)
    
    values_shift = {"Disponibilidade":Disp_shift, "Performance":Perf_shift, "Quality":Qual_shift, "OEE":OEE_shift}

    return values_shift

def values_for_a_day(values):
    """!Create a dictonary with values of availability, performance, quality and OEE of a day
    Create a dictonary with values of availability, performance, quality and OEE for that day. To do this is made a avarage of each of 4 components
    @param values Dictonary with all values of Availability, Performance, Quality and OEE for each hour in each slot. Created in fullfill_dict()
    @returns values_day Returns a dictonary with avarage values of availability, performance, quality and OEE of the day
    """
    hour = 16
    Disp_hour = []
    Perf_hour = []
    Qual_hour = []
    OEE_hour = []
    for hora in range(18):
        hour_oee_temp = 0
        hour_disp_temp = 0
        hour_perf_temp = 0
        hour_qual_temp = 0
        slots = 0
        for slot in values:
            hour_disp_temp += values[slot][str(hour) + "h"]['Disp']
            hour_perf_temp += values[slot][str(hour) + "h"]['Perf']
            hour_qual_temp += values[slot][str(hour) + "h"]['Qual']
            hour_oee_temp += values[slot][str(hour) + "h"]['OEE']
            slots += 1

        hour_disp = round((hour_disp_temp/slots),2)
        hour_perf = round((hour_perf_temp/slots),2)
        hour_qual = round((hour_qual_temp/slots),2)
        hour_oee = round((hour_oee_temp/slots),2)

        if hour_oee >= 100:
            hour_oee = 100
        if hour_disp == 100:
            Disp_hour.append(int(hour_disp))
        else:
            Disp_hour.append(hour_disp)
        Perf_hour.append(hour_perf)
        Qual_hour.append(hour_qual)
        OEE_hour.append(hour_oee)

        hour += 1
        if hour == 24:
            hour = 0
        elif hour == 1:
            hour = 7

    Disp_Day = round(sum(Disp_hour)/len(Disp_hour),2)
    Perf_day = round(sum(Perf_hour)/len(Perf_hour),2)
    Qual_Day = round(sum(Qual_hour)/len(Qual_hour),2)
    OEE_Day = round(sum(OEE_hour)/len(OEE_hour),2)
    
    values_day = {"Disponibilidade":Disp_Day, "Performance":Perf_day, "Quality":Qual_Day, "OEE":OEE_Day}
 
    return values_day

def create_graphs_slots_day(platform, values):
    """!Create a list with availability, performance, quality and OE values of day in each slot
    The script will go through values dictonary and save values for each slot of all hours in list. Then will call make_graph() function to make graph of values of this list.
    @param platform Platform which dictonary will be create for 
    @param values Dictonary with all values of Availability, Performance, Quality and OEE for each slot. Created in fullfill_dict_slot_day()
    @see make_4_graphs_in_one()
    @warning platform parameter must be **"HGW"** or **"STB"**
    """ 
    platform = platform.upper()
    if platform == "STB":
        slots = slots_stb
    elif platform == "HGW":
        slots = slots_hgw
    Disp_Slot = []
    Perf_Slot = []
    Qual_Slot = []
    OEE_Slot = []
    for slot in values:
        if values[slot]['OEE'] > 100:
            oee_value = 100
        else:
            oee_value = values[slot]['OEE']

        OEE_Slot.append(round(oee_value,1)) 
        Disp_Slot.append(round(values[slot]["Disp"]))
        Perf_Slot.append(round(values[slot]["Perf"],1))
        Qual_Slot.append(round(values[slot]["Qual"],1))
    
    graphics = {"Disponibilidade":[slots, Disp_Slot], "Performance":[slots, Perf_Slot], "Quality":[slots, Qual_Slot], "OEE":[slots, OEE_Slot]}

    pic_name = str(date_today) + "_" + platform.upper() + "_Graphs_Slots_Day.png"

    make_4_graphs_in_one(graphics, "Slots", platform, pic_name)

def create_week_graph(platform, week_day, dict_oee_day, graph_name):
    """!Create a Graphic with values of OEE for each day of week and the average of the current week.
    Get OEE of the current day of week and add to "graph_name".txt file. Check the filled days in file, calculate the current average oee of the week and 'build' the graph. 
    @param platform Platform which dictonary will be create for 
    @param week_day Current day of the week
    @param dict_oee_day Dictonary of the parameters of the current day
    @param graph_name Graphic and file name
    @see values_for_a_day()
    @warning platform parameter must be **"HGW"** or **"STB"**
    """
    pic_name = str(date_today) + "_" + graph_name + ".png"
    y_value_exp_stb = [26.34, 26.34, 26.34, 26.34, 26.34, 26.34]
    y_value_exp_hgw = [31.12, 31.12, 31.12, 31.12, 31.12, 31.12]
    x_values = ["Segunda", "Terça", "Quarta", "Quinta", "Sexta", "Week"]
    y_values = []
    oee_day = dict_oee_day["OEE"]

    f = open(graph_name + ".txt", "r")
    list_of_lines = f.readlines()
    f.close()

    list_of_lines[week_day] = "day_" + str(week_day) + "-" + str(oee_day) + ";\n"

    f = open(graph_name + ".txt", "w")
    f.writelines(list_of_lines)
    f.close()

    f = open(graph_name + ".txt", "r")
    list_of_lines = f.readlines()
    f.close()


    for i, item in enumerate(list_of_lines):
        x = list_of_lines[i].split('-')
        x = x[1].split(';')
        y_values.append(float(x[0]))
    total = 0
    for j in range(week_day + 1):
        total += y_values[j]
    
    #print(total)
    avg_week = total/(week_day + 1)

    y_values.append(round(avg_week,2))

    #print(y_values)

    platform = platform.upper()

    graph_name = pic_name.replace(".png","")

    plt.style.use('classic')
    #Set the size of graph in inches
    plt.figure(figsize=(13,6))
    if platform == "STB":
        plt.plot(x_values, y_value_exp_stb, color='red')
    elif platform == "HGW":
        plt.plot(x_values, y_value_exp_hgw, color='red')
    #Create Bar_Graph with (x, y)
    plt.bar(x_values, y_values)
    #Create DOT_Graph with (x, y)
    plt.scatter(x_values, y_values, color='red')
    #Graph title
    plt.title(graph_name)
    #Graph label x axis
    plt.xlabel('Dias')
    #Graph label y axis
    plt.ylabel('OEE (%)')
    #Create min, max and intervals for x axis
    plt.yticks(np.arange(0, 120, 20))
    #Create limits for graph y axis 
    plt.ylim(0, 120)

    #Add all values in all intercections
    for i, v in enumerate(y_values):
        plt.text(i, v+5, str(v), rotation=90, color='blue', ha="center")

    # #plt.show()
    plt.savefig('C:\\Users\\joao.gomes\\Desktop\\OEE\\' + str(pic_name))
    if week_day == 4:
        f = open(graph_name + ".txt", "w")
        for i in range (5):
            f.write("day_" + str(i) + "-" + str(0) + ";\n")
        f.close()

def create_data_csv_file(dict_values, file_name):
    """!Create a csv file with used data in project.
    Create a csv file with all data used in project. This data is present in dictonary passed in argument dict_values
    @param dict_values Dictonary with values to be written in csv file. Created in create_dictonary_for_csv()
    @param file_name Name of file
    """ 
    with open(file_name, 'w') as f:
        f.write("Slot,Hour,Nº Testes,OKs,Nº Testes Expextaveis,Availability,Performance,Quality,OEE\n")
        for slot, slot_dict in dict_values.items():
            for hour in slot_dict:
                f.write(slot.replace("Slot_","") + "," + hour + "," + str(dict_values[slot][hour]['N_Testes']) + "," + str(dict_values[slot][hour]['OKs']) + "," + str(dict_values[slot][hour]['N_Testes_Expextaveis']) + "," + str(dict_values[slot][hour]['Disp']) + "," + str(dict_values[slot][hour]['Perf']) + "," + str(dict_values[slot][hour]['Qual']) + "," + str(dict_values[slot][hour]['OEE']) + "\n")

def create_numeric_parameters_csv_file(dict_exp_stb, dict_expc_hgw, morn_shift, aftern_shift):
    """!Create a csv file with parameters used in project
    Create a csv file with all parameters used in project.
    @param dict_exp_stb Dictonary with the expected number of tests per hour for each model of STBs
    @param dict_expc_hgw Dictonary with the expected number of tests per hour for each model of STBs
    @param morn_shift List with all hours of morning shift
    @param aftern_shift List with all hours of afternoon shift
    """ 
    with open(str(date_today) + "_" + "Numeric_Parameters.csv", 'w') as f:
        i=0
        line = ""
        line_1 = ""
        for key in dict_exp_stb:
            if i == len(dict_exp_stb):
                line += str(key)
                line_1 += str(dict_exp_stb[key])
            else:
                line += str(key) + ","
                line_1 += str(dict_exp_stb[key]) + ","
            i += 1
        f.write(line + "\n")
        f.write(line_1 + "\n\n")
       
        i=0
        line = ""
        line_1 = ""
        for key in dict_expc_hgw:
            if i == len(dict_expc_hgw):
                line += str(key)
                line_1 += str(dict_expc_hgw[key])
            else:
                line += str(key) + ","
                line_1 += str(dict_expc_hgw[key]) + ","
            i += 1
        f.write(line + "\n")
        f.write(line_1 + "\n\n")

        line = ""
        f.write("Morning_Shift" + "\n")
        for j in morn_shift:
            line += j + ","
        f.write(line + "\n\n")

        line = ""
        f.write("Afternoon_Shift" + "\n")
        for k in aftern_shift:
            line += k + ","
        f.write(line + "\n\n")

def zipdir(path, ziph):
    """!Create zip file
    This function goes through all files in path and zip them in one zip file.
    @param path zipfile wanted path
    @param ziph ziph is zipfile handle
    """
    for root, dirs, files in os.walk(path):
        for file in files:
            ziph.write(os.path.join(root, file))

if __name__ == '__main__':
    now = datetime.now()
    #now = datetime.strptime("2020-09-14", '%Y-%m-%d')
    week_day = datetime.today().weekday()
    if week_day == 0:
        yesterday = datetime.today() - timedelta(3)
    else:
        yesterday = datetime.today() - timedelta(1)
    date_today = str(now.year) + "-" + str(now.month) + "-" +str(now.day)
    date_yesterday = str(yesterday.year) + "-" + str(yesterday.month) + "-" +str(yesterday.day)

    for idx,i in enumerate(files):
        files[idx] = str(date_today) + "_" + str(files[idx])

    current_dir = os.getcwd()
    graphs_dir = os.path.join(current_dir, "Graphs")
    if not os.path.exists(graphs_dir):
        os.makedirs(graphs_dir)

    db = Connect_to_db()
    db.connect()
    results_temp_STB = db.get_perf_qual_all_in_one("STB", date_yesterday, date_today, week_day)
    results_temp_HGW = db.get_perf_qual_all_in_one("HGW", date_yesterday, date_today, week_day)
    db.close_connection()

    #Create Dictionary for each Slot and hour
    dictio_stb = create_dictonary("STB")
    dictio_hgw = create_dictonary("HGW")

    un_dict = create_unavailability_dict('stb')
    #print(un_dict)
    unvai_slots = fullfill_unavailability_dict('stb', un_dict, '2020-09-17', '2020-09-16')
    #print(unvai_slots)

    #Full fill Dictionary for each Slot and hour
    dictio_full_stb = fullfill_dict("STB", dictio_stb, results_temp_STB, unvai_slots) 
    #dictio_full_hgw = fullfill_dict("HGW", dictio_hgw, results_temp_HGW, un_dict)

    # dict_stb_csv = create_dictonary_for_csv("STB")
    # dict_hgw_csv = create_dictonary_for_csv("HGW")

    # dict_stb_csv_full = fullfill_dict_for_csv("STB", dict_stb_csv, results_temp_STB)
    # dict_hgw_csv_full = fullfill_dict_for_csv("HGW", dict_hgw_csv, results_temp_HGW)    

    #Create CSV Data files
    # create_data_csv_file(dict_stb_csv_full, str(date_today) + "_" + "STB_Data.csv")
    # create_data_csv_file(dict_hgw_csv_full, str(date_today) + "_" + "HGW_Data.csv")

    # stb_day = create_dictonary_slot_day("STB")
    # hgw_day = create_dictonary_slot_day("HGW")

    # dict_day_stb = fullfill_dict_slot_day("STB", stb_day, results_temp_STB)
    # dict_day_hgw = fullfill_dict_slot_day("HGW", hgw_day, results_temp_HGW)

    # create_slot_hour_graphs_4graphs("STB", dictio_full_stb)
    # create_slot_hour_graphs_4graphs("HGW", dictio_full_hgw)

    # create_graphs_by_day("STB", dictio_full_stb)
    # create_graphs_by_day("HGW", dictio_full_hgw)

    # create_graphs_slots_day("STB", dict_day_stb)
    # create_graphs_slots_day("HGW", dict_day_hgw)

    # create_numeric_parameters_csv_file(STB_Expec_Perf, HGW_Expec_Perf, morning_shift, aftern_shift)

    # final_oee_stb = values_for_a_day(dictio_full_stb)
    # final_oee_hgw = values_for_a_day(dictio_full_hgw)

    # zipf = zipfile.ZipFile(str(date_today) + "_" + 'graphs.zip', 'w', zipfile.ZIP_DEFLATED)
    # zipdir('C:\\Users\\joao.gomes\\Desktop\\OEE\\Graphs', zipf)
    # zipf.close()
    # final_oee_stb_morning_shift = values_for_shift(dictio_full_stb, "manha")
    # final_oee_stb_afternoon_shift = values_for_shift(dictio_full_stb, "tarde")

    # create_week_graph("STB", week_day, final_oee_stb,"OEE_Day_STB")
    # create_week_graph("HGW", week_day, final_oee_hgw,"OEE_Day_HGW")
    # create_week_graph("STB", week_day, final_oee_stb_morning_shift, "OEE_Morning_Shift_STB")
    # create_week_graph("STB", week_day, final_oee_stb_afternoon_shift, "OEE_Afternoon_Shift_STB")
    # create_week_graph("HGW", week_day, final_oee_stb_morning_shift, "OEE_Morning_Shift_HGW")
    # create_week_graph("HGW", week_day, final_oee_stb_afternoon_shift, "OEE_Afternoon_Shift_HGW")

    # mens = mail_message(final_oee_stb, final_oee_hgw)
    # send_mail(mens)

    ########################Check OEE between to dates#################
    #########################Put this in a funtion#####################
    # STB = []
    # HGW = []
    # data = "2020-06-01"
    # data = datetime.strptime(data, '%Y-%m-%d')
    # i=0
    # while data != datetime.strptime("2020-07-27", '%Y-%m-%d'):
    #     week_day = data.weekday()
    #     print(week_day)
    #     if week_day == 0:
    #         yesterday = data - timedelta(3)
    #     else:
    #         yesterday = data - timedelta(1)
    #     date_today = str(data.year) + "-" + str(data.month) + "-" +str(data.day)
    #     date_yesterday = str(yesterday.year) + "-" + str(yesterday.month) + "-" +str(yesterday.day)


    #     db = Connect_to_db()
    #     db.connect()
    #     results_temp_STB = db.get_perf_qual_all_in_one("STB", date_yesterday, date_today)
    #     results_temp_HGW = db.get_perf_qual_all_in_one("HGW", date_yesterday, date_today)
    #     db.close_connection()

    #     #Create Dictionary for each Slot and hour
    #     dictio_stb = create_dictonary("STB")
    #     dictio_hgw = create_dictonary("HGW")

    #     #Full fill Dictionary for each Slot and hour
    #     dictio_full_stb = fullfill_dict("STB", dictio_stb, results_temp_STB) 
    #     dictio_full_hgw = fullfill_dict("HGW", dictio_hgw, results_temp_HGW)

    #     final_oee_stb = values_for_a_day(dictio_full_stb)
    #     final_oee_hgw = values_for_a_day(dictio_full_hgw)

    #     STB.append(final_oee_stb["OEE"])
    #     HGW.append(final_oee_hgw["OEE"])

    #     data = data + timedelta(1)

    #     if week_day == 4:
    #         data = data + timedelta(2)
    #         data = str(data.year) + "-" + str(data.month) + "-" +str(data.day)
    #         data = datetime.strptime(data, '%Y-%m-%d')

    # print(STB)
    # print(HGW)

    # print("STB_OEE = " + str(sum(STB)/len(STB)))
    # print("HGW_OEE = " + str(sum(HGW)/len(HGW)))
            

    
