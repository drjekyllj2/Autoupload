from downloadfromemail import get_link
from consolidate import consolidate_data
import datetime
import schedule
import time
from enum import Enum

class Schedule(Enum):
     daily = 1
     weekly =2
     biweekly = 3
     monthly = 4
 
# config_out= config.ConfigParser()
# config_out.read ("Config.ini")
# run_schedule = config_out.getint('schedule','run_schedule')
# del config_out
# def do_task():
#         processexcel.process_excel()
#         processexcel.process_file()
          
#         print ("it is process")
 
get_link()
consolidate_data()
      #do_task()
        
        #    if run_schedule==1:#Schedule.daily:
        #         print ("ddd")
        #         schedule.every(1).minute.do(do_task)
                
        #    if run_schedule==2:#Schedule.weekly:
        #         schedule.every(7).day.do(do_task)
        #    if run_schedule==3:#Schedule.biweekly:
        #       schedule.every(14).day.do(do_task)
        #    if run_schedule==4:#Schedule.monthly:
        #       schedule.every(30).day.do(do_task)
                
                
                

        #    while True:
        #       schedule.run_pending()
        #       time.sleep(1)
        # else:
        #         print ("Error")
        
#process()

    
                
