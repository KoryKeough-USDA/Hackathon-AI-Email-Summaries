message("Loading the packages")
library(taskscheduleR)
library(lubridate)

task_scheduler_emails <- function (freq,
                                   start_time,
                                   start_taskscheduler_date,
                                   start_taskscheduler_day,
                                   schtasks_extra){
message("Starting Outlook application, getting the MAPI namespace, and accessing inbox messages")
OutApp <- COMCreate("Outlook.Application")
outlookNameSpace <- OutApp$GetNameSpace("MAPI")

message("Defining the path to my script")
my_script <- Sys.getenv("HACKATHON_R_SCRIPT_PATH")

message("Creating the automated task")
taskscheduler_create(
  taskname = "hackathon_r_file_scheduler", 
  rscript = my_script, 
  schedule = freq, 
  starttime = start_time,
  startdate = start_taskscheduler_date, # format(Sys.Date() + (6 - wday(Sys.Date())) %% 7, "%m/%d/%Y")  Next Friday
  days = start_taskscheduler_day,
  Rexe = file.path(R.home("bin"), "Rscript.exe"),
  schtasks_extra = schtasks_extra
)
}