message("Loading the packages")
library(taskscheduleR)
library(lubridate)

task_scheduler_emails <- function (freq,
                                   start_time,
                                   start_taskscheduler_date,
                                   start_taskscheduler_day,
                                   last_day){
message("Starting Outlook application, getting the MAPI namespace, and accessing inbox messages")
OutApp <- COMCreate("Outlook.Application")
outlookNameSpace <- OutApp$GetNameSpace("MAPI")

message("Defining the path to my script")
my_script <- Sys.getenv("HACKATHON_R_SCRIPT_PATH")

if (last_day == "") {

message("Creating the automated task")
taskscheduler_create(
  taskname = "Hackathon_Task_Scheduler", 
  rscript = my_script, 
  schedule = freq, 
  starttime = start_time,
  startdate = format(start_taskscheduler_date,"%m/%d/%Y"),
  days = start_taskscheduler_day,
  Rexe = file.path(R.home("bin"), "Rscript.exe")
)
} else if (last_day != "") {
  
    message("Creating the automated task")
  cmd <- paste0(
    'schtasks /Create /F /TN "Hackathon_Task_Scheduler" ',
    '/TR "cmd /c ', file.path(R.home("bin"), "Rscript.exe"), ' \\', paste0(Sys.getenv("HACKATHON_R_SCRIPT_PATH"), '\\'), ' >> \\', paste0(str_sub(Sys.getenv("HACKATHON_R_SCRIPT_PATH"), 1, -2), 'log'), '\\" 2>&1" ',
    '/SC MONTHLY /MO LASTDAY /ST ', start_time, ' /M * /SD ', format(start_taskscheduler_date,"%m/%d/%Y"), '"'
  )
  status <- system(cmd, intern = FALSE)
  }
}