# Schedule Retriever
Python program to automatically retrieve the most recent schedule emailed to me as an excel file, parse and save it, 
and add the days I work to my google calendar.  

## Parts
- Uses Gmail API to retrieve the most recent schedule received.   
- Uses Panda to parse the attached excel schedule into a DataFrame.
- Cleans the schedule data and saves it as a JSON locally.
  - Planning to use the JSON with IFTTT to create custom google assistant prompts like "Hey Google who do I work with today?".    
- Uses the Google Calendar API to add the days I work to my calendar. *(In progress)*
