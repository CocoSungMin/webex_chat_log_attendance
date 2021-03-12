# webex_chat_log_attendance

--------------------------
This python code is use to check student attendance based on chat log

You need Webex log text file as a input and class time is based on Gachon Uni. system

----------------------------
## Required Library
1. pandas
2. numpy
3. getopt
4. openpyxl
5. etc

---------------------------

##Code Usage
1. In Powershell or Terminal Run the python code with this format
2. class time is based on like this

1st class : 09:00 am 
A class : 09:30 am

```python
python attendance.py -i <inputfile> -c <classtime>
```

Order example : 
```python
python attendance.py -i mobileProgramming_1_class.txt -c 3
```

After you received excel file you have to check the absent student manually (logic is not perfect)

----------------
##Attendance Standard
1. +- 5minutes from first class start time check true
2. +- 5minutes from last class end time check true
3. both 2 value is true  -> attend 
4. either 1 value is false -> tardy
5. both value is false -> absent
----------------
This Program belongs to VMR-LAB in Gachon Uni.

Code developer : Bachelor course Research Assistant Sungmin Lee
    
    
  
