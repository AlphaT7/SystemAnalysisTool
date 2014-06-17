//Creates a Copy of "High Performance" Power Profile
c:\windows\system32\powercfg.exe -duplicatescheme 8c5e7fda-e8bf-4a96-9a85-a6e23a8c635c

//Changes Name and Description of "High Performance" Power Profile
c:\windows\system32\powercfg.exe -changename 8c5e7fda-e8bf-4a96-9a85-a6e23a8c635c SecureDoc "Power settings required for hard disk encryption"

//Sets The New "SecureDoc" Power Profile as Active
c:\windows\system32\powercfg.exe -setactive 8c5e7fda-e8bf-4a96-9a85-a6e23a8c635c

//Disable Sleep and Hibernate
c:\windows\system32\powercfg.exe -change -monitor-timeout-ac 0
c:\windows\system32\powercfg.exe -change -monitor-timeout-dc 0
c:\windows\system32\powercfg.exe -change -disk-timeout-ac 0
c:\windows\system32\powercfg.exe -change -disk-timeout-dc 0
c:\windows\system32\powercfg.exe -change -standby-timeout-ac 0
c:\windows\system32\powercfg.exe -change -standby-timeout-dc 0
c:\windows\system32\powercfg.exe -change -hibernate-timeout-ac 0
c:\windows\system32\powercfg.exe -change -hibernate-timeout-dc 0

//Lid Close Do Nothing
c:\windows\system32\powercfg.exe -setacvalueindex 8c5e7fda-e8bf-4a96-9a85-a6e23a8c635c 4f971e89-eebd-4455-a8de-9e59040e7347 5ca83367-6e45-459f-a27b-476b1d01c936 0
